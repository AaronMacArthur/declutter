from __future__ import annotations

import ctypes
import heapq
import json
import os
import queue
import subprocess
import sys
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from tkinter import (
    BooleanVar,
    Button,
    Canvas,
    Checkbutton,
    END,
    Entry,
    Frame,
    Label,
    LabelFrame,
    Listbox,
    PanedWindow,
    StringVar,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
    ttk,
)
from tkinter.constants import BOTH, BOTTOM, DISABLED, HORIZONTAL, LEFT, NORMAL, RIGHT, VERTICAL, X, Y


APP_NAME = "Declutter"
APP_VERSION = "1.0"
DECLUTTER_FOLDER_NAME = "Decluttered Desktop"
MANIFEST_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "DeclutterApp" / "manifests"
PROJECT_ROOT = Path(__file__).resolve().parent
TREEMAP_CHILD_LIMIT = 10
TREEMAP_MAX_DEPTH = 4
LARGEST_LIST_LIMIT = 40
MOVE_TIMEOUT_SECONDS = 12
TREEMAP_LABEL_MIN_AREA = 18000
TREEMAP_LEAF_LABEL_MIN_AREA = 32000
RISKY_DELETE_MESSAGE = (
    "This could very well corrupt your system... Are you sure?\n\n"
    "Declutter is looking at this file with deep suspicion. It lives somewhere "
    "system-ish or has a spicy extension, and deleting it may make Windows start "
    "speaking in riddles.\n\n"
    "Selected file:\n{path}\n\n"
    "Why it looks bad:\n{reasons}"
)


EXTENSION_CATEGORIES = {
    "Documents": {
        ".doc",
        ".docx",
        ".odt",
        ".rtf",
        ".txt",
        ".md",
        ".pages",
        ".epub",
    },
    "PDFs": {".pdf"},
    "Spreadsheets": {".xls", ".xlsx", ".xlsm", ".csv", ".tsv", ".ods", ".numbers"},
    "Presentations": {".ppt", ".pptx", ".key", ".odp"},
    "Images": {
        ".png",
        ".jpg",
        ".jpeg",
        ".gif",
        ".bmp",
        ".webp",
        ".tif",
        ".tiff",
        ".heic",
        ".svg",
        ".ico",
    },
    "Videos": {".mp4", ".mov", ".mkv", ".avi", ".wmv", ".webm", ".m4v"},
    "Audio": {".mp3", ".wav", ".aiff", ".aac", ".flac", ".m4a", ".ogg"},
    "Archives": {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"},
    "Installers": {".exe", ".msi", ".dmg", ".pkg", ".appx", ".msix"},
    "Code": {
        ".py",
        ".js",
        ".ts",
        ".tsx",
        ".jsx",
        ".html",
        ".css",
        ".json",
        ".xml",
        ".yaml",
        ".yml",
        ".java",
        ".cs",
        ".cpp",
        ".c",
        ".h",
        ".hpp",
        ".rs",
        ".go",
        ".php",
        ".rb",
        ".sql",
        ".sh",
        ".ps1",
        ".bat",
        ".cmd",
    },
    "Shortcuts": {".lnk", ".url", ".website"},
    "Fonts": {".ttf", ".otf", ".woff", ".woff2"},
    "Design": {".psd", ".ai", ".xd", ".fig", ".sketch", ".indd"},
}

NAME_RULES = {
    "Screenshots": ("screenshot", "screen shot", "snip", "capture"),
    "Receipts": ("receipt", "invoice", "bill", "statement"),
    "Work": ("resume", "cv", "cover letter", "contract", "proposal"),
}

COLORS = [
    "#2f6f73",
    "#c45f4c",
    "#6c7a2e",
    "#715b94",
    "#b4762f",
    "#3f5f8f",
    "#8b4f64",
    "#447a4f",
    "#936f39",
    "#58666f",
    "#9b6b89",
    "#4d7570",
]


@dataclass
class DesktopItem:
    path: Path
    category: str
    size: int
    kind: str
    target: Path
    note: str = ""


@dataclass
class ScanNode:
    path: Path
    name: str
    size: int = 0
    files: int = 0
    dirs: int = 0
    children: list["ScanNode"] = field(default_factory=list)
    error: str = ""
    is_file: bool = False


@dataclass
class SkippedMove:
    source: Path
    category: str
    reason: str


@dataclass
class OrganizeSummary:
    manifest_path: Path
    moved_count: int
    skipped: list[SkippedMove]


def get_desktop_path() -> Path:
    """Return the current user's Desktop path, including OneDrive redirection when Windows reports it."""
    try:
        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        desktop_guid = GUID(
            0xB4BFCC3A,
            0xDB2C,
            0x424C,
            (ctypes.c_ubyte * 8)(0xB0, 0x29, 0x7F, 0xE9, 0x9A, 0x87, 0xC6, 0x41),
        )
        path_ptr = ctypes.c_void_p()
        ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(desktop_guid),
            0,
            None,
            ctypes.byref(path_ptr),
        )
        if path_ptr.value:
            desktop = ctypes.wstring_at(path_ptr.value)
            ctypes.windll.ole32.CoTaskMemFree(path_ptr)
            return Path(desktop)
    except Exception:
        pass

    candidates = [
        Path(os.environ.get("OneDrive", "")) / "Desktop",
        Path.home() / "OneDrive" / "Desktop",
        Path.home() / "Desktop",
    ]
    for candidate in candidates:
        if str(candidate) and candidate.exists():
            return candidate
    return Path.home() / "Desktop"


def human_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB", "TB", "PB"):
        if value < 1024 or unit == "PB":
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def folder_size_quick(path: Path, limit: int = 500) -> tuple[int, str]:
    total = 0
    seen = 0
    try:
        for root, dirs, files in os.walk(path):
            dirs[:] = [name for name in dirs if not name.startswith(".")]
            for file_name in files:
                try:
                    total += (Path(root) / file_name).stat().st_size
                except OSError:
                    continue
                seen += 1
                if seen >= limit:
                    return total, "estimated"
    except OSError:
        return 0, "permission denied"
    return total, ""


def unique_destination(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 2
    while True:
        candidate = parent / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def move_with_timeout(source: Path, destination: Path, timeout_seconds: int = MOVE_TIMEOUT_SECONDS) -> tuple[bool, str]:
    if not source.exists():
        return False, "source no longer exists"

    script = (
        "import shutil, sys\n"
        "shutil.move(sys.argv[1], sys.argv[2])\n"
    )
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0

    try:
        result = subprocess.run(
            [sys.executable, "-c", script, str(source), str(destination)],
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            creationflags=creationflags,
        )
    except subprocess.TimeoutExpired:
        if destination.exists() and not source.exists():
            return True, "moved after timeout"
        return False, f"timed out after {timeout_seconds}s"

    if result.returncode == 0:
        return True, ""

    if destination.exists() and not source.exists():
        return True, "moved with warning"

    detail = (result.stderr or result.stdout or "move failed").strip().splitlines()
    return False, detail[-1] if detail else "move failed"


def shorten_text(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    if max_chars <= 3:
        return text[:max_chars]
    return text[: max_chars - 3] + "..."


def delete_file_to_recycle_bin(path: Path) -> tuple[bool, str]:
    if not path.exists():
        return False, "file no longer exists"
    if not path.is_file():
        return False, "only files can be deleted from the visualizer"

    if os.name != "nt":
        path.unlink()
        return True, "deleted"

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [
            ("hwnd", wintypes.HWND),
            ("wFunc", wintypes.UINT),
            ("pFrom", wintypes.LPCWSTR),
            ("pTo", wintypes.LPCWSTR),
            ("fFlags", wintypes.USHORT),
            ("fAnyOperationsAborted", wintypes.BOOL),
            ("hNameMappings", wintypes.LPVOID),
            ("lpszProgressTitle", wintypes.LPCWSTR),
        ]

    operation = SHFILEOPSTRUCTW()
    operation.hwnd = None
    operation.wFunc = 3  # FO_DELETE
    operation.pFrom = str(path) + "\0\0"
    operation.pTo = None
    operation.fFlags = 0x0040 | 0x0010 | 0x0400  # recycle, no extra confirm dialog, no error UI

    result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(operation))
    if result != 0:
        return False, f"Windows delete failed ({result})"
    if operation.fAnyOperationsAborted:
        return False, "delete canceled"
    return True, "moved to Recycle Bin"


def risky_delete_reasons(path: Path) -> list[str]:
    parts = [part.lower() for part in path.parts]
    path_text = str(path).lower()
    reasons: list[str] = []

    risky_fragments = {
        "windows": "inside the Windows folder",
        "system32": "inside System32",
        "syswow64": "inside SysWOW64",
        "winsxs": "inside the Windows component store",
        "program files": "inside Program Files",
        "program files (x86)": "inside Program Files (x86)",
        "programdata": "inside ProgramData",
        "appdata": "inside AppData",
        "system volume information": "inside System Volume Information",
        "boot": "inside a boot-related folder",
    }
    for fragment, reason in risky_fragments.items():
        if fragment in parts or f"\\{fragment}\\" in path_text:
            reasons.append(reason)

    risky_extensions = {
        ".sys": "system driver file",
        ".dll": "shared Windows/program library",
        ".drv": "driver file",
        ".ocx": "system component file",
        ".efi": "boot firmware file",
        ".mui": "Windows resource file",
        ".cat": "security catalog file",
        ".inf": "driver/install information file",
        ".exe": "executable program",
        ".msi": "installer package",
    }
    if path.suffix.lower() in risky_extensions:
        reasons.append(risky_extensions[path.suffix.lower()])

    if path.parent == Path(path.anchor):
        reasons.append("sitting directly at the root of a drive")

    return sorted(set(reasons))


class DesktopOrganizer:
    def __init__(self, desktop_path: Path):
        self.desktop_path = desktop_path
        self.target_root = desktop_path / DECLUTTER_FOLDER_NAME

    def category_for(self, path: Path) -> str:
        if path.is_dir():
            lowered = path.name.lower()
            for category, tokens in NAME_RULES.items():
                if any(token in lowered for token in tokens):
                    return category
            return "Folders"

        lowered_name = path.name.lower()
        for category, tokens in NAME_RULES.items():
            if any(token in lowered_name for token in tokens):
                return category

        extension = path.suffix.lower()
        for category, extensions in EXTENSION_CATEGORIES.items():
            if extension in extensions:
                return category
        return "Other"

    def is_protected(self, path: Path) -> bool:
        protected_names = {
            "desktop.ini",
            "$recycle.bin",
            "recycle bin",
            DECLUTTER_FOLDER_NAME.lower(),
        }
        if path.name.lower() in protected_names:
            return True
        try:
            if path.resolve() == PROJECT_ROOT or PROJECT_ROOT in path.resolve().parents:
                return True
        except OSError:
            pass
        return False

    def scan(self, include_folders: bool) -> list[DesktopItem]:
        items: list[DesktopItem] = []
        if not self.desktop_path.exists():
            return items

        for entry in sorted(self.desktop_path.iterdir(), key=lambda p: p.name.lower()):
            if self.is_protected(entry):
                continue
            if entry.is_dir() and not include_folders:
                continue

            category = self.category_for(entry)
            if entry.is_dir():
                size, note = folder_size_quick(entry)
                kind = "Folder"
            else:
                try:
                    size = entry.stat().st_size
                    note = ""
                except OSError:
                    size = 0
                    note = "unreadable"
                kind = entry.suffix.lower() or "File"

            target = self.target_root / category / entry.name
            items.append(DesktopItem(entry, category, size, kind, target, note))
        return items

    def organize(
        self,
        items: list[DesktopItem],
        progress_callback=None,
        timeout_seconds: int = MOVE_TIMEOUT_SECONDS,
    ) -> OrganizeSummary:
        if not items:
            raise ValueError("No items selected to organize.")

        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        manifest = {
            "app": APP_NAME,
            "version": APP_VERSION,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "desktop": str(self.desktop_path),
            "moves": [],
            "skipped": [],
        }

        skipped: list[SkippedMove] = []
        total = len(items)
        for index, item in enumerate(items, start=1):
            target_dir = item.target.parent
            target_dir.mkdir(parents=True, exist_ok=True)
            destination = unique_destination(item.target)
            if progress_callback:
                progress_callback(index - 1, total, item, destination, "moving", "")

            moved, reason = move_with_timeout(item.path, destination, timeout_seconds=timeout_seconds)
            if not moved:
                skipped_item = SkippedMove(item.path, item.category, reason)
                skipped.append(skipped_item)
                manifest["skipped"].append(
                    {
                        "source": str(item.path),
                        "category": item.category,
                        "size": item.size,
                        "kind": item.kind,
                        "reason": reason,
                    }
                )
                if progress_callback:
                    progress_callback(index, total, item, destination, "skipped", reason)
                continue

            manifest["moves"].append(
                {
                    "source": str(item.path),
                    "destination": str(destination),
                    "category": item.category,
                    "size": item.size,
                    "kind": item.kind,
                }
            )
            if progress_callback:
                progress_callback(index, total, item, destination, "moved", reason)

        MANIFEST_DIR.mkdir(parents=True, exist_ok=True)
        manifest_path = MANIFEST_DIR / f"declutter-{timestamp}.json"
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        return OrganizeSummary(manifest_path, len(manifest["moves"]), skipped)

    def latest_manifest(self) -> Path | None:
        if not MANIFEST_DIR.exists():
            return None
        manifests = sorted(MANIFEST_DIR.glob("declutter-*.json"), key=lambda p: p.stat().st_mtime)
        return manifests[-1] if manifests else None

    def undo(self, manifest_path: Path, progress_callback=None) -> tuple[int, list[str]]:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        restored = 0
        problems: list[str] = []
        moves = list(reversed(manifest.get("moves", [])))
        total = len(moves)

        for index, move in enumerate(moves, start=1):
            source = Path(move["source"])
            destination = Path(move["destination"])
            if not destination.exists():
                problems.append(f"Missing: {destination}")
                if progress_callback:
                    progress_callback(index, total, destination.name)
                continue

            source.parent.mkdir(parents=True, exist_ok=True)
            restore_to = source if not source.exists() else unique_destination(source)
            moved, reason = move_with_timeout(destination, restore_to)
            if moved:
                restored += 1
            else:
                problems.append(f"{destination.name}: {reason}")
            if progress_callback:
                progress_callback(index, total, destination.name)

        return restored, problems


class DiskScanner:
    def __init__(self, stop_event: threading.Event, progress: queue.Queue):
        self.stop_event = stop_event
        self.progress = progress

    def scan(self, root: Path, max_depth: int = 5, max_children: int = 120) -> ScanNode:
        return self._scan_dir(root, depth=0, max_depth=max_depth, max_children=max_children)

    def _scan_dir(self, path: Path, depth: int, max_depth: int, max_children: int) -> ScanNode:
        node = ScanNode(path=path, name=path.name or str(path))
        if self.stop_event.is_set():
            return node

        try:
            entries = list(os.scandir(path))
        except OSError as exc:
            node.error = str(exc)
            return node

        child_nodes: list[ScanNode] = []
        for entry in entries:
            if self.stop_event.is_set():
                break
            try:
                if entry.is_symlink():
                    continue
                if entry.is_dir(follow_symlinks=False):
                    node.dirs += 1
                    if depth < max_depth:
                        child = self._scan_dir(Path(entry.path), depth + 1, max_depth, max_children)
                        node.size += child.size
                        node.files += child.files
                        node.dirs += child.dirs
                        child_nodes.append(child)
                    else:
                        size, files = self._estimate_dir(Path(entry.path))
                        node.size += size
                        node.files += files
                        child_nodes.append(
                            ScanNode(
                                path=Path(entry.path),
                                name=entry.name,
                                size=size,
                                files=files,
                                dirs=0,
                                error="depth limit",
                            )
                        )
                elif entry.is_file(follow_symlinks=False):
                    size = entry.stat(follow_symlinks=False).st_size
                    node.size += size
                    node.files += 1
                    if depth < 2:
                        child_nodes.append(
                            ScanNode(path=Path(entry.path), name=entry.name, size=size, files=1, is_file=True)
                        )
            except OSError:
                continue

        child_nodes.sort(key=lambda item: item.size, reverse=True)
        if len(child_nodes) > max_children:
            kept = child_nodes[:max_children]
            remainder = child_nodes[max_children:]
            kept.append(
                ScanNode(
                    path=path,
                    name=f"Other {len(remainder)} items",
                    size=sum(item.size for item in remainder),
                    files=sum(item.files for item in remainder),
                    dirs=sum(item.dirs for item in remainder),
                )
            )
            child_nodes = kept
        node.children = child_nodes

        if depth <= 1:
            self.progress.put(("progress", f"Scanned {path} ({human_size(node.size)})"))
        return node

    def _estimate_dir(self, path: Path, limit: int = 300) -> tuple[int, int]:
        total = 0
        files = 0
        try:
            for root, dirs, names in os.walk(path):
                if self.stop_event.is_set():
                    break
                dirs[:] = [name for name in dirs if not name.startswith("$")]
                for name in names:
                    try:
                        total += (Path(root) / name).stat().st_size
                    except OSError:
                        continue
                    files += 1
                    if files >= limit:
                        return total, files
        except OSError:
            return total, files
        return total, files


class DeclutterApp(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1180x760")
        self.minsize(980, 640)

        self.desktop_path = get_desktop_path()
        self.organizer = DesktopOrganizer(self.desktop_path)
        self.desktop_items: list[DesktopItem] = []
        self.organizer_busy = False
        self.organizer_thread: threading.Thread | None = None
        self.scan_root = StringVar(value=str(self.desktop_path.anchor or self.desktop_path))
        self.desktop_var = StringVar(value=str(self.desktop_path))
        self.include_folders = BooleanVar(value=False)
        self.status_var = StringVar(value="Ready")
        self.disk_status_var = StringVar(value="Choose a drive or folder, then scan.")
        self.progress_queue: queue.Queue = queue.Queue()
        self.stop_event = threading.Event()
        self.scan_thread: threading.Thread | None = None
        self.disk_root_node: ScanNode | None = None
        self.hover_node: ScanNode | None = None
        self.selected_node: ScanNode | None = None
        self.delete_in_progress = False
        self.largest_nodes_cache: list[ScanNode] = []
        self.rect_nodes: list[tuple[float, float, float, float, ScanNode, int]] = []

        self.configure(bg="#f5f4f0")
        self._configure_styles()
        self._build_ui()
        self._poll_queue()

    def _configure_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TNotebook", background="#f5f4f0", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(18, 10), font=("Segoe UI", 10))
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10), fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10))
        style.configure("TLabelframe", background="#f5f4f0")
        style.configure("TLabelframe.Label", background="#f5f4f0", font=("Segoe UI Semibold", 10))
        style.configure("TButton", padding=(10, 7), font=("Segoe UI", 10))

    def _build_ui(self) -> None:
        shell = Frame(self, bg="#f5f4f0", padx=16, pady=14)
        shell.pack(fill=BOTH, expand=True)

        header = Frame(shell, bg="#f5f4f0")
        header.pack(fill=X, pady=(0, 12))
        Label(header, text="Declutter", bg="#f5f4f0", fg="#1d2428", font=("Segoe UI Semibold", 20)).pack(side=LEFT)
        Label(
            header,
            text="Smart desktop folders and drive space maps",
            bg="#f5f4f0",
            fg="#5d666b",
            font=("Segoe UI", 10),
        ).pack(side=LEFT, padx=(14, 0), pady=(7, 0))

        notebook = ttk.Notebook(shell)
        notebook.pack(fill=BOTH, expand=True)

        organizer_tab = Frame(notebook, bg="#f5f4f0")
        visualizer_tab = Frame(notebook, bg="#f5f4f0")
        notebook.add(organizer_tab, text="Desktop Organizer")
        notebook.add(visualizer_tab, text="Drive Graph")

        self._build_organizer_tab(organizer_tab)
        self._build_visualizer_tab(visualizer_tab)

        status = Label(shell, textvariable=self.status_var, anchor="w", bg="#f5f4f0", fg="#5d666b", font=("Segoe UI", 9))
        status.pack(fill=X, pady=(10, 0))

    def _build_organizer_tab(self, parent: Frame) -> None:
        top = LabelFrame(parent, text="Desktop source", padx=12, pady=10)
        top.pack(fill=X, pady=(8, 12))

        Entry(top, textvariable=self.desktop_var, font=("Segoe UI", 10)).pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        self.browse_button = Button(top, text="Browse", command=self.choose_desktop)
        self.browse_button.pack(side=LEFT, padx=(0, 8))
        self.analyze_button = Button(top, text="Analyze", command=self.analyze_desktop)
        self.analyze_button.pack(side=LEFT)
        self.include_folders_check = Checkbutton(
            top,
            text="Include folders",
            variable=self.include_folders,
            bg="#f5f4f0",
            font=("Segoe UI", 10),
        )
        self.include_folders_check.pack(side=LEFT, padx=(12, 0))

        body = PanedWindow(parent, orient=HORIZONTAL, sashwidth=8, bg="#f5f4f0", bd=0)
        body.pack(fill=BOTH, expand=True)

        table_frame = Frame(body, bg="#f5f4f0")
        body.add(table_frame, stretch="always")

        columns = ("name", "category", "kind", "size", "target", "note")
        self.desktop_tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
        headings = {
            "name": "Name",
            "category": "Folder",
            "kind": "Type",
            "size": "Size",
            "target": "Destination",
            "note": "Note",
        }
        widths = {"name": 220, "category": 120, "kind": 90, "size": 90, "target": 300, "note": 110}
        for column in columns:
            self.desktop_tree.heading(column, text=headings[column])
            self.desktop_tree.column(column, width=widths[column], anchor="w")
        self.desktop_tree.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.desktop_tree.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.desktop_tree.configure(yscrollcommand=scrollbar.set)

        side = Frame(body, bg="#f5f4f0", padx=12)
        body.add(side, width=280)

        summary = LabelFrame(side, text="Summary", padx=12, pady=10)
        summary.pack(fill=X)
        self.summary_label = Label(
            summary,
            text="Analyze your desktop to see suggested folders.",
            justify=LEFT,
            bg="#f5f4f0",
            fg="#293238",
            font=("Segoe UI", 10),
            wraplength=235,
        )
        self.summary_label.pack(anchor="w")

        actions = LabelFrame(side, text="Actions", padx=12, pady=10)
        actions.pack(fill=X, pady=(12, 0))
        self.organize_selected_button = Button(actions, text="Organize selected", command=self.organize_selected)
        self.organize_selected_button.pack(fill=X, pady=(0, 8))
        self.organize_all_button = Button(actions, text="Organize all", command=self.organize_all)
        self.organize_all_button.pack(fill=X, pady=(0, 8))
        self.undo_button = Button(actions, text="Undo last run", command=self.undo_last)
        self.undo_button.pack(fill=X)
        self.organizer_progress = ttk.Progressbar(actions, mode="determinate", maximum=100)
        self.organizer_progress.pack(fill=X, pady=(12, 0))

        notes = LabelFrame(side, text="Rules", padx=12, pady=10)
        notes.pack(fill=BOTH, expand=True, pady=(12, 0))
        Label(
            notes,
            text=(
                "Items are grouped by file type and helpful name clues like "
                "screenshot, invoice, receipt, resume, and proposal. Existing folders "
                "stay put unless you enable Include folders."
            ),
            justify=LEFT,
            bg="#f5f4f0",
            fg="#4d585f",
            font=("Segoe UI", 9),
            wraplength=235,
        ).pack(anchor="w")

    def _build_visualizer_tab(self, parent: Frame) -> None:
        top = LabelFrame(parent, text="Scan target", padx=12, pady=10)
        top.pack(fill=X, pady=(8, 12))

        Entry(top, textvariable=self.scan_root, font=("Segoe UI", 10)).pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        Button(top, text="Drive", command=self.choose_drive).pack(side=LEFT, padx=(0, 8))
        Button(top, text="Folder", command=self.choose_scan_folder).pack(side=LEFT, padx=(0, 8))
        self.scan_button = Button(top, text="Scan", command=self.start_disk_scan)
        self.scan_button.pack(side=LEFT, padx=(0, 8))
        self.stop_button = Button(top, text="Stop", command=self.stop_disk_scan, state=DISABLED)
        self.stop_button.pack(side=LEFT)

        body = PanedWindow(parent, orient=HORIZONTAL, sashwidth=8, bg="#f5f4f0", bd=0)
        body.pack(fill=BOTH, expand=True)

        graph_panel = Frame(body, bg="#f5f4f0")
        body.add(graph_panel, stretch="always")

        self.disk_canvas = Canvas(graph_panel, bg="#ffffff", highlightthickness=1, highlightbackground="#d8d5cb")
        self.disk_canvas.pack(fill=BOTH, expand=True)
        self.disk_canvas.bind("<Configure>", lambda _event: self.draw_treemap())
        self.disk_canvas.bind("<Motion>", self.on_canvas_motion)
        self.disk_canvas.bind("<Button-1>", self.on_canvas_click)
        self.disk_canvas.bind("<Leave>", lambda _event: self.set_hover(None))

        disk_status = Label(
            graph_panel,
            textvariable=self.disk_status_var,
            anchor="w",
            bg="#f5f4f0",
            fg="#5d666b",
            font=("Segoe UI", 9),
        )
        disk_status.pack(fill=X, pady=(8, 0))

        side = Frame(body, bg="#f5f4f0", padx=12)
        body.add(side, width=315)

        detail = LabelFrame(side, text="Selected area", padx=12, pady=10)
        detail.pack(fill=X)
        self.hover_label = Label(
            detail,
            text="Select an item in the chart or largest list.",
            justify=LEFT,
            bg="#f5f4f0",
            fg="#293238",
            font=("Segoe UI", 10),
            wraplength=260,
        )
        self.hover_label.pack(anchor="w")
        self.delete_selected_button = Button(
            detail,
            text="Delete selected file",
            command=self.delete_selected_file,
            state=DISABLED,
            fg="#8a2d24",
        )
        self.delete_selected_button.pack(fill=X, pady=(10, 0))

        largest = LabelFrame(side, text="Largest items", padx=12, pady=10)
        largest.pack(fill=BOTH, expand=True, pady=(12, 0))
        self.largest_list = Listbox(largest, font=("Segoe UI", 10), activestyle="none", height=18)
        self.largest_list.pack(side=LEFT, fill=BOTH, expand=True)
        largest_scroll = ttk.Scrollbar(largest, orient=VERTICAL, command=self.largest_list.yview)
        largest_scroll.pack(side=RIGHT, fill=Y)
        self.largest_list.configure(yscrollcommand=largest_scroll.set)
        self.largest_list.bind("<<ListboxSelect>>", self.on_largest_select)

    def choose_desktop(self) -> None:
        selected = filedialog.askdirectory(initialdir=str(self.desktop_path), title="Choose desktop or folder to organize")
        if selected:
            self.desktop_path = Path(selected)
            self.desktop_var.set(selected)
            self.organizer = DesktopOrganizer(self.desktop_path)
            self.status_var.set(f"Desktop source set to {selected}")

    def analyze_desktop(self) -> None:
        if self.organizer_busy:
            self.status_var.set("Desktop work is already running.")
            return

        self.desktop_path = Path(self.desktop_var.get()).expanduser()
        if not self.desktop_path.exists():
            messagebox.showerror(APP_NAME, f"{self.desktop_path} does not exist.")
            return

        desktop_path = self.desktop_path
        self.organizer = DesktopOrganizer(self.desktop_path)
        include_folders = self.include_folders.get()
        self.desktop_items = []
        self._clear_desktop_tree()
        self.summary_label.configure(text="Analyzing your desktop...")
        self._set_organizer_busy(True, "Analyzing desktop...", mode="indeterminate")

        def worker() -> None:
            try:
                organizer = DesktopOrganizer(desktop_path)
                items = organizer.scan(include_folders=include_folders)
            except Exception as exc:
                self.progress_queue.put(("desktop_analysis_error", exc))
            else:
                self.progress_queue.put(("desktop_analysis_done", desktop_path, items))

        self.organizer_thread = threading.Thread(target=worker, daemon=True)
        self.organizer_thread.start()

    def _clear_desktop_tree(self) -> None:
        for row in self.desktop_tree.get_children():
            self.desktop_tree.delete(row)

    def _render_desktop_items(self, items: list[DesktopItem]) -> None:
        self._clear_desktop_tree()
        for index, item in enumerate(items):
            self.desktop_tree.insert(
                "",
                END,
                iid=str(index),
                values=(
                    item.path.name,
                    item.category,
                    item.kind,
                    human_size(item.size),
                    str(item.target.parent),
                    item.note,
                ),
            )

        categories: dict[str, int] = {}
        total_size = 0
        for item in items:
            categories[item.category] = categories.get(item.category, 0) + 1
            total_size += item.size

        if items:
            category_text = ", ".join(f"{name}: {count}" for name, count in sorted(categories.items()))
            self.summary_label.configure(
                text=f"{len(items)} movable items found ({human_size(total_size)}).\n\n{category_text}"
            )
            self.status_var.set("Analysis complete. Select rows to organize a subset, or organize all.")
        else:
            self.summary_label.configure(text="No movable desktop items found with the current settings.")
            self.status_var.set("No movable items found.")

    def _set_organizer_busy(
        self,
        busy: bool,
        message: str = "",
        mode: str = "determinate",
        maximum: int = 100,
    ) -> None:
        self.organizer_busy = busy
        state = DISABLED if busy else NORMAL
        for control in (
            self.browse_button,
            self.analyze_button,
            self.include_folders_check,
            self.organize_selected_button,
            self.organize_all_button,
            self.undo_button,
        ):
            control.configure(state=state)

        self.organizer_progress.stop()
        if busy:
            self.organizer_progress.configure(mode=mode, maximum=max(1, maximum), value=0)
            if mode == "indeterminate":
                self.organizer_progress.start(12)
        else:
            self.organizer_progress.configure(mode="determinate", maximum=100, value=0)

        if message:
            self.status_var.set(message)

    def selected_items(self) -> list[DesktopItem]:
        selected = self.desktop_tree.selection()
        if not selected:
            return []
        return [self.desktop_items[int(item_id)] for item_id in selected]

    def organize_selected(self) -> None:
        items = self.selected_items()
        if not items:
            messagebox.showinfo(APP_NAME, "Select one or more rows first.")
            return
        self._organize_items(items)

    def organize_all(self) -> None:
        if not self.desktop_items:
            messagebox.showinfo(APP_NAME, "Analyze your desktop first.")
            return
        self._organize_items(self.desktop_items)

    def _organize_items(self, items: list[DesktopItem]) -> None:
        if self.organizer_busy:
            self.status_var.set("Desktop work is already running.")
            return

        items = list(items)
        response = messagebox.askyesno(
            APP_NAME,
            f"Move {len(items)} item(s) into {self.organizer.target_root}?\n\n"
            f"Items that take longer than {MOVE_TIMEOUT_SECONDS} seconds will be skipped. "
            "An undo manifest will be saved before this finishes.",
        )
        if not response:
            return

        total = len(items)
        self._set_organizer_busy(True, f"Organizing 0/{total} item(s)...", maximum=total)

        def progress(
            index: int,
            count: int,
            item: DesktopItem,
            destination: Path,
            status: str,
            reason: str,
        ) -> None:
            self.progress_queue.put(("organize_progress", index, count, item.path.name, destination, status, reason))

        def worker() -> None:
            try:
                summary = self.organizer.organize(items, progress_callback=progress)
            except Exception as exc:
                self.progress_queue.put(("organize_error", exc))
            else:
                self.progress_queue.put(("organize_done", total, summary))

        self.organizer_thread = threading.Thread(target=worker, daemon=True)
        self.organizer_thread.start()

    def undo_last(self) -> None:
        if self.organizer_busy:
            self.status_var.set("Desktop work is already running.")
            return

        manifest = self.organizer.latest_manifest()
        if not manifest:
            messagebox.showinfo(APP_NAME, "No prior organize run was found.")
            return

        response = messagebox.askyesno(APP_NAME, f"Undo the moves in this manifest?\n\n{manifest}")
        if not response:
            return

        try:
            manifest_data = json.loads(manifest.read_text(encoding="utf-8"))
            total = len(manifest_data.get("moves", []))
        except Exception:
            total = 1

        self._set_organizer_busy(True, f"Undoing 0/{total} item(s)...", maximum=total)

        def progress(index: int, count: int, name: str) -> None:
            self.progress_queue.put(("undo_progress", index, count, name))

        def worker() -> None:
            try:
                restored, problems = self.organizer.undo(manifest, progress_callback=progress)
            except Exception as exc:
                self.progress_queue.put(("undo_error", exc))
            else:
                self.progress_queue.put(("undo_done", restored, problems))

        self.organizer_thread = threading.Thread(target=worker, daemon=True)
        self.organizer_thread.start()

    def choose_drive(self) -> None:
        drives = self.available_drives()
        if not drives:
            messagebox.showinfo(APP_NAME, "No drives were detected.")
            return

        chooser = Toplevel(self)
        chooser.title("Choose drive")
        chooser.geometry("320x300")
        chooser.configure(bg="#f5f4f0")
        Label(chooser, text="Choose a drive to scan", bg="#f5f4f0", font=("Segoe UI Semibold", 11)).pack(
            anchor="w", padx=14, pady=(14, 8)
        )
        drive_list = Listbox(chooser, font=("Segoe UI", 10), height=8)
        drive_list.pack(fill=BOTH, expand=True, padx=14)
        for drive in drives:
            drive_list.insert(END, drive)

        def use_selection() -> None:
            selection = drive_list.curselection()
            if selection:
                self.scan_root.set(drive_list.get(selection[0]))
                chooser.destroy()

        Button(chooser, text="Use selected", command=use_selection).pack(pady=12)
        drive_list.bind("<Double-Button-1>", lambda _event: use_selection())

    def choose_scan_folder(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.scan_root.get() or str(Path.home()), title="Choose folder to graph")
        if selected:
            self.scan_root.set(selected)

    def available_drives(self) -> list[str]:
        drives = []
        bitmask = ctypes.windll.kernel32.GetLogicalDrives() if os.name == "nt" else 0
        for index in range(26):
            if bitmask & (1 << index):
                drive = f"{chr(65 + index)}:\\"
                if Path(drive).exists():
                    drives.append(drive)
        if not drives:
            drives.append(str(Path.home().anchor or Path.home()))
        return drives

    def start_disk_scan(self) -> None:
        root = Path(self.scan_root.get()).expanduser()
        if not root.exists():
            messagebox.showerror(APP_NAME, f"{root} does not exist.")
            return

        self.stop_event.clear()
        self.disk_root_node = None
        self.select_node(None)
        self.largest_nodes_cache = []
        self.rect_nodes.clear()
        self.largest_list.delete(0, END)
        self.disk_canvas.delete("all")
        self.disk_status_var.set(f"Scanning {root}. Large drives can take a while.")
        self.scan_button.configure(state=DISABLED)
        self.stop_button.configure(state=NORMAL)

        def worker() -> None:
            started = time.time()
            scanner = DiskScanner(self.stop_event, self.progress_queue)
            node = scanner.scan(root)
            elapsed = time.time() - started
            self.progress_queue.put(("done", node, elapsed))

        self.scan_thread = threading.Thread(target=worker, daemon=True)
        self.scan_thread.start()

    def stop_disk_scan(self) -> None:
        self.stop_event.set()
        self.disk_status_var.set("Stopping scan...")

    def _poll_queue(self) -> None:
        processed = 0
        try:
            while processed < 100:
                event = self.progress_queue.get_nowait()
                processed += 1
                if event[0] == "progress":
                    self.disk_status_var.set(event[1])
                elif event[0] == "desktop_analysis_done":
                    _tag, desktop_path, items = event
                    self.desktop_path = desktop_path
                    self.organizer = DesktopOrganizer(self.desktop_path)
                    self.desktop_items = items
                    self._set_organizer_busy(False)
                    self._render_desktop_items(items)
                elif event[0] == "desktop_analysis_error":
                    _tag, exc = event
                    self._set_organizer_busy(False, "Desktop analysis failed.")
                    self.summary_label.configure(text="Analysis could not finish.")
                    messagebox.showerror(APP_NAME, f"Could not analyze desktop:\n{exc}")
                elif event[0] == "organize_progress":
                    _tag, index, total, name, _destination, status, reason = event
                    self.organizer_progress.configure(maximum=max(1, total), value=index)
                    if status == "moving":
                        self.status_var.set(f"Moving {index + 1}/{total}: {name}")
                    elif status == "skipped":
                        self.status_var.set(f"Skipped {index}/{total}: {name} ({reason})")
                    else:
                        self.status_var.set(f"Organized {index}/{total}: {name}")
                elif event[0] == "organize_done":
                    _tag, total, summary = event
                    skipped_count = len(summary.skipped)
                    message = (
                        f"Organized {summary.moved_count} of {total} item(s).\n\n"
                        f"Undo manifest:\n{summary.manifest_path}"
                    )
                    if skipped_count:
                        skipped_lines = "\n".join(
                            f"- {item.source.name}: {item.reason}" for item in summary.skipped[:8]
                        )
                        message += f"\n\nSkipped {skipped_count} item(s):\n{skipped_lines}"
                        if skipped_count > 8:
                            message += f"\n...and {skipped_count - 8} more."
                    self._set_organizer_busy(
                        False,
                        f"Organized {summary.moved_count} of {total}; skipped {skipped_count}.",
                    )
                    messagebox.showinfo(APP_NAME, message)
                    self.analyze_desktop()
                elif event[0] == "organize_error":
                    _tag, exc = event
                    self._set_organizer_busy(False, "Organizing failed.")
                    messagebox.showerror(APP_NAME, f"Could not organize desktop:\n{exc}")
                elif event[0] == "undo_progress":
                    _tag, index, total, name = event
                    self.organizer_progress.configure(maximum=max(1, total), value=index)
                    self.status_var.set(f"Undoing {index}/{total}: {name}")
                elif event[0] == "undo_done":
                    _tag, restored, problems = event
                    self._set_organizer_busy(False)
                    message = f"Restored {restored} item(s)."
                    if problems:
                        message += "\n\nSome items need attention:\n" + "\n".join(problems[:8])
                    self.status_var.set(message)
                    messagebox.showinfo(APP_NAME, message)
                    self.analyze_desktop()
                elif event[0] == "undo_error":
                    _tag, exc = event
                    self._set_organizer_busy(False, "Undo failed.")
                    messagebox.showerror(APP_NAME, f"Could not undo:\n{exc}")
                elif event[0] == "delete_done":
                    _tag, path, ok, detail = event
                    self.delete_in_progress = False
                    if ok:
                        self.remove_node_from_scan(path)
                        self.select_node(None)
                        self.populate_largest_list()
                        self.disk_status_var.set(f"Deleted {path.name}.")
                    else:
                        self.update_delete_button()
                        self.disk_status_var.set(f"Could not delete {path.name}: {detail}")
                        messagebox.showerror(APP_NAME, f"Could not delete this file:\n{detail}")
                elif event[0] == "done":
                    _tag, node, elapsed = event
                    self.disk_root_node = node
                    self.scan_button.configure(state=NORMAL)
                    self.stop_button.configure(state=DISABLED)
                    suffix = "Stopped" if self.stop_event.is_set() else "Scan complete"
                    self.disk_status_var.set(
                        f"{suffix}: {node.path} | {human_size(node.size)} | "
                        f"{node.files:,} files | {node.dirs:,} folders | {elapsed:.1f}s"
                    )
                    self.draw_treemap()
                    self.populate_largest_list()
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    def populate_largest_list(self) -> None:
        self.largest_list.delete(0, END)
        self.largest_nodes_cache = []
        if not self.disk_root_node:
            return
        self.largest_nodes_cache = self.largest_nodes(self.disk_root_node, LARGEST_LIST_LIMIT)
        for node in self.largest_nodes_cache:
            kind = "FILE" if node.is_file else "FOLDER"
            if node is self.disk_root_node:
                kind = "ROOT"
            self.largest_list.insert(END, f"{human_size(node.size):>9}  {kind:<6} {node.path}")

    def largest_nodes(self, root: ScanNode, limit: int) -> list[ScanNode]:
        largest: list[ScanNode] = []
        heap: list[tuple[int, int, ScanNode]] = [(-root.size, 0, root)]
        sequence = 1

        while heap and len(largest) < limit:
            _size, _sequence, node = heapq.heappop(heap)
            largest.append(node)
            for child in node.children[:LARGEST_LIST_LIMIT]:
                heapq.heappush(heap, (-child.size, sequence, child))
                sequence += 1

        return largest

    def draw_treemap(self) -> None:
        self.disk_canvas.delete("all")
        self.rect_nodes.clear()
        if not self.disk_root_node:
            width = max(self.disk_canvas.winfo_width(), 400)
            height = max(self.disk_canvas.winfo_height(), 300)
            self.disk_canvas.create_text(
                width / 2,
                height / 2,
                text="Scan a drive or folder to render a treemap.",
                fill="#5d666b",
                font=("Segoe UI", 12),
            )
            return

        width = max(self.disk_canvas.winfo_width(), 520)
        height = max(self.disk_canvas.winfo_height(), 340)
        root_label = shorten_text(str(self.disk_root_node.path), max(20, int((width - 170) / 7)))

        self.disk_canvas.create_text(
            14,
            12,
            text=root_label,
            anchor="nw",
            fill="#1d2428",
            font=("Segoe UI Semibold", 11),
        )
        self.disk_canvas.create_text(
            width - 14,
            12,
            text=human_size(self.disk_root_node.size),
            anchor="ne",
            fill="#4d585f",
            font=("Segoe UI", 10),
        )
        self.disk_canvas.create_line(12, 42, width - 12, 42, fill="#ddd8cd")

        if self.disk_root_node.size <= 0:
            self.disk_canvas.create_text(
                width / 2,
                height / 2,
                text="No files were found in this scan.",
                fill="#5d666b",
                font=("Segoe UI", 12),
            )
            return

        pad = 12
        self._draw_treemap_node(
            self.disk_root_node,
            pad,
            54,
            width - pad,
            height - pad,
            depth=0,
            vertical=False,
            color_index=0,
        )

        visible_children = len([child for child in self.disk_root_node.children if child.size > 0])
        if visible_children > TREEMAP_CHILD_LIMIT:
            self.disk_canvas.create_text(
                14,
                height - 18,
                text=f"Showing the largest {TREEMAP_CHILD_LIMIT} items at each level; inspect details in the side panel.",
                anchor="w",
                fill="#ffffff",
                font=("Segoe UI", 9),
            )

    def _draw_treemap_node(
        self,
        node: ScanNode,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        depth: int,
        vertical: bool,
        color_index: int,
    ) -> None:
        width = x2 - x1
        height = y2 - y1
        if width < 4 or height < 4 or node.size <= 0:
            return

        selected = self.selected_node is not None and node.path == self.selected_node.path and node.name == self.selected_node.name
        color = COLORS[color_index % len(COLORS)]
        outline = "#1d2428" if selected else "#ffffff"
        outline_width = 3 if selected else 1
        self.disk_canvas.create_rectangle(x1, y1, x2, y2, fill=color, outline=outline, width=outline_width)
        self.rect_nodes.append((x1, y1, x2, y2, node, depth))

        area = width * height
        can_label = area >= TREEMAP_LABEL_MIN_AREA and width >= 92 and height >= 52
        leaf_label = node.is_file and area >= TREEMAP_LEAF_LABEL_MIN_AREA and width >= 120 and height >= 70
        if can_label or leaf_label or depth == 0:
            label_width = max(8, int((width - 14) / 7))
            label_name = shorten_text(node.name or str(node.path), label_width)
            label = f"{label_name}\n{human_size(node.size)}"
            self.disk_canvas.create_text(
                x1 + 7,
                y1 + 7,
                text=label,
                anchor="nw",
                fill="#ffffff",
                font=("Segoe UI Semibold", 9 if depth else 10),
                width=max(40, int(width - 14)),
            )

        if depth >= TREEMAP_MAX_DEPTH:
            return

        children = [child for child in node.children if child.size > 0][:TREEMAP_CHILD_LIMIT]
        total = sum(child.size for child in children)
        if total <= 0:
            return

        inset = 5 if depth else 8
        header = 26 if (can_label or depth == 0) and height > 72 else inset
        cx1, cy1, cx2, cy2 = x1 + inset, y1 + header, x2 - inset, y2 - inset
        if cx2 - cx1 < 8 or cy2 - cy1 < 8:
            return

        cursor = cy1 if vertical else cx1
        for child_index, child in enumerate(children):
            ratio = child.size / total
            if vertical:
                next_cursor = cursor + (cy2 - cy1) * ratio
                self._draw_treemap_node(
                    child,
                    cx1,
                    cursor,
                    cx2,
                    next_cursor,
                    depth + 1,
                    not vertical,
                    color_index + child_index + 1,
                )
                cursor = next_cursor
            else:
                next_cursor = cursor + (cx2 - cx1) * ratio
                self._draw_treemap_node(
                    child,
                    cursor,
                    cy1,
                    next_cursor,
                    cy2,
                    depth + 1,
                    not vertical,
                    color_index + child_index + 1,
                )
                cursor = next_cursor

    def on_canvas_motion(self, event) -> None:
        node = self.node_at(event.x, event.y)
        self.disk_canvas.configure(cursor="hand2" if node else "")
        if self.selected_node is None:
            self.set_hover(node)

    def on_canvas_click(self, event) -> None:
        self.select_node(self.node_at(event.x, event.y))

    def node_at(self, x: int, y: int) -> ScanNode | None:
        match: tuple[int, ScanNode] | None = None
        for x1, y1, x2, y2, node, depth in self.rect_nodes:
            if x1 <= x <= x2 and y1 <= y <= y2:
                if match is None or depth > match[0]:
                    match = (depth, node)
        return match[1] if match else None

    def on_largest_select(self, _event) -> None:
        selection = self.largest_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index < len(self.largest_nodes_cache):
            self.select_node(self.largest_nodes_cache[index])

    def set_hover(self, node: ScanNode | None) -> None:
        if self.selected_node is not None:
            return
        if node is self.hover_node:
            return
        self.hover_node = node
        if not node:
            self.hover_label.configure(text="Select an item in the chart or largest list.")
            return
        self.display_node_detail(node, "Hovering")

    def select_node(self, node: ScanNode | None) -> None:
        self.selected_node = node
        self.hover_node = None
        if node:
            self.display_node_detail(node, "Selected")
        else:
            self.hover_label.configure(text="Select an item in the chart or largest list.")
        self.update_delete_button()
        self.draw_treemap()

    def display_node_detail(self, node: ScanNode, label: str) -> None:
        kind = "File" if node.is_file else "Folder"
        if node is self.disk_root_node:
            kind = "Scan root"
        percent = ""
        if self.disk_root_node and self.disk_root_node.size:
            percent = f"\nShare of scan: {(node.size / self.disk_root_node.size) * 100:.1f}%"
        delete_note = "" if node.is_file else "\n\nDelete is available for files only."
        if node.is_file and risky_delete_reasons(node.path):
            delete_note = "\n\nThis file looks system-ish. Delete will ask the scary question."
        self.hover_label.configure(
            text=(
                f"{label}: {node.name}\n"
                f"{kind}\n"
                f"{human_size(node.size)}\n"
                f"{node.files:,} files | {node.dirs:,} folders"
                f"{percent}\n\n{node.path}"
                f"{delete_note}"
            )
        )

    def update_delete_button(self) -> None:
        can_delete = self.selected_node is not None and self.selected_node.is_file and not self.delete_in_progress
        self.delete_selected_button.configure(state=NORMAL if can_delete else DISABLED)

    def delete_selected_file(self) -> None:
        node = self.selected_node
        if not node:
            messagebox.showinfo(APP_NAME, "Select a file first.")
            return
        if not node.is_file:
            messagebox.showinfo(APP_NAME, "Only files can be deleted from the visualizer.")
            return
        if not node.path.exists():
            messagebox.showinfo(APP_NAME, "That file is no longer on disk.")
            self.remove_node_from_scan(node.path)
            self.select_node(None)
            self.populate_largest_list()
            return

        risky_reasons = risky_delete_reasons(node.path)
        if risky_reasons:
            confirmed = messagebox.askyesno(
                APP_NAME,
                RISKY_DELETE_MESSAGE.format(
                    path=node.path,
                    reasons="\n".join(f"- {reason}" for reason in risky_reasons),
                ),
            )
        else:
            confirmed = messagebox.askyesno(APP_NAME, f"Move this file to the Recycle Bin?\n\n{node.path}")
        if not confirmed:
            return

        self.delete_in_progress = True
        self.update_delete_button()
        self.disk_status_var.set(f"Deleting {node.path.name}...")

        def worker() -> None:
            try:
                ok, detail = delete_file_to_recycle_bin(node.path)
            except Exception as exc:
                ok, detail = False, str(exc)
            self.progress_queue.put(("delete_done", node.path, ok, detail))

        threading.Thread(target=worker, daemon=True).start()

    def remove_node_from_scan(self, path: Path) -> bool:
        if not self.disk_root_node:
            return False
        removed, _size, _files, _dirs = self._remove_node_from_parent(self.disk_root_node, path)
        return removed

    def _remove_node_from_parent(self, node: ScanNode, path: Path) -> tuple[bool, int, int, int]:
        for child in list(node.children):
            if child.path == path:
                node.children.remove(child)
                node.size = max(0, node.size - child.size)
                node.files = max(0, node.files - child.files)
                node.dirs = max(0, node.dirs - child.dirs)
                return True, child.size, child.files, child.dirs

            removed, size, files, dirs = self._remove_node_from_parent(child, path)
            if removed:
                node.size = max(0, node.size - size)
                node.files = max(0, node.files - files)
                node.dirs = max(0, node.dirs - dirs)
                return True, size, files, dirs

        return False, 0, 0, 0


def main() -> None:
    app = DeclutterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
