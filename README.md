# Declutter

Declutter is a local Windows-friendly desktop app that helps organize desktop icons into folders and visualize disk usage as a treemap.

## Features

- Scans your Desktop and suggests category folders for documents, screenshots, PDFs, shortcuts, installers, archives, code, media, and more.
- Moves selected or all suggested items into `Decluttered Desktop/<Category>`.
- Keeps desktop analysis, organizing, undo, and drive scanning off the UI thread so the window stays responsive.
- Shows a progress bar while organizing and undoing desktop moves.
- Skips individual desktop items that take too long to move, then reports what was skipped.
- Skips the app's own project folder and `desktop.ini`.
- Stores an undo manifest in `%LOCALAPPDATA%\DeclutterApp\manifests`.
- Can undo the most recent organize run.
- Scans any drive or folder and renders a treemap of disk usage.
- Shows a treemap, largest scanned items, and selected-item details.
- Lets you select a file in the visualizer and move it to the Recycle Bin.
- Gives a skeptical extra warning before deleting files that look system-critical.

## Run

Double-click `Launch Declutter.bat`, or run it from PowerShell:

```powershell
.\Launch Declutter.bat
```

You can also launch the Python file directly:

```powershell
python .\declutter_app.py
```

If `python` is not registered on the machine, try:

```powershell
py .\declutter_app.py
```

## Notes

The drive graph intentionally skips symlinks and ignores folders it cannot read. A full-drive scan can take a while, especially the first time.
