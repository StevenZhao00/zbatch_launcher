# zbatch_launcher (Z Launch Assistant)

## Introduction

**Z Launch Assistant** is a simple and practical tool for batch launching multiple programs. It supports organizing your frequently used software into groups and launching all items in a group with a single click, greatly improving your productivity for work, study, and development.
Supports adding shortcuts and `.exe` files by drag & drop, automatic deduplication, and configuration saving. Designed for Windows platform.

---

## Architecture

* **UI Layer:**
  Built with `Tkinter` and `tkinterdnd2` for a clean main interface and intuitive drag-and-drop functionality.
* **Logic Layer:**
  Supports group management (add/delete groups), batch add/delete/paste/drag executables (`.exe`/`.lnk`).
* **Persistence:**
  Software group and path data are automatically saved to a local file `apps_groups.json`.
* **Batch Launch:**
  One-click to launch all programs in the current group, with automatic handling of executables and shortcuts.
* **Dependencies:**
  Relies only on Python standard library and `tkinterdnd2`; optionally, `pywin32` for shortcut (`.lnk`) parsing.

---

## Installation

1. **Prepare Python environment (Windows, Python 3.8+ recommended)**
2. **Install dependencies**

   ```bash
   pip install tkinterdnd2 pywin32
   ```
3. **Download source code or release package**

   * Clone this repository or download and extract the zip file.
4. **Run the main program**

   ```bash
   python zbatch_launcher.py
   ```

   Or double-click the packaged exe file (if available).

---

## Usage

* **Create a Group:**
  Click the "Add Group" button to create different program groups for various scenarios.
* **Add Programs:**
  Click "Add Program" to select `.exe` or `.lnk` files, or drag files directly into the list. Shortcuts are automatically recognized and resolved.
* **One-click Launch:**
  Click the "Launch All in Current Group" button to start all programs in the current group in sequence.
* **Paste Path:**
  Right-click the list to paste an `.exe` or `.lnk` path from the clipboard.
* **Group/Program Management:**
  Supports deleting groups, removing single or multiple programs, and automatically saves the order of groups and programs.
* **Exit the Program:**
  Click the "Exit" button to close the application.

---

## Contributing

1. Fork this repository
2. Create a `Feat_xxx` branch
3. Commit your code
4. Open a Pull Request

---

## Other Notes

* **Shortcut parsing** requires `pywin32`; if errors occur, you may manually add the `.exe` path.
* **Data file** is saved as `apps_groups.json` and will be auto-created if missing.
* **Supports custom extension**—feel free to submit Issues or PRs for new feature suggestions.
