## Description

Add cross-platform guidance and headless/CLI operation while preserving the existing GUI workflow.

## Extended description

This change makes the tool more reliable on macOS/Linux and in non-GUI environments.

### What changed
- Added a CLI/headless mode: you can now run `python process_consumers_itemized_statement.py /path/to/statement.pdf`.
  - When no argument is provided, the original Tk file picker is still used.
- Made `tkinter` imports lazy/optional so importing the module and running in CLI mode does not fail on platforms where Tk is not installed.
  - If the GUI picker cannot open (missing Tk or no display), the program prints a clear error message and suggests CLI mode.
- Updated README with Linux/macOS notes:
  - How to activate venv on those platforms.
  - How to install Tk on common Linux distros.
  - How to use CLI mode for headless/server runs.
- Added unit tests covering the new argument parsing helper.

### Why
- Some Linux Python installs omit Tk by default (`python3-tk`/`python3-tkinter`).
- Headless environments (CI, servers) cannot open a GUI dialog.

## How to verify
- `python -m pytest` (all tests should pass)
- GUI path: `python process_consumers_itemized_statement.py` (file dialog opens on systems with Tk)
- CLI path: `python process_consumers_itemized_statement.py path/to/statement.pdf`

## Notes
- Documentation continues to describe PDF as the supported input format; Excel/XLSX import remains available in the code but is intentionally not advertised in the README.