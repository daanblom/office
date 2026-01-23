# Office Automations

A collection of small, pragmatic Python utilities for automating repetitive Microsoft Office and document-related tasks.

The repo mainly focuses on **PowerPoint** and **Word** workflows, with scripts designed to run as standalone tools.

> **Status:** Actively evolving — scripts may be added, renamed, or refined as new use-cases come up.

---

## Contents

### PowerPoint utilities

- **`PPTXtoPOTX.py`**  
  Recursively converts `.pptx` files into `.potx` templates. Useful for standardising decks or extracting reusable templates.

- **`countSlides.py`**  
  Counts the number of slides across one or more PowerPoint files.

- **`footerMover.py`**  
  GUI automation script to assist with repositioning footers on slides in bulk.

- **`footerRemover.py`**  
  GUI automation script to remove footers from slides in bulk.

- **`searchBinding.py`**  
  Searches PowerPoint files for specific bindings / placeholders and reports matches.

- **`extractTheme.py`**  
  Extracts theme information from Word or PowerPoint files for inspection or reuse.

- **`insertTheme.py`**  
  Applies or inserts a theme into Word or PowerPoint files programmatically.

### Word utilities

- **`mergeWord.py`**  
  Merges multiple Word documents into a single output document.

- **`findAssetID.py`**  
  Searches Word `.docx` / `.dotx` packages for a given **asset ID** and reports which internal parts contain it.

### Other

- **`setSharedValues.py`**  
  Updates form-like XML content by ensuring fields have `shareValue = true` (supports a `--dry-run` mode).

- **`fillForm/`**  
  GUI automation tooling for automated form filling. Includes `fillForm.py` and an `example.txt` instruction file.

---

## Requirements

- Python 3.9+
- Install dependencies from `requirements.txt`

Some scripts rely on:
- `python-pptx`
- `python-docx`
- ZIP/XML parsing utilities
- GUI automation libraries (e.g. `pyautogui`)

> ⚠️ **Note:** GUI automation scripts interact with your mouse and keyboard. Close other applications and use with care.

---

## Quick start

```bash
# 1) Create a virtual environment (recommended)
python -m venv .venv

# 2) Activate it
# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

# 3) Install dependencies
pip install -r requirements.txt
```

Run a script directly, for example:

```bash
python PPTXtoPOTX.py /path/to/presentations
```

Most scripts either:
- accept CLI arguments (via `argparse`), or
- can be adjusted via constants at the top of the file.

---

## Tips & troubleshooting

- **Scripts running too fast / too slow**  
  Adjust `time.sleep(...)` values inside the script.

- **Accidentally spamming inputs (GUI automation)**  
  Stop the script with `Ctrl + C`. Consider adding an initial delay before automation starts.

- **Unexpected PowerPoint / Word behaviour**  
  Make sure no modal dialogs are open and that the target window is focused.

---

## Contributing

Pull requests are welcome. If you have additional Office or document-automation chores to add:

1. Open an issue describing the use-case, or
2. Submit a PR with a clearly named script and brief inline documentation.

---

## License

MIT — use at your own risk. No warranties provided.
