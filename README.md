# Office Automations

A collection of small, pragmatic Python utilities for automating repetitive Microsoft Office and document-related tasks. The tools focus on PowerPoint and Word workflows and are designed to be used as standalone scripts.

> **Status:** Actively evolving. Scripts may be added, renamed, or refined as new use‑cases come up.

---

## Contents

### PowerPoint utilities

- **`PPTXtoPOTX.py`**  
  Recursively converts `.pptx` files into `.potx` templates. Useful for standardising decks or extracting reusable templates. Does **not** require Microsoft Office to be installed.

- **`countSlides.py`**  
  Counts the number of slides across one or more PowerPoint files. Handy for audits, quick reporting, or validating large slide collections.

- **`footerMover.py`**  
  GUI automation script that steps through slides and assists with repositioning or adjusting footers. Intended for bulk footer clean‑up when PowerPoint’s native tools fall short.

- **`footerRemover.py`**  
  Companion GUI automation script to remove footers from slides in bulk.

- **`searchBinding.py`**  
  Searches PowerPoint files for specific bindings / placeholders and reports matches. Useful when cleaning up templates or validating brand assets.

- **`extractTheme.py`**  
  Extracts theme information from PowerPoint files for inspection or reuse.

- **`insertTheme.py`**  
  Applies or inserts a theme into PowerPoint files programmatically.

### Word utilities

- **`mergeWord.py`**  
  Merges multiple Word documents into a single file while preserving structure as much as possible.

### Other

- **`fillForm/`**  
  Experimental or work‑in‑progress scripts related to automated form filling.

---

## Requirements

- Python 3.9+
- Dependencies listed in `requirements.txt`

Some scripts rely on:
- `python-pptx`
- `python-docx`
- GUI automation libraries (for footer scripts)

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

(See individual scripts for configurable paths, constants, or arguments.)

---

## Tips & troubleshooting

- **Scripts running too fast / too slow**  
  Adjust `time.sleep(...)` values inside the script.

- **Accidentally spamming inputs (GUI automation)**  
  Press `q` in the terminal to exit. Consider adding an initial delay before automation starts.

- **Unexpected PowerPoint behaviour**  
  Make sure no modal dialogs are open and that the target window is focused.

---

## Contributing

Pull requests are welcome. If you have additional Office or document‑automation chores to add:

1. Open an issue describing the use‑case, or
2. Submit a PR with a clearly named script and brief inline documentation.

---

## License

MIT — use at your own risk. No warranties provided.
