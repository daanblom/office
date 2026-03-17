# Office Automations

A collection of small, pragmatic Python utilities for automating repetitive Microsoft Office and document-related tasks.

The repo mainly focuses on **PowerPoint** and **Word** workflows, with scripts designed to run as standalone tools.

> **Status:** Actively evolving — scripts may be added, renamed, or refined as new use-cases come up.

---

## Contents

### PowerPoint utilities

* **`PPTXtoPOTX.py`**
  Recursively converts `.pptx` files into `.potx` templates.

* **`countSlides.py`**
  Counts the number of slides across PowerPoint files.

* **`pptFindFont.py`**
  Searches PowerPoint files for specific fonts.

* **`pptImageReplacer/`**
  Utilities for replacing images inside PowerPoint files.

* **`footerMover.py`**
  GUI automation script to reposition footers.

* **`footerRemover.py`**
  GUI automation script to remove footers.

* **`searchBinding.py`**
  Searches for bindings/placeholders in PowerPoint files.

* **`extractTheme.py`**
  Extracts `theme1.xml` from `.pptx` and `.docx` files.

* **`insertTheme.py`**
  Inserts or replaces themes in PowerPoint/Word files.

---

### Word utilities

* **`mergeWord.py`**
  Merges multiple Word documents into one.

* **`findAssetID.py`**
  Searches `.docx` / `.dotx` files for Templafy asset IDs.

---

### XML / package utilities

* **`setSharedValues.py`**
  Ensures form XML fields have `shareValue = true`.

---

### Automation & helpers

* **`fillForm/`**
  GUI automation for filling forms based on instruction files.

* **`humanTyping.py`**
  Simulates human-like typing input (useful while creating screenflows).

* **`openAll.py`**
  Opens multiple files in bulk (useful for manual inspection workflows).

---

### Experimental (`alpha/`)

Work-in-progress scripts for Word document styling:

* **`replaceWordColors.py`**
  Replace colors inside Word documents.

* **`replaceWordFonts.py`**
  Replace fonts in Word files.

* **`replaceWordFontsize.py`**
  Adjust font sizes in Word documents.

---

## Requirements

* Python 3.9+
* Install dependencies from `requirements.txt`

Common dependencies include:

* `python-pptx`
* `python-docx`
* XML / ZIP handling
* GUI automation libraries (e.g. `pyautogui`)

> ⚠️ **Note:** GUI automation scripts interact with your mouse and keyboard. Close other applications and use with care.

---

## Quick start

```bash
python -m venv .venv

# Activate
.venv\Scripts\activate  # Windows
source .venv/bin/activate   # macOS/Linux

pip install -r requirements.txt
```

Run a script:

```bash
python extractTheme.py input.pptx out
```

Most scripts:

* support CLI arguments, or
* use configurable constants at the top of the file.

---

## Tips & troubleshooting

* **GUI automation behaving weirdly**
  Ensure the correct window is focused and no dialogs are open.

* **Scripts too fast/slow**
  Adjust `time.sleep()` values.

* **Automation safety**
  Keep your mouse away during execution or add startup delays.

---

## Contributing

Feel free to add scripts for repetitive Office tasks.

---

## License

MIT — use at your own risk.

# Bonus scripts / funtions

add this bash funtion to your `.bashrc` or `.zshrc` to jump to a directory from a windows path (great if you are using WSL)

usage: `gop "WINDOWS_PATH"`

```bash
gop() {
  if [ -z "$1" ]; then
    echo "Usage: gop <windows-path>"
    return 1
  fi

  local og="$1"
  local np
  np="$(printf '%s\n' "$og" | sed -E 's#^([A-Za-z]):#\/mnt\/\L\1\/#; s#\\#\/#g')"

  cd "$np"
}
