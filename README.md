# Office Automations

A collection of small, pragmatic Python utilities for automating repetitive Microsoft Office and document-related tasks.

The repo mainly focuses on **PowerPoint** and **Word** workflows, with scripts designed to run as standalone tools.

> **Status:** Actively evolving — scripts may be added, renamed, or refined as new use-cases come up.

---

## Contents

### PowerPoint utilities

- **`PPTXtoPOTX.py`**  
  Recursively converts `.pptx` files into `.potx` templates.

- **`countSlides.py`**  
  Counts the number of slides across PowerPoint files.

- **`pptFindFont.py`**  
  Searches PowerPoint files for specific fonts.

- **`pptImageReplacer/`**  
  Utilities for replacing images inside PowerPoint files.

- **`extractTheme.py`**  
  Extracts `theme1.xml` from both `.pptx` and `.docx` files.

- **`insertTheme.py`**  
  Inserts or replaces themes in PowerPoint and Word files.

- **`searchBinding.py`**  
  Searches PowerPoint files for specific bindings / placeholders.

- **`footerMover.py`**  
  GUI automation utility for repositioning slide footers.

- **`footerRemover.py`**  
  GUI automation utility for removing slide footers.

---

### Word utilities

- **`mergeWord.py`**  
  Merges multiple Word documents into one.

- **`findAssetID.py`**  
  Searches `.docx` / `.dotx` packages for asset IDs.

- **`replaceWordColors.py`**  
  Replaces color values inside Word documents.

- **`replaceWordFonts.py`**  
  Replaces fonts in Word documents.

- **`replaceWordFontsize.py`**  
  Adjusts font sizes in Word documents.

---

### XML / package utilities

- **`setSharedValues.py`**  
  Ensures XML form fields use `shareValue = true`.

---

### Automation & helpers

- **`fillForm/`**  
  GUI automation tooling for filling forms using scripted instructions.

- **`humanTyping.py`**  
  Simulates human-like typing for automation flows.

- **`openAll.py`**  
  Opens multiple Office files in bulk.

- **`alpha/`**  
  Experimental and work-in-progress scripts.

---

## Requirements

- Python 3.9+
- Install dependencies from `requirements.txt`

Common dependencies include:
- `python-pptx`
- `python-docx`
- XML / ZIP handling
- GUI automation libraries (e.g. `pyautogui`)

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
- support CLI arguments, or
- use configurable constants at the top of the file.

---

## Tips & troubleshooting

- **GUI automation behaving weirdly**  
  Ensure the correct window is focused and no dialogs are open.

- **Scripts too fast/slow**  
  Adjust `time.sleep()` values.

- **Automation safety**  
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
