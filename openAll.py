from __future__ import annotations

import os
import time
from pathlib import Path


EXTS = {".doc", ".docx", ".ppt", ".pptx"}
DELAY_SECONDS = 0.5  # 👈 change this if you want (e.g. 1.0)


def open_file(path: Path) -> bool:
    """Open a file using the default associated application on Windows."""
    try:
        os.startfile(str(path))  # Windows only
        return True
    except OSError as e:
        print(f"[ERROR] Could not open: {path}\n        {e}")
        return False


def main() -> int:
    root = Path.cwd()
    print(f"Scanning under: {root}\n")

    opened = 0
    matched = 0

    for p in root.rglob("*"):
        if not p.is_file():
            continue

        if p.suffix.lower() in EXTS:
            matched += 1
            print(f"Opening: {p}")

            if open_file(p):
                opened += 1
                time.sleep(DELAY_SECONDS)

    print("\nDone.")
    print(f"Matched files: {matched}")
    print(f"Opened files:  {opened}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
