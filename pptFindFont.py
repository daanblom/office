#!/usr/bin/env python3
"""
pptFindFont.py

Find (and optionally replace) font references inside a .pptx by scanning XML parts.

Examples:

  # Find occurrences only
  python pptFindFont.py deck.pptx "System Font Regular"

  # Replace font → write to new file
  python pptFindFont.py deck.pptx "System Font Regular" -r "Calibri" --out deck_fixed.pptx

  # Replace font → overwrite original file
  python pptFindFont.py deck.pptx "System Font Regular" -r "Calibri" -o
"""

import argparse
import os
import re
import sys
import zipfile
import tempfile
from collections import defaultdict
from typing import Dict, List, Tuple


def iter_xml_parts(z: zipfile.ZipFile) -> List[str]:
    return [n for n in z.namelist() if n.lower().endswith(".xml")]


def find_matches(text: str, needle: str, case_sensitive: bool) -> List[int]:
    flags = 0 if case_sensitive else re.IGNORECASE
    return [m.start() for m in re.finditer(re.escape(needle), text, flags=flags)]


def snippet(text: str, idx: int, radius: int = 90) -> str:
    start = max(0, idx - radius)
    end = min(len(text), idx + radius)
    s = text[start:end].replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return f"…{s}…"


def scan_pptx(pptx_path: str, font_name: str, case_sensitive: bool):
    hits: Dict[str, List[str]] = defaultdict(list)

    with zipfile.ZipFile(pptx_path, "r") as z:
        for part in iter_xml_parts(z):
            text = z.read(part).decode("utf-8", errors="ignore")
            positions = find_matches(text, font_name, case_sensitive)

            if positions:
                for pos in positions[:20]:
                    hits[part].append(snippet(text, pos))
                if len(positions) > 20:
                    hits[part].append(f"(+{len(positions)-20} more matches)")

    return hits


def replace_font(
    pptx_path: str,
    font_name: str,
    replace_with: str,
    out_path: str,
    case_sensitive: bool,
) -> Tuple[int, int]:
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(re.escape(font_name), flags)

    parts_changed = 0
    total_replacements = 0

    with zipfile.ZipFile(pptx_path, "r") as zin:
        with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename.lower().endswith(".xml"):
                    text = data.decode("utf-8", errors="ignore")
                    new_text, n = pattern.subn(replace_with, text)

                    if n > 0:
                        parts_changed += 1
                        total_replacements += n
                        data = new_text.encode("utf-8")

                zout.writestr(item, data)

    return parts_changed, total_replacements


def main():
    parser = argparse.ArgumentParser(
        description="Find (and optionally replace) font references inside PPTX XML."
    )

    parser.add_argument("pptx", help="Input .pptx file")
    parser.add_argument("font", help="Font name to search for")

    parser.add_argument(
        "-r",
        dest="replace_with",
        metavar="NEWFONT",
        help="Replace the font with NEWFONT",
    )

    parser.add_argument(
        "--out",
        help="Write output to this file (required unless using -o)",
    )

    parser.add_argument(
        "-o",
        action="store_true",
        help="Overwrite the original PPTX file (only valid with -r)",
    )

    parser.add_argument(
        "--case-sensitive",
        action="store_true",
        help="Case-sensitive search/replace (default: case-insensitive)",
    )

    args = parser.parse_args()

    if not os.path.isfile(args.pptx):
        print("ERROR: File not found:", args.pptx)
        sys.exit(2)

    # --------------------
    # Scan first
    # --------------------
    hits = scan_pptx(args.pptx, args.font, args.case_sensitive)

    if not hits:
        print(f"No XML references found for: {args.font!r}")
    else:
        print(f"Found references to {args.font!r} in {len(hits)} part(s):\n")
        for part in sorted(hits):
            print(f"- {part}")
            for s in hits[part]:
                print(" ", s)
            print()

    # --------------------
    # Replacement mode
    # --------------------
    if args.replace_with:

        if args.o:
            # overwrite mode → write temp file first
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pptx")
            os.close(tmp_fd)

            out_path = tmp_path
            final_path = args.pptx

        else:
            if not args.out:
                print("ERROR: Must provide --out unless using -o overwrite mode.")
                sys.exit(2)

            out_path = args.out
            final_path = out_path

        parts_changed, total_replacements = replace_font(
            args.pptx,
            args.font,
            args.replace_with,
            out_path,
            args.case_sensitive,
        )

        # If overwrite, move temp file into place
        if args.o:
            os.replace(out_path, final_path)

        print("\nReplacement complete:")
        print("  File written:", final_path)
        print("  Parts changed:", parts_changed)
        print("  Total replacements:", total_replacements)


if __name__ == "__main__":
    main()
