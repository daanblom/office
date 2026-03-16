#!/usr/bin/env python3
"""
Replace Arial with Aptos in Word .docx style definitions.

What it changes:
- word/styles.xml
  - style run fonts (<w:rFonts ...>)
  - default run properties
- word/theme/*.xml
  - theme font entries that reference Arial

What it does NOT fully rewrite:
- direct per-run formatting inside document content (word/document.xml, headers, footers, etc.)
  unless that formatting is inherited from styles/theme.

Usage:
    python replace_arial_with_aptos.py /path/to/folder
    python replace_arial_with_aptos.py /path/to/folder --recursive
    python replace_arial_with_aptos.py /path/to/folder --no-backup
"""

from __future__ import annotations

import argparse
import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

ET.register_namespace("w", W_NS)
ET.register_namespace("a", A_NS)

ARIAL_NAMES = {
    "Arial",
    "ArialMT",
}
REPLACEMENT_FONT = "Aptos"


def replace_if_arial(value: str | None) -> tuple[str | None, bool]:
    if value in ARIAL_NAMES:
        return REPLACEMENT_FONT, True
    return value, False


def process_styles_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    changed = 0

    # Replace font declarations in all <w:rFonts> elements
    for rfonts in root.findall(f".//{{{W_NS}}}rFonts"):
        for attr_name in (
            f"{{{W_NS}}}ascii",
            f"{{{W_NS}}}hAnsi",
            f"{{{W_NS}}}eastAsia",
            f"{{{W_NS}}}cs",
        ):
            old = rfonts.get(attr_name)
            new, did_change = replace_if_arial(old)
            if did_change:
                rfonts.set(attr_name, new)
                changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_theme_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    changed = 0

    # Common theme font elements like:
    # <a:latin typeface="Arial"/>
    # <a:ea typeface="Arial"/>
    # <a:cs typeface="Arial"/>
    for elem in root.iter():
        typeface = elem.get("typeface")
        new, did_change = replace_if_arial(typeface)
        if did_change:
            elem.set("typeface", new)
            changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_docx(docx_path: Path, make_backup: bool) -> tuple[bool, int]:
    total_changes = 0
    modified = False

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        with zipfile.ZipFile(docx_path, "r") as zin:
            zin.extractall(tmpdir_path)

        styles_path = tmpdir_path / "word" / "styles.xml"
        if styles_path.exists():
            xml_bytes = styles_path.read_bytes()
            new_bytes, changes = process_styles_xml(xml_bytes)
            if changes > 0:
                styles_path.write_bytes(new_bytes)
                total_changes += changes
                modified = True

        theme_dir = tmpdir_path / "word" / "theme"
        if theme_dir.exists():
            for theme_file in theme_dir.glob("*.xml"):
                xml_bytes = theme_file.read_bytes()
                new_bytes, changes = process_theme_xml(xml_bytes)
                if changes > 0:
                    theme_file.write_bytes(new_bytes)
                    total_changes += changes
                    modified = True

        if modified:
            if make_backup:
                backup_path = docx_path.with_suffix(docx_path.suffix + ".bak")
                if not backup_path.exists():
                    shutil.copy2(docx_path, backup_path)

            new_docx = docx_path.with_suffix(docx_path.suffix + ".tmp")
            with zipfile.ZipFile(new_docx, "w", zipfile.ZIP_DEFLATED) as zout:
                for file_path in tmpdir_path.rglob("*"):
                    if file_path.is_file():
                        arcname = file_path.relative_to(tmpdir_path).as_posix()
                        zout.write(file_path, arcname)

            new_docx.replace(docx_path)

    return modified, total_changes


def find_docx_files(folder: Path, recursive: bool) -> list[Path]:
    if recursive:
        return sorted(
            p for p in folder.rglob("*.docx")
            if p.is_file() and not p.name.startswith("~$")
        )
    return sorted(
        p for p in folder.glob("*.docx")
        if p.is_file() and not p.name.startswith("~$")
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Replace Arial with Aptos in Word .docx style definitions."
    )
    parser.add_argument("folder", type=Path, help="Folder containing .docx files")
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search subfolders too",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Do not create .bak backup files",
    )
    args = parser.parse_args()

    folder = args.folder
    if not folder.exists() or not folder.is_dir():
        raise SystemExit(f"Folder does not exist or is not a directory: {folder}")

    docx_files = find_docx_files(folder, args.recursive)
    if not docx_files:
        print("No .docx files found.")
        return

    total_files_changed = 0
    total_replacements = 0

    for docx_path in docx_files:
        try:
            modified, changes = process_docx(
                docx_path=docx_path,
                make_backup=not args.no_backup,
            )
            if modified:
                total_files_changed += 1
                total_replacements += changes
                print(f"[CHANGED] {docx_path} ({changes} replacements)")
            else:
                print(f"[OK]      {docx_path} (no Arial found in styles/theme)")
        except Exception as exc:
            print(f"[ERROR]   {docx_path}: {exc}")

    print()
    print(f"Files scanned:   {len(docx_files)}")
    print(f"Files changed:   {total_files_changed}")
    print(f"Replacements:    {total_replacements}")


if __name__ == "__main__":
    main()
