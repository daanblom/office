#!/usr/bin/env python3

import argparse
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

ET.register_namespace("w", W_NS)
ET.register_namespace("a", A_NS)


def normalize_hex(value):
    value = value.strip().lstrip("#").upper()
    if len(value) != 6:
        raise ValueError(f"Invalid hex color '{value}'. Use 6-digit hex like FF0000.")
    int(value, 16)  # validate
    return value


def replace_if_match(value, source_hex, target_hex):
    if value is None:
        return value, False
    if value.upper() == source_hex:
        return target_hex, True
    return value, False


def process_styles_xml(xml_bytes, source_hex, target_hex):
    root = ET.fromstring(xml_bytes)
    changed = 0

    # Common Word style color attributes
    style_attrs = [
        (f".//{{{W_NS}}}color", f"{{{W_NS}}}val"),
        (f".//{{{W_NS}}}shd", f"{{{W_NS}}}fill"),
        (f".//{{{W_NS}}}shd", f"{{{W_NS}}}color"),
        (f".//{{{W_NS}}}top", f"{{{W_NS}}}color"),
        (f".//{{{W_NS}}}bottom", f"{{{W_NS}}}color"),
        (f".//{{{W_NS}}}left", f"{{{W_NS}}}color"),
        (f".//{{{W_NS}}}right", f"{{{W_NS}}}color"),
        (f".//{{{W_NS}}}insideH", f"{{{W_NS}}}color"),
        (f".//{{{W_NS}}}insideV", f"{{{W_NS}}}color"),
    ]

    for xpath, attr in style_attrs:
        for elem in root.findall(xpath):
            old = elem.get(attr)
            new, did_change = replace_if_match(old, source_hex, target_hex)
            if did_change:
                elem.set(attr, new)
                changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_theme_xml(xml_bytes, source_hex, target_hex):
    root = ET.fromstring(xml_bytes)
    changed = 0

    # Theme color definitions often use attributes like rgb / lastClr
    for elem in root.iter():
        for attr in ("lastClr", "val", "rgb"):
            old = elem.get(attr)
            new, did_change = replace_if_match(old, source_hex, target_hex)
            if did_change:
                elem.set(attr, new)
                changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_docx(docx_path, source_hex, target_hex, make_backup=False, dry_run=False):
    total_changes = 0
    modified = False

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        with zipfile.ZipFile(docx_path, "r") as zin:
            zin.extractall(tmpdir)

        styles_path = tmpdir / "word" / "styles.xml"
        if styles_path.exists():
            new_bytes, changes = process_styles_xml(
                styles_path.read_bytes(),
                source_hex,
                target_hex,
            )
            if changes:
                styles_path.write_bytes(new_bytes)
                total_changes += changes
                modified = True

        theme_dir = tmpdir / "word" / "theme"
        if theme_dir.exists():
            for theme_file in theme_dir.glob("*.xml"):
                new_bytes, changes = process_theme_xml(
                    theme_file.read_bytes(),
                    source_hex,
                    target_hex,
                )
                if changes:
                    theme_file.write_bytes(new_bytes)
                    total_changes += changes
                    modified = True

        if modified and not dry_run:
            if make_backup:
                backup = docx_path.with_suffix(docx_path.suffix + ".bak")
                if not backup.exists():
                    shutil.copy2(docx_path, backup)

            new_docx = docx_path.with_suffix(docx_path.suffix + ".tmp")

            with zipfile.ZipFile(new_docx, "w", zipfile.ZIP_DEFLATED) as zout:
                for f in tmpdir.rglob("*"):
                    if f.is_file():
                        zout.write(f, f.relative_to(tmpdir).as_posix())

            new_docx.replace(docx_path)

    return modified, total_changes


def find_docx(folder, recursive):
    if recursive:
        return sorted(
            p for p in folder.rglob("*.docx")
            if p.is_file() and not p.name.startswith("~$")
        )
    return sorted(
        p for p in folder.glob("*.docx")
        if p.is_file() and not p.name.startswith("~$")
    )


def main():
    parser = argparse.ArgumentParser(
        description="Replace a hex color with another hex color in Word .docx styles and theme files."
    )

    parser.add_argument("folder", type=Path, help="Folder containing .docx files")
    parser.add_argument("source_color", help="Source color, e.g. 000000")
    parser.add_argument("target_color", help="Target color, e.g. FF0000")

    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="Scan subfolders"
    )
    parser.add_argument(
        "-b",
        action="store_true",
        help="Create backup (.bak) files"
    )
    parser.add_argument(
        "-d",
        action="store_true",
        help="Dry run (show what would change without modifying files)"
    )

    args = parser.parse_args()

    folder = args.folder
    source_hex = normalize_hex(args.source_color)
    target_hex = normalize_hex(args.target_color)

    if not folder.exists() or not folder.is_dir():
        raise SystemExit(f"Folder not found or not a directory: {folder}")

    files = find_docx(folder, args.recursive)
    if not files:
        print("No .docx files found.")
        return

    changed_files = 0
    replacements = 0

    for docx in files:
        try:
            modified, changes = process_docx(
                docx,
                source_hex=source_hex,
                target_hex=target_hex,
                make_backup=args.b,
                dry_run=args.d,
            )

            if modified:
                changed_files += 1
                replacements += changes
                if args.d:
                    print(f"[DRY RUN] {docx} ({changes} replacements)")
                else:
                    print(f"[CHANGED] {docx} ({changes} replacements)")
            else:
                print(f"[OK]      {docx}")

        except Exception as e:
            print(f"[ERROR]   {docx}: {e}")

    print("\nSummary")
    print(f"Files scanned: {len(files)}")
    print(f"Files changed: {changed_files}")
    print(f"Total replacements: {replacements}")


if __name__ == "__main__":
    main()
