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

ARIAL_NAMES = {"BlockGothicRR LightExtraCond", "ArialMT"}
REPLACEMENT_FONT = "Alright Sans Medium"


def replace_if_arial(value):
    if value in ARIAL_NAMES:
        return REPLACEMENT_FONT, True
    return value, False


def process_styles_xml(xml_bytes):
    root = ET.fromstring(xml_bytes)
    changed = 0

    for rfonts in root.findall(f".//{{{W_NS}}}rFonts"):
        for attr in (
            f"{{{W_NS}}}ascii",
            f"{{{W_NS}}}hAnsi",
            f"{{{W_NS}}}eastAsia",
            f"{{{W_NS}}}cs",
        ):
            old = rfonts.get(attr)
            new, did_change = replace_if_arial(old)
            if did_change:
                rfonts.set(attr, new)
                changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_theme_xml(xml_bytes):
    root = ET.fromstring(xml_bytes)
    changed = 0

    for elem in root.iter():
        typeface = elem.get("typeface")
        new, did_change = replace_if_arial(typeface)
        if did_change:
            elem.set("typeface", new)
            changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_docx(docx_path, make_backup=False, dry_run=False):
    total_changes = 0
    modified = False

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        with zipfile.ZipFile(docx_path, "r") as zin:
            zin.extractall(tmpdir)

        styles_path = tmpdir / "word" / "styles.xml"
        if styles_path.exists():
            new_bytes, changes = process_styles_xml(styles_path.read_bytes())
            if changes:
                styles_path.write_bytes(new_bytes)
                total_changes += changes
                modified = True

        theme_dir = tmpdir / "word" / "theme"
        if theme_dir.exists():
            for theme_file in theme_dir.glob("*.xml"):
                new_bytes, changes = process_theme_xml(theme_file.read_bytes())
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
        return [p for p in folder.rglob("*.docx") if not p.name.startswith("~$")]
    return [p for p in folder.glob("*.docx") if not p.name.startswith("~$")]


def main():
    parser = argparse.ArgumentParser(
        description="Replace Arial with Aptos in Word .docx styles."
    )

    parser.add_argument("folder", type=Path, help="Folder containing .docx files")

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

    if not folder.exists():
        raise SystemExit(f"Folder not found: {folder}")

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
                make_backup=args.b,
                dry_run=args.d
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
