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


def w_tag(name):
    return f"{{{W_NS}}}{name}"


def normalize_hex(value):
    value = value.strip().lstrip("#").upper()
    if len(value) != 6:
        raise ValueError(f"Invalid hex color '{value}'. Use 6-digit hex like FF0000.")
    int(value, 16)
    return value


def clear_theme_attrs(elem):
    for attr in (
        w_tag("themeColor"),
        w_tag("themeTint"),
        w_tag("themeShade"),
        w_tag("themeFill"),
        w_tag("themeFillTint"),
        w_tag("themeFillShade"),
    ):
        elem.attrib.pop(attr, None)


def is_explicit_hex_color(color_elem):
    if color_elem is None:
        return False

    val = color_elem.get(w_tag("val"))
    if not val:
        return False
    if val.lower() == "auto":
        return False
    if color_elem.get(w_tag("themeColor")) is not None:
        return False

    try:
        int(val, 16)
        return len(val) == 6
    except ValueError:
        return False


def is_explicit_auto_or_theme(color_elem):
    """
    AUTO mode should only touch styles that explicitly declare
    automatic/theme-based color in this style node itself.

    It should NOT treat a missing <w:color> as AUTO, because that may
    intentionally inherit a real color from basedOn/defaults.
    """
    if color_elem is None:
        return False

    if color_elem.get(w_tag("themeColor")) is not None:
        return True

    val = color_elem.get(w_tag("val"))
    if val is not None and val.lower() == "auto":
        return True

    return False


def set_color_elem(color_elem, target_hex):
    changed = False

    old_val = color_elem.get(w_tag("val"))
    if old_val != target_hex:
        color_elem.set(w_tag("val"), target_hex)
        changed = True

    before = dict(color_elem.attrib)
    clear_theme_attrs(color_elem)
    if dict(color_elem.attrib) != before:
        changed = True

    return changed


def replace_auto_color_in_rpr(rpr, target_hex):
    color = rpr.find(w_tag("color"))
    if is_explicit_auto_or_theme(color):
        return set_color_elem(color, target_hex)
    return False


def replace_exact_color_in_rpr(rpr, source_hex, target_hex):
    color = rpr.find(w_tag("color"))
    if color is None:
        return False

    val = color.get(w_tag("val"))
    if val and val.upper() == source_hex and color.get(w_tag("themeColor")) is None:
        return set_color_elem(color, target_hex)

    return False


def replace_matching_attr(elem, attr_name, source_hex, target_hex):
    old = elem.get(attr_name)
    if old and old.upper() == source_hex:
        elem.set(attr_name, target_hex)
        clear_theme_attrs(elem)
        return True
    return False


def process_styles_xml(xml_bytes, target_hex, source_hex=None, auto_mode=False):
    root = ET.fromstring(xml_bytes)
    changed = 0

    # docDefaults: only change if color is explicitly auto/theme-based
    doc_defaults = root.find(w_tag("docDefaults"))
    if doc_defaults is not None:
        rpr_default = doc_defaults.find(w_tag("rPrDefault"))
        if rpr_default is not None:
            rpr = rpr_default.find(w_tag("rPr"))
            if rpr is not None:
                if auto_mode:
                    if replace_auto_color_in_rpr(rpr, target_hex):
                        changed += 1
                elif source_hex:
                    if replace_exact_color_in_rpr(rpr, source_hex, target_hex):
                        changed += 1

    # styles
    for style in root.findall(w_tag("style")):
        rpr = style.find(w_tag("rPr"))

        if rpr is not None:
            if auto_mode:
                if replace_auto_color_in_rpr(rpr, target_hex):
                    changed += 1
            elif source_hex:
                if replace_exact_color_in_rpr(rpr, source_hex, target_hex):
                    changed += 1

        # In exact-hex mode, also replace style shading/border colors
        if source_hex:
            for shd in style.findall(f".//{w_tag('shd')}"):
                for attr in (w_tag("fill"), w_tag("color")):
                    if replace_matching_attr(shd, attr, source_hex, target_hex):
                        changed += 1

            for border_name in ("top", "bottom", "left", "right", "insideH", "insideV"):
                for border in style.findall(f".//{w_tag(border_name)}"):
                    if replace_matching_attr(border, w_tag("color"), source_hex, target_hex):
                        changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_theme_xml(xml_bytes, source_hex, target_hex):
    root = ET.fromstring(xml_bytes)
    changed = 0

    for elem in root.iter():
        for attr in ("val", "lastClr", "rgb"):
            old = elem.get(attr)
            if old is not None and old.upper() == source_hex:
                elem.set(attr, target_hex)
                changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_docx(docx_path, target_hex, make_backup=False, dry_run=False, source_hex=None, auto_mode=False):
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
                target_hex=target_hex,
                source_hex=source_hex,
                auto_mode=auto_mode,
            )
            if changes:
                styles_path.write_bytes(new_bytes)
                total_changes += changes
                modified = True

        theme_dir = tmpdir / "word" / "theme"
        if theme_dir.exists() and source_hex:
            for theme_file in theme_dir.glob("*.xml"):
                new_bytes, changes = process_theme_xml(
                    theme_file.read_bytes(),
                    source_hex=source_hex,
                    target_hex=target_hex,
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
        description=(
            "Replace Word style colors. "
            "Use AUTO as source to replace only explicitly auto/theme-based font colors."
        )
    )

    parser.add_argument("folder", type=Path, help="Folder containing .docx files")
    parser.add_argument("source_color", help="Source hex color or AUTO")
    parser.add_argument("target_color", help="Target hex color, e.g. 000000")

    parser.add_argument("-r", "--recursive", action="store_true", help="Scan subfolders")
    parser.add_argument("-b", action="store_true", help="Create backup (.bak) files")
    parser.add_argument("-d", action="store_true", help="Dry run (show what would change without modifying files)")

    args = parser.parse_args()

    if not args.folder.exists() or not args.folder.is_dir():
        raise SystemExit(f"Folder not found or not a directory: {args.folder}")

    source_raw = args.source_color.strip().upper()
    target_hex = normalize_hex(args.target_color)

    if source_raw == "AUTO":
        auto_mode = True
        source_hex = None
    else:
        auto_mode = False
        source_hex = normalize_hex(source_raw)

    files = find_docx(args.folder, args.recursive)
    if not files:
        print("No .docx files found.")
        return

    changed_files = 0
    replacements = 0

    for docx in files:
        try:
            modified, changes = process_docx(
                docx,
                target_hex=target_hex,
                make_backup=args.b,
                dry_run=args.d,
                source_hex=source_hex,
                auto_mode=auto_mode,
            )

            if modified:
                changed_files += 1
                replacements += changes
                label = "[DRY RUN]" if args.d else "[CHANGED]"
                print(f"{label} {docx} ({changes} changes)")
            else:
                print(f"[OK]      {docx}")

        except Exception as e:
            print(f"[ERROR]   {docx}: {e}")

    print("\nSummary")
    print(f"Files scanned: {len(files)}")
    print(f"Files changed: {changed_files}")
    print(f"Total changes: {replacements}")


if __name__ == "__main__":
    main()
