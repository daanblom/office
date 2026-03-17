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


def parse_half_points(value):
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def format_points_from_half_points(value):
    pts = value / 2
    if pts.is_integer():
        return str(int(pts))
    return str(pts)


def parse_increment_points(value):
    try:
        points = float(value)
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"Invalid increment '{value}'. Use a number like 0.5, 1, or 2."
        )

    half_points = points * 2
    if not half_points.is_integer():
        raise argparse.ArgumentTypeError(
            "Increment must be in 0.5 point steps, e.g. 0.5, 1, 1.5, 2."
        )

    return points


def increase_size_elem(elem, amount_points):
    """
    Increase a <w:sz> or <w:szCs> element by amount_points.
    Returns (changed, old_pts, new_pts)
    """
    val_attr = w_tag("val")
    old_raw = elem.get(val_attr)
    old_half_points = parse_half_points(old_raw)

    if old_half_points is None:
        return False, None, None

    increment_half_points = int(amount_points * 2)
    new_half_points = old_half_points + increment_half_points

    if new_half_points < 1:
        new_half_points = 1

    if new_half_points != old_half_points:
        elem.set(val_attr, str(new_half_points))
        return (
            True,
            format_points_from_half_points(old_half_points),
            format_points_from_half_points(new_half_points),
        )

    return False, None, None


def process_styles_xml(xml_bytes, amount_points):
    root = ET.fromstring(xml_bytes)
    changed = 0

    for tag_name in ("sz", "szCs"):
        for elem in root.findall(f".//{w_tag(tag_name)}"):
            did_change, _, _ = increase_size_elem(elem, amount_points)
            if did_change:
                changed += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def process_docx(docx_path, make_backup=False, dry_run=False, amount_points=1.0):
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
                amount_points=amount_points,
            )
            if changes:
                styles_path.write_bytes(new_bytes)
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
        description="Increase or decrease all Word style font sizes by a configurable number of points."
    )

    parser.add_argument("folder", type=Path, help="Folder containing .docx files")
    parser.add_argument(
        "amount",
        type=parse_increment_points,
        help="Amount in points to change sizes by, e.g. 0.5, 1, 1.5, 2. Use negative values to decrease.",
    )
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

    if not args.folder.exists() or not args.folder.is_dir():
        raise SystemExit(f"Folder not found or not a directory: {args.folder}")

    files = find_docx(args.folder, args.recursive)
    if not files:
        print("No .docx files found.")
        return

    changed_files = 0
    total_changes = 0

    for docx in files:
        try:
            modified, changes = process_docx(
                docx,
                make_backup=args.b,
                dry_run=args.d,
                amount_points=args.amount,
            )

            if modified:
                changed_files += 1
                total_changes += changes
                label = "[DRY RUN]" if args.d else "[CHANGED]"
                sign = "+" if args.amount >= 0 else ""
                print(f"{label} {docx} ({changes} size changes, {sign}{args.amount} pt)")
            else:
                print(f"[OK]      {docx}")

        except Exception as e:
            print(f"[ERROR]   {docx}: {e}")

    print("\nSummary")
    print(f"Files scanned: {len(files)}")
    print(f"Files changed: {changed_files}")
    print(f"Total size changes: {total_changes}")
    print(f"Adjustment: {args.amount:+g} pt")


if __name__ == "__main__":
    main()
