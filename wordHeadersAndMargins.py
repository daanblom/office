#!/usr/bin/env python3
"""
Check and optionally update margins and header distance in .docx files.

Usage examples:
    python docx_layout_check.py /path/to/file.docx
    python docx_layout_check.py /path/to/folder -r
    python docx_layout_check.py /path/to/file.docx -m 2.5 2.5 2.5 2.5
    python docx_layout_check.py /path/to/file.docx -m 1.2 x x 1.5
    python docx_layout_check.py /path/to/file.docx -h 1.25
    python docx_layout_check.py /path/to/folder -r -m 2.5 x 3.0 2.5 -h 1.25 -d

Units:
    -m and -h values are in centimeters.

Behavior:
    - If neither -m nor -h is provided, the script only reports values.
    - -m sets margins in Word order: top bottom left right
    - Use 'x' in -m to keep the existing value for that side unchanged
    - -h sets header distance for every section
    - -d performs a dry run and shows what would change without saving
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Optional, Tuple

from docx import Document
from docx.shared import Cm


def emu_to_cm_str(value) -> str:
    """Format a python-docx length value as cm with 2 decimals."""
    if value is None:
        return "n/a"
    try:
        return f"{value.cm:.2f} cm"
    except Exception:
        return "n/a"


def collect_docx_files(path: Path, recursive: bool) -> List[Path]:
    """Return .docx files from a file or directory, excluding temporary Word lock files."""
    if not path.exists():
        raise FileNotFoundError(f"Path does not exist: {path}")

    if path.is_file():
        if path.suffix.lower() == ".docx" and not path.name.startswith("~$"):
            return [path]
        return []

    pattern = "**/*.docx" if recursive else "*.docx"
    files = [p for p in path.glob(pattern) if p.is_file() and not p.name.startswith("~$")]
    return sorted(files)


def parse_margin_value(value: str) -> Optional[float]:
    """
    Parse one margin argument.
    Returns:
        float for cm value
        None for 'x' (leave unchanged)
    """
    if value.lower() == "x":
        return None

    try:
        return float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Invalid margin value '{value}'. Use a number in cm or 'x'."
        ) from exc


def describe_section(section, index: int) -> str:
    """Return a readable summary of one section's layout settings."""
    return (
        f"  Section {index}:\n"
        f"    top_margin    = {emu_to_cm_str(section.top_margin)}\n"
        f"    bottom_margin = {emu_to_cm_str(section.bottom_margin)}\n"
        f"    left_margin   = {emu_to_cm_str(section.left_margin)}\n"
        f"    right_margin  = {emu_to_cm_str(section.right_margin)}\n"
        f"    header_dist   = {emu_to_cm_str(section.header_distance)}"
    )


def apply_changes(
    document: Document,
    margin_values: Optional[List[Optional[float]]],
    header_cm: Optional[float],
) -> bool:
    """
    Apply requested changes to all sections.
    Returns True if any value was changed.
    """
    changed = False

    for section in document.sections:
        if margin_values is not None:
            top_cm, bottom_cm, left_cm, right_cm = margin_values

            if top_cm is not None:
                new_top = Cm(top_cm)
                if section.top_margin != new_top:
                    section.top_margin = new_top
                    changed = True

            if bottom_cm is not None:
                new_bottom = Cm(bottom_cm)
                if section.bottom_margin != new_bottom:
                    section.bottom_margin = new_bottom
                    changed = True

            if left_cm is not None:
                new_left = Cm(left_cm)
                if section.left_margin != new_left:
                    section.left_margin = new_left
                    changed = True

            if right_cm is not None:
                new_right = Cm(right_cm)
                if section.right_margin != new_right:
                    section.right_margin = new_right
                    changed = True

        if header_cm is not None:
            new_header = Cm(header_cm)
            if section.header_distance != new_header:
                section.header_distance = new_header
                changed = True

    return changed


def format_requested_margin_changes(margin_values: Optional[List[Optional[float]]]) -> str:
    """Format requested margin changes for display."""
    if margin_values is None:
        return "none"

    labels = ["top", "bottom", "left", "right"]
    parts = []

    for label, value in zip(labels, margin_values):
        if value is None:
            parts.append(f"{label}=unchanged")
        else:
            parts.append(f"{label}={value:.2f} cm")

    return ", ".join(parts)


def process_file(
    file_path: Path,
    margin_values: Optional[List[Optional[float]]],
    header_cm: Optional[float],
    dry_run: bool,
) -> Tuple[bool, bool]:
    """
    Process one file.
    Returns:
        (success, changed)
    """
    try:
        doc = Document(str(file_path))
    except Exception as exc:
        print(f"[ERROR] {file_path}: could not open document: {exc}", file=sys.stderr)
        return False, False

    print(f"\nFILE: {file_path}")

    for i, section in enumerate(doc.sections, start=1):
        print(describe_section(section, i))

    report_only = margin_values is None and header_cm is None
    if report_only:
        return True, False

    print("  Requested changes:")
    if margin_values is not None:
        print(f"    margins     = {format_requested_margin_changes(margin_values)}")
    if header_cm is not None:
        print(f"    header_dist = {header_cm:.2f} cm")

    changed = apply_changes(doc, margin_values, header_cm)

    if not changed:
        print("  No changes needed.")
        return True, False

    if dry_run:
        print("  Dry run: changes detected, file not saved.")
        return True, True

    try:
        doc.save(str(file_path))
        print("  Changes saved.")
        return True, True
    except Exception as exc:
        print(f"[ERROR] {file_path}: could not save document: {exc}", file=sys.stderr)
        return False, False


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx_layout_check.py",
        description="Check and optionally update margins and header distance in Word .docx files.",
        add_help=False,  # keep -h available for header size as requested
        formatter_class=argparse.RawTextHelpFormatter,
    )

    parser.add_argument(
        "path",
        help="Path to a .docx file or a directory containing .docx files.",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Recurse into subdirectories.",
    )
    parser.add_argument(
        "-d",
        "--dry-run",
        action="store_true",
        help="Show what would change without saving.",
    )
    parser.add_argument(
        "-h",
        "--header-size",
        type=float,
        metavar="CM",
        help="Set header distance for all sections, in centimeters.",
    )
    parser.add_argument(
        "-m",
        "--margin-size",
        nargs=4,
        type=parse_margin_value,
        metavar=("TOP", "BOTTOM", "LEFT", "RIGHT"),
        help=(
            "Set margins in Word order: top bottom left right.\n"
            "Use values in centimeters, or 'x' to leave a side unchanged.\n"
            "Example: -m 1.2 x x 1.5"
        ),
    )
    parser.add_argument(
        "--help",
        action="help",
        help="Show this help message and exit.",
    )

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.path)

    try:
        files = collect_docx_files(input_path, args.recursive)
    except FileNotFoundError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    if not files:
        print("No .docx files found.")
        return 1

    total = 0
    ok = 0
    changed = 0

    for file_path in files:
        total += 1
        success, was_changed = process_file(
            file_path=file_path,
            margin_values=args.margin_size,
            header_cm=args.header_size,
            dry_run=args.dry_run,
        )
        if success:
            ok += 1
        if was_changed:
            changed += 1

    print("\nSUMMARY")
    print(f"  Files found    : {total}")
    print(f"  Files processed: {ok}")
    print(f"  Files changed  : {changed}")

    return 0 if ok == total else 1


if __name__ == "__main__":
    raise SystemExit(main())
