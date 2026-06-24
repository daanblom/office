#!/usr/bin/env python3
"""
extract_custom_props.py

Extracts Word custom document properties from .docx files and writes
a sidecar .csv file next to each document.

Usage:
    python extract_custom_props.py <folder> [-r]

Arguments:
    folder      Path to the folder containing .docx files.
    -r          Recurse into subfolders.

Output:
    For each file, e.g. "Report.docx", a "Report.docx.custom_props.csv"
    is written in the same directory.
"""

import argparse
import csv
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Namespace used in docProps/custom.xml
NS_VTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"

# All possible vt:* leaf types in the OOXML VTypes namespace
VTYPE_TAGS = [
    "lpwstr", "lpstr", "bstr", "bool", "i1", "i2", "i4", "i8",
    "ui1", "ui2", "ui4", "ui8", "r4", "r8", "decimal",
    "date", "filetime", "cy", "error", "empty", "null",
    "clsid", "cf",
]


def extract_custom_properties(docx_path: Path) -> list[dict]:
    """
    Returns a list of dicts with keys: name, type, value.
    Returns an empty list if no custom properties exist.
    Raises on corrupt/unreadable files.
    """
    properties = []

    with zipfile.ZipFile(docx_path, "r") as zf:
        if "docProps/custom.xml" not in zf.namelist():
            return properties  # No custom properties defined

        with zf.open("docProps/custom.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()

    for prop in root.findall(f"{{{NS_CUSTOM}}}property"):
        name = prop.get("name", "")
        prop_type = ""
        value = ""

        for vt in VTYPE_TAGS:
            el = prop.find(f"{{{NS_VTYPE}}}{vt}")
            if el is not None:
                prop_type = vt
                # Handle empty/null types (no text content)
                value = el.text.strip() if el.text else ""
                break

        properties.append({"name": name, "type": prop_type, "value": value})

    return properties


def write_csv(docx_path: Path, properties: list[dict]) -> Path:
    """
    Writes a CSV file next to the .docx file.
    Returns the path to the CSV file.
    """
    csv_path = docx_path.with_suffix(docx_path.suffix + ".custom_props.csv")

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["name", "value", "type"])
        writer.writeheader()
        writer.writerows(properties)

    return csv_path


def find_docx_files(folder: Path, recursive: bool) -> list[Path]:
    pattern = "**/*.docx" if recursive else "*.docx"
    return sorted(folder.glob(pattern))


def main():
    parser = argparse.ArgumentParser(
        description="Extract Word custom document properties to sidecar CSV files."
    )
    parser.add_argument(
        "folder",
        type=Path,
        help="Folder containing .docx files to process.",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Recurse into subfolders.",
    )
    args = parser.parse_args()

    folder = args.folder.resolve()
    if not folder.is_dir():
        print(f"ERROR: '{folder}' is not a valid directory.", file=sys.stderr)
        sys.exit(1)

    files = find_docx_files(folder, args.recursive)

    if not files:
        print("No .docx files found.")
        return

    print(f"Found {len(files)} .docx file(s).\n")

    processed = 0
    skipped = 0
    errors = 0

    for docx_path in files:
        try:
            props = extract_custom_properties(docx_path)

            if not props:
                print(f"  [SKIP]  {docx_path} — no custom properties found")
                skipped += 1
                continue

            csv_path = write_csv(docx_path, props)
            print(f"  [OK]    {docx_path} → {csv_path.name} ({len(props)} propert{'y' if len(props) == 1 else 'ies'})")
            processed += 1

        except zipfile.BadZipFile:
            print(f"  [ERROR] {docx_path} — not a valid .docx (bad ZIP)", file=sys.stderr)
            errors += 1
        except Exception as e:
            print(f"  [ERROR] {docx_path} — {e}", file=sys.stderr)
            errors += 1

    print(f"\nDone. Processed: {processed} | Skipped (no props): {skipped} | Errors: {errors}")


if __name__ == "__main__":
    main()
