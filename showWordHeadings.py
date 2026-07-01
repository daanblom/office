#!/usr/bin/env python3

import argparse
import re
import sys
from pathlib import Path

from docx import Document


def get_heading_level(style_name: str):
    """
    Return heading level as int if style name looks like '#_Heading 1',
    otherwise return None.
    """
    if not style_name:
        return None

    match = re.fullmatch(r"#_Heading\s+([1-9])", style_name.strip(), re.IGNORECASE)
    if match:
        return int(match.group(1))

    return None


def extract_headings(docx_path: Path, selected_level: int):
    """
    Yield tuples of (heading_level, heading_text) for headings
    matching exactly the selected level.
    """
    document = Document(docx_path)

    for paragraph in document.paragraphs:
        style_name = paragraph.style.name if paragraph.style else ""
        heading_level = get_heading_level(style_name)

        if heading_level is None:
            continue

        if heading_level == selected_level:
            text = paragraph.text.strip()
            if text:
                yield heading_level, text


def main():
    parser = argparse.ArgumentParser(
        description="Show headings from a Word .docx file for a specific heading level."
    )
    parser.add_argument(
        "-l",
        "--list-heading",
        dest="heading_level",
        type=int,
        required=True,
        choices=range(1, 10),
        metavar="1-9",
        help="Heading level to show exactly (e.g. 4 shows only Heading 4).",
    )
    parser.add_argument(
        "file",
        help="Path to the Word .docx file."
    )

    args = parser.parse_args()

    docx_path = Path(args.file)

    if not docx_path.exists():
        print(f"Error: file does not exist: {docx_path}", file=sys.stderr)
        sys.exit(1)

    if docx_path.suffix.lower() != ".docx":
        print(
            f"Error: only .docx files are supported, got: {docx_path.suffix}",
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        found_any = False
        for level, text in extract_headings(docx_path, args.heading_level):
            print(f"H{level}: {text}")
            found_any = True

        if not found_any:
            print("No headings found for the selected level.")
    except Exception as exc:
        print(f"Error reading document: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
