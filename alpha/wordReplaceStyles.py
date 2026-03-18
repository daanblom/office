#!/usr/bin/env python3
"""
Overwrite the complete Word font/style setup in one or more .docx files
with that from a source .docx.

This script copies the source document's style-related package parts into
target document(s) so they get the same:

- default document fonts
- named styles
- heading/body style font settings
- theme fonts
- font table mappings

Managed parts:
- word/styles.xml
- word/theme/theme1.xml
- word/fontTable.xml

It also synchronizes:
- word/_rels/document.xml.rels
- [Content_Types].xml

By default, target files are overwritten in place.

Options:
- -b / --backup     Create a backup before overwriting
- -r / --recursive  Recursively process all .docx files in a directory

Examples:
    python overwrite_word_font_setup.py source.docx target.docx
    python overwrite_word_font_setup.py -b source.docx target.docx
    python overwrite_word_font_setup.py -r source.docx ./documents
    python overwrite_word_font_setup.py -r -b source.docx ./documents
"""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Iterable

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_TYPE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_TYPE_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
REL_TYPE_FONT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"

CONTENT_TYPE_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
CONTENT_TYPE_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
CONTENT_TYPE_FONT_TABLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"

STYLE_PARTS = {
    "word/styles.xml": {
        "rel_type": REL_TYPE_STYLES,
        "rel_target": "styles.xml",
        "content_type": CONTENT_TYPE_STYLES,
        "part_name": "/word/styles.xml",
    },
    "word/theme/theme1.xml": {
        "rel_type": REL_TYPE_THEME,
        "rel_target": "theme/theme1.xml",
        "content_type": CONTENT_TYPE_THEME,
        "part_name": "/word/theme/theme1.xml",
    },
    "word/fontTable.xml": {
        "rel_type": REL_TYPE_FONT_TABLE,
        "rel_target": "fontTable.xml",
        "content_type": CONTENT_TYPE_FONT_TABLE,
        "part_name": "/word/fontTable.xml",
    },
}

ET.register_namespace("", PKG_REL_NS)
ET.register_namespace("", CT_NS)


def read_docx(path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def write_docx(path: Path, files: Dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)


def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def to_xml_bytes(elem: ET.Element) -> bytes:
    return ET.tostring(elem, encoding="utf-8", xml_declaration=True)


def sync_document_relationships(
    rels_xml: bytes,
    source_files: Dict[str, bytes],
) -> bytes:
    root = parse_xml(rels_xml)
    ns = {"r": PKG_REL_NS}

    rels = root.findall("r:Relationship", ns)
    rel_types_to_manage = {cfg["rel_type"] for cfg in STYLE_PARTS.values()}

    for rel in list(rels):
        if rel.attrib.get("Type") in rel_types_to_manage:
            root.remove(rel)

    existing_ids = {
        rel.attrib.get("Id", "")
        for rel in root.findall("r:Relationship", ns)
        if rel.attrib.get("Id")
    }

    def next_rel_id() -> str:
        i = 1
        while f"rId{i}" in existing_ids:
            i += 1
        rid = f"rId{i}"
        existing_ids.add(rid)
        return rid

    for part_path, cfg in STYLE_PARTS.items():
        if part_path in source_files:
            rel = ET.Element(f"{{{PKG_REL_NS}}}Relationship")
            rel.set("Id", next_rel_id())
            rel.set("Type", cfg["rel_type"])
            rel.set("Target", cfg["rel_target"])
            root.append(rel)

    return to_xml_bytes(root)


def sync_content_types(
    content_types_xml: bytes,
    source_files: Dict[str, bytes],
) -> bytes:
    root = parse_xml(content_types_xml)
    ns = {"ct": CT_NS}

    managed_part_names = {cfg["part_name"] for cfg in STYLE_PARTS.values()}

    for override in list(root.findall("ct:Override", ns)):
        if override.attrib.get("PartName") in managed_part_names:
            root.remove(override)

    for part_path, cfg in STYLE_PARTS.items():
        if part_path in source_files:
            override = ET.Element(f"{{{CT_NS}}}Override")
            override.set("PartName", cfg["part_name"])
            override.set("ContentType", cfg["content_type"])
            root.append(override)

    return to_xml_bytes(root)


def sync_style_parts(
    source_files: Dict[str, bytes],
    target_files: Dict[str, bytes],
) -> Dict[str, bytes]:
    output_files = dict(target_files)

    for part_path in STYLE_PARTS:
        if part_path in source_files:
            output_files[part_path] = source_files[part_path]
        else:
            output_files.pop(part_path, None)

    return output_files


def remove_empty_theme_folder_entries(files: Dict[str, bytes]) -> Dict[str, bytes]:
    output = dict(files)
    if "word/theme/theme1.xml" not in output:
        for key in list(output.keys()):
            if key in ("word/theme/", "word/theme"):
                output.pop(key, None)
    return output


def build_output_files(source_files: Dict[str, bytes], target_files: Dict[str, bytes]) -> Dict[str, bytes]:
    if "word/document.xml" not in target_files:
        raise FileNotFoundError("Missing word/document.xml")
    if "word/_rels/document.xml.rels" not in target_files:
        raise FileNotFoundError("Missing word/_rels/document.xml.rels")
    if "[Content_Types].xml" not in target_files:
        raise FileNotFoundError("Missing [Content_Types].xml")

    output_files = sync_style_parts(source_files, target_files)

    output_files["word/_rels/document.xml.rels"] = sync_document_relationships(
        output_files["word/_rels/document.xml.rels"],
        source_files,
    )

    output_files["[Content_Types].xml"] = sync_content_types(
        output_files["[Content_Types].xml"],
        source_files,
    )

    output_files = remove_empty_theme_folder_entries(output_files)
    return output_files


def make_backup(path: Path) -> Path:
    backup = path.with_name(path.name + ".bak")
    counter = 1
    while backup.exists():
        backup = path.with_name(f"{path.name}.bak{counter}")
        counter += 1
    shutil.copy2(path, backup)
    return backup


def overwrite_target(source_files: Dict[str, bytes], target_path: Path, backup: bool = False) -> None:
    if backup:
        backup_path = make_backup(target_path)
        print(f"[backup] {target_path} -> {backup_path}")

    target_files = read_docx(target_path)
    output_files = build_output_files(source_files, target_files)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", dir=str(target_path.parent)) as tmp:
        tmp_path = Path(tmp.name)

    try:
        write_docx(tmp_path, output_files)
        tmp_path.replace(target_path)
        print(f"[updated] {target_path}")
    finally:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)


def find_docx_files(root: Path) -> Iterable[Path]:
    for path in root.rglob("*.docx"):
        if path.is_file():
            yield path


def validate_source(source: Path) -> None:
    if not source.exists():
        raise FileNotFoundError(f"Source file not found: {source}")
    if not source.is_file():
        raise ValueError(f"Source is not a file: {source}")
    if source.suffix.lower() != ".docx":
        raise ValueError(f"Source is not a .docx file: {source}")


def process_single(source: Path, target: Path, backup: bool) -> int:
    if not target.exists():
        print(f"[error] Target file not found: {target}", file=sys.stderr)
        return 1
    if not target.is_file():
        print(f"[error] Target is not a file: {target}", file=sys.stderr)
        return 1
    if target.suffix.lower() != ".docx":
        print(f"[error] Target is not a .docx file: {target}", file=sys.stderr)
        return 1

    if source.resolve() == target.resolve():
        print(f"[skip] Source and target are the same file: {target}")
        return 0

    source_files = read_docx(source)
    try:
        overwrite_target(source_files, target, backup=backup)
        return 0
    except Exception as exc:
        print(f"[error] {target}: {exc}", file=sys.stderr)
        return 1


def process_recursive(source: Path, target_dir: Path, backup: bool) -> int:
    if not target_dir.exists():
        print(f"[error] Target directory not found: {target_dir}", file=sys.stderr)
        return 1
    if not target_dir.is_dir():
        print(f"[error] Target is not a directory: {target_dir}", file=sys.stderr)
        return 1

    source_resolved = source.resolve()
    source_files = read_docx(source)

    failures = 0
    matched = 0

    for docx_path in find_docx_files(target_dir):
        matched += 1

        try:
            if docx_path.resolve() == source_resolved:
                print(f"[skip] Source file inside tree: {docx_path}")
                continue

            overwrite_target(source_files, docx_path, backup=backup)
        except Exception as exc:
            failures += 1
            print(f"[error] {docx_path}: {exc}", file=sys.stderr)

    if matched == 0:
        print(f"[warn] No .docx files found in: {target_dir}")

    return 1 if failures else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Overwrite Word font/style setup in-place from a source .docx."
    )
    parser.add_argument(
        "-b",
        "--backup",
        action="store_true",
        help="Create backup file(s) before overwriting",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Recursively process all .docx files in the target directory",
    )
    parser.add_argument(
        "source",
        help="Source .docx to copy font/style setup from",
    )
    parser.add_argument(
        "target",
        help="Target .docx file, or target directory when using -r",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    source = Path(args.source)
    target = Path(args.target)

    try:
        validate_source(source)
    except Exception as exc:
        print(f"[error] {exc}", file=sys.stderr)
        return 1

    if args.recursive:
        return process_recursive(source, target, backup=args.backup)

    return process_single(source, target, backup=args.backup)


if __name__ == "__main__":
    raise SystemExit(main())
