#!/usr/bin/env python3
"""
Sync custom table styles from a source Word document into one or more target
Word documents, with optional removal of attached template references.

Usage:
    python sync_table_styles_and_strip_template.py source.docx target.docx
    python sync_table_styles_and_strip_template.py source.docx ./folder
    python sync_table_styles_and_strip_template.py -b --strip-template source.docx ./folder

Behavior:
- Reads all custom table styles from the source document.
- Removes all custom table styles from each target document.
- Inserts the source custom table styles into each target document.
- Optionally removes attached template references from:
    - word/settings.xml
    - word/_rels/settings.xml.rels
- Overwrites target files in place unless --backup is used.

Notes:
- Only custom table styles are replaced 1:1.
- Non-table styles are untouched.
- Built-in table styles are untouched.
- Supports .docx and .docm.
"""

from __future__ import annotations

import argparse
import copy
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import List, Tuple
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


def w_tag(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def r_tag(tag: str) -> str:
    return f"{{{R_NS}}}{tag}"


def pkg_rel_tag(tag: str) -> str:
    return f"{{{PKG_REL_NS}}}{tag}"


def is_word_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in {".docx", ".docm"}


def collect_target_files(target: Path) -> List[Path]:
    if target.is_file():
        if not is_word_file(target):
            raise ValueError(f"Target file is not a .docx or .docm file: {target}")
        return [target]

    if target.is_dir():
        files = [p for p in target.rglob("*") if is_word_file(p)]
        if not files:
            raise ValueError(f"No .docx or .docm files found in folder: {target}")
        return sorted(files)

    raise ValueError(f"Target does not exist: {target}")


def read_styles_xml_from_docx(doc_path: Path) -> bytes:
    try:
        with zipfile.ZipFile(doc_path, "r") as zf:
            return zf.read("word/styles.xml")
    except KeyError:
        raise ValueError(f"'word/styles.xml' not found in {doc_path}")
    except zipfile.BadZipFile:
        raise ValueError(f"Not a valid Word/Office zip file: {doc_path}")


def extract_custom_table_styles(styles_xml: bytes) -> List[ET.Element]:
    root = ET.fromstring(styles_xml)
    styles = []

    for style in root.findall(w_tag("style")):
        style_type = style.attrib.get(w_tag("type"))
        custom_style = style.attrib.get(w_tag("customStyle"))
        if style_type == "table" and custom_style in {"1", "true", "on"}:
            styles.append(copy.deepcopy(style))

    return styles


def remove_custom_table_styles(root: ET.Element) -> int:
    removed = 0
    to_remove = []

    for style in root.findall(w_tag("style")):
        style_type = style.attrib.get(w_tag("type"))
        custom_style = style.attrib.get(w_tag("customStyle"))
        if style_type == "table" and custom_style in {"1", "true", "on"}:
            to_remove.append(style)

    for style in to_remove:
        root.remove(style)
        removed += 1

    return removed


def insert_styles(root: ET.Element, styles_to_insert: List[ET.Element]) -> None:
    for style in styles_to_insert:
        root.append(copy.deepcopy(style))


def strip_attached_template(
    settings_xml: bytes | None,
    settings_rels_xml: bytes | None,
) -> Tuple[bytes | None, bytes | None, bool]:
    """
    Remove attached template references from:
    - word/settings.xml           -> <w:attachedTemplate r:id="..."/>
    - word/_rels/settings.xml.rels -> matching relationship(s), and any attachedTemplate rels

    Returns:
        (new_settings_xml, new_settings_rels_xml, changed)
    """
    changed = False
    rel_ids_to_remove: set[str] = set()

    # Update word/settings.xml
    new_settings_xml = settings_xml
    if settings_xml is not None:
        try:
            settings_root = ET.fromstring(settings_xml)
            attached_nodes = settings_root.findall(w_tag("attachedTemplate"))
            for node in attached_nodes:
                rel_id = node.attrib.get(r_tag("id"))
                if rel_id:
                    rel_ids_to_remove.add(rel_id)
                settings_root.remove(node)
                changed = True

            if changed:
                new_settings_xml = ET.tostring(
                    settings_root,
                    encoding="utf-8",
                    xml_declaration=True,
                )
        except ET.ParseError:
            raise ValueError("Failed to parse word/settings.xml")

    # Update word/_rels/settings.xml.rels
    new_settings_rels_xml = settings_rels_xml
    if settings_rels_xml is not None:
        try:
            rels_root = ET.fromstring(settings_rels_xml)
            to_remove = []

            for rel in rels_root.findall(pkg_rel_tag("Relationship")):
                rel_id = rel.attrib.get("Id")
                rel_type = rel.attrib.get("Type", "")
                target = rel.attrib.get("Target", "")

                if rel_id in rel_ids_to_remove:
                    to_remove.append(rel)
                    continue

                if rel_type.endswith("/attachedTemplate"):
                    to_remove.append(rel)
                    continue

                # Defensive extra rule: if target points at a template-like file.
                target_lower = target.lower()
                if target_lower.endswith(".dotx") or target_lower.endswith(".dotm"):
                    to_remove.append(rel)
                    continue

            if to_remove:
                for rel in to_remove:
                    rels_root.remove(rel)
                new_settings_rels_xml = ET.tostring(
                    rels_root,
                    encoding="utf-8",
                    xml_declaration=True,
                )
                changed = True
        except ET.ParseError:
            raise ValueError("Failed to parse word/_rels/settings.xml.rels")

    return new_settings_xml, new_settings_rels_xml, changed


def replace_custom_table_styles_in_doc(
    source_styles: List[ET.Element],
    target_path: Path,
    make_backup: bool = False,
    strip_template: bool = False,
) -> Tuple[int, int, bool]:
    """
    Returns:
        (removed_style_count, inserted_style_count, template_stripped)
    """
    if make_backup:
        backup_path = target_path.with_suffix(target_path.suffix + ".bak")
        shutil.copy2(target_path, backup_path)

    with zipfile.ZipFile(target_path, "r") as zin:
        try:
            styles_xml = zin.read("word/styles.xml")
        except KeyError:
            raise ValueError(f"'word/styles.xml' not found in {target_path}")

        styles_root = ET.fromstring(styles_xml)
        removed_count = remove_custom_table_styles(styles_root)
        insert_styles(styles_root, source_styles)
        new_styles_xml = ET.tostring(
            styles_root,
            encoding="utf-8",
            xml_declaration=True,
        )

        settings_xml = None
        settings_rels_xml = None

        if strip_template:
            try:
                settings_xml = zin.read("word/settings.xml")
            except KeyError:
                settings_xml = None

            try:
                settings_rels_xml = zin.read("word/_rels/settings.xml.rels")
            except KeyError:
                settings_rels_xml = None

            new_settings_xml, new_settings_rels_xml, template_stripped = strip_attached_template(
                settings_xml,
                settings_rels_xml,
            )
        else:
            new_settings_xml = settings_xml
            new_settings_rels_xml = settings_rels_xml
            template_stripped = False

        fd, tmp_name = tempfile.mkstemp(suffix=target_path.suffix)
        os.close(fd)
        tmp_path = Path(tmp_name)

        try:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/styles.xml":
                        zout.writestr(item, new_styles_xml)
                    elif strip_template and item.filename == "word/settings.xml" and new_settings_xml is not None:
                        zout.writestr(item, new_settings_xml)
                    elif strip_template and item.filename == "word/_rels/settings.xml.rels" and new_settings_rels_xml is not None:
                        zout.writestr(item, new_settings_rels_xml)
                    else:
                        zout.writestr(item, zin.read(item.filename))

            shutil.move(str(tmp_path), str(target_path))
        finally:
            if tmp_path.exists():
                tmp_path.unlink(missing_ok=True)

    return removed_count, len(source_styles), template_stripped


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Copy all custom table styles from a source Word document into "
            "a target file or folder, optionally stripping attached template references."
        )
    )
    parser.add_argument(
        "source",
        help="Source Word file (.docx or .docm) containing the desired custom table styles.",
    )
    parser.add_argument(
        "target",
        help="Target Word file, or a folder containing Word files to update.",
    )
    parser.add_argument(
        "-b",
        "--backup",
        action="store_true",
        help="Create a .bak backup next to each target before overwriting it.",
    )
    parser.add_argument(
        "--strip-template",
        action="store_true",
        help="Remove attached template references from target documents.",
    )
    return parser


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()

    source = Path(args.source).expanduser().resolve()
    target = Path(args.target).expanduser().resolve()

    if not source.exists():
        print(f"Error: source does not exist: {source}", file=sys.stderr)
        return 1

    if not is_word_file(source):
        print(f"Error: source must be a .docx or .docm file: {source}", file=sys.stderr)
        return 1

    try:
        target_files = collect_target_files(target)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    try:
        source_styles_xml = read_styles_xml_from_docx(source)
        source_custom_table_styles = extract_custom_table_styles(source_styles_xml)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    print(f"Source: {source}")
    print(f"Custom table styles found in source: {len(source_custom_table_styles)}")
    print(f"Targets to update: {len(target_files)}")
    print(f"Strip attached template: {'yes' if args.strip_template else 'no'}")
    print()

    failures = 0

    for doc in target_files:
        try:
            removed, inserted, template_stripped = replace_custom_table_styles_in_doc(
                source_styles=source_custom_table_styles,
                target_path=doc,
                make_backup=args.backup,
                strip_template=args.strip_template,
            )
            print(f"[OK] {doc}")
            print(f"     removed existing custom table styles: {removed}")
            print(f"     inserted source custom table styles: {inserted}")
            print(f"     attached template stripped: {'yes' if template_stripped else 'no'}")
        except Exception as e:
            failures += 1
            print(f"[FAIL] {doc}", file=sys.stderr)
            print(f"       {e}", file=sys.stderr)

    print()
    if failures:
        print(f"Completed with {failures} failure(s).", file=sys.stderr)
        return 2

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
