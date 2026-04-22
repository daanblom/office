#!/usr/bin/env python3
"""
Clean a .pptx by removing:
- unused slide layouts
- slide masters left with no remaining layouts

Behavior:
- Default: overwrite the original file in place
- With -b / --backup: rename original to backup_<filename> and write cleaned file
  at the original path

Usage:
    python cleanPowerPointMasters.py deck.pptx
    python cleanPowerPointMasters.py -b deck.pptx
"""

from __future__ import annotations

import argparse
import os
import posixpath
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional, Set


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

REL_NS = NS["pr"]

RELTYPE_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
RELTYPE_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
RELTYPE_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
RELTYPE_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

PRESENTATION_XML = "ppt/presentation.xml"
PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels"
CONTENT_TYPES_XML = "[Content_Types].xml"


@dataclass(frozen=True)
class Relationship:
    r_id: str
    rel_type: str
    target: str
    target_mode: Optional[str] = None


def norm_partname(path: str) -> str:
    path = path.replace("\\", "/")
    if path.startswith("/"):
        path = path[1:]
    return posixpath.normpath(path)


def rels_path_for_part(partname: str) -> str:
    folder = posixpath.dirname(partname)
    filename = posixpath.basename(partname)
    return posixpath.join(folder, "_rels", f"{filename}.rels")


def resolve_target(source_part: str, target: str) -> str:
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return norm_partname(target)
    base_dir = posixpath.dirname(source_part)
    return norm_partname(posixpath.join(base_dir, target))


def parse_xml(blob: bytes) -> ET.Element:
    return ET.fromstring(blob)


def tostring_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def read_relationships(zf: zipfile.ZipFile, rels_partname: str) -> Dict[str, Relationship]:
    if rels_partname not in zf.namelist():
        return {}

    root = parse_xml(zf.read(rels_partname))
    rels: Dict[str, Relationship] = {}

    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        r_id = rel.attrib["Id"]
        rels[r_id] = Relationship(
            r_id=r_id,
            rel_type=rel.attrib["Type"],
            target=rel.attrib["Target"],
            target_mode=rel.attrib.get("TargetMode"),
        )
    return rels


def remove_relationships_by_id(root: ET.Element, rel_ids: Set[str]) -> None:
    for rel in list(root.findall(f"{{{REL_NS}}}Relationship")):
        if rel.attrib.get("Id") in rel_ids:
            root.remove(rel)


def remove_content_type_overrides(root: ET.Element, partnames_to_remove: Set[str]) -> None:
    absolute_targets = {f"/{norm_partname(p)}" for p in partnames_to_remove}
    for override in list(root.findall(f"{{{NS['ct']}}}Override")):
        if override.attrib.get("PartName") in absolute_targets:
            root.remove(override)


def validate_pptx(zf: zipfile.ZipFile, input_path: Path) -> None:
    names = set(zf.namelist())

    if CONTENT_TYPES_XML not in names:
        raise ValueError(
            f"{input_path} is not a valid Office Open XML file: missing {CONTENT_TYPES_XML}"
        )

    if PRESENTATION_XML not in names:
        sample = sorted(list(names))[:20]
        raise ValueError(
            f"{input_path} is not a valid .pptx presentation: missing {PRESENTATION_XML}.\n"
            f"First archive entries: {sample}"
        )


def find_slide_parts(zf: zipfile.ZipFile) -> Set[str]:
    pres_rels = read_relationships(zf, PRESENTATION_RELS)
    return {
        resolve_target(PRESENTATION_XML, rel.target)
        for rel in pres_rels.values()
        if rel.rel_type == RELTYPE_SLIDE and rel.target_mode is None
    }


def find_used_layouts(zf: zipfile.ZipFile) -> Set[str]:
    used_layouts: Set[str] = set()

    for slide_part in find_slide_parts(zf):
        slide_rels = read_relationships(zf, rels_path_for_part(slide_part))
        for rel in slide_rels.values():
            if rel.rel_type == RELTYPE_SLIDE_LAYOUT and rel.target_mode is None:
                used_layouts.add(resolve_target(slide_part, rel.target))

    return used_layouts


def find_presentation_master_relationships(zf: zipfile.ZipFile) -> Dict[str, str]:
    pres_rels = read_relationships(zf, PRESENTATION_RELS)
    return {
        r_id: resolve_target(PRESENTATION_XML, rel.target)
        for r_id, rel in pres_rels.items()
        if rel.rel_type == RELTYPE_SLIDE_MASTER and rel.target_mode is None
    }


def find_master_layout_relationships(zf: zipfile.ZipFile, master_part: str) -> Dict[str, str]:
    master_rels = read_relationships(zf, rels_path_for_part(master_part))
    return {
        r_id: resolve_target(master_part, rel.target)
        for r_id, rel in master_rels.items()
        if rel.rel_type == RELTYPE_SLIDE_LAYOUT and rel.target_mode is None
    }


def find_layout_master(zf: zipfile.ZipFile, layout_part: str) -> Optional[str]:
    layout_rels = read_relationships(zf, rels_path_for_part(layout_part))
    for rel in layout_rels.values():
        if rel.rel_type == RELTYPE_SLIDE_MASTER and rel.target_mode is None:
            return resolve_target(layout_part, rel.target)
    return None


def find_theme_used_by_master(zf: zipfile.ZipFile, master_part: str) -> Optional[str]:
    master_rels = read_relationships(zf, rels_path_for_part(master_part))
    for rel in master_rels.values():
        if rel.rel_type == RELTYPE_THEME and rel.target_mode is None:
            return resolve_target(master_part, rel.target)
    return None


def collect_deletions(zf: zipfile.ZipFile) -> tuple[Set[str], Set[str], Set[str]]:
    """
    Returns:
        parts_to_delete,
        deleted_layout_parts,
        deleted_master_parts
    """
    used_layouts = find_used_layouts(zf)
    pres_master_rels = find_presentation_master_relationships(zf)
    all_masters = set(pres_master_rels.values())

    all_layouts: Set[str] = set()
    master_to_layouts: Dict[str, Set[str]] = {}

    for master_part in all_masters:
        rels = find_master_layout_relationships(zf, master_part)
        layouts = set(rels.values())
        master_to_layouts[master_part] = layouts
        all_layouts |= layouts

    deleted_layouts = all_layouts - used_layouts

    remaining_layouts_by_master: Dict[str, Set[str]] = {}
    for master_part, layouts in master_to_layouts.items():
        remaining_layouts_by_master[master_part] = layouts - deleted_layouts

    deleted_masters = {
        master_part
        for master_part, remaining_layouts in remaining_layouts_by_master.items()
        if not remaining_layouts
    }

    parts_to_delete: Set[str] = set()

    # Delete unused layouts and their rels
    for layout_part in deleted_layouts:
        parts_to_delete.add(layout_part)
        layout_rels_part = rels_path_for_part(layout_part)
        if layout_rels_part in zf.namelist():
            parts_to_delete.add(layout_rels_part)

    # Delete masters with no remaining layouts and their rels
    for master_part in deleted_masters:
        parts_to_delete.add(master_part)
        master_rels_part = rels_path_for_part(master_part)
        if master_rels_part in zf.namelist():
            parts_to_delete.add(master_rels_part)

    # Delete themes no remaining master uses
    remaining_masters = all_masters - deleted_masters
    remaining_themes = {
        theme
        for master in remaining_masters
        if (theme := find_theme_used_by_master(zf, master)) is not None
    }
    removed_master_themes = {
        theme
        for master in deleted_masters
        if (theme := find_theme_used_by_master(zf, master)) is not None
    }
    unused_themes = removed_master_themes - remaining_themes

    for theme_part in unused_themes:
        parts_to_delete.add(theme_part)
        theme_rels_part = rels_path_for_part(theme_part)
        if theme_rels_part in zf.namelist():
            parts_to_delete.add(theme_rels_part)

    return parts_to_delete, deleted_layouts, deleted_masters


def rewrite_presentation_xml(zf: zipfile.ZipFile, deleted_master_parts: Set[str]) -> bytes:
    root = parse_xml(zf.read(PRESENTATION_XML))

    pres_master_rels = find_presentation_master_relationships(zf)
    rel_ids_to_remove = {
        r_id
        for r_id, master_part in pres_master_rels.items()
        if master_part in deleted_master_parts
    }

    sld_master_id_lst = root.find("p:sldMasterIdLst", NS)
    if sld_master_id_lst is not None:
        for child in list(sld_master_id_lst):
            rid = child.attrib.get(f"{{{NS['r']}}}id")
            if rid in rel_ids_to_remove:
                sld_master_id_lst.remove(child)

    return tostring_xml(root)


def rewrite_presentation_rels_xml(zf: zipfile.ZipFile, deleted_master_parts: Set[str]) -> bytes:
    root = parse_xml(zf.read(PRESENTATION_RELS))

    pres_master_rels = find_presentation_master_relationships(zf)
    rel_ids_to_remove = {
        r_id
        for r_id, master_part in pres_master_rels.items()
        if master_part in deleted_master_parts
    }

    remove_relationships_by_id(root, rel_ids_to_remove)
    return tostring_xml(root)


def rewrite_master_xml(zf: zipfile.ZipFile, master_part: str, deleted_layout_parts: Set[str]) -> bytes:
    root = parse_xml(zf.read(master_part))

    master_layout_rels = find_master_layout_relationships(zf, master_part)
    rel_ids_to_remove = {
        r_id
        for r_id, layout_part in master_layout_rels.items()
        if layout_part in deleted_layout_parts
    }

    sld_layout_id_lst = root.find("p:sldLayoutIdLst", NS)
    if sld_layout_id_lst is not None:
        for child in list(sld_layout_id_lst):
            rid = child.attrib.get(f"{{{NS['r']}}}id")
            if rid in rel_ids_to_remove:
                sld_layout_id_lst.remove(child)

    return tostring_xml(root)


def rewrite_master_rels_xml(zf: zipfile.ZipFile, master_part: str, deleted_layout_parts: Set[str]) -> bytes:
    rels_part = rels_path_for_part(master_part)
    if rels_part not in zf.namelist():
        return b""

    root = parse_xml(zf.read(rels_part))

    master_layout_rels = find_master_layout_relationships(zf, master_part)
    rel_ids_to_remove = {
        r_id
        for r_id, layout_part in master_layout_rels.items()
        if layout_part in deleted_layout_parts
    }

    remove_relationships_by_id(root, rel_ids_to_remove)
    return tostring_xml(root)


def rewrite_content_types_xml(zf: zipfile.ZipFile, parts_to_delete: Set[str]) -> bytes:
    root = parse_xml(zf.read(CONTENT_TYPES_XML))
    remove_content_type_overrides(root, parts_to_delete)
    return tostring_xml(root)


def clean_pptx_to_file(input_path: Path, output_path: Path) -> None:
    with zipfile.ZipFile(input_path, "r") as zin:
        validate_pptx(zin, input_path)

        parts_to_delete, deleted_layouts, deleted_masters = collect_deletions(zin)
        all_masters = set(find_presentation_master_relationships(zin).values())
        surviving_masters = all_masters - deleted_masters

        rewritten_parts: Dict[str, bytes] = {
            PRESENTATION_XML: rewrite_presentation_xml(zin, deleted_masters),
            PRESENTATION_RELS: rewrite_presentation_rels_xml(zin, deleted_masters),
            CONTENT_TYPES_XML: rewrite_content_types_xml(zin, parts_to_delete),
        }

        for master_part in surviving_masters:
            rewritten_parts[master_part] = rewrite_master_xml(zin, master_part, deleted_layouts)

            master_rels_part = rels_path_for_part(master_part)
            if master_rels_part in zin.namelist():
                rewritten_parts[master_rels_part] = rewrite_master_rels_xml(
                    zin, master_part, deleted_layouts
                )

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if norm_partname(name) in parts_to_delete:
                    continue

                if name in rewritten_parts:
                    zout.writestr(name, rewritten_parts[name])
                else:
                    zout.writestr(item, zin.read(name))


def inplace_clean(input_path: Path, make_backup: bool) -> None:
    input_path = input_path.resolve()
    parent = input_path.parent

    fd, temp_name = tempfile.mkstemp(prefix=".pptx_clean_", suffix=".pptx", dir=parent)
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        clean_pptx_to_file(input_path, temp_path)

        if make_backup:
            backup_path = input_path.with_name(f"backup_{input_path.name}")
            if backup_path.exists():
                raise FileExistsError(f"Backup file already exists: {backup_path}")

            input_path.rename(backup_path)
            temp_path.replace(input_path)
            print(f"Backup written to: {backup_path}")
            print(f"Cleaned presentation written to: {input_path}")
        else:
            temp_path.replace(input_path)
            print(f"Overwrote original presentation: {input_path}")

    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Remove unused slide layouts and orphaned slide masters from a .pptx file."
    )
    parser.add_argument(
        "-b",
        "--backup",
        action="store_true",
        help="Rename original to backup_<filename> and write cleaned file at original path.",
    )
    parser.add_argument("pptx", help="Path to the .pptx file to clean.")
    args = parser.parse_args(argv[1:])

    input_path = Path(args.pptx)

    if not input_path.exists():
        print(f"Error: file not found: {input_path}", file=sys.stderr)
        return 1
    if not input_path.is_file():
        print(f"Error: not a file: {input_path}", file=sys.stderr)
        return 1

    try:
        inplace_clean(input_path, args.backup)
        return 0
    except zipfile.BadZipFile:
        print(
            f"Error: {input_path} is not a ZIP-based Office file, so it is not a valid .pptx.",
            file=sys.stderr,
        )
        return 2
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
