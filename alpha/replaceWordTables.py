#!/usr/bin/env python3
"""
Deep Word table-style diagnostic and sync tool.

This script syncs custom table styles from a clean source Word document into
target Word documents, and can also inspect/clean related XML parts that may
affect how styles reappear in downstream systems.

Supported targets:
- .docx
- .docm

Main features:
- Replace all custom table styles in word/styles.xml with those from source
- Optionally do the same for word/stylesWithEffects.xml
- Optionally strip attached template references
- Diagnostic reporting for unexpected styles such as CA*
- Dry-run mode
- Backup mode

Examples:
    python replaceWordTables.py source.docx target.docx
    python replaceWordTables.py -b source.docx ./folder
    python replaceWordTables.py --dry-run --report-unexpected --unexpected-prefix CA source.docx target.docx
    python replaceWordTables.py --dry-run --check-effects --deep-report source.docx target.docx
    python replaceWordTables.py -b --sync-effects --strip-template source.docx ./folder
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
from typing import Dict, List, Optional, Set, Tuple
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


def read_zip_member(zf: zipfile.ZipFile, name: str) -> Optional[bytes]:
    try:
        return zf.read(name)
    except KeyError:
        return None


def read_styles_xml_from_docx(doc_path: Path, part_name: str = "word/styles.xml") -> bytes:
    try:
        with zipfile.ZipFile(doc_path, "r") as zf:
            data = zf.read(part_name)
            return data
    except KeyError:
        raise ValueError(f"'{part_name}' not found in {doc_path}")
    except zipfile.BadZipFile:
        raise ValueError(f"Not a valid Word/Office zip file: {doc_path}")


def extract_custom_table_styles(styles_xml: bytes) -> List[ET.Element]:
    root = ET.fromstring(styles_xml)
    styles: List[ET.Element] = []

    for style in root.findall(w_tag("style")):
        style_type = style.attrib.get(w_tag("type"))
        custom_style = style.attrib.get(w_tag("customStyle"))
        if style_type == "table" and custom_style in {"1", "true", "on"}:
            styles.append(copy.deepcopy(style))

    return styles


def classify_table_styles(root: ET.Element) -> Tuple[List[str], List[str]]:
    """
    Returns:
        (all_table_style_ids, custom_table_style_ids)
    """
    all_names: List[str] = []
    custom_names: List[str] = []

    for style in root.findall(w_tag("style")):
        style_type = style.attrib.get(w_tag("type"))
        if style_type != "table":
            continue

        style_id = style.attrib.get(w_tag("styleId"), "UNKNOWN")
        all_names.append(style_id)

        custom_style = style.attrib.get(w_tag("customStyle"))
        if custom_style in {"1", "true", "on"}:
            custom_names.append(style_id)

    return sorted(all_names), sorted(custom_names)


def remove_custom_table_styles(root: ET.Element) -> Tuple[int, List[str]]:
    removed = 0
    removed_names: List[str] = []
    to_remove: List[ET.Element] = []

    for style in root.findall(w_tag("style")):
        style_type = style.attrib.get(w_tag("type"))
        custom_style = style.attrib.get(w_tag("customStyle"))

        if style_type == "table" and custom_style in {"1", "true", "on"}:
            style_id = style.attrib.get(w_tag("styleId"), "UNKNOWN")
            removed_names.append(style_id)
            to_remove.append(style)

    for style in to_remove:
        root.remove(style)
        removed += 1

    return removed, sorted(removed_names)


def insert_styles(root: ET.Element, styles_to_insert: List[ET.Element]) -> None:
    for style in styles_to_insert:
        root.append(copy.deepcopy(style))


def transform_style_part(
    xml_bytes: bytes,
    source_styles: List[ET.Element],
) -> Dict[str, object]:
    """
    Replace all custom table styles in a style part with source styles.
    """
    root = ET.fromstring(xml_bytes)

    before_all, before_custom = classify_table_styles(root)
    removed_count, removed_names = remove_custom_table_styles(root)
    insert_styles(root, source_styles)
    after_all, after_custom = classify_table_styles(root)

    new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    inserted_names = sorted(
        [s.attrib.get(w_tag("styleId"), "UNKNOWN") for s in source_styles]
    )

    return {
        "before_all": before_all,
        "before_custom": before_custom,
        "removed_count": removed_count,
        "removed_names": removed_names,
        "inserted_count": len(source_styles),
        "inserted_names": inserted_names,
        "after_all": after_all,
        "after_custom": after_custom,
        "new_xml": new_xml,
    }


def strip_attached_template(
    settings_xml: Optional[bytes],
    settings_rels_xml: Optional[bytes],
) -> Tuple[Optional[bytes], Optional[bytes], bool, List[str]]:
    """
    Remove attached template references from:
    - word/settings.xml
    - word/_rels/settings.xml.rels

    Returns:
        (new_settings_xml, new_settings_rels_xml, changed, notes)
    """
    changed = False
    notes: List[str] = []
    rel_ids_to_remove: Set[str] = set()

    new_settings_xml = settings_xml
    if settings_xml is not None:
        try:
            settings_root = ET.fromstring(settings_xml)
            attached_nodes = settings_root.findall(w_tag("attachedTemplate"))
            local_changed = False

            for node in attached_nodes:
                rel_id = node.attrib.get(r_tag("id"))
                if rel_id:
                    rel_ids_to_remove.add(rel_id)
                    notes.append(f"settings.xml attachedTemplate r:id={rel_id}")
                else:
                    notes.append("settings.xml attachedTemplate without relationship id")
                settings_root.remove(node)
                local_changed = True

            if local_changed:
                new_settings_xml = ET.tostring(
                    settings_root,
                    encoding="utf-8",
                    xml_declaration=True,
                )
                changed = True
        except ET.ParseError:
            raise ValueError("Failed to parse word/settings.xml")

    new_settings_rels_xml = settings_rels_xml
    if settings_rels_xml is not None:
        try:
            rels_root = ET.fromstring(settings_rels_xml)
            to_remove: List[ET.Element] = []

            for rel in rels_root.findall(pkg_rel_tag("Relationship")):
                rel_id = rel.attrib.get("Id", "")
                rel_type = rel.attrib.get("Type", "")
                target = rel.attrib.get("Target", "")

                should_remove = False

                if rel_id in rel_ids_to_remove:
                    should_remove = True

                if rel_type.endswith("/attachedTemplate"):
                    should_remove = True

                target_lower = target.lower()
                if target_lower.endswith(".dotx") or target_lower.endswith(".dotm"):
                    should_remove = True

                if should_remove:
                    notes.append(
                        f"settings.xml.rels removed Id={rel_id or 'UNKNOWN'} "
                        f"Type={rel_type or 'UNKNOWN'} Target={target or 'UNKNOWN'}"
                    )
                    to_remove.append(rel)

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

    return new_settings_xml, new_settings_rels_xml, changed, notes


def inspect_attached_template(
    settings_xml: Optional[bytes],
    settings_rels_xml: Optional[bytes],
) -> Dict[str, object]:
    """
    Report attached-template related metadata without changing anything.
    """
    found_settings_node = False
    found_relationships: List[str] = []
    found_targets: List[str] = []
    parse_errors: List[str] = []

    if settings_xml is not None:
        try:
            settings_root = ET.fromstring(settings_xml)
            attached_nodes = settings_root.findall(w_tag("attachedTemplate"))
            if attached_nodes:
                found_settings_node = True
        except ET.ParseError:
            parse_errors.append("Failed to parse word/settings.xml")

    if settings_rels_xml is not None:
        try:
            rels_root = ET.fromstring(settings_rels_xml)
            for rel in rels_root.findall(pkg_rel_tag("Relationship")):
                rel_id = rel.attrib.get("Id", "")
                rel_type = rel.attrib.get("Type", "")
                target = rel.attrib.get("Target", "")
                target_lower = target.lower()

                if rel_type.endswith("/attachedTemplate"):
                    found_relationships.append(rel_id or "UNKNOWN")
                    found_targets.append(target or "UNKNOWN")
                elif target_lower.endswith(".dotx") or target_lower.endswith(".dotm"):
                    found_relationships.append(rel_id or "UNKNOWN")
                    found_targets.append(target or "UNKNOWN")
        except ET.ParseError:
            parse_errors.append("Failed to parse word/_rels/settings.xml.rels")

    return {
        "has_attached_template_node": found_settings_node,
        "relationship_ids": sorted(found_relationships),
        "targets": sorted(found_targets),
        "parse_errors": parse_errors,
        "has_any_template_reference": bool(found_settings_node or found_relationships or found_targets),
    }


def style_matches_prefixes(style_id: str, prefixes: List[str]) -> bool:
    if not prefixes:
        return False
    style_id_lower = style_id.lower()
    return any(style_id_lower.startswith(prefix.lower()) for prefix in prefixes)


def analyze_unexpected_styles(
    all_table_styles: List[str],
    custom_table_styles: List[str],
    source_custom_style_ids: Set[str],
    unexpected_prefixes: List[str],
) -> Dict[str, List[str]]:
    prefixed_all = sorted(
        [s for s in all_table_styles if style_matches_prefixes(s, unexpected_prefixes)]
    )
    prefixed_custom = sorted(
        [s for s in custom_table_styles if style_matches_prefixes(s, unexpected_prefixes)]
    )
    custom_not_in_source = sorted(
        [s for s in custom_table_styles if s not in source_custom_style_ids]
    )
    all_not_in_source = sorted(
        [s for s in all_table_styles if s not in source_custom_style_ids]
    )

    return {
        "prefixed_all": prefixed_all,
        "prefixed_custom": prefixed_custom,
        "custom_not_in_source": custom_not_in_source,
        "all_not_in_source": all_not_in_source,
    }


def analyze_style_part_for_report(
    xml_bytes: bytes,
    source_custom_style_ids: Set[str],
    unexpected_prefixes: List[str],
) -> Dict[str, object]:
    root = ET.fromstring(xml_bytes)
    all_names, custom_names = classify_table_styles(root)

    return {
        "all_names": all_names,
        "custom_names": custom_names,
        "unexpected": analyze_unexpected_styles(
            all_table_styles=all_names,
            custom_table_styles=custom_names,
            source_custom_style_ids=source_custom_style_ids,
            unexpected_prefixes=unexpected_prefixes,
        ),
    }


def analyze_and_prepare_doc(
    source_styles: List[ET.Element],
    source_custom_style_ids: Set[str],
    target_path: Path,
    strip_template: bool = False,
    unexpected_prefixes: Optional[List[str]] = None,
    check_effects: bool = False,
    sync_effects: bool = False,
) -> Dict[str, object]:
    unexpected_prefixes = unexpected_prefixes or []

    with zipfile.ZipFile(target_path, "r") as zin:
        styles_xml = read_zip_member(zin, "word/styles.xml")
        if styles_xml is None:
            raise ValueError(f"'word/styles.xml' not found in {target_path}")

        styles_result = transform_style_part(styles_xml, source_styles)
        styles_before_unexpected = analyze_unexpected_styles(
            all_table_styles=styles_result["before_all"],
            custom_table_styles=styles_result["before_custom"],
            source_custom_style_ids=source_custom_style_ids,
            unexpected_prefixes=unexpected_prefixes,
        )
        styles_after_unexpected = analyze_unexpected_styles(
            all_table_styles=styles_result["after_all"],
            custom_table_styles=styles_result["after_custom"],
            source_custom_style_ids=source_custom_style_ids,
            unexpected_prefixes=unexpected_prefixes,
        )

        styles_with_effects_xml = read_zip_member(zin, "word/stylesWithEffects.xml")
        effects_exists = styles_with_effects_xml is not None
        effects_report = None
        effects_transform = None

        if effects_exists and (check_effects or sync_effects):
            effects_report = analyze_style_part_for_report(
                styles_with_effects_xml,
                source_custom_style_ids,
                unexpected_prefixes,
            )

        if effects_exists and sync_effects:
            effects_transform = transform_style_part(styles_with_effects_xml, source_styles)
            effects_after_unexpected = analyze_unexpected_styles(
                all_table_styles=effects_transform["after_all"],
                custom_table_styles=effects_transform["after_custom"],
                source_custom_style_ids=source_custom_style_ids,
                unexpected_prefixes=unexpected_prefixes,
            )
        else:
            effects_after_unexpected = None

        settings_xml = read_zip_member(zin, "word/settings.xml")
        settings_rels_xml = read_zip_member(zin, "word/_rels/settings.xml.rels")
        template_inspection = inspect_attached_template(settings_xml, settings_rels_xml)

        new_settings_xml = None
        new_settings_rels_xml = None
        template_stripped = False
        template_strip_notes: List[str] = []

        if strip_template:
            (
                new_settings_xml,
                new_settings_rels_xml,
                template_stripped,
                template_strip_notes,
            ) = strip_attached_template(settings_xml, settings_rels_xml)

        return {
            "styles": {
                **styles_result,
                "unexpected_before": styles_before_unexpected,
                "unexpected_after": styles_after_unexpected,
            },
            "effects_exists": effects_exists,
            "effects_report": effects_report,
            "effects_transform": effects_transform,
            "effects_after_unexpected": effects_after_unexpected,
            "template_inspection": template_inspection,
            "template_stripped": template_stripped,
            "template_strip_notes": template_strip_notes,
            "new_settings_xml": new_settings_xml,
            "new_settings_rels_xml": new_settings_rels_xml,
            "parts_present": {
                "word/styles.xml": styles_xml is not None,
                "word/stylesWithEffects.xml": styles_with_effects_xml is not None,
                "word/settings.xml": settings_xml is not None,
                "word/_rels/settings.xml.rels": settings_rels_xml is not None,
            },
        }


def write_updated_doc(
    target_path: Path,
    new_styles_xml: bytes,
    make_backup: bool = False,
    new_styles_with_effects_xml: Optional[bytes] = None,
    new_settings_xml: Optional[bytes] = None,
    new_settings_rels_xml: Optional[bytes] = None,
) -> None:
    if make_backup:
        backup_path = target_path.with_suffix(target_path.suffix + ".bak")
        shutil.copy2(target_path, backup_path)

    with zipfile.ZipFile(target_path, "r") as zin:
        fd, tmp_name = tempfile.mkstemp(suffix=target_path.suffix)
        os.close(fd)
        tmp_path = Path(tmp_name)

        try:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/styles.xml":
                        zout.writestr(item, new_styles_xml)
                    elif item.filename == "word/stylesWithEffects.xml" and new_styles_with_effects_xml is not None:
                        zout.writestr(item, new_styles_with_effects_xml)
                    elif item.filename == "word/settings.xml" and new_settings_xml is not None:
                        zout.writestr(item, new_settings_xml)
                    elif item.filename == "word/_rels/settings.xml.rels" and new_settings_rels_xml is not None:
                        zout.writestr(item, new_settings_rels_xml)
                    else:
                        zout.writestr(item, zin.read(item.filename))

            shutil.move(str(tmp_path), str(target_path))
        finally:
            if tmp_path.exists():
                tmp_path.unlink(missing_ok=True)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Sync custom table styles from a source Word document into one or more "
            "target Word documents, with optional deep inspection of style-related parts."
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
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Analyze and report changes without modifying any files.",
    )
    parser.add_argument(
        "--report-all-table-styles",
        action="store_true",
        help="Report all table style IDs before and after processing for word/styles.xml.",
    )
    parser.add_argument(
        "--report-unexpected",
        action="store_true",
        help="Report potentially unexpected table styles.",
    )
    parser.add_argument(
        "--unexpected-prefix",
        action="append",
        default=[],
        help="Prefix to flag as unexpected in diagnostics. Repeatable.",
    )
    parser.add_argument(
        "--fail-on-unexpected",
        action="store_true",
        help="Return a non-zero exit code if unexpected styles remain after processing.",
    )
    parser.add_argument(
        "--check-effects",
        action="store_true",
        help="Inspect word/stylesWithEffects.xml too, if present.",
    )
    parser.add_argument(
        "--sync-effects",
        action="store_true",
        help="Also replace custom table styles in word/stylesWithEffects.xml, if present.",
    )
    parser.add_argument(
        "--deep-report",
        action="store_true",
        help="Report which style-related parts exist and whether attached-template references are present.",
    )
    return parser


def print_list(label: str, items: List[str], indent: str = "     ") -> None:
    print(f"{indent}{label}: {', '.join(items) if items else 'none'}")


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
        source_styles_xml = read_styles_xml_from_docx(source, "word/styles.xml")
        source_custom_table_styles = extract_custom_table_styles(source_styles_xml)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    source_style_names = sorted(
        [s.attrib.get(w_tag("styleId"), "UNKNOWN") for s in source_custom_table_styles]
    )
    source_style_set = set(source_style_names)

    print(f"Source: {source}")
    print(f"Custom table styles found in source: {len(source_custom_table_styles)}")
    print(f"Source custom table style IDs: {', '.join(source_style_names) if source_style_names else 'none'}")
    print(f"Targets to update: {len(target_files)}")
    print(f"Strip attached template: {'yes' if args.strip_template else 'no'}")
    print(f"Dry run: {'yes' if args.dry_run else 'no'}")
    print(f"Report all table styles: {'yes' if args.report_all_table_styles else 'no'}")
    print(f"Report unexpected styles: {'yes' if args.report_unexpected else 'no'}")
    print(f"Unexpected prefixes: {', '.join(args.unexpected_prefix) if args.unexpected_prefix else 'none'}")
    print(f"Fail on unexpected: {'yes' if args.fail_on_unexpected else 'no'}")
    print(f"Check effects part: {'yes' if args.check_effects else 'no'}")
    print(f"Sync effects part: {'yes' if args.sync_effects else 'no'}")
    print(f"Deep report: {'yes' if args.deep_report else 'no'}")
    print()

    failures = 0
    changed_files = 0
    matched_files = 0
    unexpected_failures = 0

    for doc in target_files:
        try:
            result = analyze_and_prepare_doc(
                source_styles=source_custom_table_styles,
                source_custom_style_ids=source_style_set,
                target_path=doc,
                strip_template=args.strip_template,
                unexpected_prefixes=args.unexpected_prefix,
                check_effects=args.check_effects,
                sync_effects=args.sync_effects,
            )

            styles = result["styles"]
            before_custom = styles["before_custom"]
            after_custom = styles["after_custom"]
            before_all = styles["before_all"]
            after_all = styles["after_all"]
            removed_count = styles["removed_count"]
            removed_names = styles["removed_names"]
            inserted_count = styles["inserted_count"]
            inserted_names = styles["inserted_names"]
            unexpected_before = styles["unexpected_before"]
            unexpected_after = styles["unexpected_after"]

            final_match = sorted(after_custom) == sorted(inserted_names)
            will_change_styles = sorted(before_custom) != sorted(after_custom)

            will_change_effects = False
            if args.sync_effects and result["effects_transform"] is not None:
                eff = result["effects_transform"]
                will_change_effects = sorted(eff["before_custom"]) != sorted(eff["after_custom"])

            will_change_template = result["template_stripped"]
            will_change_file = will_change_styles or will_change_effects or will_change_template

            remaining_unexpected = (
                bool(unexpected_after["prefixed_all"])
                or bool(unexpected_after["prefixed_custom"])
                or bool(unexpected_after["custom_not_in_source"])
            )

            if args.sync_effects and result["effects_after_unexpected"] is not None:
                eff_unexp = result["effects_after_unexpected"]
                remaining_unexpected = remaining_unexpected or bool(
                    eff_unexp["prefixed_all"]
                    or eff_unexp["prefixed_custom"]
                    or eff_unexp["custom_not_in_source"]
                )

            print(f"[OK] {doc}")
            print(f"     before custom styles: {len(before_custom)}")
            print(f"     removed custom styles: {removed_count}")
            print(f"     inserted custom styles: {inserted_count}")
            print(f"     after custom styles: {len(after_custom)}")
            print(f"     attached template stripped: {'yes' if result['template_stripped'] else 'no'}")
            print(f"     file would change: {'yes' if will_change_file else 'no'}")

            print_list("removed custom style IDs", removed_names)
            print_list("inserted custom style IDs", inserted_names)
            print_list("final custom style IDs", after_custom)

            if args.report_all_table_styles:
                print_list("all table style IDs before", before_all)
                print_list("all table style IDs after", after_all)

            if args.report_unexpected:
                print_list("unexpected all-table styles before (by prefix)", unexpected_before["prefixed_all"])
                print_list("unexpected custom styles before (by prefix)", unexpected_before["prefixed_custom"])
                print_list("unexpected custom styles before (not in source)", unexpected_before["custom_not_in_source"])
                print_list("unexpected all-table styles after (by prefix)", unexpected_after["prefixed_all"])
                print_list("unexpected custom styles after (by prefix)", unexpected_after["prefixed_custom"])
                print_list("unexpected custom styles after (not in source)", unexpected_after["custom_not_in_source"])

            if final_match:
                print("     final custom style match to source: yes")
                matched_files += 1
            else:
                print("     final custom style match to source: NO")
                print("     WARNING: final custom table styles do not match source exactly")

            if args.deep_report:
                parts_present = result["parts_present"]
                print("     parts present:")
                for part_name in [
                    "word/styles.xml",
                    "word/stylesWithEffects.xml",
                    "word/settings.xml",
                    "word/_rels/settings.xml.rels",
                ]:
                    print(f"       {part_name}: {'yes' if parts_present[part_name] else 'no'}")

                ti = result["template_inspection"]
                print(f"     attached-template reference present: {'yes' if ti['has_any_template_reference'] else 'no'}")
                if ti["parse_errors"]:
                    print_list("template inspection parse errors", ti["parse_errors"])
                if ti["relationship_ids"]:
                    print_list("template relationship IDs", ti["relationship_ids"])
                if ti["targets"]:
                    print_list("template targets", ti["targets"])
                if result["template_strip_notes"]:
                    print_list("template strip notes", result["template_strip_notes"])

            if result["effects_exists"] and (args.check_effects or args.sync_effects):
                if args.sync_effects and result["effects_transform"] is not None:
                    eff = result["effects_transform"]
                    print("     stylesWithEffects.xml:")
                    print(f"       before custom styles: {len(eff['before_custom'])}")
                    print(f"       removed custom styles: {eff['removed_count']}")
                    print(f"       inserted custom styles: {eff['inserted_count']}")
                    print(f"       after custom styles: {len(eff['after_custom'])}")
                    if args.report_all_table_styles:
                        print(f"       all table style IDs before: {', '.join(eff['before_all']) if eff['before_all'] else 'none'}")
                        print(f"       all table style IDs after: {', '.join(eff['after_all']) if eff['after_all'] else 'none'}")
                    if args.report_unexpected and result["effects_after_unexpected"] is not None:
                        eff_u = result["effects_after_unexpected"]
                        print(f"       unexpected all-table styles after (by prefix): {', '.join(eff_u['prefixed_all']) if eff_u['prefixed_all'] else 'none'}")
                        print(f"       unexpected custom styles after (by prefix): {', '.join(eff_u['prefixed_custom']) if eff_u['prefixed_custom'] else 'none'}")
                        print(f"       unexpected custom styles after (not in source): {', '.join(eff_u['custom_not_in_source']) if eff_u['custom_not_in_source'] else 'none'}")
                elif args.check_effects and result["effects_report"] is not None:
                    eff = result["effects_report"]
                    print("     stylesWithEffects.xml:")
                    print(f"       custom styles: {len(eff['custom_names'])}")
                    if args.report_all_table_styles:
                        print(f"       all table style IDs: {', '.join(eff['all_names']) if eff['all_names'] else 'none'}")
                    if args.report_unexpected:
                        eff_u = eff["unexpected"]
                        print(f"       unexpected all-table styles (by prefix): {', '.join(eff_u['prefixed_all']) if eff_u['prefixed_all'] else 'none'}")
                        print(f"       unexpected custom styles (by prefix): {', '.join(eff_u['prefixed_custom']) if eff_u['prefixed_custom'] else 'none'}")
                        print(f"       unexpected custom styles (not in source): {', '.join(eff_u['custom_not_in_source']) if eff_u['custom_not_in_source'] else 'none'}")
            elif args.check_effects or args.sync_effects:
                print("     stylesWithEffects.xml: not present")

            if remaining_unexpected:
                print("     unexpected styles remain after processing: YES")
                if args.fail_on_unexpected:
                    unexpected_failures += 1
            else:
                print("     unexpected styles remain after processing: no")

            if will_change_file:
                changed_files += 1

            if not args.dry_run and will_change_file:
                new_styles_with_effects_xml = None
                if args.sync_effects and result["effects_transform"] is not None:
                    new_styles_with_effects_xml = result["effects_transform"]["new_xml"]

                write_updated_doc(
                    target_path=doc,
                    new_styles_xml=styles["new_xml"],
                    make_backup=args.backup,
                    new_styles_with_effects_xml=new_styles_with_effects_xml,
                    new_settings_xml=result["new_settings_xml"],
                    new_settings_rels_xml=result["new_settings_rels_xml"],
                )
                print("     write action: applied")
            elif args.dry_run:
                print("     write action: skipped (dry run)")
            else:
                print("     write action: skipped (no changes needed)")

        except Exception as e:
            failures += 1
            print(f"[FAIL] {doc}", file=sys.stderr)
            print(f"       {e}", file=sys.stderr)

        print()

    print("Summary:")
    print(f"  total targets: {len(target_files)}")
    print(f"  files matching source custom styles after processing: {matched_files}")
    print(f"  files that would change: {changed_files}")
    print(f"  files with unexpected-style policy failures: {unexpected_failures}")
    print(f"  failures: {failures}")
    print(f"  mode: {'dry run' if args.dry_run else 'write'}")

    if failures:
        return 2
    if args.fail_on_unexpected and unexpected_failures:
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
