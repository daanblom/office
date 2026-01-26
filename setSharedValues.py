import zipfile
import json
import re
import sys
import argparse
from pathlib import Path
import shutil

SKIP_TYPES = {"heading", "textElementPlaceholder"}
CDATA_JSON_RE = re.compile(r"<!\[CDATA\[(\{.*?\})\]\]>", re.DOTALL)


def find_cdata_json(xml_text):
    m = CDATA_JSON_RE.search(xml_text)
    if not m:
        return None, None
    return m.span(1), m.group(1)


def update_sharevalue(json_obj):
    changed = []
    if "formFields" not in json_obj:
        return json_obj, changed

    for f in json_obj["formFields"]:
        if not isinstance(f, dict):
            continue
        if f.get("type") in SKIP_TYPES:
            continue

        if f.get("shareValue") is not True:
            f["shareValue"] = True
            changed.append(f.get("name") or f.get("label") or "<unnamed>")

    return json_obj, changed


def process(input_path, dry_run=False):
    changed_fields = []
    changed_files = []
    modified_data = {}

    with zipfile.ZipFile(input_path, "r") as zin:
        for info in zin.infolist():
            data = zin.read(info.filename)

            if info.filename.startswith("customXml/") and info.filename.endswith(".xml") and "itemProps" not in info.filename:
                xml_text = data.decode("utf-8", errors="ignore")
                span, json_str = find_cdata_json(xml_text)

                if span and json_str:
                    try:
                        obj = json.loads(json_str)
                    except Exception:
                        obj = None

                    if isinstance(obj, dict) and "formFields" in obj:
                        updated_obj, changed = update_sharevalue(obj)

                        if changed:
                            new_json = json.dumps(updated_obj, ensure_ascii=False, separators=(",", ":"))
                            start, end = span
                            xml_text = xml_text[:start] + new_json + xml_text[end:]
                            data = xml_text.encode("utf-8")

                            changed_fields.extend(changed)
                            changed_files.append(info.filename)

            modified_data[info.filename] = data

    return changed_fields, changed_files, modified_data


def overwrite_docx(path, modified_data):
    tmp_path = path.with_suffix(path.suffix + ".tmp")

    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in modified_data.items():
            zout.writestr(name, data)

    tmp_path.replace(path)


def main():
    parser = argparse.ArgumentParser(
        description="Force Templafy formFields[*].shareValue=true in a Word template",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="""Examples:
  Dry run (audit only):
    python templafy_force_sharevalue_true.py template.docx --dry-run

  Overwrite file:
    python templafy_force_sharevalue_true.py template.docx

  Overwrite with backup:
    python templafy_force_sharevalue_true.py template.docx -b
"""
    )

    parser.add_argument("file", help="Path to .docx or extracted Word .zip")
    parser.add_argument("--dry-run", action="store_true", help="Show changes without modifying file")
    parser.add_argument("-b", "--backup", action="store_true", help="Create .bak backup before overwriting")

    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        print("❌ File not found")
        sys.exit(1)

    changed_fields, changed_files, modified_data = process(path, dry_run=args.dry_run)

    if not changed_fields:
        print("✅ No fields required updating.")
        return

    print(f"🔧 {len(changed_fields)} field(s) will be updated:")
    for f in changed_fields:
        print(f"  - {f}")

    print("\nAffected files:")
    for f in sorted(set(changed_files)):
        print(f"  - {f}")

    if args.dry_run:
        print("\n🧪 Dry run enabled — no files were modified.")
        return

    if args.backup:
        backup_path = path.with_suffix(path.suffix + ".bak")
        shutil.copy2(path, backup_path)
        print(f"📦 Backup created: {backup_path}")

    overwrite_docx(path, modified_data)
    print(f"\n✅ File overwritten: {path}")


if __name__ == "__main__":
    main()
