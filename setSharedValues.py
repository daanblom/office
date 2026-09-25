import argparse
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

SKIP_TYPES = {"heading", "textElementPlaceholder"}
SUPPORTED_EXTENSIONS = {".docx", ".zip"}

CDATA_JSON_RE = re.compile(
    r"<!\[CDATA\[(\{.*?\})\]\]>",
    re.DOTALL,
)


def find_cdata_json(xml_text):
    match = CDATA_JSON_RE.search(xml_text)

    if not match:
        return None, None

    return match.span(1), match.group(1)


def update_sharevalue(json_obj, target_value=True):
    changed = []

    if "formFields" not in json_obj:
        return json_obj, changed

    for field in json_obj["formFields"]:
        if not isinstance(field, dict):
            continue

        if field.get("type") in SKIP_TYPES:
            continue

        if field.get("shareValue") is not target_value:
            field["shareValue"] = target_value

            changed.append(
                field.get("name")
                or field.get("label")
                or "<unnamed>"
            )

    return json_obj, changed


def process(input_path, target_value=True):
    changed_fields = []
    changed_files = []
    modified_data = {}

    try:
        with zipfile.ZipFile(input_path, "r") as zip_file:
            for info in zip_file.infolist():
                data = zip_file.read(info.filename)

                is_custom_xml = (
                    info.filename.startswith("customXml/")
                    and info.filename.endswith(".xml")
                    and "itemProps" not in info.filename
                )

                if is_custom_xml:
                    xml_text = data.decode("utf-8", errors="ignore")
                    span, json_str = find_cdata_json(xml_text)

                    if span and json_str:
                        try:
                            obj = json.loads(json_str)
                        except json.JSONDecodeError:
                            obj = None

                        if isinstance(obj, dict) and "formFields" in obj:
                            updated_obj, changed = update_sharevalue(
                                obj,
                                target_value=target_value,
                            )

                            if changed:
                                new_json = json.dumps(
                                    updated_obj,
                                    ensure_ascii=False,
                                    separators=(",", ":"),
                                )

                                start, end = span
                                xml_text = (
                                    xml_text[:start]
                                    + new_json
                                    + xml_text[end:]
                                )

                                data = xml_text.encode("utf-8")

                                changed_fields.extend(changed)
                                changed_files.append(info.filename)

                modified_data[info.filename] = data

    except zipfile.BadZipFile:
        raise ValueError(f"Not a valid Word ZIP file: {input_path}")

    return changed_fields, changed_files, modified_data


def overwrite_docx(path, modified_data):
    tmp_path = path.with_suffix(path.suffix + ".tmp")

    try:
        with zipfile.ZipFile(
            tmp_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as zip_file:
            for name, data in modified_data.items():
                zip_file.writestr(name, data)

        tmp_path.replace(path)

    except Exception:
        if tmp_path.exists():
            tmp_path.unlink()
        raise


def find_input_files(path, recursive=False):
    """
    Return files to process.

    Without --recursive:
        The supplied path must be a single .docx or .zip file.

    With --recursive:
        The supplied path must be a directory. All .docx and .zip files
        below that directory are returned recursively.
    """
    if recursive:
        if not path.is_dir():
            raise ValueError(
                "When using --recursive, the input path must be a directory."
            )

        files = [
            file_path
            for file_path in path.rglob("*")
            if (
                file_path.is_file()
                and file_path.suffix.lower() in SUPPORTED_EXTENSIONS
                and not file_path.name.endswith(".bak")
                and not file_path.name.endswith(".tmp")
            )
        ]

        return sorted(files)

    if not path.is_file():
        raise ValueError(f"File not found: {path}")

    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file type: {path.suffix}. "
            f"Expected one of: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )

    return [path]


def process_file(path, args, target_value, action_label):
    print(f"\nProcessing: {path}")

    try:
        changed_fields, changed_files, modified_data = process(
            path,
            target_value=target_value,
        )
    except ValueError as error:
        print(f"❌ {error}")
        return False

    if not changed_fields:
        print(
            f"✅ No fields required updating "
            f"(shareValue={action_label} already set on all eligible fields)."
        )
        return True

    print(
        f"🔧 {len(changed_fields)} field(s) will have "
        f"shareValue set to {action_label}:"
    )

    for field_name in changed_fields:
        print(f"  - {field_name}")

    print("\nAffected files inside archive:")
    for filename in sorted(set(changed_files)):
        print(f"  - {filename}")

    if args.dry_run:
        print("\n🧪 Dry run enabled — no files were modified.")
        return True

    if args.backup:
        backup_path = path.with_suffix(path.suffix + ".bak")
        shutil.copy2(path, backup_path)
        print(f"📦 Backup created: {backup_path}")

    overwrite_docx(path, modified_data)

    print(
        f"\n✅ File overwritten: {path} "
        f"(shareValue={action_label})"
    )

    return True


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Set Templafy formFields[*].shareValue in a Word template "
            "or multiple templates"
        ),
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="""Examples:

  Dry run on one file:
    python setSharedValues.py template.docx --dry-run

  Set shareValue=true on one file:
    python setSharedValues.py template.docx

  Set shareValue=false on one file:
    python setSharedValues.py template.docx -u

  Process all Word files recursively in a folder:
    python setSharedValues.py ./templates -r

  Recursively process files, create backups, and show changes:
    python setSharedValues.py ./templates -r -b

  Recursively process files without modifying them:
    python setSharedValues.py ./templates -r --dry-run
""",
    )

    parser.add_argument(
        "file",
        help=(
            "Path to a .docx or extracted Word .zip file. "
            "With --recursive, provide a directory."
        ),
    )

    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help=(
            "Recursively process all .docx and .zip files "
            "under the supplied directory."
        ),
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show changes without modifying files.",
    )

    parser.add_argument(
        "-b",
        "--backup",
        action="store_true",
        help="Create a .bak backup before overwriting each file.",
    )

    parser.add_argument(
        "-u",
        "--unshare",
        action="store_true",
        help="Set shareValue=false instead of true.",
    )

    args = parser.parse_args()

    target_value = not args.unshare
    action_label = "true" if target_value else "false"
    input_path = Path(args.file)

    try:
        files_to_process = find_input_files(
            input_path,
            recursive=args.recursive,
        )
    except ValueError as error:
        print(f"❌ {error}")
        sys.exit(1)

    if not files_to_process:
        print(
            f"⚠️ No .docx or .zip files found under: {input_path}"
        )
        return

    print(f"Found {len(files_to_process)} file(s) to process.")

    successful = 0
    failed = 0

    for path in files_to_process:
        try:
            result = process_file(
                path,
                args,
                target_value,
                action_label,
            )

            if result:
                successful += 1
            else:
                failed += 1

        except Exception as error:
            failed += 1
            print(f"❌ Failed to process {path}: {error}")

    print("\nSummary:")
    print(f"  Successfully processed: {successful}")
    print(f"  Failed: {failed}")

    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
