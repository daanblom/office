import zipfile
import sys
import csv
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path
import tempfile
import os

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
DEFAULT_CSV = "SearchBinding.csv"
SUPPORTED_EXTS = {".docx", ".zip"}


def text_of(elem):
    if elem is None:
        return ""
    return "".join(t.text for t in elem.findall(".//w:t", NS) if t.text).strip()


def is_heading(p):
    pPr = p.find("w:pPr", NS)
    if pPr is None:
        return False
    pStyle = pPr.find("w:pStyle", NS)
    return pStyle is not None and "Heading" in pStyle.attrib.get(f"{{{NS['w']}}}val", "")


def extract_expressions(xml):
    out, i = [], 0
    while "{{" in xml[i:]:
        i = xml.find("{{", i)
        j = xml.find("}}", i)
        if i == -1 or j == -1:
            break
        out.append(xml[i + 2:j])
        i = j + 2
    return list(set(out))


def replace_in_zip_entry(zip_path: Path, entry_name: str, old: str, new: str):
    """
    Rewrites a single entry inside a zip/docx package.
    Returns the number of replacements made in the entry.
    """
    total_replacements = 0
    tmp_path = None

    try:
        # Create temp file in the same directory as the target to avoid cross-device move errors
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=zip_path.suffix,
            dir=str(zip_path.parent)
        ) as tmp_file:
            tmp_path = Path(tmp_file.name)

        with zipfile.ZipFile(zip_path, "r") as zin, zipfile.ZipFile(tmp_path, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == entry_name:
                    try:
                        text = data.decode("utf-8")
                    except UnicodeDecodeError:
                        text = data.decode("utf-8", errors="replace")

                    count = text.count(old)
                    if count:
                        text = text.replace(old, new)
                        data = text.encode("utf-8")
                        total_replacements += count

                zout.writestr(item, data)

        # Preserve file mode where possible
        try:
            st = zip_path.stat()
            os.chmod(tmp_path, st.st_mode)
        except OSError:
            pass

        # Atomic replace in same directory/filesystem
        os.replace(tmp_path, zip_path)

    finally:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass

    return total_replacements


def audit(zip_path: Path, search_term: str, replace_old: str = None, replace_new: str = None, dry_run: bool = False):
    """
    Audits a .docx or .zip (Word package) for SDTs containing search_term.
    Optionally replaces replace_old -> replace_new inside word/document.xml.
    Returns:
      - results: list of result dicts
      - replacement_count: number of actual replacements made in file
      - potential_replacement_count: number of matching occurrences found in matching SDTs
    """
    results = []
    replacement_count = 0
    potential_replacement_count = 0

    try:
        with zipfile.ZipFile(zip_path) as z:
            if "word/document.xml" not in z.namelist():
                return results, replacement_count, potential_replacement_count

            xml_bytes = z.read("word/document.xml")
            root = ET.fromstring(xml_bytes)
            paragraphs = root.findall(".//w:body//w:p", NS)

            paragraph_map = []
            last_heading = ""

            for idx, p in enumerate(paragraphs, 1):
                txt = text_of(p)
                if is_heading(p) and txt:
                    last_heading = txt
                paragraph_map.append((idx, p, last_heading))

            matching_sdts_for_replace = 0

            for sdt_idx, sdt in enumerate(root.findall(".//w:sdt", NS), 1):
                xml = ET.tostring(sdt, encoding="unicode")
                if search_term not in xml:
                    continue

                visible_text = text_of(sdt.find("w:sdtContent", NS))
                expressions = extract_expressions(xml)

                para_index = ""
                heading = ""

                for idx, p, h in paragraph_map:
                    if sdt in list(p.iter()):
                        para_index = idx
                        heading = h
                        break

                row = {
                    "File": str(zip_path),
                    "SearchTerm": search_term,
                    "SDTIndex": sdt_idx,
                    "ParagraphIndex": para_index,
                    "HeadingContext": heading,
                    "VisibleText": visible_text or "[NO VISIBLE TEXT]",
                    "VisibilityState": "Hidden/Placeholder" if not visible_text else "Visible",
                    "Expressions": " | ".join(expressions)
                }

                if replace_old is not None:
                    occurrences = xml.count(replace_old)
                    potential_replacement_count += occurrences
                    if occurrences > 0:
                        matching_sdts_for_replace += 1

                    row["ReplacementOld"] = replace_old
                    row["ReplacementNew"] = replace_new
                    row["OccurrencesInSDT"] = occurrences
                    row["DryRun"] = "Yes" if dry_run else "No"

                results.append(row)

        if replace_old is not None and matching_sdts_for_replace > 0 and not dry_run:
            replacement_count = replace_in_zip_entry(
                zip_path,
                "word/document.xml",
                replace_old,
                replace_new
            )

    except zipfile.BadZipFile:
        return results, replacement_count, potential_replacement_count

    return results, replacement_count, potential_replacement_count


def iter_targets(input_path: Path):
    """
    If input_path is a file: yield it.
    If it's a directory: recursively yield *.docx and *.zip files.
    """
    if input_path.is_file():
        yield input_path
        return

    if input_path.is_dir():
        for p in input_path.rglob("*"):
            if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                yield p


def print_results(results):
    if not results:
        print("⚠️ No matching visibility bindings found.")
        return

    for r in results:
        print("=" * 80)
        for k, v in r.items():
            print(f"{k:17}: {v}")
    print("=" * 80)
    print(f"✅ {len(results)} matching SDT(s) found")


def write_csv(results, filename):
    with open(filename, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=results[0].keys())
        writer.writeheader()
        writer.writerows(results)
    print(f"📄 Results written to {filename}")


def main():
    parser = argparse.ArgumentParser(
        description="Inspect Templafy visibility bindings inside Word (.docx/.zip) files or recursively in a directory.",
        epilog="""Examples:
  Search only:
    python searchBinding.py "Form.Country" template.docx
    python searchBinding.py "Form.Country" ./templates/ -l

  Replace using OLD as the search term:
    python searchBinding.py -r "NL-nl" "nl-NL" .
    python searchBinding.py -r "Form.Country" "Form.CountryCode" ./templates/ -l

  Replace using an explicit search term:
    python searchBinding.py "IfElse(" . -r "NL-nl" "nl-NL"

  Dry run:
    python searchBinding.py -r "NL-nl" "nl-NL" . -d
    python searchBinding.py "IfElse(" . -r "NL-nl" "nl-NL" --dry-run -l
""",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    parser.add_argument(
        "args",
        nargs="+",
        help="Either [search_term path] or [path] when using -r"
    )
    parser.add_argument(
        "-r", "--replace",
        nargs=2,
        metavar=("OLD", "NEW"),
        help="Replace OLD with NEW inside word/document.xml for files where matching SDTs are found"
    )
    parser.add_argument(
        "-d", "--dry-run",
        action="store_true",
        help="Preview replacements without modifying files"
    )
    parser.add_argument(
        "-l", "--log",
        action="store_true",
        help=f"Write results to CSV ({DEFAULT_CSV})"
    )
    parser.add_argument(
        "--no-skip-badzip",
        action="store_true",
        help="If set, prints a warning when a file is not a valid .docx/.zip package"
    )

    args = parser.parse_args()
    args_list = args.args

    if args.replace:
        replace_old, replace_new = args.replace

        if len(args_list) == 1:
            search_term = replace_old
            path_str = args_list[0]
        elif len(args_list) == 2:
            search_term, path_str = args_list
        else:
            print("❌ Invalid arguments for replace mode\n")
            parser.print_help()
            sys.exit(1)
    else:
        if len(args_list) != 2:
            print("❌ search_term and path are required\n")
            parser.print_help()
            sys.exit(1)

        search_term, path_str = args_list
        replace_old = None
        replace_new = None

        if args.dry_run:
            print("⚠️ -d/--dry-run has no effect unless --replace is used")

    input_path = Path(path_str)
    if not input_path.exists():
        print("❌ Path does not exist\n")
        parser.print_help()
        sys.exit(1)

    targets = list(iter_targets(input_path))
    if not targets:
        print("⚠️ No .docx or .zip files found.")
        sys.exit(0)

    all_results = []
    scanned = 0
    changed_files = 0
    total_replacements = 0
    total_potential_replacements = 0

    for target in targets:
        scanned += 1
        res, replacement_count, potential_count = audit(
            target,
            search_term,
            replace_old,
            replace_new,
            dry_run=args.dry_run
        )

        if replace_old is not None:
            total_potential_replacements += potential_count

            if args.dry_run and potential_count > 0:
                print(f"🧪 Dry run: would update {target} ({potential_count} replacement(s))")
                changed_files += 1
            elif replacement_count > 0:
                print(f"✏️ Updated: {target} ({replacement_count} replacement(s))")
                changed_files += 1
                total_replacements += replacement_count

        if not res and args.no_skip_badzip and target.suffix.lower() in SUPPORTED_EXTS:
            try:
                with zipfile.ZipFile(target):
                    pass
            except zipfile.BadZipFile:
                print(f"⚠️ Skipping invalid zip/docx: {target}")

        all_results.extend(res)

    print(f"🔎 Scanned {scanned} file(s)")
    print_results(all_results)

    if args.replace:
        if args.dry_run:
            print(f"🧪 Dry run summary: {changed_files} file(s) would be updated")
            print(f"🧪 Potential replacements found: {total_potential_replacements}")
        else:
            print(f"🔁 Replacements applied in {changed_files} file(s)")
            print(f"🔁 Total replacements made: {total_replacements}")

    if args.log and all_results:
        write_csv(all_results, DEFAULT_CSV)


if __name__ == "__main__":
    main()
