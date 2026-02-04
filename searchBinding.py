import zipfile
import sys
import csv
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

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


def audit(zip_path: Path, search_term: str):
    """
    Audits a .docx or .zip (Word package) for SDTs containing search_term.
    Returns list of result dicts.
    """
    results = []

    try:
        with zipfile.ZipFile(zip_path) as z:
            if "word/document.xml" not in z.namelist():
                return results

            root = ET.fromstring(z.read("word/document.xml"))
            paragraphs = root.findall(".//w:body//w:p", NS)

            paragraph_map = []
            last_heading = ""

            for idx, p in enumerate(paragraphs, 1):
                txt = text_of(p)
                if is_heading(p) and txt:
                    last_heading = txt
                paragraph_map.append((idx, p, last_heading))

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

                results.append({
                    "File": str(zip_path),
                    "SearchTerm": search_term,
                    "SDTIndex": sdt_idx,
                    "ParagraphIndex": para_index,
                    "HeadingContext": heading,
                    "VisibleText": visible_text or "[NO VISIBLE TEXT]",
                    "VisibilityState": "Hidden/Placeholder" if not visible_text else "Visible",
                    "Expressions": " | ".join(expressions)
                })

    except zipfile.BadZipFile:
        # Not a valid zip/docx package; ignore quietly or print warning in caller
        return results

    return results


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
        description="Inspect Templafy visibility bindings inside Word (.docx/.zip) or recursively in a directory",
        epilog="""Examples:
  searchBinding.py "IfElse(Equals(Form.Counterparty.Name" template.docx
  searchBinding.py "Form.Country" template.zip -l
  searchBinding.py "Form.Country" ./templates/ -l
"""
    )

    parser.add_argument(
        "search_term",
        help="Templafy expression or fragment to search for (e.g. IfElse(...))"
    )
    parser.add_argument(
        "path",
        help="Path to a .docx/.zip file OR a directory (will be searched recursively)"
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

    input_path = Path(args.path)
    if not input_path.exists():
        print("❌ Path does not exist\n")
        parser.print_help()
        sys.exit(1)

    # Collect targets
    targets = list(iter_targets(input_path))
    if not targets:
        print("⚠️ No .docx or .zip files found.")
        sys.exit(0)

    all_results = []
    scanned = 0

    for target in targets:
        scanned += 1
        res = audit(target, args.search_term)
        if not res and args.no_skip_badzip and target.suffix.lower() in SUPPORTED_EXTS:
            # If it was a supported extension but not actually a zip package, audit() returns []
            # This warning is optional; enable via --no-skip-badzip
            try:
                # Quick check to differentiate "valid but no match" from "not a zip"
                with zipfile.ZipFile(target):
                    pass
            except zipfile.BadZipFile:
                print(f"⚠️ Skipping invalid zip/docx: {target}")
        all_results.extend(res)

    print(f"🔎 Scanned {scanned} file(s)")
    print_results(all_results)

    if args.log and all_results:
        write_csv(all_results, DEFAULT_CSV)


if __name__ == "__main__":
    main()
