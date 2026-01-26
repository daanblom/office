import zipfile
import sys
import csv
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
DEFAULT_CSV = "SearchBinding.csv"


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


def audit(zip_path, search_term):
    results = []

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
                "SearchTerm": search_term,
                "SDTIndex": sdt_idx,
                "ParagraphIndex": para_index,
                "HeadingContext": heading,
                "VisibleText": visible_text or "[NO VISIBLE TEXT]",
                "VisibilityState": "Hidden/Placeholder" if not visible_text else "Visible",
                "Expressions": " | ".join(expressions)
            })

    return results


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
        description="Inspect Templafy visibility bindings inside Word (.docx or extracted .zip)",
        epilog="""Examples:
  searchBinding.py "IfElse(Equals(Form.Counterparty.Name" template.docx
  searchBinding.py "Form.Country" template.zip -l
"""
    )

    parser.add_argument(
        "search_term",
        help="Templafy expression or fragment to search for (e.g. IfElse(...))"
    )
    parser.add_argument(
        "file",
        help="Path to .docx file or .zip"
    )
    parser.add_argument(
        "-l", "--log",
        action="store_true",
        help="Write results to CSV (searchBinding.csv)"
    )

    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        print("❌ File does not exist\n")
        parser.print_help()
        sys.exit(1)

    results = audit(path, args.search_term)
    print_results(results)

    if args.log and results:
        write_csv(results, DEFAULT_CSV)


if __name__ == "__main__":
    main()

