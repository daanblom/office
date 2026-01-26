import argparse
import sys
import zipfile
from pathlib import Path

DEFAULT_EXTS = {".docx", ".dotx"}  # add ".docm" if you want (but macros)


def is_word_package(path: Path) -> bool:
    return path.suffix.lower() in DEFAULT_EXTS and path.is_file()


def find_asset_in_docx(docx_path: Path, asset_id: str):
    """
    Returns list of internal zip part names that contain asset_id.
    """
    hits = []
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            for name in z.namelist():
                # Most useful places are XML; but sometimes it could be in rels or json-like text
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                data = z.read(name)
                text = data.decode("utf-8", errors="ignore")
                if asset_id in text:
                    hits.append(name)
    except zipfile.BadZipFile:
        return ["[NOT A VALID DOCX ZIP]"]
    except Exception as e:
        return [f"[ERROR: {e}]"]
    return hits


def iter_targets(input_path: Path):
    if input_path.is_file():
        yield input_path
        return

    # Directory: recursive scan
    for p in input_path.rglob("*"):
        if is_word_package(p):
            yield p


def main():
    parser = argparse.ArgumentParser(
        description="Recursively scan Word templates for a Templafy assetId reference."
    )
    parser.add_argument("path", help="File or directory to scan")
    parser.add_argument("asset_id", help="Templafy assetId (digits only)")
    parser.add_argument(
        "--show-parts",
        action="store_true",
        help="Also print which internal docx parts contain the assetId",
    )
    args = parser.parse_args()

    input_path = Path(args.path)

    if not input_path.exists():
        print("❌ Path not found.")
        sys.exit(1)

    asset_id = args.asset_id.strip()
    if not asset_id.isdigit():
        print("❌ asset_id must be digits only (e.g. 123456).")
        sys.exit(1)

    total = 0
    matched = 0

    for file_path in iter_targets(input_path):
        if not is_word_package(file_path):
            continue

        total += 1
        hits = find_asset_in_docx(file_path, asset_id)
        if hits:
            # If returned error markers, treat as "hit" only if actually found.
            if hits == ["[NOT A VALID DOCX ZIP]"] or (hits and hits[0].startswith("[ERROR:")):
                print(f"⚠️ {file_path}  ->  {hits[0]}")
                continue

            matched += 1
            print(f"✅ MATCH: {file_path}")
            if args.show_parts:
                for h in hits:
                    print(f"   - {h}")

    print("\n" + "-" * 60)
    print(f"Scanned: {total} Word file(s)")
    print(f"Matched: {matched} file(s)")
    print("-" * 60)


if __name__ == "__main__":
    main()
