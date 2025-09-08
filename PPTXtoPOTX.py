#!/usr/bin/env python3
import argparse
import os
import sys
import tempfile
import zipfile
from pathlib import Path

PRESENTATION_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
TEMPLATE_CT     = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
CONTENT_TYPES   = "[Content_Types].xml"

def convert_pptx_to_potx(src: Path, dest: Path) -> None:
    """
    Create a POTX at 'dest' from PPTX 'src' by rewriting [Content_Types].xml.
    Writes via a temp file in dest's directory to avoid cross-device link errors.
    """
    # Make dest absolute & ensure parent exists
    dest = dest.resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src, "r") as zin:
        # load and tweak [Content_Types].xml
        try:
            with zin.open(CONTENT_TYPES, "r") as f:
                xml = f.read().decode("utf-8")
        except KeyError:
            raise RuntimeError(f"{src} is missing {CONTENT_TYPES}")

        if PRESENTATION_CT in xml:
            xml = xml.replace(PRESENTATION_CT, TEMPLATE_CT)
        # else: already template or unusual declaration; proceed anyway

        # write to temp file in the SAME directory as dest
        with tempfile.NamedTemporaryFile(
            dir=dest.parent, prefix=".tmp_pptx2potx_", suffix=".zip", delete=False
        ) as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == CONTENT_TYPES:
                        zout.writestr(CONTENT_TYPES, xml.encode("utf-8"))
                    else:
                        zout.writestr(item, zin.read(item.filename))
            # atomic replace within same filesystem
            os.replace(tmp_path, dest)
        except Exception:
            try:
                if tmp_path.exists():
                    tmp_path.unlink()
            finally:
                raise

def should_skip(name: str) -> bool:
    return name.startswith("~$")  # Office temp files

def main():
    ap = argparse.ArgumentParser(description="Convert all .pptx to .potx next to source (no Office/COM).")
    ap.add_argument("root", type=Path, help="Root folder to scan (recurses)")
    ap.add_argument("--overwrite", action="store_true", help="Overwrite existing .potx")
    args = ap.parse_args()

    root = args.root.resolve()
    if not root.is_dir():
        print(f"Root not found or not a directory: {root}", file=sys.stderr)
        sys.exit(2)

    converted = skipped = errors = 0

    for dirpath, _, filenames in os.walk(root):
        dir_abs = Path(dirpath)
        for fname in filenames:
            if not fname.lower().endswith(".pptx"):
                continue
            if should_skip(fname):
                continue

            src = (dir_abs / fname).resolve()
            dest = src.with_suffix(".potx")

            if dest.exists() and not args.overwrite:
                skipped += 1
                print(f"Skip (exists): {dest}")
                continue

            try:
                convert_pptx_to_potx(src, dest)
                if dest.exists() and dest.stat().st_size > 0:
                    converted += 1
                    print(f"OK: {dest}")
                else:
                    errors += 1
                    print(f"ERROR: Output missing/empty: {dest}", file=sys.stderr)
            except Exception as e:
                errors += 1
                print(f"ERROR on {src}: {e}", file=sys.stderr)

    print(f"Done. Converted: {converted} | Skipped: {skipped} | Errors: {errors}")

if __name__ == "__main__":
    main()
