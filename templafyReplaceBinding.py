import argparse
from pathlib import Path
import zipfile
import shutil
import tempfile

REPLACEMENTS = {
    "LogoLeft": "LeftLogo",
    "LogoRight": "RightLogo",
    "LogoCenter": "CenterLogo",
}

XML_EXTENSIONS = {".xml", ".rels"}


def replace_in_docx(docx_path: Path, write_changes: bool = True) -> int:
    total_replacements = 0

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        extract_dir = tmpdir / "unzipped"
        extract_dir.mkdir()

        with zipfile.ZipFile(docx_path, "r") as zin:
            zin.extractall(extract_dir)

        for file_path in extract_dir.rglob("*"):
            if not file_path.is_file():
                continue

            if file_path.suffix.lower() not in XML_EXTENSIONS:
                continue

            try:
                original = file_path.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                continue

            updated = original
            file_replacements = 0

            for old, new in REPLACEMENTS.items():
                count = updated.count(old)
                if count:
                    updated = updated.replace(old, new)
                    file_replacements += count

            if file_replacements:
                total_replacements += file_replacements
                if write_changes:
                    file_path.write_text(updated, encoding="utf-8")

        if write_changes and total_replacements > 0:
            rebuilt = tmpdir / "rebuilt.docx"
            with zipfile.ZipFile(rebuilt, "w", zipfile.ZIP_DEFLATED) as zout:
                for file_path in extract_dir.rglob("*"):
                    if file_path.is_file():
                        arcname = file_path.relative_to(extract_dir)
                        zout.write(file_path, arcname)

            shutil.copy2(rebuilt, docx_path)

    return total_replacements


def process_folder(folder: Path, recursive: bool, backup: bool, dry_run: bool):
    pattern = "**/*.docx" if recursive else "*.docx"
    files = list(folder.glob(pattern))

    if not files:
        print("No .docx files found.")
        return

    for docx_file in files:
        try:
            if backup and not dry_run:
                backup_path = docx_file.with_suffix(docx_file.suffix + ".bak")
                if not backup_path.exists():
                    shutil.copy2(docx_file, backup_path)

            replacements = replace_in_docx(
                docx_file,
                write_changes=not dry_run
            )

            action = "DRY-RUN" if dry_run else "UPDATED"
            print(f"{action}: {docx_file} → {replacements} replacement(s)")

        except Exception as e:
            print(f"ERROR: {docx_file} → {e}")


def main():
    parser = argparse.ArgumentParser(
        description="Replace Templafy bindings in DOCX files"
    )

    parser.add_argument(
        "path",
        help="Path to folder containing .docx files"
    )

    parser.add_argument(
        "-b", "--backup",
        action="store_true",
        help="Create .bak backup files"
    )

    parser.add_argument(
        "-d", "--dry-run",
        action="store_true",
        help="Preview changes without modifying files"
    )

    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="Process subfolders recursively"
    )

    args = parser.parse_args()

    folder = Path(args.path)

    if not folder.exists() or not folder.is_dir():
        print("Invalid folder path.")
        return

    process_folder(
        folder=folder,
        recursive=args.recursive,
        backup=args.backup,
        dry_run=args.dry_run
    )


if __name__ == "__main__":
    main()
