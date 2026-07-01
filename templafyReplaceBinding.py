import argparse
from pathlib import Path
import zipfile
import shutil
import tempfile

REPLACEMENTS = {
    'CRMNumber':
    'CRM_Number'
}

XML_EXTENSIONS = {".xml", ".rels"}


def variants(s: str) -> set[str]:
    """
    Generate common on-disk variants of a binding string as it appears in DOCX XML.
    Includes:
      - literal
      - XML-escaped quotes (&quot;)
      - double-escaped quotes (\\&quot;) as seen inside JSON stored in attributes
    """
    v = set()
    v.add(s)

    xml = s.replace('"', "&quot;")
    v.add(xml)

    # Seen in w:tag JSON payloads: \" becomes \\&quot; in the XML text
    v.add(s.replace('"', r"\""))
    v.add(s.replace('"', r"\\\""))

    # Combine with &quot; then escape slashes
    v.add(xml.replace("&quot;", r"\\&quot;"))
    v.add(xml.replace("&quot;", r"\&quot;"))

    return v


def replace_many(text: str, replacements: dict[str, str]) -> tuple[str, int]:
    """
    Apply replacements and return (updated_text, total_count).
    Counts actual occurrences replaced across all variants.
    """
    updated = text
    total = 0

    for old, new in replacements.items():
        old_vars = variants(old)
        new_vars = {}

        # Map each old-variant to the corresponding new-variant in the same encoding style.
        for ov in old_vars:
            if "&quot;" in ov:
                nv = new.replace('"', "&quot;")
                # preserve \\&quot; style if present in ov
                if r"\\&quot;" in ov:
                    nv = nv.replace("&quot;", r"\\&quot;")
                elif r"\&quot;" in ov:
                    nv = nv.replace("&quot;", r"\&quot;")
                new_vars[ov] = nv
            elif r"\\\"" in ov:
                new_vars[ov] = new.replace('"', r"\\\"")
            elif r"\"" in ov:
                new_vars[ov] = new.replace('"', r"\"")
            else:
                new_vars[ov] = new

        # Apply all variant replacements
        for ov, nv in new_vars.items():
            c = updated.count(ov)
            if c:
                updated = updated.replace(ov, nv)
                total += c

    return updated, total


def replace_in_docx(docx_path: Path, write_changes: bool = True) -> int:
    total_replacements = 0

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        extract_dir = tmpdir / "unzipped"
        extract_dir.mkdir()

        with zipfile.ZipFile(docx_path, "r") as zin:
            zin.extractall(extract_dir)

        # Process XML-ish parts
        for file_path in extract_dir.rglob("*"):
            if not file_path.is_file():
                continue
            if file_path.suffix.lower() not in XML_EXTENSIONS:
                continue

            try:
                original = file_path.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                continue

            updated, file_repls = replace_many(original, REPLACEMENTS)

            if file_repls:
                total_replacements += file_repls
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


def process_docx_file(docx_file: Path, backup: bool, dry_run: bool):
    if docx_file.suffix.lower() != ".docx":
        print(f"Skipped (not .docx): {docx_file}")
        return

    try:
        if backup and not dry_run:
            backup_path = docx_file.with_suffix(docx_file.suffix + ".bak")
            if not backup_path.exists():
                shutil.copy2(docx_file, backup_path)

        replacements = replace_in_docx(docx_file, write_changes=not dry_run)

        action = "DRY-RUN" if dry_run else "UPDATED"
        print(f"{action}: {docx_file} → {replacements} replacement(s)")

    except zipfile.BadZipFile:
        print(f"ERROR: {docx_file} → not a valid .docx (BadZipFile)")
    except Exception as e:
        print(f"ERROR: {docx_file} → {e}")


def process_folder(folder: Path, recursive: bool, backup: bool, dry_run: bool):
    pattern = "**/*.docx" if recursive else "*.docx"
    files = list(folder.glob(pattern))

    if not files:
        print("No .docx files found.")
        return

    for docx_file in files:
        process_docx_file(docx_file=docx_file, backup=backup, dry_run=dry_run)


def main():
    parser = argparse.ArgumentParser(description="Replace Templafy bindings in DOCX files")

    parser.add_argument("path", help="Path to a .docx file OR a folder containing .docx files")

    parser.add_argument("-b", "--backup", action="store_true", help="Create .bak backup files")
    parser.add_argument("-d", "--dry-run", action="store_true", help="Preview changes without modifying files")
    parser.add_argument("-r", "--recursive", action="store_true", help="Process subfolders recursively (folders only)")

    args = parser.parse_args()
    target = Path(args.path)

    if not target.exists():
        print("Invalid path: does not exist.")
        return

    if target.is_file():
        process_docx_file(docx_file=target, backup=args.backup, dry_run=args.dry_run)
        return

    if target.is_dir():
        process_folder(folder=target, recursive=args.recursive, backup=args.backup, dry_run=args.dry_run)
        return

    print("Invalid path: not a file or directory.")


if __name__ == "__main__":
    main()
