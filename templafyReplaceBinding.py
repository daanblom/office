import argparse
import re
import zipfile
import shutil
import tempfile
from pathlib import Path

REPLACEMENTS = {
    'CRMNumber': 'CRM_Number'
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

    v.add(s.replace('"', r"\""))
    v.add(s.replace('"', r"\\\""))

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

        for ov in old_vars:
            if "&quot;" in ov:
                nv = new.replace('"', "&quot;")
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

        for ov, nv in new_vars.items():
            c = updated.count(ov)
            if c:
                updated = updated.replace(ov, nv)
                total += c

    return updated, total


# ---------------------------------------------------------------------------
# Binding listing
# ---------------------------------------------------------------------------

# Matches binding names as they appear in Templafy w:tag JSON payloads and
# similar XML attribute values, e.g.:
#   "fieldName":"CRMNumber"
#   "name":"SomeBinding"
# Captures the value between the quotes after the colon.
# Matches any {{ ... }} expression in raw XML text.
# Handles both literal {{ }} and XML-escaped {{ (&amp;amp; etc. are unlikely here,
# but &amp; in attribute values is possible — we normalise before matching).
_BINDING_RE = re.compile(r'\{\{([^}]+)\}\}')


def list_bindings_in_docx(docx_path: Path) -> list[str]:
    """
    Return all {{ }} binding expressions found in the XML parts of a .docx file.
    Results are in order of appearance; duplicates are preserved.
    Strips surrounding whitespace from each match.
    """
    found = []

    try:
        with zipfile.ZipFile(docx_path, "r") as zin:
            for entry in zin.namelist():
                suffix = Path(entry).suffix.lower()
                if suffix not in XML_EXTENSIONS:
                    continue
                try:
                    content = zin.read(entry).decode("utf-8")
                except (UnicodeDecodeError, KeyError):
                    continue

                # Normalise XML escaping so {{ }} aren't obscured
                # Normalise XML and JSON escape variants before matching
                content_clean = (
                    content
                    .replace("&quot;", '"')
                    .replace("&amp;", "&")
                    .replace('\\"', '"')      # \" → " (JSON-escaped quotes in XML attributes)
                    .replace('\\\\"', '"')    # \\\" → " (double-escaped variant)
)
                for match in _BINDING_RE.finditer(content_clean):
                    found.append("{{" + match.group(1).strip() + "}}")

    except zipfile.BadZipFile:
        print(f"ERROR: {docx_path} → not a valid .docx (BadZipFile)")

    return found

def print_bindings(docx_file: Path, unique: bool):
    bindings = list_bindings_in_docx(docx_file)

    if not bindings:
        print(f"{docx_file}: no bindings found")
        return

    if unique:
        bindings = sorted(set(bindings))

    print(f"\n{docx_file} — {len(bindings)} binding(s){'  [unique]' if unique else ''}:")
    for b in bindings:
        print(f"  {b}")


# ---------------------------------------------------------------------------
# Replacement logic (unchanged)
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Replace or inspect Templafy bindings in DOCX files"
    )

    parser.add_argument("path", help="Path to a .docx file OR a folder containing .docx files")

    parser.add_argument("-b", "--backup", action="store_true", help="Create .bak backup files before modifying")
    parser.add_argument("-d", "--dry-run", action="store_true", help="Preview replacements without modifying files")
    parser.add_argument("-r", "--recursive", action="store_true", help="Process subfolders recursively (folders only)")
    parser.add_argument("-l", "--list", action="store_true", help="List all bindings found in the file(s) without making changes")
    parser.add_argument("-u", "--unique", action="store_true", help="When used with -l: show only unique bindings, sorted alphabetically")

    args = parser.parse_args()

    # Validate -u dependency
    if args.unique and not args.list:
        parser.error("-u / --unique requires -l / --list")

    target = Path(args.path)

    if not target.exists():
        print("Invalid path: does not exist.")
        return

    # -l mode: list bindings, skip replacement logic
    if args.list:
        if target.is_file():
            if target.suffix.lower() != ".docx":
                print(f"Skipped (not .docx): {target}")
                return
            print_bindings(target, unique=args.unique)

        elif target.is_dir():
            pattern = "**/*.docx" if args.recursive else "*.docx"
            files = list(target.glob(pattern))
            if not files:
                print("No .docx files found.")
                return
            for docx_file in files:
                print_bindings(docx_file, unique=args.unique)
        return

    # Default mode: replacement
    if target.is_file():
        process_docx_file(docx_file=target, backup=args.backup, dry_run=args.dry_run)
        return

    if target.is_dir():
        process_folder(folder=target, recursive=args.recursive, backup=args.backup, dry_run=args.dry_run)
        return

    print("Invalid path: not a file or directory.")


if __name__ == "__main__":
    main()
