import zipfile
import os
import argparse
from typing import Optional


def get_theme_zip_path(file_path: str) -> Optional[str]:
    """Return the internal zip path to theme1.xml for supported Office files."""
    ext = os.path.splitext(file_path.lower())[1]
    if ext == ".pptx":
        return "ppt/theme/theme1.xml"
    if ext == ".docx":
        return "word/theme/theme1.xml"
    return None


def extract_theme_xml(office_path: str, output_xml_path: str) -> bool:
    """Extract theme1.xml from a single .pptx or .docx file."""
    theme_zip_path = get_theme_zip_path(office_path)
    if not theme_zip_path:
        print(f"Skipping unsupported file: {office_path}")
        return False

    if not os.path.isfile(office_path):
        print(f"Error: File '{office_path}' not found.")
        return False

    with zipfile.ZipFile(office_path, "r") as zip_ref:
        if theme_zip_path not in zip_ref.namelist():
            print(f"theme1.xml not found in: {office_path}")
            return False

        os.makedirs(os.path.dirname(output_xml_path) or ".", exist_ok=True)
        with zip_ref.open(theme_zip_path) as theme_file, open(output_xml_path, "wb") as out_f:
            out_f.write(theme_file.read())

    print(f"✔ Extracted theme1.xml -> {output_xml_path}")
    return True


def process_path(input_path: str, output_folder: str, recursive: bool = False) -> None:
    """Process either a single Office file or a directory of Office files."""
    valid_exts = (".pptx", ".docx")

    # ✅ Single file: ALWAYS write theme1.xml
    if os.path.isfile(input_path):
        if not input_path.lower().endswith(valid_exts):
            print(f"Skipping unsupported file: {input_path}")
            return

        out_path = os.path.join(output_folder, "theme1.xml")
        extract_theme_xml(input_path, out_path)
        return

    # Folder: write per-file outputs to avoid overwriting
    if os.path.isdir(input_path):
        if recursive:
            walker = os.walk(input_path)
        else:
            walker = [(input_path, [], os.listdir(input_path))]

        for root, dirs, files in walker:
            dirs[:] = [d for d in dirs if not d.startswith("_")]

            for file in files:
                if not file.lower().endswith(valid_exts):
                    continue

                source_file = os.path.join(root, file)

                rel_path = os.path.relpath(root, input_path)
                target_dir = os.path.join(output_folder, rel_path)

                stem = os.path.splitext(file)[0]
                target_file = os.path.join(target_dir, f"{stem}_theme1.xml")

                extract_theme_xml(source_file, target_file)
        return

    print(f"Error: Path '{input_path}' does not exist.")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract theme1.xml from .pptx/.docx files."
    )
    parser.add_argument("input_path", help="A .pptx/.docx file or a directory")
    parser.add_argument(
        "output_folder",
        nargs="?",
        default=".",
        help="Output folder (default: current directory)",
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="If input_path is a directory, process recursively"
    )

    args = parser.parse_args()
    process_path(args.input_path, args.output_folder, recursive=args.recursive)


if __name__ == "__main__":
    main()
