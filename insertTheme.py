import zipfile
import sys
import os
import shutil
import tempfile
import argparse


def get_theme_zip_path(file_path):
    ext = os.path.splitext(file_path.lower())[1]
    if ext == '.pptx':
        return 'ppt/theme/theme1.xml'
    elif ext == '.docx':
        return 'word/theme/theme1.xml'
    else:
        return None


def replace_theme_xml(office_path, theme_xml_path, output_path):
    """Replace theme1.xml in a single pptx or docx file."""
    theme_zip_path = get_theme_zip_path(office_path)
    if not theme_zip_path:
        print(f"Skipping unsupported file: {office_path}")
        return

    temp_fd, temp_path = tempfile.mkstemp(suffix=os.path.splitext(office_path)[1])
    os.close(temp_fd)

    with zipfile.ZipFile(office_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w') as zout:
            for item in zin.infolist():
                if item.filename != theme_zip_path:
                    zout.writestr(item, zin.read(item.filename))

            with open(theme_xml_path, 'rb') as theme_file:
                zout.writestr(theme_zip_path, theme_file.read())

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.move(temp_path, output_path)
    print(f"✔ Updated theme: {output_path}")


def process_path(input_path, theme_xml_path, output_folder, recursive=False):
    """Process either a single Office file or a directory."""
    valid_exts = ('.pptx', '.docx')

    if os.path.isfile(input_path):
        if not input_path.lower().endswith(valid_exts):
            print(f"Skipping unsupported file: {input_path}")
            return

        output_path = os.path.join(output_folder, os.path.basename(input_path))
        replace_theme_xml(input_path, theme_xml_path, output_path)

    elif os.path.isdir(input_path):
        if recursive:
            walker = os.walk(input_path)
        else:
            # non-recursive: only one level
            walker = [(input_path, [], os.listdir(input_path))]

        for root, dirs, files in walker:
            # 🔒 Blacklist directories starting with "_"
            dirs[:] = [d for d in dirs if not d.startswith('_')]

            for file in files:
                if file.lower().endswith(valid_exts):
                    source_file = os.path.join(root, file)

                    rel_path = os.path.relpath(root, input_path)
                    target_dir = os.path.join(output_folder, rel_path)
                    target_file = os.path.join(target_dir, file)

                    replace_theme_xml(source_file, theme_xml_path, target_file)

    else:
        print(f"Error: Path '{input_path}' does not exist.")


def main():
    parser = argparse.ArgumentParser(
        description="Replace theme1.xml inside pptx/docx files."
    )
    parser.add_argument("input_path", help="A .pptx/.docx file or a directory")
    parser.add_argument("theme_xml", help="Path to theme1.xml replacement file")
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

    if not os.path.isfile(args.theme_xml):
        print(f"Error: Theme XML file '{args.theme_xml}' not found.")
        sys.exit(1)

    process_path(args.input_path, args.theme_xml, args.output_folder, recursive=args.recursive)


if __name__ == "__main__":
    main()
