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


def log_skip(path, reason):
    print(f"⚠ Skipping: {path} ({reason})")


def replace_theme_xml(office_path, theme_xml_path, output_path):
    """Replace theme1.xml in a single pptx or docx file."""
    theme_zip_path = get_theme_zip_path(office_path)
    if not theme_zip_path:
        log_skip(office_path, "unsupported extension")
        return

    # Quick sanity check: empty files will always fail
    try:
        if os.path.getsize(office_path) == 0:
            log_skip(office_path, "file is empty (0 bytes)")
            return
    except OSError as e:
        log_skip(office_path, f"cannot stat file: {e}")
        return

    temp_fd, temp_path = tempfile.mkstemp(suffix=os.path.splitext(office_path)[1])
    os.close(temp_fd)

    try:
        with zipfile.ZipFile(office_path, 'r') as zin:
            # Ensures the zip central directory is readable
            zin.testzip()

            with zipfile.ZipFile(temp_path, 'w') as zout:
                for item in zin.infolist():
                    if item.filename != theme_zip_path:
                        zout.writestr(item, zin.read(item.filename))

                with open(theme_xml_path, 'rb') as theme_file:
                    zout.writestr(theme_zip_path, theme_file.read())

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        shutil.move(temp_path, output_path)
        print(f"✔ Updated theme: {output_path}")

    except (zipfile.BadZipFile, zipfile.LargeZipFile) as e:
        log_skip(office_path, f"not a valid Office zip/package: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)

    except PermissionError as e:
        log_skip(office_path, f"permission error: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)

    except Exception as e:
        # Catch-all so one bad file never stops the batch
        log_skip(office_path, f"unexpected error: {type(e).__name__}: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)


def process_path(input_path, theme_xml_path, output_folder):
    """Process either a single Office file or a directory (recursively)."""
    valid_exts = ('.pptx', '.docx')

    if os.path.isfile(input_path):
        if not input_path.lower().endswith(valid_exts):
            log_skip(input_path, "unsupported extension")
            return
        output_path = os.path.join(output_folder, os.path.basename(input_path))
        replace_theme_xml(input_path, theme_xml_path, output_path)
        return

    if not os.path.isdir(input_path):
        print(f"Error: Path '{input_path}' does not exist.")
        return

    for root, dirs, files in os.walk(input_path):
        dirs[:] = [d for d in dirs if not d.startswith('_')]

        for file in files:
            if file.lower().endswith(valid_exts):
                source_file = os.path.join(root, file)
                rel_path = os.path.relpath(root, input_path)
                target_dir = os.path.join(output_folder, rel_path)
                target_file = os.path.join(target_dir, file)
                replace_theme_xml(source_file, theme_xml_path, target_file)


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

    args = parser.parse_args()

    if not os.path.isfile(args.theme_xml):
        print(f"Error: Theme XML file '{args.theme_xml}' not found.")
        sys.exit(1)

    process_path(args.input_path, args.theme_xml, args.output_folder)


if __name__ == "__main__":
    main()
