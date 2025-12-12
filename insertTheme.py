import zipfile
import sys
import os
import shutil
import tempfile

THEME_ZIP_PATH = 'ppt/theme/theme1.xml'


def replace_theme_xml(pptx_path, theme_xml_path, output_path):
    """Replace theme1.xml in a single pptx file."""
    temp_fd, temp_path = tempfile.mkstemp(suffix='.pptx')
    os.close(temp_fd)

    with zipfile.ZipFile(pptx_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w') as zout:
            for item in zin.infolist():
                if item.filename != THEME_ZIP_PATH:
                    zout.writestr(item, zin.read(item.filename))

            with open(theme_xml_path, 'rb') as theme_file:
                zout.writestr(THEME_ZIP_PATH, theme_file.read())

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.move(temp_path, output_path)
    print(f"✔ Updated theme: {output_path}")


def process_path(input_path, theme_xml_path, output_folder):
    """Process either a single pptx file or a directory recursively."""
    if os.path.isfile(input_path):
        if not input_path.lower().endswith('.pptx'):
            print(f"Skipping non-pptx file: {input_path}")
            return

        output_path = os.path.join(output_folder, os.path.basename(input_path))
        replace_theme_xml(input_path, theme_xml_path, output_path)

    elif os.path.isdir(input_path):
        for root, dirs, files in os.walk(input_path):
            # 🔒 Blacklist directories starting with "_"
            dirs[:] = [d for d in dirs if not d.startswith('_')]

            for file in files:
                if file.lower().endswith('.pptx'):
                    source_pptx = os.path.join(root, file)

                    rel_path = os.path.relpath(root, input_path)
                    target_dir = os.path.join(output_folder, rel_path)
                    target_pptx = os.path.join(target_dir, file)

                    replace_theme_xml(source_pptx, theme_xml_path, target_pptx)
    else:
        print(f"Error: Path '{input_path}' does not exist.")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(
            "Usage:\n"
            "  python replace_theme.py <pptx_or_directory> <theme1.xml> [output_folder]\n\n"
            "Examples:\n"
            "  python replace_theme.py slides.pptx theme1.xml out/\n"
            "  python replace_theme.py presentations/ theme1.xml themed_presentations/"
        )
        sys.exit(1)

    input_path = sys.argv[1]
    theme_xml_path = sys.argv[2]
    output_folder = sys.argv[3] if len(sys.argv) > 3 else '.'

    if not os.path.isfile(theme_xml_path):
        print(f"Error: Theme XML file '{theme_xml_path}' not found.")
        sys.exit(1)

    process_path(input_path, theme_xml_path, output_folder)

