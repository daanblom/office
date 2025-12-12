import zipfile
import sys
import os

def extract_theme_xml(pptx_path, output_folder='.'):
    # Check if file exists
    if not os.path.isfile(pptx_path):
        print(f"Error: File '{pptx_path}' not found.")
        return

    # PPTX files are ZIP archives
    with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
        theme_path = 'word/theme/theme1.xml'
        
        if theme_path in zip_ref.namelist():
            output_path = os.path.join(output_folder, 'theme1.xml')
            with zip_ref.open(theme_path) as theme_file:
                with open(output_path, 'wb') as output_file:
                    output_file.write(theme_file.read())
            print(f"Extracted theme1.xml to: {output_path}")
        else:
            print("theme1.xml not found in the presentation.")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_theme.py <path_to_pptx> [output_folder]")
    else:
        pptx_path = sys.argv[1]
        output_folder = sys.argv[2] if len(sys.argv) > 2 else '.'
        extract_theme_xml(pptx_path, output_folder)


