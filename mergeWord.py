from docx import Document
from docxcompose.composer import Composer
import argparse
import os
import sys

from docx import Document
from docxcompose.composer import Composer
from docx.enum.section import WD_SECTION

def combine_word_documents(input_files, output_file):
    if len(input_files) < 1:
        raise ValueError("At least one input file is required.")

    master = Document(input_files[0])
    composer = Composer(master)

    for file in input_files[1:]:
        # Force a new section before appending
        master.add_section(WD_SECTION.NEW_PAGE)

        doc = Document(file)
        composer.append(doc)

    composer.save(output_file)

def main():
    parser = argparse.ArgumentParser(
        description="Combine multiple Word (.docx) files into one without changing layout or design."
    )

    parser.add_argument(
        "inputs",
        nargs="+",
        help="Input Word (.docx) files (order matters)"
    )

    parser.add_argument(
        "-o", "--output",
        required=True,
        help="Output Word (.docx) file"
    )

    args = parser.parse_args()

    # Basic validation
    for file in args.inputs:
        if not file.lower().endswith(".docx"):
            print(f"Error: '{file}' is not a .docx file.")
            sys.exit(1)
        if not os.path.exists(file):
            print(f"Error: File not found: {file}")
            sys.exit(1)

    if not args.output.lower().endswith(".docx"):
        print("Error: Output file must have a .docx extension.")
        sys.exit(1)

    combine_word_documents(args.inputs, args.output)
    print("Documents combined successfully.")


if __name__ == "__main__":
    main()

