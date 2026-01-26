import os
import csv
import argparse
from pptx import Presentation

def count_slides_in_pptx(file_path):
    """Returns number of slides in a PPTX file."""
    try:
        prs = Presentation(file_path)
        return len(prs.slides)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return 0

def scan_directory(directory):
    """Recursively scans directory for PPTX files and returns list of results."""
    results = []  # Each item: (file_path, slide_count)
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(".pptx"):
                full_path = os.path.join(root, file)
                slides = count_slides_in_pptx(full_path)
                print(f"{full_path}: {slides} slides")
                results.append((full_path, slides))
    return results

def write_csv(results, output_file="slide_counts.csv"):
    """Writes results to a CSV file using semicolon delimiter."""
    with open(output_file, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile, delimiter=";")
        writer.writerow(["File Path", "Slide Count"])
        for path, count in results:
            writer.writerow([path, count])
    print(f"\nCSV exported to: {output_file}")

def main():
    parser = argparse.ArgumentParser(description="Count slides in PPTX files.")
    parser.add_argument("-d", "--dir", required=True, help="Directory to scan recursively")
    parser.add_argument("-c", "--csv", default="slide_counts.csv", help="Output CSV filename")
    args = parser.parse_args()

    directory = args.dir
    output_csv = args.csv

    print(f"Scanning directory: {directory}\n")
    results = scan_directory(directory)

    total_slides = sum(count for _, count in results)
    print(f"\nTotal slides across all PPTX files: {total_slides}")

    write_csv(results, output_csv)

if __name__ == "__main__":
    main()

