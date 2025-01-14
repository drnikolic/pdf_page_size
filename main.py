import os
import csv
import PyPDF2
import argparse

# Conversion factor from points to millimeters
POINTS_TO_MM = 0.3528

# Set up argument parsing
parser = argparse.ArgumentParser(description="Extract PDF page sizes and save to a CSV file.")
parser.add_argument("folder_path", help="Path to the folder containing PDF files")
parser.add_argument("output_csv", help="Path to the output CSV file")

# Parse the command-line arguments
args = parser.parse_args()
folder_path = args.folder_path
output_csv = args.output_csv

# Open the CSV file for writing
with open(output_csv, "w", newline="", encoding="utf-8") as csvfile:
    csv_writer = csv.writer(csvfile)
    # Write the header
    csv_writer.writerow(["File Name", "Page Number", "Width (mm)", "Height (mm)"])

    # Loop through all PDF files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            
            # Open the PDF file
            with open(file_path, "rb") as pdf_file:
                reader = PyPDF2.PdfReader(pdf_file)
                
                # Iterate through pages
                for page_number, page in enumerate(reader.pages, start=1):
                    width = float(page.mediabox.width)
                    height = float(page.mediabox.height)
                    width_mm = round(width * POINTS_TO_MM)
                    height_mm = round(height * POINTS_TO_MM)

                    # Write the row to CSV
                    csv_writer.writerow([file_name, page_number, width_mm, height_mm])

print(f"Page size information saved to '{output_csv}'")
