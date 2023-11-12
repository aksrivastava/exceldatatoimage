# Excel to PDF to Cropped Images

This Python script reads data from an Excel file, generates PDFs with the data centered on each page, converts the PDFs to PNG images, and finally crops the top and bottom portions of each image.

## Requirements

- Python 3.x
- Required Python packages: `openpyxl`, `reportlab`, `PyMuPDF` (via `fitz`), `Pillow`

Install the required packages using:

```bash


Usage
Place your Excel file (excel.xlsx) in the same directory as the script.
Run the script using:
pip install openpyxl reportlab PyMuPDF Pillow
Python excel.py

Output
The script will generate PDFs and PNG images in the output_images/ directory. Cropped images will be saved in the cropped_images/ directory.
excel_file = "excel.xlsx"

Column Name: The column containing the data in the Excel file.
column_name = "A"  # Change this to the column letter that contains your data

Output Folder: The directory to save generated images and cropped images.
output_folder = "output_images/"

Cropping Parameters: Adjust these values based on your requirements.


top_crop_cm = 9  # Crop 5 cm from the top
bottom_crop_cm = 9  # Crop 5 cm from the bottom
dpi = 96  # Dots per inch
