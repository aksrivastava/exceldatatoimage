# Excel to PDF to Cropped Images

This Python script reads data from an Excel file, generates PDFs with the data centered on each page, converts the PDFs to PNG images, and finally crops the top and bottom portions of each image.

## Requirements

- Python 3.x
- Required Python packages: `openpyxl`, `reportlab`, `PyMuPDF` (via `fitz`), `Pillow`

Install the required packages using:

```bash
pip install openpyxl reportlab PyMuPDF Pillow
python excel.py
excel_file = "excel.xlsx"
column_name = "A"  # Change this to the column letter that contains your data
output_folder = "output_images/"
top_crop_cm = 9  # Crop 5 cm from the top
bottom_crop_cm = 9  # Crop 5 cm from the bottom
dpi = 96  # Dots per inch
