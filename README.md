# Excel to PDF to Cropped Images

This script reads data from an Excel file, generates PDFs with the data centered on each page, converts the PDFs to PNG images, and finally crops the top and bottom portions of each image.

## Requirements

- Python 3.x
- Required Python packages: `openpyxl`, `reportlab`, `PyMuPDF` (via `fitz`), `Pillow`

Install the required packages using:

```bash
pip install openpyxl reportlab PyMuPDF Pillow
