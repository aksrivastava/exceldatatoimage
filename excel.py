import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os
import fitz  # PyMuPDF
from PIL import Image

# Load the Excel file
excel_file = "excel.xlsx"
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active  # Assuming you're working with the active sheet

# Specify the column containing the data
column_name = "A"  # Change this to the column letter that contains your data

# Create an output folder to save the images
output_folder = "output_images/"

# Ensure the output folder exists, or create it
os.makedirs(output_folder, exist_ok=True)

# Initialize a counter for the filenames
file_counter = 1

# Page size
page_width, page_height = letter

# Iterate through the cells in the specified column
for cell in sheet[column_name]:
    cell_value = str(cell.value)

    # Create a PDF
    pdf_filename = f"{output_folder}s{str(file_counter).zfill(3)}.pdf"
    c = canvas.Canvas(pdf_filename, pagesize=letter)
    
    # Set the font size and style (bold)
    font_size = 40
    c.setFont("Helvetica-Bold", font_size)
    
    # Calculate text width and height for the current font size
    text_width = c.stringWidth(cell_value, "Helvetica-Bold", font_size)
    text_height = font_size
    
    # Calculate text position to center it within the page
    x = (page_width - text_width) / 2
    y = (page_height - text_height) / 2
    
    c.drawString(x, y, cell_value)
    c.save()

    # Convert the PDF to an image using PyMuPDF
    pdf_document = fitz.open(pdf_filename)
    first_page = pdf_document.load_page(0)
    pixmap = first_page.get_pixmap()
    image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
    image_filename = f"{output_folder}S{str(file_counter).zfill(3)}.png"
    image.save(image_filename)

    # Increment the counter for the next filename
    file_counter += 1

print("Images generated and saved.")

# Directory containing the images
input_folder = "output_images/"

# Create an output folder to save the cropped images
output_folder = "cropped_images/"
os.makedirs(output_folder, exist_ok=True)

# List all image files in the input folder
image_files = [f for f in os.listdir(input_folder) if f.endswith(".png")]

from PIL import Image
import os

# Create an output folder to save the cropped images
output_folder = "cropped_images/"
os.makedirs(output_folder, exist_ok=True)

# Define the cropping parameters in pixels
top_crop_cm = 9  # Crop 5 cm from the top
bottom_crop_cm = 9  # Crop 5 cm from the bottom
dpi = 96  # Dots per inch

for image_file in image_files:
    image_path = os.path.join(input_folder, image_file)
    image = Image.open(image_path)

    # Get the image dimensions
    width, height = image.size

    # Calculate the cropping parameters based on cm to pixels and DPI
    top_crop = top_crop_cm * dpi / 2.54  # 1 inch = 2.54 cm
    bottom_crop = bottom_crop_cm * dpi / 2.54

    # Crop the image
    cropped_image = image.crop((0, top_crop, width, height - bottom_crop))

    # Save the cropped image
    cropped_image.save(os.path.join(output_folder, image_file))

print("Images cropped and saved.")
