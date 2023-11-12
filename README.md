# Excel Column A cell data to  desired sized Images

Convert Excel data to beautifully centered PDFs and create cropped images with ease!

## 🚀 Quick Start

1. **Clone the repository:**

   ```bash
https://github.com/aksrivastava/exceldatatoimage

2. Navigate to the project folder:

cd exceldatatoimage


3. Install dependencies:
   
   pip install -r requirements.txt


# Place your Excel file (excel.xlsx) in the same directory as the script.
python excel.py

Run the script:

# 🌈 Output
The script will generate PDFs and PNG images in the output_images/ directory. Cropped images will be saved in the cropped_images/ directory.

# ⚙️ Customization
You can easily customize the script by adjusting the following parameters in the script:
Excel File:

excel_file = "excel.xlsx"

Column Name:
column_name = "A"  # Change this to the column letter that contains your data


Output Folder:
output_folder = "output_images/"


# Cropping Parameters:

top_crop_cm = 9  # Crop 5 cm from the top
bottom_crop_cm = 9  # Crop 5 cm from the bottom
dpi = 96  # Dots per inch










