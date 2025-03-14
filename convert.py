import openpyxl
import os
from PIL import Image, ImageDraw, ImageFont
import traceback
import logging

# Paths
angel_file_path = "angel.xlsx"
photos_folder = "photos"  # All files (design images) are in this single folder

def read_excel(file_path):
    """
    Reads the angel.xlsx file and returns a list of SKU numbers and their quantities.
    """
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        log_error(f"Error: Excel file '{file_path}' not found.")
        return [], []

    sheet = wb.active
    skus = []
    qtys = []

    # Read the SKU numbers and quantities from the Excel file
    for sku_cell, qty_cell in sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        if sku_cell and qty_cell:
            skus.append(str(sku_cell))  # Use SKU exactly as it appears, ensure it's a string
            qtys.append(int(qty_cell))  # Add corresponding quantity to the list
    
    return skus, qtys

def open_images(skus, qtys):
    """
    Opens the image files corresponding to each SKU number and displays them with the quantity.
    """
    for sku, qty in zip(skus, qtys):
        image_path = os.path.join(photos_folder, f"{sku}.png")  # Create the path to the image file (e.g., `1.png`)

        if not os.path.exists(image_path):
            log_error(f"The file {image_path} does not exist.")
            continue
        
        try:
            img = Image.open(image_path)  # Open the image file
            draw = ImageDraw.Draw(img)

            try:
                font_path = "arial.ttf"  # Path to the .ttf font file (you can replace with the correct path)
                font = ImageFont.truetype(font_path, 36)
            except IOError:
                font = ImageFont.load_default()  # Fallback to default font if Arial is not found

            text = f"Qty: {qty}"

            # Use textbbox for Pillow versions >= 8.0.0
            text_bbox = draw.textbbox((0, 0), text, font=font)
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]

            width, height = img.size
            x = width - text_width - 10
            y = height - text_height - 10

            draw.text((x, y), text, font=font, fill=(255, 255, 255))

            img.show()  # Display the image
            print(f"Displayed {image_path} with quantity {qty}")

        except Exception as e:
            log_error(f"Error opening {image_path}: {e}")

def log_error(message):
    """
    Logs error messages to a file.
    """
    logging.basicConfig(filename="error_log.txt", level=logging.ERROR)
    logging.error(f"{message}\n{traceback.format_exc()}")

def main():
    """
    Main function that processes the Excel file and displays the images with quantity text.
    """
    try:
        skus, qtys = read_excel(angel_file_path)  # Read the SKUs and quantities from the Excel file
        if skus:
            open_images(skus, qtys)  # Open the images corresponding to each SKU
            print("Processing complete!")
        else:
            print("No data found in the selected Excel file.")
    except Exception as e:
        log_error(f"Unexpected error: {e}")
        print("An unexpected error occurred. Check the log file for details.")

if __name__ == "__main__":
    main()
