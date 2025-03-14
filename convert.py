import openpyxl
import os
from PIL import Image, ImageDraw, ImageFont
import traceback
import logging

# Paths
angel_file_path = "angel.xlsx"
reference_file_path = "reference.xlsx"
photos_folder = "photos"  # All files (base images and patches) are in this single folder

def read_excel(file_path):
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        log_error(f"Error: Excel file '{file_path}' not found.")
        return [], []

    sheet = wb.active
    skus = []
    qtys = []

    for sku_cell, qty_cell in sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        if sku_cell and qty_cell:
            sku_cell += '.png'  # Append .png to SKU to match file names
            skus.append(sku_cell)  # Add SKU file name to the list
            qtys.append(int(qty_cell))  # Add corresponding quantity to the list
    
    return skus, qtys

def read_reference(file_path):
    """
    Reads the reference file (reference.xlsx) and creates a dictionary mapping SKU numbers to patches.
    """
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        log_error(f"Error: Reference file '{file_path}' not found.")
        return {}

    sheet = wb.active
    reference_data = {}

    for sku_number, patch_name in sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        if sku_number:
            if sku_number not in reference_data:
                reference_data[sku_number] = []

            if patch_name:
                reference_data[sku_number].append(patch_name)

    return reference_data

def open_images(skus, qtys, reference_data):
    for sku, qty in zip(skus, qtys):
        base_sku = os.path.basename(sku).replace('.png', '')  # Get the SKU number without extension
        associated_patches = reference_data.get(base_sku, [])  # Get any patches for the SKU
        images_to_open = [sku] + associated_patches  # Base image + associated patches

        for image_path in images_to_open:
            if not os.path.exists(os.path.join(photos_folder, image_path)):  # Check if file exists in the folder
                log_error(f"The file {image_path} does not exist.")
                continue
            
            try:
                img = Image.open(os.path.join(photos_folder, image_path))  # Open the image from the folder
                draw = ImageDraw.Draw(img)

                try:
                    font_path = "arial.ttf"  # Path to the .ttf font file
                    font = ImageFont.truetype(font_path, 36)
                except IOError:
                    font = ImageFont.load_default()

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
    logging.basicConfig(filename="error_log.txt", level=logging.ERROR)
    logging.error(f"{message}\n{traceback.format_exc()}")

def main():
    try:
        skus, qtys = read_excel(angel_file_path)
        if skus:
            reference_data = read_reference(reference_file_path)
            open_images(skus, qtys, reference_data)
            print("Processing complete!")
        else:
            print("No data found in the selected Excel file.")
    except Exception as e:
        log_error(f"Unexpected error: {e}")
        print("An unexpected error occurred. Check the log file for details.")

if __name__ == "__main__":
    main()
