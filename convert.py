import openpyxl
import os
from PIL import Image, ImageDraw, ImageFont
import traceback
import logging

# Paths
angel_file_path = r"C:/Users/jrsea/Desktop/Custom Software/Project Angel/angel.xlsx"
reference_file_path = r"C:/Users/jrsea/Desktop/Custom Software/Project Angel/reference.xlsx"
photos_folder = r"C:/Users/jrsea/Desktop/Custom Software/Project Angel/TS"

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
            sku_cell += '.png'
            sku_path = os.path.join(photos_folder, sku_cell)
            skus.append(sku_path)
            qtys.append(int(qty_cell))
    
    return skus, qtys

def read_reference(file_path):
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        log_error(f"Error: Reference file '{file_path}' not found.")
        return {}

    sheet = wb.active
    reference_data = {}
    for sku_cell, patch_cell in sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        if sku_cell:
            patches = []
            if patch_cell:
                patches.append(os.path.join(photos_folder, patch_cell + '.png'))
            reference_data[sku_cell] = patches

    return reference_data

def open_images(skus, qtys, reference_data):
    for sku, qty in zip(skus, qtys):
        base_sku = os.path.basename(sku).replace('.png', '')
        associated_patches = reference_data.get(base_sku, [])
        images_to_open = [sku] + associated_patches

        for image_path in images_to_open:
            if not os.path.exists(image_path):
                log_error(f"The file {image_path} does not exist.")
                continue
            
            try:
                img = Image.open(image_path)
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
                
                img.show()
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