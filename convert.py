import openpyxl
import os
from PIL import Image, ImageDraw, ImageFont
import traceback
import logging
from openpyxl.styles import PatternFill

# Paths
angel_file_path = "angel.xlsx"
photos_folder = "photos"  # All files (design images) are in this single folder

def read_excel(file_path):
    """
    Reads the angel.xlsx file and returns a list of SKU numbers, their quantities, and the workbook.
    """
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        log_error(f"Error: Excel file '{file_path}' not found.")
        return [], [], None

    sheet = wb.active
    skus = []
    qtys = []

    # Read the SKU numbers and quantities from the Excel file
    for sku_cell, qty_cell in sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        if sku_cell and qty_cell:
            skus.append(str(sku_cell))  # Use SKU exactly as it appears, ensure it's a string
            qtys.append(int(qty_cell))  # Add corresponding quantity to the list
    
    return skus, qtys, wb  # Return the workbook so we can save changes

def extract_color_size_combinations(skus):
    """
    Extracts and returns a set of unique color and size combinations from the SKU list.
    The 5th part is the color (WHT, BLK, GRN) and the 6th part is the size (S, M, L, XL).
    """
    color_choices = {"BLK", "GRN", "WHT"}
    size_choices = {"S", "M", "L", "XL"}
    
    available_colors = set()
    available_sizes = set()

    for sku in skus:
        sku_parts = sku.split('-')
        if len(sku_parts) >= 6:
            color = sku_parts[4]  # 5th part: color
            size = sku_parts[5]   # 6th part: size
            
            if color in color_choices:
                available_colors.add(color)
            if size in size_choices:
                available_sizes.add(size)

    return list(available_colors), list(available_sizes)

def open_images_for_combination(skus, qtys, selected_color, selected_size, sheet):
    """
    Opens the image files corresponding to the selected color and size combination and modifies the Excel file.
    """
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    for idx, (sku, qty) in enumerate(zip(skus, qtys)):
        # Extract the 5th and 6th parts of the SKU for color and size
        sku_parts = sku.split('-')
        if len(sku_parts) >= 6:
            color = sku_parts[4]  # 5th part: color
            size = sku_parts[5]   # 6th part: size

            # Check if this SKU matches the selected color and size
            if color == selected_color and size == selected_size:
                # Construct the image file path using the third part of SKU
                image_sku = sku_parts[2]  # The third part of the SKU to use as the filename
                image_path = os.path.join(photos_folder, f"{image_sku}.png")

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

                    # Mark the QTY cell with yellow background after displaying the image
                    sheet.cell(row=idx + 2, column=2).fill = yellow_fill  # Color the QTY cell

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
    Main function that processes the Excel file, allows the user to choose color/size, shows images, and updates Excel.
    """
    try:
        skus, qtys, wb = read_excel(angel_file_path)  # Read the SKUs, quantities, and workbook
        if skus:
            sheet = wb.active  # Get the active sheet from the workbook

            # Get available colors and sizes from SKUs
            available_colors, available_sizes = extract_color_size_combinations(skus)

            # Show available colors
            print("Available colors:")
            for idx, color in enumerate(available_colors, start=1):
                print(f"{idx}. {color}")

            # Get user input for color
            color_choice = int(input(f"Enter the number corresponding to your color choice (1-{len(available_colors)}): "))
            selected_color = available_colors[color_choice - 1]

            # Show available sizes
            print("Available sizes:")
            for idx, size in enumerate(available_sizes, start=1):
                print(f"{idx}. {size}")

            # Get user input for size
            size_choice = int(input(f"Enter the number corresponding to your size choice (1-{len(available_sizes)}): "))
            selected_size = available_sizes[size_choice - 1]

            print(f"Displaying images for Color: {selected_color}, Size: {selected_size}...")

            # Display images for the selected combination and mark QTY cells yellow
            open_images_for_combination(skus, qtys, selected_color, selected_size, sheet)

            # Save the modified Excel file directly to the original file
            wb.save(angel_file_path)  # Save the changes to the same file
            print("Processing complete! Excel file has been updated.")
        else:
            print("No data found in the selected Excel file.")
    except Exception as e:
        log_error(f"Unexpected error: {e}")
        print("An unexpected error occurred. Check the log file for details.")

if __name__ == "__main__":
    main()
