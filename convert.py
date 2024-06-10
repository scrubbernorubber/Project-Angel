import openpyxl
import os

def read_excel(file_path):
    # Open the workbook
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        print(f"Error: Excel file '{file_path}' not found.")
        return [], []

    # Select the first sheet
    sheet = wb.active
    
    # Read the data from the SKUs and QTYs columns
    skus = []
    qtys = []
    Photos_folder = os.path.join(os.path.dirname(file_path), 'TS')

    for sku_cell, qty_cell in sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        if sku_cell is not None and qty_cell is not None:

            # Append .png extension to the SKU value
            sku_cell += '.png'

            # Construct the full path to the PNG file
            sku_path = os.path.join(Photos_folder, sku_cell)
            print(f"Constructed path: {sku_path}")  # Debug print
            skus.append(sku_path)
            
            # Ensure qty_cell is treated as an integer
            qty_cell = int(qty_cell)
            qtys.append(qty_cell)
    
    return skus, qtys

def open_images(skus, qtys):
    for sku, qty in zip(skus, qtys):
        # Check if the file exists
        if not os.path.exists(sku):
            print(f"The file {sku} does not exist.")  # Debug print
            continue
        
        # Debug print before opening the images
        print(f"Will open {sku} {qty} times.")
        
        # Open the image with the default viewer the specified number of times
        for i in range(qty):
            print(f"Opening {sku} ({i+1}/{qty})")  # Debug print
            os.startfile(sku)

if __name__ == "__main__":
    # Excel path
    excel_path = "angel.xlsx"
    
    # Read SKUs and QTYs from the Excel file
    skus, qtys = read_excel(excel_path)
    
    # Open each image file associated with its quantity
    open_images(skus, qtys)