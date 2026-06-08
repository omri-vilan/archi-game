import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import streamlit as st

# Setup directories
os.makedirs("images", exist_ok=True)

# Load workbook using openpyxl to access drawing layer
wb = load_workbook("data.xlsx")
breakpoint()  # 👈 This pauses your code right here!
ws = wb.active

# Load the normal text data
df = pd.read_excel("data.xlsx")
df.columns = df.columns.astype(str).str.strip().str.lower()

# Clear out whatever is currently in the photo columns so we can fill them with paths
photo_cols = ['photo 1', 'photo 2', 'photo 3', 'photo 4']
for col in photo_cols:
    df[col] = ""

# Track columns by index (A=1, B=2, etc.)
# Based on your structure:
# Column 1 (A): building name, 2 (B): architect name, 3 (C): location, 4 (D): year
# Column 5 (E): photo 1, 6 (F): photo 2, 7 (G): photo 3, 8 (H): photo 4
col_mapping = {5: 'photo 1', 6: 'photo 2', 7: 'photo 3', 8: 'photo 4'}

print("Extracting embedded images...")

image_count = 0
# Iterate through floating images on the spreadsheet drawing layer
for image in ws._images:
    print(image)
    # Get the cell coordinate where the top-left corner of the image sits
    row = image.anchor.row  # 0-indexed
    col = image.anchor.column  # 0-indexed (e.g., Column E is 4)
    
    # openpyxl coordinates are 0-indexed, but mapping uses 1-indexed Excel notation
    excel_col_num = col + 1
    excel_row_num = row + 1 # row 1 is header, row 2 is data index 0
    
    if excel_col_num in col_mapping and excel_row_num > 1:
        target_col_name = col_mapping[excel_col_num]
        df_row_idx = excel_row_num - 2  # Convert back to pandas dataframe 0-indexed row
        
        if df_row_idx < len(df):
            building_name = str(df.iloc[df_row_idx]['building name']).replace('/', '_').strip()
            
            # Form a unique filename for the image
            img_filename = f"images/{building_name}_{target_col_name.replace(' ', '')}.png"
            
            # Save raw image data to disk
            with open(img_filename, "wb") as f:
                f.write(image._data())
                
            # Write the path string into our pandas dataframe
            df.at[df_row_idx, target_col_name] = img_filename
            image_count += 1

# Save the newly linked data back over your spreadsheet
df.to_excel("data_new.xlsx", index=False)
print(f"🎉 Success! Extracted {image_count} images and updated architecture_game.xlsx with local file paths.")