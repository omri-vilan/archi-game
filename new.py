import os
import re
import pandas as pd
from bs4 import BeautifulSoup

# Define paths
html_file = "data.htm"
excel_file = "data.xlsx"
output_image_dir = "images_new"

os.makedirs(output_image_dir, exist_ok=True)

if not os.path.exists(html_file):
    print(f"❌ Error: Could not find '{html_file}'. Make sure you saved your spreadsheet as a Web Page.")
    exit()

print("📖 Reading and parsing HTML map file...")
with open(html_file, "r", encoding="utf-8", errors="ignore") as f:
    soup = BeautifulSoup(f.read(), "html.parser")

# Load your clean text sheet rows
df = pd.read_excel(excel_file)
df.columns = df.columns.astype(str).str.strip().str.lower()

photo_cols = ['photo 1', 'photo 2', 'photo 3', 'photo 4']
for col in photo_cols:
    df[col] = ""

# Locate the table rows inside the generated webpage
html_rows = soup.find_all("tr")

# Filter out rows to find where actual spreadsheet data starts
data_rows = []
for tr in html_rows:
    # Excel rows containing true data typically match standard classes or layout styles
    cells = tr.find_all("td")
    if len(cells) >= 4:
        data_rows.append(cells)

# Clean and iterate through rows to map images directly to structural positions
mapped_records = []
excel_row_index = 0

for cells in data_rows:
    # Skip header lines by checking if the first cell matches your column names
    first_cell_text = cells[0].get_text(strip=True).lower()
    if "building name" in first_cell_text or excel_row_index >= len(df):
        continue
        
    row_data = df.iloc[excel_row_index]
    unified_option = f"{row_data.get('building name', '')} / {row_data.get('architect name', '')} / {row_data.get('year', '')}"
    
    # Photo columns start at index 4 (Column E)
    for i, col_name in enumerate(photo_cols):
        cell_position = 4 + i
        if cell_position < len(cells):
            target_cell = cells[cell_position]
            img_tag = target_cell.find("img")
            
            if img_tag and img_tag.get("src"):
                src_path = img_tag["src"]
                
                # Double check that the file referenced exists locally
                if os.path.exists(src_path):
                    # Form a pristine path reference
                    record = {
                        'unified_option': unified_option,
                        'photo_path': src_path
                    }
                    mapped_records.append(record)
                    
    excel_row_index += 1

# Convert our structural data to a clean game pool dataframe
game_pool_df = pd.DataFrame(mapped_records)

# Save this robust pool out to a hidden CSV so the game app can load it instantly
game_pool_df.to_csv("verified_game_pool.csv", index=False)
print(f"🎉 Success! Mapped exactly {len(game_pool_df)} images using the underlying HTML layout indices.")