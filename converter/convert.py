import pandas as pd
import json
import re

INPUT_EXCEL = "ikea_prices.xlsx"
OUTPUT_FILE = "data.json"

def clean_price(value):
    """
    Handles strings "22.499,00" and Excel numbers 22499.0
    """
    if pd.isna(value):
        return 0.0
    
    # If Excel already sees it as a number, return it
    if isinstance(value, (int, float)):
        return float(value)
    
    # Clean string format
    price_str = str(value)
    # Remove 'TL', dots, spaces
    clean = price_str.replace("TL", "").replace("tl", "").replace(".", "").replace(" ", "").strip()
    # Convert comma to dot
    clean = clean.replace(",", ".")
    # Remove junk
    clean = re.sub(r'[^\d.]', '', clean)
    
    try:
        return float(clean)
    except ValueError:
        return 0.0

def convert_excel():
    print(f"Reading {INPUT_EXCEL}...")
    
    # Read without header assumption
    try:
        df = pd.read_excel(INPUT_EXCEL, header=None)
    except FileNotFoundError:
        print(f"Error: Could not find {INPUT_EXCEL}")
        return

    data_list = []

    for index, row in df.iterrows():
        # 1. Get row as a list
        raw_values = row.values.tolist()
        
        # 2. FILTER: Remove empty columns (NaN, None, or empty strings)
        valid_values = []
        for val in raw_values:
            if pd.isna(val): continue
            if str(val).strip() == "": continue
            valid_values.append(val)
            
        # We need at least 4 items: Barcode, Name, OldPrice, NewPrice
        if len(valid_values) < 4:
            continue
            
        # 3. SMART MAPPING
        # First item is Barcode
        raw_barcode = str(valid_values[0]).strip()
        # Skip header rows if they exist
        if "Article" in raw_barcode or "Barkod" in raw_barcode:
            continue
            
        barcode = raw_barcode.replace(".", "") # Clean barcode
        
        # Last item is ALWAYS New Price
        new_price = clean_price(valid_values[-1])
        
        # Second to last item is ALWAYS Old Price
        old_price = clean_price(valid_values[-2])
        
        # Everything else in the middle is the Name 
        # (This handles if the Name got split into 2 columns too!)
        name_parts = valid_values[1:-2]
        name = " ".join([str(p).strip() for p in name_parts])

        item = {
            "barcode": barcode,
            "name": name,
            "old_price": old_price,
            "new_price": new_price
        }
        data_list.append(item)

    # Save
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data_list, f, ensure_ascii=False, indent=4)
        
    print(f"Success! Converted {len(data_list)} items.")
    print(f"Saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    convert_excel()
