import requests
import json
from urllib.parse import urlparse
import os
import pandas as pd
import time

# Constants
API_ENDPOINT = 'https://app.rakuten.co.jp/services/api/IchibaItem/Search/20220601'
APP_ID = '1085274442429696242'
AFFILIATE_ID = '494cbb7b.da2d8105.494cbb7c.3832fe1e'

def read_skus_from_excel(excel_path='araki.xlsx'):
    """Read SKUs from Excel file"""
    engines = ['openpyxl', 'xlrd', 'odf']
    for engine in engines:
        try:
            print(f"Trying with {engine} engine...")
            df = pd.read_excel(excel_path, engine=engine)
            # Get SKUs from column B (index 1)
            skus = df.iloc[:, 1].dropna().astype(str).tolist()
            valid_skus = [(idx + 2, sku) for idx, sku in enumerate(skus) if sku.strip()]
            print(f"Found {len(valid_skus)} SKUs in Excel")
            return valid_skus, df
        except Exception as e:
            print(f"Error with {engine} engine: {e}")
            continue
    return [], None

def fetch_item_details(item_code):
    """Fetch item details using Rakuten Ichiba Item Search API"""
    params = {
        'applicationId': APP_ID,
        'affiliateId': AFFILIATE_ID,
        'keyword': item_code,
        'hits': 1,
        'format': 'json'
    }
    
    try:
        print(f"\nFetching data for SKU: {item_code}")
        response = requests.get(API_ENDPOINT, params=params)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        if hasattr(e.response, 'text'):
            print(f"Response text: {e.response.text}")
        return None

def update_excel(df, updates, excel_path='araki.xlsx'):
    """Update Excel file with new data"""
    try:
        # Apply updates to dataframe
        for row_idx, item_name, price in updates:
            df.iloc[row_idx-2, 2] = item_name  # Column C (index 2)
            df.iloc[row_idx-2, 9] = price      # Column J (index 9)
        
        # Save updated dataframe to Excel
        df.to_excel(excel_path, index=False)
        print(f"\nSuccessfully updated Excel file: {excel_path}")
    except Exception as e:
        print(f"Error updating Excel file: {e}")

def main():
    # Read SKUs from Excel
    skus, df = read_skus_from_excel()
    if not skus or df is None:
        print("No SKUs found in Excel file")
        return
    
    updates = []
    # Process each SKU
    for row, sku in skus:
        # Fetch item data
        data = fetch_item_details(sku)
        if not data or 'Items' not in data or not data['Items']:
            print(f"No data found for SKU: {sku}")
            continue
            
        item = data['Items'][0]['Item']
        
        # Store update
        updates.append((
            row,
            item.get('itemName', 'N/A'),
            item.get('itemPrice', 'N/A')
        ))
        print(f"Got data for row {row}: {item.get('itemName', 'N/A')}, Price: {item.get('itemPrice', 'N/A')}")
        
        # Add delay to avoid API rate limits
        time.sleep(1)
    
    # Update Excel file with all changes at once
    if updates:
        update_excel(df, updates)

if __name__ == "__main__":
    main()