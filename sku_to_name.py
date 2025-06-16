import requests
import json
import time
import re
import pandas as pd
from typing import List, Dict
from datetime import datetime

# Constants
API_ENDPOINT = 'https://app.rakuten.co.jp/services/api/IchibaItem/Search/20220601'
APP_ID = "1085274442429696242"  # Rakuten Application ID

# Define shop configurations
SHOPS = {
    'waste': {  # e-life＆work shop
        'name': 'e-life＆work shop',
        'shop_code': 'waste',
        'base_url': 'https://item.rakuten.co.jp/waste/',
        'description': 'Main shop for safety equipment and work gear',
        'sku_format': lambda sku: sku.split('-')[1]  # Extract model number (e.g., 1271a029)
    },
    'kougushop': {  # 工具ショップ
        'name': '工具ショップ',
        'shop_code': 'kougushop',
        'base_url': 'https://item.rakuten.co.jp/kougushop/',
        'description': 'Specialized in tools and safety equipment',
        'sku_format': lambda sku: f"{sku.split('-')[1]}-{sku.split('-')[2]}"  # Format: 1271a029-025
    }
}

def truncate_str(s: str, length: int) -> str:
    """Truncate string to specified length, adding ... if truncated"""
    if len(s) <= length:
        return s
    return s[:length-3] + "..."

def clean_text(text):
    """Clean text by removing extra whitespace and normalizing newlines"""
    if not text:
        return ""
    # Replace multiple spaces with single space
    text = re.sub(r'\s+', ' ', text)
    # Replace multiple newlines with single newline
    text = re.sub(r'\n+', '\n', text)
    return text.strip()

def extract_sizes(caption):
    """Extract size information from item caption"""
    sizes = []
    if not caption:
        return sizes
        
    # Look for size patterns like "25.5cm" or "25.5 cm"
    size_pattern = r'(\d{2}\.?\d?)\s*cm'
    matches = re.finditer(size_pattern, caption)
    for match in matches:
        size = match.group(1)
        if 20 <= float(size) <= 31:  # Reasonable shoe size range
            sizes.append(f"{size}cm")
    return sorted(list(set(sizes)))  # Remove duplicates and sort

def extract_specs(caption):
    """Extract specifications from item caption"""
    specs = {}
    if not caption:
        return specs
    
    # Common Japanese spec patterns
    known_specs = {
        '幅/ラスト': r'幅[/／]?ラスト[：:]?\s*([^。\n]+)',
        'アッパー素材': r'アッパー素材[：:]?\s*([^。\n]+)',
        'アウター素材': r'アウター素材[：:]?\s*([^。\n]+)',
        'インナーソール': r'インナーソール[：:]?\s*([^。\n]+)',
        '品番': r'品番\s*[：:]\s*([^。\n]+)',
        'サイズ': r'サイズ[：:]?\s*([^。\n]+)',
        '重量': r'重量[：:]?\s*([^。\n]+)',
        '生産国': r'Made in ([^。\n]+)',
    }
    
    # Extract each specification
    for key, pattern in known_specs.items():
        match = re.search(pattern, caption)
        if match:
            value = match.group(1).strip()
            # Clean up the value
            value = re.sub(r'\s+', ' ', value)  # Normalize spaces
            value = re.sub(r'[：:]\s*', ': ', value)  # Normalize colons
            specs[key] = value
    
    return specs

def format_sku_for_shop(sku: str, shop_info: dict) -> str:
    """Format SKU based on shop's pattern"""
    try:
        return shop_info['sku_format'](sku)
    except Exception as e:
        print(f"Error formatting SKU {sku} for {shop_info['name']}: {e}")
        return sku

def fetch_item_details(code: str) -> dict:
    """Fetch detailed item information from multiple shops"""
    all_items = []
    
    for shop_code, shop_info in SHOPS.items():
        search_code = format_sku_for_shop(code, shop_info)
        
        params = {
            'applicationId': APP_ID,
            'shopCode': shop_code,
            'keyword': search_code,
            'hits': 10,
            'format': 'json',
            'availability': 1  # Include stock status
        }
        
        try:
            print(f"Searching in {shop_info['name']} with code: {search_code}")
            response = requests.get(API_ENDPOINT, params=params)
            response.raise_for_status()
            data = response.json()
            
            if data.get('Items'):
                for item_data in data['Items']:
                    item = item_data['Item']
                    
                    # Calculate values
                    price = item.get('itemPrice', 0)
                    points = item.get('points', 0)
                    tax_rate = 1.1  # 10% tax rate
                    
                    details = {
                        'original_sku': code,
                        'shop_name': shop_info['name'],
                        'shop_code': shop_code,
                        'search_code_used': search_code,
                        'product_info': {
                            '商品管理番号': item.get('itemCode', 'N/A').replace(f"{shop_code}:", ''),
                            '商品名': item.get('itemName', 'N/A'),
                            '検索条件': code,
                            '検索除外': '-',
                            '在庫': '○' if item.get('availability', 0) == 1 else '×',
                            '定価': '-',
                            '仕入金額': '-',
                            '平均単価': '-',
                            'FA売価(税抜)': int(price / tax_rate) if price else 0,
                            '粗利': '-',
                            'RT後の利益': '-',
                            'FA売価(税込)': price
                        },
                        'shop_info': {
                            '価格': price,
                            'ポイント': points,
                            'クーポン': item.get('couponPrice', 0),
                            '在庫状況': '在庫あり' if item.get('availability', 0) == 1 else '在庫なし',
                            'URL': item.get('itemUrl', '')
                        }
                    }
                    all_items.append(details)
            else:
                print(f"No items found in {shop_info['name']} for code: {search_code}")
                    
        except Exception as e:
            print(f"Error fetching data for {code} from {shop_info['name']}: {e}")
    
    return {'items': all_items}

def read_skus_from_excel(excel_path):
    """Read SKUs from first column of Excel file"""
    try:
        df = pd.read_excel(excel_path)
        # Clean SKUs - remove header row if it exists
        skus = df.iloc[:, 0].dropna().astype(str).tolist()
        return [sku for sku in skus if sku.lower() != 'skuコード']
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def print_table_header(columns: List[str], widths: List[int]) -> None:
    """Print table header with proper formatting"""
    # Print top border
    print("+" + "+".join("-" * (w + 2) for w in widths) + "+")
    
    # Print column headers
    row = "|"
    for col, width in zip(columns, widths):
        row += f" {col.center(width)} |"
    print(row)
    
    # Print separator
    print("+" + "+".join("=" * (w + 2) for w in widths) + "+")

def print_table_row(data: List[str], widths: List[int]) -> None:
    """Print a table row with proper formatting"""
    # Split long data into multiple rows if needed
    max_lines = max(len(str(d).split('\n')) for d in data)
    
    for line_num in range(max_lines):
        row = "|"
        for value, width in zip(data, widths):
            value_str = str(value).split('\n')[line_num] if line_num < len(str(value).split('\n')) else ''
            value_str = truncate_str(value_str, width)
            
            # Right-align numbers (including those with ¥), left-align text
            if value_str.startswith("¥") or value_str.replace(",", "").isdigit():
                row += f" {value_str.rjust(width)} |"
            else:
                row += f" {value_str.ljust(width)} |"
        print(row)
    
    # Print bottom border
    print("+" + "+".join("-" * (w + 2) for w in widths) + "+")

def main():
    start_time = time.time()
    
    # Read SKUs from Excel
    excel_path = "New folder/araki.xlsx"
    skus = read_skus_from_excel(excel_path)
    
    if not skus:
        print("No SKUs found in Excel file")
        return
        
    print(f"Processing {len(skus)} SKUs...")
    print(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("\nSearching in the following shops:")
    for shop_code, shop_info in SHOPS.items():
        print(f"- {shop_info['name']} ({shop_code})")
        print(f"  URL: {shop_info['base_url']}")
        print(f"  Description: {shop_info['description']}\n")
    
    # Process all SKUs
    all_results = {
        'items': [],
        'shops_searched': {
            shop_code: {
                'name': info['name'],
                'base_url': info['base_url'],
                'description': info['description']
            } for shop_code, info in SHOPS.items()
        }
    }
    processed_count = 0
    items_per_shop = {shop_code: 0 for shop_code in SHOPS.keys()}
    
    for sku in skus:
        start_item_time = time.time()
        print(f"\nFetching data for SKU: {sku}")
        result = fetch_item_details(sku)
        
        # Count items per shop
        for item in result['items']:
            shop_name = item['shop_name']
            shop_code = next((code for code, info in SHOPS.items() if info['name'] == shop_name), None)
            if shop_code:
                items_per_shop[shop_code] += 1
        
        all_results['items'].extend(result['items'])
        processed_count += len(result['items'])
        
        # Calculate and show progress
        elapsed_item_time = time.time() - start_item_time
        print(f"Time for this SKU: {elapsed_item_time:.2f} seconds")
        time.sleep(1)  # Be nice to Rakuten's servers
    
    # Calculate total time
    total_time = time.time() - start_time
    
    # Add timing info to results
    all_results['metadata'] = {
        'total_time_seconds': round(total_time, 2),
        'average_time_per_sku': round(total_time / len(skus), 2) if skus else 0,
        'total_skus_processed': len(skus),
        'total_items_found': processed_count,
        'items_per_shop': items_per_shop,
        'start_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    # Save results to JSON file
    with open('results.json', 'w', encoding='utf-8') as f:
        json.dump(all_results, ensure_ascii=False, indent=2, fp=f)
    
    print(f"\nProcessing completed:")
    print(f"Total time: {total_time:.2f} seconds")
    print(f"Average time per SKU: {total_time / len(skus):.2f} seconds")
    print(f"Total SKUs processed: {len(skus)}")
    print(f"Items found per shop:")
    for shop_code, count in items_per_shop.items():
        print(f"- {SHOPS[shop_code]['name']}: {count} items")
    print(f"Results saved to results.json")

if __name__ == "__main__":
    main() 