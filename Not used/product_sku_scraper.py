import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re

# Define shop URLs and their SKU patterns
SHOP_PATTERNS = {
    'e-life＆work shop': {
        'base_url': 'https://item.rakuten.co.jp/waste/',
        'sku_patterns': [
            lambda sku: "cp209/?l-id=shoptop_widget_in_shop_ranking&s-id=shoptop_in_shop_ranking"
        ]
    },
    '工具ショップ': {
        'base_url': 'https://item.rakuten.co.jp/kougushop/',
        'sku_patterns': [
            lambda sku: f"{sku.split('-')[1]}-{sku.split('-')[2]}/",  # 1271a029-025/
            lambda sku: "1271a029-602/",
            lambda sku: "1271a029-400/",
            lambda sku: "1271a029-025/"
        ]
    },
    '晃栄産業　楽天市場店': {
        'base_url': 'https://item.rakuten.co.jp/kouei-sangyou/',
        'sku_patterns': [
            lambda sku: "fcp209/"
        ]
    },
    'Dear worker': {
        'base_url': 'https://item.rakuten.co.jp/dear-worker/',
        'sku_patterns': [
            lambda sku: "cp209boa/",
            lambda sku: "2360345/?variantId=2360345750300"
        ]
    }
}

def read_skus_from_excel(excel_path):
    """Read SKUs from first column of Excel file"""
    try:
        df = pd.read_excel(excel_path)
        # Get first column values, drop any empty/NA values
        skus = df.iloc[:, 0].dropna().astype(str).tolist()
        return skus
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def check_url_exists(url):
    """Check if URL exists and returns product page"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Connection': 'keep-alive',
        }
        response = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
        
        # Check if it's a valid product page
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Look for common elements in Rakuten product pages
            price_element = soup.find('span', class_='price2') or soup.find('span', class_='price')
            title_element = soup.find('span', class_='item_name') or soup.find('h1', class_='item_name')
            
            return bool(price_element or title_element)
            
    except Exception as e:
        print(f"Error checking URL {url}: {e}")
        return False
    
    return False

def search_sku_in_shop(sku, shop_name, shop_info):
    """Try all possible URL patterns for a shop"""
    base_url = shop_info['base_url']
    
    for pattern_func in shop_info['sku_patterns']:
        try:
            # Generate URL using pattern
            product_path = pattern_func(sku)
            full_url = base_url + product_path
            
            print(f"Trying URL for {shop_name}: {full_url}")
            
            if check_url_exists(full_url):
                return {
                    'sku': sku,
                    'shop': shop_name,
                    'url': full_url,
                    'found': True
                }
        except Exception as e:
            print(f"Error with pattern for SKU {sku} in {shop_name}: {e}")
            continue
    
    return {
        'sku': sku,
        'shop': shop_name,
        'url': None,
        'found': False
    }

def search_sku_all_shops(sku):
    """Search SKU across all shops"""
    results = []
    for shop_name, shop_info in SHOP_PATTERNS.items():
        result = search_sku_in_shop(sku, shop_name, shop_info)
        if result['found']:
            results.append(result)
            print(f"Found SKU {sku} in {shop_name}")
    return results

def main():
    excel_path = "New folder/araki.xlsx"
    skus = read_skus_from_excel(excel_path)
    
    if not skus:
        print("No SKUs found in Excel file")
        return
    
    print(f"Found {len(skus)} SKUs")
    print("Starting search across all shops...")
    
    all_results = []
    for sku in skus:
        print(f"\nProcessing SKU: {sku}")
        results = search_sku_all_shops(sku)
        all_results.extend(results)
        time.sleep(1)  # Be nice to Rakuten servers
    
    # Print results
    print("\n=== Search Results ===")
    for result in all_results:
        print(f"\nSKU: {result['sku']}")
        print(f"Found in: {result['shop']}")
        print(f"URL: {result['url']}")
    
    # Print summary
    print(f"\nTotal SKUs processed: {len(skus)}")
    found_skus = len(set(r['sku'] for r in all_results))
    print(f"SKUs found: {found_skus}")
    print(f"SKUs not found: {len(skus) - found_skus}")

if __name__ == "__main__":
    main() 