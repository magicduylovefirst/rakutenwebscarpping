import requests
import json
from urllib.parse import urlparse
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# Constants
URL = "https://item.rakuten.co.jp/waste/cp209/?l-id=shoptop_widget_in_shop_ranking&s-id=shoptop_in_shop_ranking"
API_ENDPOINT = 'https://app.rakuten.co.jp/services/api/IchibaItem/Search/20220601'

def parse_url(url):
    """Parse Rakuten item URL to get shop code and item code"""
    print(f"\nParsing URL: {url}")
    parsed = urlparse(url)
    print(f"Parsed components: {parsed}")
    path_parts = [part for part in parsed.path.split('/') if part]
    print(f"Path parts: {path_parts}")
    if len(path_parts) >= 2:
        return path_parts[0], path_parts[1]
    return None, None

def fetch_item_details(shop_code, item_code, app_id):
    """Fetch item details using Rakuten Ichiba Item Search API"""
    params = {
        'applicationId': app_id,
        'keyword': item_code,
        'shopCode': shop_code,
        'hits': 1,
        'format': 'json'
    }
    
    try:
        print(f"\nFetching data from: {API_ENDPOINT}")
        print(f"Parameters: {params}")
        response = requests.get(API_ENDPOINT, params=params)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        if hasattr(e.response, 'text'):
            print(f"Response text: {e.response.text}")
        return None

def fetch_variations(url):
    """Fetch size variations using Selenium"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--allow-running-insecure-content')
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get(url)
        
        # Wait for the page to load
        time.sleep(5)
        
        # Try to find the size selector
        try:
            # Try different selectors
            selectors = [
                (By.NAME, "size"),
                (By.CSS_SELECTOR, "select[name='size']"),
                (By.CSS_SELECTOR, "select.inventory_choice_name"),
                (By.XPATH, "//select[contains(@name, 'size')]"),
                (By.XPATH, "//select[contains(@class, 'inventory_choice_name')]")
            ]
            
            size_select = None
            for selector_type, selector_value in selectors:
                try:
                    size_select = WebDriverWait(driver, 2).until(
                        EC.presence_of_element_located((selector_type, selector_value))
                    )
                    if size_select:
                        break
                except:
                    continue
            
            if size_select:
                options = size_select.find_elements(By.TAG_NAME, "option")
                sizes = []
                for option in options:
                    size = option.text.strip()
                    if size and size != '未選択':
                        sizes.append(size)
                return sorted(sizes)
            
            # If no size selector found, try looking in the page source
            page_source = driver.page_source
            with open('page_source.html', 'w', encoding='utf-8') as f:
                f.write(page_source)
            print("\nSaved page source to page_source.html")
            
            # Look for size patterns in the source
            import re
            patterns = [
                r'data-size="(2[5-8]\.[05])"',
                r'value="(2[5-8]\.[05])"',
                r'サイズ：(2[5-8]\.[05])',
                r'size=(2[5-8]\.[05])',
                r'size">(2[5-8]\.[05])<'
            ]
            
            all_matches = []
            for pattern in patterns:
                matches = re.findall(pattern, page_source)
                all_matches.extend(matches)
            
            if all_matches:
                return sorted(set(all_matches))
            
            return []
            
        except Exception as e:
            print(f"Error finding size selector: {e}")
            return []
            
    except Exception as e:
        print(f"Error setting up Selenium: {e}")
        return []
    finally:
        if 'driver' in locals():
            driver.quit()

def main():
    # Get Application ID from environment variable
    app_id = os.environ.get('RAKUTEN_APP_ID')
    if not app_id:
        print("Error: RAKUTEN_APP_ID environment variable not set")
        return
        
    # Parse URL and get item details
    shop_code, item_code = parse_url(URL)
    if not shop_code or not item_code:
        print(f"Could not parse Rakuten URL: {URL}")
        return
    
    print(f"\nShop Code: {shop_code}")
    print(f"Item Code: {item_code}")
    
    # Fetch item data
    data = fetch_item_details(shop_code, item_code, app_id)
    if not data or 'Items' not in data or not data['Items']:
        print("No item data found")
        return
        
    item = data['Items'][0]['Item']
    
    print("\n=== Item Information ===")
    print(f"Name: {item.get('itemName', 'N/A')}")
    print(f"Price: ¥{item.get('itemPrice', 'N/A'):,}")
    
    # Get variations from web scraping
    print("\n=== Size Variations ===")
    variations = fetch_variations(URL)
    
    if variations:
        for size in variations:
            print(f"Size: {size}")
            variation_url = f"https://item.rakuten.co.jp/{shop_code}/{item_code}/?size={size}"
            print(f"URL: {variation_url}\n")
    else:
        print("No size variations found")

if __name__ == "__main__":
    main() 