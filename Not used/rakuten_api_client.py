import requests
import json
import base64
from urllib.parse import urlparse
import sys
import os

# Constants for Rakuten RMS API
RAKUTEN_API_ITEM_ENDPOINT_TEMPLATE = 'https://api.rms.rakuten.co.jp/es/2.0/items/manage-numbers/{}'

def get_rakuten_api_auth(service_secret=None, license_key=None):
    """
    Creates authorization header using serviceSecret and licenseKey.
    Checks environment variables if not provided.
    """
    # Try environment variables if not provided
    service_secret = service_secret or os.environ.get('RAKUTEN_SERVICE_SECRET')
    license_key = license_key or os.environ.get('RAKUTEN_LICENSE_KEY')
    
    if not service_secret or not license_key:
        print("Error: Rakuten API credentials not found.")
        print("Please provide them either as arguments or set environment variables:")
        print("  RAKUTEN_SERVICE_SECRET - Your Service Secret")
        print("  RAKUTEN_LICENSE_KEY - Your License Key")
        print("\nYou can find these in RMS: メインメニュー > 拡張サービス一覧 > WEB APIサービス > 利用設定")
        return None
        
    try:
        auth_str = f"{service_secret}:{license_key}"
        base64_auth = base64.b64encode(auth_str.encode()).decode()
        return f"ESA {base64_auth}"
    except Exception as e:
        print(f"Error creating authorization header: {e}")
        return None

def parse_rakuten_url(item_url):
    """
    Parses a Rakuten item URL to extract shop_code and item_code.
    Example: "https://item.rakuten.co.jp/shopname/itemcode/" -> ("shopname", "itemcode")
    """
    try:
        parsed = urlparse(item_url)
        if parsed.hostname != 'item.rakuten.co.jp' and not parsed.hostname.endswith('.item.rakuten.co.jp'):
            print(f"Invalid hostname: {parsed.hostname}")
            return None
        
        # Path is typically /shop_code/item_code/ or /shop_code/item_code
        path_parts = [part for part in parsed.path.split('/') if part]
        
        if len(path_parts) >= 2:
            shop_code = path_parts[0]
            item_code = path_parts[1]
            return shop_code, item_code
        
        print(f"Could not extract shop_code and item_code from path: {parsed.path}")
        return None
    except Exception as e:
        print(f"Error parsing URL '{item_url}': {e}")
        return None

def fetch_rakuten_item_details(manage_number, auth_token):
    """
    Fetches item details from Rakuten API using the manage_number.
    """
    api_url = RAKUTEN_API_ITEM_ENDPOINT_TEMPLATE.format(manage_number)
    headers = {
        'Authorization': auth_token
    }
    
    try:
        print(f"Fetching data from: {api_url}")
        response = requests.get(api_url, headers=headers, timeout=15)
        response.raise_for_status()  # Raise an exception for HTTP errors
        return response.json()
    except requests.exceptions.HTTPError as http_err:
        if response.status_code == 404:
            print(f"Item not found (404) for manage number: {manage_number}")
        else:
            print(f"HTTP error occurred: {http_err} - Status: {response.status_code} - Response: {response.text}")
        return None
    except requests.exceptions.RequestException as req_err:
        print(f"Request error occurred: {req_err}")
        return None
    except json.JSONDecodeError:
        print(f"Error decoding JSON response from API for {manage_number}. Response: {response.text}")
        return None

def main():
    # Get credentials from command line arguments or environment variables
    service_secret = None
    license_key = None
    item_url = None
    
    # Parse command line arguments
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] == '--service-secret':
            i += 1
            if i < len(args):
                service_secret = args[i]
        elif args[i] == '--license-key':
            i += 1
            if i < len(args):
                license_key = args[i]
        elif args[i].startswith('http'):
            item_url = args[i]
        i += 1
    
    # Use default URL if none provided
    if not item_url:
        item_url = "https://item.rakuten.co.jp/waste/cp209/?l-id=shoptop_widget_in_shop_ranking&s-id=shoptop_in_shop_ranking"
        print(f"No URL provided, using default: {item_url}")

    # Get authorization token
    auth_token = get_rakuten_api_auth(service_secret, license_key)
    if not auth_token:
        return
        
    # Parse URL and get item details
    parsed_url_info = parse_rakuten_url(item_url)
    if not parsed_url_info:
        print(f"Could not parse Rakuten URL: {item_url}")
        return
    
    shop_code, item_code = parsed_url_info
    manage_number = f"{shop_code}:{item_code}"
    print(f"Parsed manage_number: {manage_number}")
    
    item_data = fetch_rakuten_item_details(manage_number, auth_token)
    
    if not item_data:
        print(f"Failed to fetch item details for {manage_number}")
        return

    # Extracting information
    item_name = item_data.get('title', 'N/A')
    
    price = 'N/A'
    variants_data = item_data.get('variants')
    if variants_data and isinstance(variants_data, dict):
        # Get price from the first SKU found in the variants dictionary
        first_sku_key = next(iter(variants_data), None)
        if first_sku_key and isinstance(variants_data[first_sku_key], dict):
            price = variants_data[first_sku_key].get('standardPrice', 'N/A')

    size_info_parts = []
    variant_selectors = item_data.get('variantSelectors', [])
    if isinstance(variant_selectors, list):
        for selector in variant_selectors:
            selector_name = selector.get('displayName', '')
            values = selector.get('values', [])
            value_names = [val.get('displayValue', '') for val in values if isinstance(val, dict)]
            if selector_name and value_names:
                size_info_parts.append(f"{selector_name}: {', '.join(value_names)}")
    size_str = '; '.join(size_info_parts) if size_info_parts else 'N/A'

    canonical_url = f"https://item.rakuten.co.jp/{shop_code}/{item_code}/"

    print("\n--- Item Information ---")
    print(f"Name: {item_name}")
    print(f"Price: {price}")
    print(f"Size/Variations: {size_str}")
    print(f"URL: {canonical_url}")

if __name__ == "__main__":
    main() 