import pandas as pd
import os
import requests
import json
import base64
import sys
import argparse

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

def process_excel(service_secret=None, license_key=None):
    # Read Excel file
    try:
        print("Reading Excel file...")
        df = pd.read_excel('araki.xlsx', engine='openpyxl')
        
        # Print column names to verify structure
        print("\nAvailable columns in Excel:")
        for col in df.columns:
            print(f"- {col}")
        
        # Get Rakuten API auth
        auth_token = get_rakuten_api_auth(service_secret, license_key)
        if not auth_token:
            return
        
        # Create output columns based on the headers
        result_df = pd.DataFrame(columns=[
            '商品管理番号商品名', '検索条件 検索除外', '在庫', '定価', '仕入金額',
            '平均単価（税込）', 'FA売価(税込)', '粗利', 'RT後の利益FA売価(税込)',
            '価格', 'ポイント', 'クーポン', '在庫'
        ])
        
        # Look for URLs in the URL columns
        url_columns = ['URL', 'URL.1', 'URL.2']
        processed_urls = set()
        
        for url_col in url_columns:
            if url_col not in df.columns:
                continue
                
            print(f"\nProcessing URLs from column: {url_col}")
            for index, row in df.iterrows():
                url = str(row.get(url_col, ''))
                if not url or pd.isna(url) or url in processed_urls:
                    continue
                    
                processed_urls.add(url)
                print(f"Processing URL: {url}")
                
                # Extract shop and item code from URL
                if 'item.rakuten.co.jp' in url:
                    parts = url.split('/')
                    try:
                        shop_code = parts[parts.index('item.rakuten.co.jp') + 1]
                        item_code = parts[parts.index('item.rakuten.co.jp') + 2].split('?')[0]
                        manage_number = f"{shop_code}:{item_code}"
                        
                        item_data = fetch_rakuten_item_details(manage_number, auth_token)
                        if item_data:
                            # Map the data to our result columns
                            new_row = {
                                '商品管理番号商品名': item_data.get('title', ''),
                                '在庫': item_data.get('variants', {}).get(manage_number, {}).get('inventoryCount', 0),
                                '価格': item_data.get('variants', {}).get(manage_number, {}).get('standardPrice', 0),
                                # Other columns would be filled based on business logic
                            }
                            result_df = pd.concat([result_df, pd.DataFrame([new_row])], ignore_index=True)
                    except (ValueError, IndexError) as e:
                        print(f"Could not parse URL: {url} - Error: {e}")
        
        # Save results
        output_file = 'processed_results.xlsx'
        result_df.to_excel(output_file, index=False, encoding='utf-8')
        print(f"\nProcessing complete. Results saved to '{output_file}'")
        
    except Exception as e:
        print(f"Error processing Excel file: {e}")

def main():
    parser = argparse.ArgumentParser(description='Process Excel file with Rakuten URLs')
    parser.add_argument('--service-secret', help='Rakuten API Service Secret')
    parser.add_argument('--license-key', help='Rakuten API License Key')
    
    args = parser.parse_args()
    process_excel(args.service_secret, args.license_key)

if __name__ == "__main__":
    main() 