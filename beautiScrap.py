import pandas as pd
import requests
from bs4 import BeautifulSoup
import json
import time
from datetime import datetime
import random

def read_skus_from_excel(excel_path):
    """Read SKUs from first column of Excel file"""
    try:
        df = pd.read_excel(excel_path)
        skus = df.iloc[:, 0].dropna().astype(str).tolist()
        return [sku for sku in skus if sku.lower() != 'skuコード']
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def get_headers():
    """Get random user agent headers to avoid blocking"""
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15',
    ]
    return {
        'User-Agent': random.choice(user_agents),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Connection': 'keep-alive',
    }

def get_price_from_soup(soup):
    """Extract price from soup with multiple selector attempts"""
    try:
        # Try all possible price selectors in order
        price_selectors = [
            "div.value--1oSD_.layout-inline--2z490.size-x-large--DyMl5.style-bold500--1X0Xl.color-crimson--2uc0e.align-right--3POGa",
            "span.price--OX_YW",
            "span[data-test='price']",
            "span.price--3xbIC",
            "span.price_color",
            "span.price"
        ]
        
        for selector in price_selectors:
            price_element = soup.select_one(selector)
            if price_element:
                price_text = price_element.text.strip()
                # Remove currency symbols and commas
                price_text = ''.join(filter(str.isdigit, price_text))
                if price_text:
                    return int(price_text)
        return None
    except Exception as e:
        print(f"Error extracting price: {e}")
        return None

def get_points_from_soup(soup):
    """Extract points from soup with multiple selector attempts"""
    try:
        # Try all possible points selectors in order
        points_selectors = [
            "div.point-summary__total___3rYYD span",
            "span.price--point-badge_item",
            "span[data-test='points']",
            "span.points--3A8tR",
            "span.point_value",
            "div.point-summary span"
        ]
        
        for selector in points_selectors:
            points_element = soup.select_one(selector)
            if points_element:
                points_text = points_element.text.strip()
                # Remove non-digit characters and handle percentage
                if '%' in points_text:
                    # Handle percentage points
                    percentage = float(''.join(filter(lambda x: x.isdigit() or x == '.', points_text)))
                    # If we have price, calculate points
                    price = get_price_from_soup(soup)
                    if price:
                        return int(price * (percentage / 100))
                else:
                    # Direct point value
                    points_text = ''.join(filter(str.isdigit, points_text))
                    if points_text:
                        return int(points_text)
        return None
    except Exception as e:
        print(f"Error extracting points: {e}")
        return None

def get_variant_info(session, base_url):
    """Get all variants information including colors and sizes"""
    variants = []
    try:
        response = session.get(base_url, headers=get_headers())
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Get color variants
        color_buttons = soup.select("div.grid-cols-2--1uI00 button.type-sku-button--BJoVv")
        size_buttons = soup.select("div.grid-cols-5--3wKbc button.type-sku-button--BJoVv")
        
        # Extract color and size information
        colors = []
        for btn in color_buttons:
            color_name = btn.get('aria-label')
            is_selected = 'selected--Mg4iu' in btn.get('class', [])
            colors.append({
                'name': color_name,
                'selected': is_selected
            })
            
        sizes = []
        for btn in size_buttons:
            size_name = btn.get('aria-label')
            is_selected = 'selected--Mg4iu' in btn.get('class', [])
            sizes.append({
                'name': size_name,
                'selected': is_selected
            })

        # Generate all possible combinations
        variant_id = 1
        for color in colors:
            for size in sizes:
                variant_url = f"{base_url}&variantId=r-sku{variant_id:08d}"
                variants.append({
                    'color': color['name'],
                    'size': size['name'],
                    'variant_id': f"r-sku{variant_id:08d}",
                    'url': variant_url,
                    '価格': None,
                    'ポイント': None,
                    'クーポン': None,
                    '在庫状況': None
                })
                variant_id += 1

                # Get variant information
                variant_response = session.get(variant_url, headers=get_headers())
                variant_soup = BeautifulSoup(variant_response.text, 'html.parser')

                # Get price and points using new helper functions
                variants[-1]['価格'] = get_price_from_soup(variant_soup)
                variants[-1]['ポイント'] = get_points_from_soup(variant_soup)

                try:
                    # Get coupon
                    coupon_element = variant_soup.select_one("div.coupon")
                    if coupon_element:
                        coupon_text = coupon_element.text
                        coupon_value = int(''.join(filter(str.isdigit, coupon_text)))
                        variants[-1]['クーポン'] = coupon_value
                except:
                    pass

                try:
                    # Check inventory status
                    out_of_stock = variant_soup.find(string='売り切れ')
                    if out_of_stock:
                        variants[-1]['在庫状況'] = '在庫なし'
                    else:
                        in_stock = variant_soup.find(string='在庫あり')
                        if in_stock:
                            variants[-1]['在庫状況'] = '在庫あり'
                except:
                    pass

                # Random delay between requests
                time.sleep(random.uniform(0.5, 1))  # Reduced delay

    except Exception as e:
        print(f"Error getting variants: {e}")
    
    return variants

def get_kougushop_variant_info(session, base_url):
    """Get variants information for kougushop"""
    variants = []
    try:
        # For kougushop, variants are numbered from 8021 to 8034 for sizes 22.5 to 30
        sizes = ['22.5', '23', '23.5', '24', '24.5', '25', '25.5', '26', '26.5', '27', '27.5', '28', '29', '30']
        for idx, size in enumerate(sizes, 8021):
            variant_url = f"{base_url}&variantId={idx}"
            variants.append({
                'size': size,
                'variant_id': str(idx),
                'url': variant_url,
                '価格': None,
                'ポイント': None,
                'クーポン': None,
                '在庫状況': None
            })

            # Get variant information
            variant_response = session.get(variant_url, headers=get_headers())
            variant_soup = BeautifulSoup(variant_response.text, 'html.parser')

            # Get price and points using new helper functions
            variants[-1]['価格'] = get_price_from_soup(variant_soup)
            variants[-1]['ポイント'] = get_points_from_soup(variant_soup)

            try:
                # Get coupon
                coupon_element = variant_soup.select_one("div.coupon")
                if coupon_element:
                    coupon_text = coupon_element.text
                    coupon_value = int(''.join(filter(str.isdigit, coupon_text)))
                    variants[-1]['クーポン'] = coupon_value
                else:
                    variants[-1]['クーポン'] = None
            except:
                pass

            try:
                # Check inventory status
                out_of_stock = variant_soup.find(string='売り切れ')
                if out_of_stock:
                    variants[-1]['在庫状況'] = '在庫なし'
                else:
                    in_stock = variant_soup.find(string='在庫あり')
                    if in_stock:
                        variants[-1]['在庫状況'] = '在庫あり'
            except:
                pass

            # Random delay between requests
            time.sleep(random.uniform(0.5, 1))  # Reduced delay

    except Exception as e:
        print(f"Error getting kougushop variants: {e}")
    
    return variants

def scrape_product_info(session, url, is_waste_shop=False, is_kougushop=False, is_kouei_shop=False, is_dear_worker=False):
    """Scrape product information from a given URL"""
    try:
        response = session.get(url, headers=get_headers())
        time.sleep(random.uniform(1, 2))  # Random delay

        # Initialize product info dictionary
        product_info = {
            'url': url,
            'variants': [] if (is_waste_shop or is_kougushop or is_kouei_shop or is_dear_worker) else None
        }

        # Get variants based on shop type
        if is_waste_shop or is_kougushop or is_kouei_shop or is_dear_worker:
            base_url = url.split('&variantId=')[0] if '&variantId=' in url else url
            if is_waste_shop or is_kouei_shop or is_dear_worker:
                product_info['variants'] = get_variant_info(session, base_url)
            elif is_kougushop:
                product_info['variants'] = get_kougushop_variant_info(session, base_url)

        return product_info

    except Exception as e:
        print(f"Error scraping URL {url}: {e}")
        return None

def main():
    start_time = time.time()
    
    # Read SKUs from Excel
    excel_path = "New folder/araki.xlsx"
    skus = read_skus_from_excel(excel_path)
    
    if not skus:
        print("No SKUs found in Excel file")
        return

    print(f"Found {len(skus)} SKUs to process")

    # Create new results structure
    results = {
        'items': [],
        'metadata': {
            'total_skus': len(skus),
            'start_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
    }

    # Create a session for connection pooling
    session = requests.Session()

    try:
        # Process each SKU
        for idx, sku in enumerate(skus, 1):
            print(f"\nProcessing SKU {idx}/{len(skus)}: {sku}")
            sku_start_time = time.time()
            
            # Initialize item structure
            item = {
                'original_sku': sku,
                'search_code_used': sku.split('-')[1] if '-' in sku else sku,
                'product_info': {
                    '商品管理番号': None,
                    '商品名': 'N/A',
                    '検索条件': sku,
                    '検索除外': '-',
                    '在庫': None,
                    '定価': '-',
                    '仕入金額': '-',
                    '平均単価': '-',
                    'FA売価(税抜)': None,
                    '粗利': '-',
                    'RT後の利益': '-',
                    'FA売価(税込)': None
                },
                'shop_info': {}
            }
            
            # Add shop information
            for shop_code, shop_info in {
                'waste': {
                    'name': 'e-life＆work shop',
                    'base_url': 'https://item.rakuten.co.jp/waste/'
                },
                'kougushop': {
                    'name': '工具ショップ',
                    'base_url': 'https://item.rakuten.co.jp/kougushop/'
                },
                'kouei-sangyou': {
                    'name': '晃栄産業',
                    'base_url': 'https://item.rakuten.co.jp/kouei-sangyou/'
                },
                'dear-worker': {
                    'name': 'dear-worker',
                    'base_url': 'https://item.rakuten.co.jp/dear-worker/'
                }
            }.items():
                # Generate shop-specific URL
                if shop_code == 'waste':
                    url = f"{shop_info['base_url']}cp209/?rafcid=wsc_i_is_1085274442429696242"
                elif shop_code == 'kougushop':
                    url = f"{shop_info['base_url']}{sku.split('-')[1]}-{sku.split('-')[2]}/?rafcid=wsc_i_is_1085274442429696242"
                elif shop_code == 'kouei-sangyou':
                    url = f"{shop_info['base_url']}fcp209/?rafcid=wsc_i_is_1085274442429696242"
                else:  # dear-worker
                    url = f"{shop_info['base_url']}cp209boa/?rafcid=wsc_i_is_1085274442429696242"

                item['shop_info'][shop_code] = {
                    'shop_name': shop_info['name'],
                    'shop_code': shop_code,
                    'URL': url
                }

                print(f"  Scraping {shop_info['name']} ({url})")
                
                # Check shop type
                is_waste_shop = shop_code == 'waste'
                is_kougushop = shop_code == 'kougushop'
                is_kouei_shop = shop_code == 'kouei-sangyou'
                is_dear_worker = shop_code == 'dear-worker'
                
                # Scrape product info
                product_info = scrape_product_info(session, url, is_waste_shop, is_kougushop, is_kouei_shop, is_dear_worker)
                
                if product_info:
                    if (is_waste_shop or is_kouei_shop or is_dear_worker) and product_info['variants']:
                        item['shop_info'][shop_code]['variants'] = product_info['variants']
                    elif is_kougushop and product_info['variants']:
                        item['shop_info'][shop_code]['variants'] = product_info['variants']

                # Random delay between shops
                time.sleep(random.uniform(1, 2))

            # Add item to results
            results['items'].append(item)
            
            # Calculate and show progress
            elapsed_time = time.time() - sku_start_time
            total_elapsed = time.time() - start_time
            avg_time = total_elapsed / idx
            remaining_time = (len(skus) - idx) * avg_time
            
            print(f"  Time for this SKU: {elapsed_time:.2f}s")
            print(f"  Progress: {idx}/{len(skus)} ({idx/len(skus)*100:.1f}%)")
            print(f"  Estimated time remaining: {remaining_time/60:.1f} minutes")

            # Save results after each item (in case of interruption)
            results['metadata'].update({
                'processed_skus': idx,
                'current_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'elapsed_time_seconds': round(total_elapsed, 2),
                'average_time_per_sku': round(avg_time, 2)
            })
            
            with open('results_beautifulsoup.json', 'w', encoding='utf-8') as f:
                json.dump(results, ensure_ascii=False, indent=2, fp=f)

        # Final metadata update
        results['metadata'].update({
            'completion_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'total_time_seconds': round(time.time() - start_time, 2),
            'total_items_processed': len(skus)
        })

        # Final save
        with open('results_beautifulsoup.json', 'w', encoding='utf-8') as f:
            json.dump(results, ensure_ascii=False, indent=2, fp=f)

        print("\nScraping completed successfully!")
        print(f"Total time: {time.time() - start_time:.2f} seconds")
        print(f"Average time per SKU: {(time.time() - start_time)/len(skus):.2f} seconds")

    finally:
        session.close()

if __name__ == "__main__":
    main() 