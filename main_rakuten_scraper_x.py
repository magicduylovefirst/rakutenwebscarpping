import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import urllib.parse
import json
import time
from datetime import datetime
import threading
from queue import Queue
from concurrent.futures import ThreadPoolExecutor

def read_skus_from_excel(excel_path):
    """Read SKUs from first column of Excel file"""
    try:
        df = pd.read_excel(excel_path)
        skus = df.iloc[:, 0].dropna().astype(str).tolist()
        return [sku for sku in skus if sku.lower() != 'skuコード']
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def setup_webdriver():
    """Setup and return configured Chrome WebDriver"""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.stylesheets": 2
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def get_variant_info(driver, base_url):
    """Get all variants information including colors and sizes"""
    variants = []
    try:
        # Get color variants
        color_buttons = driver.find_elements(By.CSS_SELECTOR, "div.grid-cols-2--1uI00 button.type-sku-button--BJoVv")
        size_buttons = driver.find_elements(By.CSS_SELECTOR, "div.grid-cols-5--3wKbc button.type-sku-button--BJoVv")
        
        # Extract color and size information
        colors = []
        for btn in color_buttons:
            color_name = btn.get_attribute('aria-label')
            is_selected = 'selected--Mg4iu' in btn.get_attribute('class')
            colors.append({
                'name': color_name,
                'selected': is_selected
            })
            
        sizes = []
        for btn in size_buttons:
            size_name = btn.get_attribute('aria-label')
            is_selected = 'selected--Mg4iu' in btn.get_attribute('class')
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

                # Visit variant URL and get its information
                driver.get(variant_url)
                time.sleep(1)  # Wait for page to load

                try:
                    # Get price - try both new and old selectors
                    price_element = None
                    try:
                        price_element = driver.find_element(By.CSS_SELECTOR, "div.value--1oSD_.layout-inline--2z490.size-x-large--DyMl5.style-bold500--1X0Xl.color-crimson--2uc0e.align-right--3POGa")
                    except:
                        price_element = driver.find_element(By.CSS_SELECTOR, "span.price--OX_YW")
                    
                    if price_element:
                        variants[-1]['価格'] = int(price_element.text.replace('円', '').replace(',', ''))
                except:
                    pass

                try:
                    # Get points
                    points_element = driver.find_element(By.CSS_SELECTOR, "div.point-summary__total___3rYYD span")
                    if not points_element:
                        points_element = driver.find_element(By.CSS_SELECTOR, "span.price--point-badge_item")
                    
                    if points_element:
                        points_text = points_element.text.replace('ポイント', '').replace(',', '')
                        variants[-1]['ポイント'] = int(points_text) if points_text.isdigit() else None
                except:
                    pass

                try:
                    # Get coupon
                    coupon_element = driver.find_element(By.CSS_SELECTOR, "div.coupon")
                    coupon_text = coupon_element.text
                    if coupon_text:
                        coupon_value = int(''.join(filter(str.isdigit, coupon_text)))
                        variants[-1]['クーポン'] = coupon_value
                except:
                    pass

                try:
                    # Check inventory status
                    out_of_stock = driver.find_elements(By.XPATH, "//*[contains(text(), '売り切れ')]")
                    if out_of_stock:
                        variants[-1]['在庫状況'] = '在庫なし'
                    else:
                        in_stock = driver.find_elements(By.XPATH, "//*[contains(text(), '在庫あり')]")
                        if in_stock:
                            variants[-1]['在庫状況'] = '在庫あり'
                except:
                    pass

    except Exception as e:
        print(f"Error getting variants: {e}")
    
    return variants

def get_kougushop_variant_info(driver, base_url):
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

            # Visit variant URL and get its information
            driver.get(variant_url)
            time.sleep(1)  # Wait for page to load

            try:
                # Get price - try both new and old selectors
                price_element = None
                try:
                    price_element = driver.find_element(By.CSS_SELECTOR, "div.value--1oSD_.layout-inline--2z490.size-x-large--DyMl5.style-bold500--1X0Xl.color-crimson--2uc0e.align-right--3POGa")
                except:
                    price_element = driver.find_element(By.CSS_SELECTOR, "span.price--OX_YW")
                
                if price_element:
                    variants[-1]['価格'] = int(price_element.text.replace('円', '').replace(',', ''))
            except:
                pass

            try:
                # Get points
                points_element = driver.find_element(By.CSS_SELECTOR, "div.point-summary__total___3rYYD span")
                if not points_element:
                    points_element = driver.find_element(By.CSS_SELECTOR, "span.price--point-badge_item")
                
                if points_element:
                    points_text = points_element.text.replace('ポイント', '').replace(',', '')
                    variants[-1]['ポイント'] = int(points_text) if points_text.isdigit() else None
            except:
                pass

            try:
                # Get coupon
                coupon_element = driver.find_element(By.CSS_SELECTOR, "div.coupon")
                coupon_text = coupon_element.text
                if coupon_text:
                    coupon_value = int(''.join(filter(str.isdigit, coupon_text)))
                    variants[-1]['クーポン'] = coupon_value
            except:
                pass

            try:
                # Check inventory status
                out_of_stock = driver.find_elements(By.XPATH, "//*[contains(text(), '売り切れ')]")
                if out_of_stock:
                    variants[-1]['在庫状況'] = '在庫なし'
                else:
                    in_stock = driver.find_elements(By.XPATH, "//*[contains(text(), '在庫あり')]")
                    if in_stock:
                        variants[-1]['在庫状況'] = '在庫あり'
            except:
                pass

    except Exception as e:
        print(f"Error getting kougushop variants: {e}")
    
    return variants

def scrape_product_info(driver, url, is_waste_shop=False, is_kougushop=False, is_kouei_shop=False, is_dear_worker=False):
    """Scrape product information from a given URL"""
    try:
        driver.get(url)
        time.sleep(2)  # Wait for page to load

        # Initialize product info dictionary
        product_info = {
            'url': url,
            'variants': [] if (is_waste_shop or is_kougushop or is_kouei_shop or is_dear_worker) else None
        }

        # Get variants based on shop type
        if is_waste_shop or is_kougushop or is_kouei_shop or is_dear_worker:
            base_url = url.split('&variantId=')[0] if '&variantId=' in url else url
            if is_waste_shop or is_kouei_shop or is_dear_worker:
                product_info['variants'] = get_variant_info(driver, base_url)
            elif is_kougushop:
                product_info['variants'] = get_kougushop_variant_info(driver, base_url)

        return product_info

    except Exception as e:
        print(f"Error scraping URL {url}: {e}")
        return None

def process_sku(sku, idx, total_skus, result_queue, shops):
    """Process a single SKU in a separate thread"""
    driver = None
    try:
        driver = setup_webdriver()
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

        print(f"\nProcessing SKU {idx}/{total_skus}: {sku}")
        
        # Process each shop sequentially within this thread
        for shop_code, shop_info in shops.items():
            try:
                # Generate shop-specific URL
                if shop_code == 'waste':
                    url = f"{shop_info['base_url']}cp209/?rafcid=wsc_i_is_1085274442429696242"
                elif shop_code == 'kougushop':
                    url = f"{shop_info['base_url']}{sku.split('-')[1]}-{sku.split('-')[2]}/?rafcid=wsc_i_is_1085274442429696242"
                elif shop_code == 'kouei-sangyou':
                    url = f"{shop_info['base_url']}fcp209/?rafcid=wsc_i_is_1085274442429696242"
                else:  # dear-worker
                    url = f"{shop_info['base_url']}cp209boa/?rafcid=wsc_i_is_1085274442429696242"

                shop_result = {
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
                product_info = scrape_product_info(driver, url, is_waste_shop, is_kougushop, is_kouei_shop, is_dear_worker)
                
                if product_info:
                    if (is_waste_shop or is_kouei_shop or is_dear_worker) and product_info['variants']:
                        shop_result['variants'] = product_info['variants']
                    elif is_kougushop and product_info['variants']:
                        shop_result['variants'] = product_info['variants']

                item['shop_info'][shop_code] = shop_result

            except Exception as e:
                print(f"Error scraping shop {shop_code} for SKU {sku}: {e}")

        # Calculate timing information
        elapsed_time = time.time() - sku_start_time
        print(f"  Time for SKU {sku}: {elapsed_time:.2f}s")
        
        # Put result in queue
        result_queue.put((idx, item))

    except Exception as e:
        print(f"Error processing SKU {sku}: {e}")
        result_queue.put((idx, None))
    finally:
        if driver:
            driver.quit()

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

    # Shop information dictionary
    shops = {
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
    }

    try:
        # Create a queue for results
        result_queue = Queue()
        
        # Process SKUs in batches of 4 using ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=4) as executor:
            # Submit all tasks
            futures = []
            for idx, sku in enumerate(skus, 1):
                future = executor.submit(process_sku, sku, idx, len(skus), result_queue, shops)
                futures.append(future)

            # Process results as they complete
            completed_count = 0
            while completed_count < len(skus):
                # Get result from queue
                idx, item = result_queue.get()
                completed_count += 1
                
                if item:
                    results['items'].append(item)
                
                # Calculate progress
                total_elapsed = time.time() - start_time
                avg_time = total_elapsed / completed_count
                remaining_time = (len(skus) - completed_count) * avg_time
                
                print(f"Progress: {completed_count}/{len(skus)} ({completed_count/len(skus)*100:.1f}%)")
                print(f"Estimated time remaining: {remaining_time/60:.1f} minutes")

                # Save intermediate results
                results['metadata'].update({
                    'processed_skus': completed_count,
                    'current_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'elapsed_time_seconds': round(total_elapsed, 2),
                    'average_time_per_sku': round(avg_time, 2)
                })
                
                with open('results.json', 'w', encoding='utf-8') as f:
                    json.dump(results, ensure_ascii=False, indent=2, fp=f)

        # Final metadata update
        results['metadata'].update({
            'completion_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'total_time_seconds': round(time.time() - start_time, 2),
            'total_items_processed': len(skus)
        })

        # Final save
        with open('results.json', 'w', encoding='utf-8') as f:
            json.dump(results, ensure_ascii=False, indent=2, fp=f)

        print("\nScraping completed successfully!")
        print(f"Total time: {time.time() - start_time:.2f} seconds")
        print(f"Average time per SKU: {(time.time() - start_time)/len(skus):.2f} seconds")

    except Exception as e:
        print(f"Error in main process: {e}")

if __name__ == "__main__":
    main()