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
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
from functools import partial

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Constants
MAX_WORKERS = 4  # Adjust based on your system's capabilities
WAIT_TIMEOUT = 10  # Maximum wait time for elements
PAGE_LOAD_WAIT = 1  # Reduced from 2 seconds

class WebDriverPool:
    def __init__(self, pool_size):
        self.pool = Queue()
        self.pool_size = pool_size
        self.initialize_pool()

    def initialize_pool(self):
        for _ in range(self.pool_size):
            driver = setup_webdriver()
            self.pool.put(driver)

    def get_driver(self):
        return self.pool.get()

    def return_driver(self, driver):
        self.pool.put(driver)

    def close_all(self):
        while not self.pool.empty():
            driver = self.pool.get()
            driver.quit()

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
    """Setup and return configured Chrome WebDriver with optimized settings"""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.page_load_strategy = 'eager'  # Don't wait for all resources to load
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(30)
    return driver

def wait_for_element(driver, selector, by=By.CSS_SELECTOR, timeout=WAIT_TIMEOUT):
    """Wait for element with better error handling"""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
        return element
    except Exception:
        return None

def process_variant(driver, variant_url, color=None, size=None):
    """Process a single variant with optimized waits"""
    variant_info = {
        'color': color,
        'size': size,
        'variant_id': variant_url.split('variantId=')[-1],
        'url': variant_url,
        '価格': None,
        'ポイント': None,
        'クーポン': None,
        '在庫状況': None
    }

    try:
        driver.get(variant_url)
        
        # Price
        price_selectors = [
            "div.value--1oSD_.layout-inline--2z490.size-x-large--DyMl5.style-bold500--1X0Xl.color-crimson--2uc0e.align-right--3POGa",
            "span.price--OX_YW"
        ]
        for selector in price_selectors:
            element = wait_for_element(driver, selector)
            if element:
                variant_info['価格'] = int(element.text.replace('円', '').replace(',', ''))
                break

        # Points
        points_selectors = [
            "div.point-summary__total___3rYYD span",
            "span.price--point-badge_item"
        ]
        for selector in points_selectors:
            element = wait_for_element(driver, selector)
            if element:
                points_text = element.text.replace('ポイント', '').replace(',', '')
                variant_info['ポイント'] = int(points_text) if points_text.isdigit() else None
                break

        # Coupon
        coupon_element = wait_for_element(driver, "div.coupon")
        if coupon_element:
            coupon_text = coupon_element.text
            if coupon_text:
                coupon_value = int(''.join(filter(str.isdigit, coupon_text)))
                variant_info['クーポン'] = coupon_value

        # Stock status
        out_of_stock = wait_for_element(driver, "//*[contains(text(), '売り切れ')]", By.XPATH)
        if out_of_stock:
            variant_info['在庫状況'] = '在庫なし'
        else:
            in_stock = wait_for_element(driver, "//*[contains(text(), '在庫あり')]", By.XPATH)
            if in_stock:
                variant_info['在庫状況'] = '在庫あり'

    except Exception as e:
        logger.error(f"Error processing variant {variant_url}: {e}")

    return variant_info

def get_variant_info(driver, base_url):
    """Get all variants information using concurrent processing"""
    variants = []
    try:
        driver.get(base_url)
        
        # Get color and size buttons
        color_buttons = driver.find_elements(By.CSS_SELECTOR, "div.grid-cols-2--1uI00 button.type-sku-button--BJoVv")
        size_buttons = driver.find_elements(By.CSS_SELECTOR, "div.grid-cols-5--3wKbc button.type-sku-button--BJoVv")
        
        # Extract color and size information
        colors = [{'name': btn.get_attribute('aria-label')} for btn in color_buttons]
        sizes = [{'name': btn.get_attribute('aria-label')} for btn in size_buttons]

        # Generate variant URLs
        variant_tasks = []
        variant_id = 1
        
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            for color in colors:
                for size in sizes:
                    variant_url = f"{base_url}&variantId=r-sku{variant_id:08d}"
                    task = executor.submit(
                        process_variant,
                        driver,
                        variant_url,
                        color['name'],
                        size['name']
                    )
                    variant_tasks.append(task)
                    variant_id += 1

            # Collect results
            for future in as_completed(variant_tasks):
                try:
                    variant_info = future.result()
                    if variant_info:
                        variants.append(variant_info)
                except Exception as e:
                    logger.error(f"Error collecting variant result: {e}")

    except Exception as e:
        logger.error(f"Error getting variants: {e}")
    
    return variants

def get_kougushop_variant_info(driver, base_url):
    """Get variants information for kougushop using concurrent processing"""
    variants = []
    try:
        sizes = ['22.5', '23', '23.5', '24', '24.5', '25', '25.5', '26', '26.5', '27', '27.5', '28', '29', '30']
        variant_tasks = []
        
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            for idx, size in enumerate(sizes, 8021):
                variant_url = f"{base_url}&variantId={idx}"
                task = executor.submit(
                    process_variant,
                    driver,
                    variant_url,
                    None,
                    size
                )
                variant_tasks.append(task)

            # Collect results
            for future in as_completed(variant_tasks):
                try:
                    variant_info = future.result()
                    if variant_info:
                        variants.append(variant_info)
                except Exception as e:
                    logger.error(f"Error collecting kougushop variant result: {e}")

    except Exception as e:
        logger.error(f"Error getting kougushop variants: {e}")
    
    return variants

def scrape_product_info(driver, url, is_waste_shop=False, is_kougushop=False, is_kouei_shop=False, is_dear_worker=False):
    """Scrape product information with optimized variant handling"""
    try:
        base_url = url.split('&variantId=')[0] if '&variantId=' in url else url
        product_info = {
            'url': url,
            'variants': []
        }

        if is_waste_shop or is_kouei_shop or is_dear_worker:
            product_info['variants'] = get_variant_info(driver, base_url)
        elif is_kougushop:
            product_info['variants'] = get_kougushop_variant_info(driver, base_url)

        return product_info

    except Exception as e:
        logger.error(f"Error scraping URL {url}: {e}")
        return None

def scrape_shop(driver_pool, shop_code, shop_info, sku, result_queue):
    """Scrape a single shop using the driver pool"""
    driver = None
    try:
        driver = driver_pool.get_driver()
        
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

        logger.info(f"Scraping {shop_info['name']} ({url})")
        
        # Scrape product info
        product_info = scrape_product_info(
            driver, 
            url,
            is_waste_shop=(shop_code == 'waste'),
            is_kougushop=(shop_code == 'kougushop'),
            is_kouei_shop=(shop_code == 'kouei-sangyou'),
            is_dear_worker=(shop_code == 'dear-worker')
        )
        
        if product_info and product_info['variants']:
            shop_result['variants'] = product_info['variants']

        result_queue.put((shop_code, shop_result))

    except Exception as e:
        logger.error(f"Error scraping {shop_code}: {e}")
        result_queue.put((shop_code, None))
    finally:
        if driver:
            driver_pool.return_driver(driver)

def main():
    start_time = time.time()
    
    # Read SKUs from Excel
    excel_path = "New folder/araki.xlsx"
    skus = read_skus_from_excel(excel_path)
    
    if not skus:
        logger.error("No SKUs found in Excel file")
        return

    logger.info(f"Found {len(skus)} SKUs to process")

    # Initialize results structure
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

    # Initialize WebDriver pool
    driver_pool = WebDriverPool(MAX_WORKERS)

    try:
        # Process each SKU
        for idx, sku in enumerate(skus, 1):
            logger.info(f"\nProcessing SKU {idx}/{len(skus)}: {sku}")
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
            
            # Create a queue for thread results
            result_queue = Queue()
            threads = []
            
            # Start a thread for each shop
            for shop_code, shop_info in shops.items():
                thread = threading.Thread(
                    target=scrape_shop,
                    args=(driver_pool, shop_code, shop_info, sku, result_queue)
                )
                threads.append(thread)
                thread.start()

            # Wait for all threads to complete
            for thread in threads:
                thread.join()

            # Collect results from queue
            while not result_queue.empty():
                shop_code, shop_result = result_queue.get()
                if shop_result:
                    item['shop_info'][shop_code] = shop_result

            # Add item to results
            results['items'].append(item)
            
            # Calculate and show progress
            elapsed_time = time.time() - sku_start_time
            total_elapsed = time.time() - start_time
            avg_time = total_elapsed / idx
            remaining_time = (len(skus) - idx) * avg_time
            
            logger.info(f"Time for this SKU: {elapsed_time:.2f}s")
            logger.info(f"Progress: {idx}/{len(skus)} ({idx/len(skus)*100:.1f}%)")
            logger.info(f"Estimated time remaining: {remaining_time/60:.1f} minutes")

            # Save results after each item
            results['metadata'].update({
                'processed_skus': idx,
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

        logger.info("\nScraping completed successfully!")
        logger.info(f"Total time: {time.time() - start_time:.2f} seconds")
        logger.info(f"Average time per SKU: {(time.time() - start_time)/len(skus):.2f} seconds")

    except Exception as e:
        logger.error(f"Error in main process: {e}")
    finally:
        driver_pool.close_all()

if __name__ == "__main__":
    main() 