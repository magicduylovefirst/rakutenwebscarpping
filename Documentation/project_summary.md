# Rakuten Web Scraping Project Summary

## Project Structure

Main Directory: `C:\DDisk\WorkingProgress\First\Python\WebScrapping`

### Files:

```
WebScrapping/
├── rakuten_api_fetcher.py      # Uses Rakuten API to fetch item data
├── New folder/
│   ├── rakuten_item_fetcher.py # Uses RMS API for item details
│   ├── sku_scraper.py          # New file for SKU-based searching
│   └── araki.xlsx             # Contains SKUs to search
```

## File Details

### 1. rakuten_api_fetcher.py

- Uses Rakuten Item Search API
- Requires Application ID: "1085274442429696242"
- Capabilities:
  - Fetches item name
  - Gets price
  - Gets shop name
  - Gets size variations

### 2. New folder/rakuten_item_fetcher.py

- Uses Rakuten RMS API
- Requirements:
  - Service Secret
  - License Key
- Purpose: Fetches detailed item information

### 3. New folder/sku_scraper.py

- Reads SKUs from araki.xlsx
- Searches across specific shops:
  - e-life＆work shop (`waste`)
  - 工具ショップ (`kougushop`)
  - 晃栄産業 (`kouei-sangyou`)
  - Dear worker (`dear-worker`)

## Shop URL Patterns

Example SKU: asc-1271a029-025-250

```
e-life＆work shop: 'https://item.rakuten.co.jp/waste/cp209/'
工具ショップ: 'https://item.rakuten.co.jp/kougushop/1271a029-025/'
晃栄産業: 'https://item.rakuten.co.jp/kouei-sangyou/fcp209/'
Dear worker: 'https://item.rakuten.co.jp/dear-worker/cp209boa/'
```

Additional variants:

- 工具ショップ:
  - https://item.rakuten.co.jp/kougushop/1271a029-602/
  - https://item.rakuten.co.jp/kougushop/1271a029-400/
- Dear worker:
  - https://item.rakuten.co.jp/dear-worker/2360345/?variantId=2360345750300

## Current Status

### Working Methods:

1. API-based approach (`rakuten_api_fetcher.py`)
2. Web scraping approach (in development)

### Challenges:

- Rate limiting from Rakuten
- Different URL patterns per shop
- Need to handle variants/sizes

## Next Steps

1. Choose approach:
   - Web scraping (needs rate limiting solution)
   - Rakuten API (more reliable but limited data)
2. Add size/variant extraction
3. Handle shop-specific URL patterns
