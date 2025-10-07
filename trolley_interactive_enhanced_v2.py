#!/usr/bin/env python3
"""
Enhanced Trolley.co.uk Interactive Scraper with Professional Architecture
Features: Product matching, CSV export, error handling, logging, quantity detection
"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
import random
from dataclasses import dataclass
from typing import List, Optional, Tuple
import logging
from datetime import datetime
import os
from urllib.parse import urljoin, quote

# Setup logging
def setup_logging():
    """Setup comprehensive logging with file and console output"""
    log_filename = f"trolley_scraper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logging.info(f"🚀 Enhanced Trolley Scraper started - Log file: {log_filename}")

@dataclass
class SearchResultItem:
    """Representation of a single product from Trolley search results"""
    brand: str
    description: str
    size: str
    quantity: str
    price: str
    url: str
    search_url: str = ""

def parse_sku_name(sku_name: str) -> Tuple[str, Optional[str], Optional[int]]:
    """
    Parse brand, size and quantity from a SKU name string.
    
    Enhanced parsing logic for patterns like:
    - "Heineken Premium Lager 15x440ml"
    - "Corona Extra Beer 12 x 330ml"
    - "Stella Artois 4x568ml"
    
    :param sku_name: Raw product name from Excel
    :return: Tuple of (brand, size, quantity)
    """
    name = sku_name.strip()
    logging.debug(f"🔍 Parsing SKU: '{name}'")
    
    qty = None
    size = None
    brand = ""
    
    # Enhanced patterns for quantity x size matching
    patterns = [
        r"(\d+)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*(ml|l|g|kg|cl|oz|fl\s*oz)",  # 15x440ml, 12 x 330ml
        r"(\d+)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*([a-zA-Z]+)",  # More flexible unit matching
        r"(\d+)\s*pack.*?(\d+(?:\.\d+)?)\s*(ml|l|g|kg|cl|oz|fl\s*oz)",  # 15 pack 440ml
    ]
    
    qty_size_match = None
    for pattern in patterns:
        match = re.search(pattern, name, re.IGNORECASE)
        if match:
            qty = int(match.group(1))
            size_num = match.group(2)
            size_unit = match.group(3).lower().replace(" ", "")
            size = f"{size_num}{size_unit}"
            qty_size_match = match
            logging.debug(f"✅ Pattern matched: qty={qty}, size={size}")
            break
    
    if not size:
        # Look for standalone size (e.g., "440ml", "1.5l")
        size_match = re.search(r"(\d+(?:\.\d+)?)\s*(ml|l|g|kg|cl|oz|fl\s*oz)", name, re.IGNORECASE)
        if size_match:
            size = f"{size_match.group(1)}{size_match.group(2).lower().replace(' ', '')}"
            qty = 1  # Default for single items
            logging.debug(f"✅ Size only matched: size={size}, qty={qty}")
    
    # Extract brand - improved logic to avoid including quantity/size in brand
    if qty_size_match:
        # Split at the quantity pattern and take the part before it
        brand_part = name[:qty_size_match.start()].strip()
        brand = brand_part
    else:
        # Fallback: take all words before first number
        tokens = name.split()
        brand_tokens = []
        for token in tokens:
            if re.search(r'\d', token):
                break
            brand_tokens.append(token)
        brand = " ".join(brand_tokens)
    
    # Clean up brand - remove common trailing words that might be size/quantity related
    brand = re.sub(r'\s+(x|X|×|\d+)\s*$', '', brand)
    brand = re.sub(r'\s+', ' ', brand).strip()
    
    logging.debug(f"📋 Final parsing result: Brand='{brand}', Size='{size}', Quantity={qty}")
    return brand, size, qty

def get_search_results(search_term: str, max_retries: int = 3) -> Tuple[List[SearchResultItem], str]:
    """
    Get search results from Trolley.co.uk with comprehensive HTML parsing
    
    :param search_term: Product name to search for
    :param max_retries: Maximum number of retry attempts
    :return: Tuple of (List of SearchResultItem objects, search_url)
    """
    base_url = "https://www.trolley.co.uk/search/"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
    }
    
    params = {
        'from': 'search',
        'q': search_term
    }
    
    # Construct search URL
    search_url = f"{base_url}?from=search&q={quote(search_term)}"
    
    for attempt in range(max_retries):
        try:
            logging.info(f"🔍 Searching for: {search_term} (attempt {attempt + 1})")
            logging.info(f"🌐 Search URL: {search_url}")
            
            response = requests.get(base_url, params=params, headers=headers, timeout=15)
            response.raise_for_status()
            
            # Save the complete HTML for debugging
            html_debug_file = f"search_debug_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
            with open(html_debug_file, 'w', encoding='utf-8') as f:
                f.write(response.text)
            logging.debug(f"💾 Saved search HTML to: {html_debug_file}")
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Parse search results with comprehensive logic
            results = []
            
            # Try multiple selectors to find product items
            product_selectors = [
                'div.product-item',
                'section#search-results div.product-item',
                '.product-item',
                '[class*="product"]'
            ]
            
            product_items = []
            for selector in product_selectors:
                product_items = soup.select(selector)
                if product_items:
                    logging.info(f"✅ Found {len(product_items)} items using selector: {selector}")
                    break
            
            if not product_items:
                logging.warning("❌ No product items found with any selector")
                return [], search_url
            
            logging.info(f"📦 Processing {len(product_items)} product items")
            
            for i, item in enumerate(product_items, 1):
                try:
                    # More comprehensive brand extraction
                    brand = "Unknown"
                    brand_selectors = ['div._brand', '._brand', '[class*="brand"]']
                    for selector in brand_selectors:
                        brand_elem = item.select_one(selector)
                        if brand_elem:
                            brand = brand_elem.get_text(strip=True)
                            break
                    
                    # More comprehensive description extraction
                    description = "No description"
                    desc_selectors = ['div._desc', '._desc', '[class*="desc"]', '[class*="title"]']
                    for selector in desc_selectors:
                        desc_elem = item.select_one(selector)
                        if desc_elem:
                            description = desc_elem.get_text(strip=True)
                            break
                    
                    # Comprehensive size extraction
                    size = "Unknown"
                    size_text = ""
                    
                    # Try multiple approaches to find size
                    size_elem = item.select_one('div._size')
                    if size_elem:
                        # Get all text from size element and its children
                        size_text = size_elem.get_text(strip=True)
                        
                        # Try to find size in a nested div
                        size_div = size_elem.select_one('div:not([class])')
                        if size_div:
                            size_text = size_div.get_text(strip=True)
                    
                    # Extract size with multiple patterns
                    if size_text:
                        size_patterns = [
                            r'(\d+(?:\.\d+)?\s*(?:ml|l|ML|L|g|kg|oz))',
                            r'(\d+(?:\.\d+)?\s*[a-zA-Z]+)',
                            r'([\d.]+\s*[a-zA-Z]+)'
                        ]
                        
                        for pattern in size_patterns:
                            size_match = re.search(pattern, size_text, re.IGNORECASE)
                            if size_match:
                                size = size_match.group(1).lower().replace(' ', '')
                                break
                        
                        if size == "Unknown":
                            size = size_text[:20]  # Fallback to first 20 chars
                    
                    # Comprehensive quantity extraction
                    quantity = "1"  # Default
                    
                    # Look for quantity in multiple places
                    qty_sources = [
                        item.select_one('div._qty'),
                        item.select_one('._qty'),
                        size_elem.select_one('._qty') if size_elem else None,
                        size_elem.select_one('div._qty') if size_elem else None
                    ]
                    
                    for qty_elem in qty_sources:
                        if qty_elem:
                            qty_text = qty_elem.get_text(strip=True)
                            # Extract first number from quantity text
                            qty_match = re.search(r'(\d+)', qty_text)
                            if qty_match:
                                quantity = qty_match.group(1)
                                break
                    
                    # Also check the full item text for quantity patterns like "15x", "12 x"
                    full_text = item.get_text()
                    qty_patterns = [
                        r'(\d+)\s*[xX×]\s*\d+\s*[a-zA-Z]+',  # 15x440ml
                        r'(\d+)\s*[xX×]',  # 15x
                        r'[xX×]\s*(\d+)',  # x15
                        r'(\d+)\s*pack',  # 15 pack
                        r'pack\s*of\s*(\d+)',  # pack of 15
                    ]
                    
                    for pattern in qty_patterns:
                        qty_match = re.search(pattern, full_text, re.IGNORECASE)
                        if qty_match:
                            potential_qty = qty_match.group(1)
                            # Only use if it's a reasonable quantity (1-100)
                            if 1 <= int(potential_qty) <= 100:
                                quantity = potential_qty
                                break
                    
                    # Comprehensive price extraction
                    price = "Unknown"
                    price_selectors = ['div._price', '._price', '[class*="price"]']
                    for selector in price_selectors:
                        price_elem = item.select_one(selector)
                        if price_elem:
                            price_text = price_elem.get_text(strip=True)
                            price_match = re.search(r'£[\d.,]+', price_text)
                            if price_match:
                                price = price_match.group()
                                break
                    
                    # Comprehensive URL extraction
                    url = ""
                    link_elem = item.select_one('a[href]')
                    if link_elem:
                        href = link_elem.get('href')
                        if href:
                            url = urljoin("https://www.trolley.co.uk", href)
                    
                    # Create result item with search URL
                    result = SearchResultItem(
                        brand=brand,
                        description=description,
                        size=size,
                        quantity=quantity,
                        price=price,
                        url=url,
                        search_url=search_url
                    )
                    
                    results.append(result)
                    logging.debug(f"📄 Product {i}: {brand} | {description} | {size} | x{quantity} | {price} | {url}")
                    
                except Exception as e:
                    logging.warning(f"⚠️ Error parsing product item {i}: {e}")
                    continue
            
            if results:
                logging.info(f"✅ Successfully parsed {len(results)} products")
                # Log detailed results for debugging
                for i, result in enumerate(results[:5], 1):
                    logging.info(f"  {i}. {result.brand} {result.description} {result.size} x{result.quantity} - {result.price}")
                return results, search_url
            else:
                logging.warning("❌ No products found in search results")
                
        except requests.RequestException as e:
            logging.error(f"🌐 Request failed (attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                wait_time = (attempt + 1) * 2
                logging.info(f"⏳ Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
        except Exception as e:
            logging.error(f"💥 Unexpected error during search: {e}")
            break
    
    logging.error(f"❌ Failed to get search results after {max_retries} attempts")
    return [], search_url

def find_best_match(
    expected_brand: str,
    expected_size: Optional[str],
    expected_qty: Optional[int],
    candidates: List[SearchResultItem],
) -> Optional[SearchResultItem]:
    """
    Return the candidate that best matches the expected attributes.
    
    Enhanced matching logic with multi-tier strategy:
    1. Exact match: brand + size + quantity
    2. Partial match: brand + size
    3. Brand match only
    4. Fallback to first result
    
    :param expected_brand: Parsed brand from SKU name
    :param expected_size: Parsed size from SKU name (e.g., '440ml')
    :param expected_qty: Parsed quantity from SKU name (e.g., 15)
    :param candidates: List of SearchResultItem objects from search results
    :return: Matched SearchResultItem or None
    """
    if not candidates:
        logging.warning("❌ No candidates provided for matching")
        return None
    
    # Convert expected quantity to string for comparison
    expected_qty_str = str(expected_qty) if expected_qty else ""
    
    logging.info(f"🎯 Looking for: Brand='{expected_brand}', Size='{expected_size}', Quantity='{expected_qty_str}'")
    logging.info(f"📝 Checking {len(candidates)} candidates...")
    
    # Tier 1: Exact match (brand + size + quantity)
    logging.info("🔍 Tier 1: Checking for exact matches (brand + size + quantity)...")
    for i, item in enumerate(candidates, 1):
        # More flexible brand matching - check both directions
        brand_match = False
        if expected_brand:
            # Check if expected brand contains item brand OR item brand contains expected brand
            expected_lower = expected_brand.lower()
            item_lower = item.brand.lower()
            brand_match = (expected_lower in item_lower or 
                          item_lower in expected_lower or
                          # Also check for partial word matches
                          any(word in item_lower for word in expected_lower.split() if len(word) > 3))
        else:
            brand_match = True
            
        size_match = expected_size.lower() in item.size.lower() if expected_size else True
        qty_match = expected_qty_str == item.quantity if expected_qty_str else True
        
        logging.debug(f"  Candidate {i}: {item.brand} {item.description} {item.size} x{item.quantity}")
        logging.debug(f"    Expected: brand='{expected_brand}', size='{expected_size}', qty='{expected_qty_str}'")
        logging.debug(f"    Found: brand='{item.brand}', size='{item.size}', qty='{item.quantity}'")
        logging.debug(f"    Matches: brand={brand_match}, size={size_match}, qty={qty_match}")
        
        if brand_match and size_match and qty_match:
            logging.info(f"🎉 Tier 1 PERFECT MATCH found: {item.brand} {item.description} {item.size} x{item.quantity}")
            logging.info(f"✅ Verification: Brand(✓) + Size(✓) + Quantity(✓) - Price: {item.price}")
            return item
    
    # Tier 2: Brand + size match
    logging.info("🔍 Tier 2: Checking for brand + size matches...")
    for i, item in enumerate(candidates, 1):
        # More flexible brand matching
        brand_match = False
        if expected_brand:
            expected_lower = expected_brand.lower()
            item_lower = item.brand.lower()
            brand_match = (expected_lower in item_lower or 
                          item_lower in expected_lower or
                          any(word in item_lower for word in expected_lower.split() if len(word) > 3))
        else:
            brand_match = True
            
        size_match = expected_size.lower() in item.size.lower() if expected_size else True
        
        if brand_match and size_match:
            logging.info(f"✅ Tier 2 match found: {item.brand} {item.description} {item.size} (quantity differs)")
            return item
    
    # Tier 3: Brand match only
    logging.info("🔍 Tier 3: Checking for brand-only matches...")
    for item in candidates:
        # More flexible brand matching
        brand_match = False
        if expected_brand:
            expected_lower = expected_brand.lower()
            item_lower = item.brand.lower()
            brand_match = (expected_lower in item_lower or 
                          item_lower in expected_lower or
                          any(word in item_lower for word in expected_lower.split() if len(word) > 3))
        else:
            brand_match = True
            
        if brand_match:
            logging.info(f"⚠️ Tier 3 match found: {item.brand} {item.description} (size/quantity differ)")
            return item
    
    # Tier 4: Fallback to first result
    logging.warning("⚠️ No good matches found, using first result as fallback")
    first_item = candidates[0]
    logging.info(f"📌 Fallback: {first_item.brand} {first_item.description} {first_item.size} x{first_item.quantity}")
    return first_item

def process_products(file_path: str, column_name: str, limit: int = None):
    """
    Process products from Excel file with enhanced matching and quantity detection
    
    :param file_path: Path to Excel file
    :param column_name: Column containing product names
    :param limit: Maximum number of products to process
    """
    print("🚀 Starting Enhanced Trolley Scraper...")
    setup_logging()
    
    try:
        # Read Excel file
        print(f"📖 Reading Excel file: {file_path}")
        df = pd.read_excel(file_path)
        logging.info(f"📊 Loaded {len(df)} rows from Excel file")
        
        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in Excel file. Available columns: {list(df.columns)}")
        
        # Prepare results
        results = []
        total_products = min(len(df), limit) if limit else len(df)
        current_product = 0
        
        print(f"🎯 Processing {total_products} products...")
        
        for index, row in df.head(limit).iterrows():
            product_name = row[column_name].strip()
            
            if not product_name or pd.isna(product_name):
                logging.warning(f"Empty product name at row {index + 1}, skipping")
                continue
            
            current_product += 1
            print(f"\n{'='*60}")
            print(f"🔄 Processing {current_product}/{total_products}: {product_name}")
            print('='*60)
            logging.info(f"Processing {current_product}/{total_products}: {product_name}")
            
            # Parse SKU name for better matching
            brand, size, quantity = parse_sku_name(product_name)
            print(f"📄 Parsed - Brand: '{brand}', Size: '{size}', Quantity: {quantity}")
            logging.info(f"Parsed - Brand: '{brand}', Size: '{size}', Quantity: {quantity}")
            
            # Get search results
            search_results, search_url = get_search_results(product_name)
            
            if not search_results:
                print("❌ No search results found")
                result_data = {
                    'SKU_Name': product_name,
                    'Parsed_Brand': brand,
                    'Parsed_Size': size,
                    'Parsed_Quantity': quantity,
                    'Search_URL': search_url,
                    'Match_Status': 'No Results',
                    'Matched_Brand': 'N/A',
                    'Matched_Description': 'N/A',
                    'Matched_Size': 'N/A',
                    'Matched_Quantity': 'N/A',
                    'Matched_Price': 'N/A',
                    'Matched_URL': 'N/A',
                    'Match_Tier': 'None'
                }
                results.append(result_data)
                continue
            
            print(f"🎯 Found {len(search_results)} search results")
            
            # Find best match
            best_match = find_best_match(brand, size, quantity, search_results)
            
            if best_match:
                # Determine match tier for reporting
                brand_match = brand.lower() in best_match.brand.lower() if brand else True
                size_match = size.lower() in best_match.size.lower() if size else True
                qty_match = str(quantity) == best_match.quantity if quantity else True
                
                if brand_match and size_match and qty_match:
                    match_tier = "Tier 1 (Perfect)"
                elif brand_match and size_match:
                    match_tier = "Tier 2 (Brand+Size)"
                elif brand_match:
                    match_tier = "Tier 3 (Brand Only)"
                else:
                    match_tier = "Tier 4 (Fallback)"
                
                print(f"✅ Best match found ({match_tier}):")
                print(f"   Brand: {best_match.brand}")
                print(f"   Description: {best_match.description}")
                print(f"   Size: {best_match.size}")
                print(f"   Quantity: {best_match.quantity}")
                print(f"   Price: {best_match.price}")
                print(f"   URL: {best_match.url}")
                
                result_data = {
                    'SKU_Name': product_name,
                    'Parsed_Brand': brand,
                    'Parsed_Size': size,
                    'Parsed_Quantity': quantity,
                    'Search_URL': search_url,
                    'Match_Status': 'Matched',
                    'Matched_Brand': best_match.brand,
                    'Matched_Description': best_match.description,
                    'Matched_Size': best_match.size,
                    'Matched_Quantity': best_match.quantity,
                    'Matched_Price': best_match.price,
                    'Matched_URL': best_match.url,
                    'Match_Tier': match_tier
                }
            else:
                print("❌ No suitable match found")
                result_data = {
                    'SKU_Name': product_name,
                    'Parsed_Brand': brand,
                    'Parsed_Size': size,
                    'Parsed_Quantity': quantity,
                    'Search_URL': search_url,
                    'Match_Status': 'No Match',
                    'Matched_Brand': 'N/A',
                    'Matched_Description': 'N/A',
                    'Matched_Size': 'N/A',
                    'Matched_Quantity': 'N/A',
                    'Matched_Price': 'N/A',
                    'Matched_URL': 'N/A',
                    'Match_Tier': 'None'
                }
            
            results.append(result_data)
            
            # Rate limiting
            delay = random.uniform(1, 3)
            print(f"⏳ Waiting {delay:.1f} seconds...")
            time.sleep(delay)
        
        # Save results to CSV
        if results:
            output_filename = f"trolley_enhanced_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            results_df = pd.DataFrame(results)
            results_df.to_csv(output_filename, index=False, encoding='utf-8-sig')
            
            print(f"\n{'='*60}")
            print(f"📊 SUMMARY REPORT")
            print('='*60)
            print(f"✅ Total products processed: {len(results)}")
            print(f"📁 Results saved to: {output_filename}")
            
            # Match tier statistics
            tier_counts = results_df['Match_Tier'].value_counts()
            print(f"\n📈 Match Quality Distribution:")
            for tier, count in tier_counts.items():
                print(f"   {tier}: {count} products")
            
            logging.info(f"Processing complete. Results saved to {output_filename}")
            print(f"\n🎉 Processing complete! Check {output_filename} for detailed results.")
        
    except Exception as e:
        logging.error(f"❌ Critical error in process_products: {e}")
        print(f"❌ Error: {e}")
        raise

def interactive_mode():
    """Interactive mode with selection-based interface"""
    print("🛒 Enhanced Trolley.co.uk Scraper with Quantity Detection")
    print("="*60)
    
    # Step 1: Find and select Excel files
    current_dir = os.getcwd()
    excel_files = []
    
    # Look for Excel files in current directory
    for file in os.listdir(current_dir):
        if file.endswith(('.xlsx', '.xls')):
            excel_files.append(file)
    
    if not excel_files:
        print("❌ No Excel files found in current directory.")
        # Fallback to manual input
        while True:
            excel_file = input("📁 Please enter Excel file path: ").strip().strip('"')
            if os.path.exists(excel_file):
                break
            print("❌ File not found. Please check the path and try again.")
    else:
        print(f"\n📁 Found {len(excel_files)} Excel file(s) in current directory:")
        for i, file in enumerate(excel_files, 1):
            file_size = os.path.getsize(file) / 1024  # KB
            print(f"  {i}. {file} ({file_size:.1f} KB)")
        
        while True:
            try:
                choice = input(f"\n📝 Select Excel file (1-{len(excel_files)}) or enter 0 for custom path: ").strip()
                if choice == "0":
                    excel_file = input("📁 Enter Excel file path: ").strip().strip('"')
                    if not os.path.exists(excel_file):
                        print("❌ File not found. Please try again.")
                        continue
                    break
                else:
                    choice_num = int(choice)
                    if 1 <= choice_num <= len(excel_files):
                        excel_file = excel_files[choice_num - 1]
                        break
                    else:
                        print(f"❌ Please enter a number between 1 and {len(excel_files)}")
            except ValueError:
                print("❌ Please enter a valid number.")
    
    # Step 2: Load and analyze Excel file
    try:
        df = pd.read_excel(excel_file)
        print(f"\n✅ File loaded successfully: {excel_file}")
        print(f"📊 Total rows: {len(df)}")
        print(f"📋 Total columns: {len(df.columns)}")
        
        # Show columns with sample data
        print(f"\n� Available columns with sample data:")
        for i, col in enumerate(df.columns, 1):
            sample_values = df[col].dropna().head(2).tolist()
            sample_str = ", ".join([str(v)[:30] + "..." if len(str(v)) > 30 else str(v) for v in sample_values])
            print(f"  {i}. {col}")
            print(f"     Sample: {sample_str}")
        
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return
    
    # Step 3: Select column
    while True:
        try:
            col_choice = input(f"\n📝 Select column containing product names (1-{len(df.columns)}): ").strip()
            col_num = int(col_choice)
            if 1 <= col_num <= len(df.columns):
                column_name = df.columns[col_num - 1]
                break
            else:
                print(f"❌ Please enter a number between 1 and {len(df.columns)}")
        except ValueError:
            print("❌ Please enter a valid number.")
    
    # Step 4: Show preview of selected column
    print(f"\n👀 Preview of selected column '{column_name}':")
    non_empty_products = df[column_name].dropna()
    for i, product in enumerate(non_empty_products.head(5), 1):
        print(f"  {i}. {product}")
    if len(non_empty_products) > 5:
        print(f"  ... and {len(non_empty_products) - 5} more products")
    
    # Step 5: Processing options
    print(f"\n🎯 Processing Options:")
    print(f"  1. Process first 5 products (Quick test)")
    print(f"  2. Process first 10 products (Small batch)")
    print(f"  3. Process first 50 products (Medium batch)")
    print(f"  4. Process all {len(non_empty_products)} products (Full run)")
    print(f"  5. Custom number")
    
    while True:
        try:
            proc_choice = input(f"\nSelect processing option (1-5): ").strip()
            if proc_choice == "1":
                limit = 5
                break
            elif proc_choice == "2":
                limit = 10
                break
            elif proc_choice == "3":
                limit = 50
                break
            elif proc_choice == "4":
                limit = None
                break
            elif proc_choice == "5":
                custom_limit = input(f"Enter number of products to process (max {len(non_empty_products)}): ").strip()
                custom_num = int(custom_limit)
                if 1 <= custom_num <= len(non_empty_products):
                    limit = custom_num
                    break
                else:
                    print(f"❌ Please enter a number between 1 and {len(non_empty_products)}")
            else:
                print("❌ Please enter a number between 1 and 5")
        except ValueError:
            print("❌ Please enter a valid number.")
    
    # Step 6: Confirmation
    products_to_process = limit if limit else len(non_empty_products)
    print(f"\n📋 Processing Summary:")
    print(f"  📁 File: {excel_file}")
    print(f"  📊 Column: {column_name}")
    print(f"  🎯 Products to process: {products_to_process}")
    print(f"  ⏱️ Estimated time: {products_to_process * 2:.0f}-{products_to_process * 4:.0f} seconds")
    
    confirm = input(f"\n🚀 Start processing? (y/n): ").strip().lower()
    if confirm in ['y', 'yes', '是', '确认']:
        print(f"\n🚀 Starting to process products...")
        process_products(excel_file, column_name, limit)
    else:
        print("❌ Processing cancelled by user.")

if __name__ == "__main__":
    interactive_mode()