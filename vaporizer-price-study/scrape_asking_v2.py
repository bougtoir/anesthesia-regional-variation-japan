#!/usr/bin/env python3
"""Scrape eBay active listings using headless Selenium with improved selectors."""

import csv
import time
import re
import json
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

def setup_driver():
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(options=options)
    driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
        'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
    })
    driver.implicitly_wait(5)
    return driver

def parse_price(price_text):
    """Extract numeric price from text."""
    match = re.search(r'[\$]\s*([\d,]+\.?\d*)', price_text)
    if match:
        return float(match.group(1).replace(',', ''))
    return None

def scrape_query(driver, query, agent_type):
    """Scrape eBay listings for a given query."""
    listings = []
    page = 1
    
    while True:
        url = f'https://www.ebay.com/sch/i.html?_nkw={query.replace(" ", "+")}&_sacat=0&_pgn={page}&rt=nc&LH_BIN=1'
        print(f"  Page {page}: {url}")
        driver.get(url)
        time.sleep(4)
        
        # Save page source for debugging on first query
        if page == 1 and agent_type == 'Desflurane':
            with open('/home/ubuntu/vaporizer_research/data/debug_page.html', 'w') as f:
                f.write(driver.page_source)
            print(f"  Saved debug page source")
        
        # Try multiple selectors for items
        items = []
        for selector in [
            'ul.srp-results li.s-item',
            'li.s-item',
            'div.s-item__wrapper',
            'div[data-view]',
            '.srp-river-results li',
        ]:
            items = driver.find_elements(By.CSS_SELECTOR, selector)
            if items:
                print(f"  Found {len(items)} items with selector: {selector}")
                break
        
        if not items:
            # Try via page source parsing
            source = driver.page_source
            # Look for JSON-LD or structured data
            price_matches = re.findall(r'"price":\s*"?([\d.]+)"?', source)
            title_matches = re.findall(r'"name":\s*"([^"]+)"', source)
            print(f"  Regex fallback: found {len(price_matches)} prices, {len(title_matches)} titles in source")
            
            if not price_matches:
                print(f"  No items found on page {page}")
                break
        
        page_count = 0
        for item in items:
            try:
                text = item.text
                if not text or 'Shop on eBay' in text:
                    continue
                
                lines = text.split('\n')
                title = ''
                price = None
                condition = ''
                listing_type = 'Buy It Now'
                
                for line in lines:
                    line = line.strip()
                    if not title and line and '$' not in line and not line.startswith('Brand') and not line.startswith('Free') and len(line) > 10:
                        title = line
                    if '$' in line and price is None:
                        price = parse_price(line)
                    if 'Pre-Owned' in line or 'New' in line or 'Used' in line or 'Refurbished' in line:
                        condition = line
                    if 'bid' in line.lower():
                        listing_type = 'Auction'
                    if 'Best Offer' in line:
                        listing_type = 'Best Offer'
                
                if title and price and price > 10:
                    listings.append({
                        'title': title,
                        'price_usd': price,
                        'listing_type': listing_type,
                        'condition': condition,
                        'agent_type': agent_type,
                        'scrape_date': datetime.now().strftime('%Y-%m-%d'),
                    })
                    page_count += 1
            except Exception as e:
                continue
        
        print(f"  Parsed {page_count} valid listings on page {page}")
        
        if page_count == 0:
            break
        
        # Check for next page
        try:
            next_btns = driver.find_elements(By.CSS_SELECTOR, 'a.pagination__next, a[aria-label="Next page"]')
            if not next_btns:
                break
        except:
            break
        
        page += 1
        if page > 5:
            break
        time.sleep(3)
    
    return listings

def main():
    driver = setup_driver()
    all_listings = []
    
    # Use broader queries too
    queries = [
        ('desflurane vaporizer', 'Desflurane'),
        ('Tec 6 vaporizer', 'Desflurane'),
        ('D-Vapor desflurane', 'Desflurane'),
        ('sevoflurane vaporizer', 'Sevoflurane'),
        ('Vapor 2000 sevoflurane', 'Sevoflurane'),
        ('Sevotec vaporizer', 'Sevoflurane'),
        ('isoflurane vaporizer', 'Isoflurane'),
        ('Vapor 2000 isoflurane', 'Isoflurane'),
        ('Isotec vaporizer', 'Isoflurane'),
    ]
    
    seen_titles = set()
    
    for query, agent_type in queries:
        print(f"\n=== {query} ({agent_type}) ===")
        listings = scrape_query(driver, query, agent_type)
        for l in listings:
            # Deduplicate by title+price
            key = (l['title'][:50], l['price_usd'])
            if key not in seen_titles:
                seen_titles.add(key)
                all_listings.append(l)
        print(f"  Unique total so far: {len(all_listings)}")
        time.sleep(2)
    
    driver.quit()
    
    # Save
    outpath = '/home/ubuntu/vaporizer_research/data/ebay_asking_prices.csv'
    if all_listings:
        fieldnames = ['title', 'price_usd', 'listing_type', 'condition', 'agent_type', 'scrape_date']
        with open(outpath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(all_listings)
        print(f"\nSaved {len(all_listings)} listings to {outpath}")
    
    # Summary
    print("\n=== Summary ===")
    for agent_type in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
        agent_listings = [l for l in all_listings if l['agent_type'] == agent_type]
        if agent_listings:
            prices = [l['price_usd'] for l in agent_listings]
            prices.sort()
            median = prices[len(prices)//2]
            mean = sum(prices)/len(prices)
            print(f"{agent_type}: n={len(prices)}, median=${median:.0f}, mean=${mean:.0f}, range=${min(prices):.0f}-${max(prices):.0f}")
        else:
            print(f"{agent_type}: No listings found")

if __name__ == '__main__':
    main()
