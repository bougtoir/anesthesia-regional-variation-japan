#!/usr/bin/env python3
"""Scrape eBay active listings (asking prices) for desflurane, sevoflurane, isoflurane vaporizers."""

import csv
import time
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def setup_driver():
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(10)
    return driver

def parse_price(price_text):
    """Extract numeric price from text like '$1,234.56' or 'US $1,234.56'."""
    match = re.search(r'[\$£€]\s*([\d,]+\.?\d*)', price_text)
    if match:
        return float(match.group(1).replace(',', ''))
    # Try without currency symbol
    match = re.search(r'([\d,]+\.\d{2})', price_text)
    if match:
        return float(match.group(1).replace(',', ''))
    return None

def scrape_ebay_listings(driver, query, agent_type):
    """Scrape all active eBay listings for a given query."""
    listings = []
    page = 1
    
    while True:
        # eBay search URL for active (Buy It Now + Auction) listings
        url = f'https://www.ebay.com/sch/i.html?_nkw={query.replace(" ", "+")}&_sacat=0&_pgn={page}&_ipg=240'
        print(f"  Scraping page {page}: {url}")
        
        try:
            driver.get(url)
            time.sleep(3)
            
            # Check for results
            items = driver.find_elements(By.CSS_SELECTOR, 'li.s-item')
            
            if not items:
                print(f"  No items found on page {page}, stopping.")
                break
            
            page_count = 0
            for item in items:
                try:
                    # Skip "Shop on eBay" or ad items
                    try:
                        title_el = item.find_element(By.CSS_SELECTOR, '.s-item__title span[role="heading"]')
                        title = title_el.text.strip()
                    except:
                        try:
                            title_el = item.find_element(By.CSS_SELECTOR, '.s-item__title')
                            title = title_el.text.strip()
                        except:
                            continue
                    
                    if not title or title == 'Shop on eBay' or 'shop on ebay' in title.lower():
                        continue
                    
                    # Get price
                    try:
                        price_el = item.find_element(By.CSS_SELECTOR, '.s-item__price')
                        price_text = price_el.text.strip()
                    except:
                        continue
                    
                    price = parse_price(price_text)
                    if price is None or price <= 0:
                        continue
                    
                    # Get listing type (auction vs buy it now)
                    listing_type = 'Unknown'
                    try:
                        format_el = item.find_element(By.CSS_SELECTOR, '.s-item__purchaseOptions, .s-item__formatBuyItNow, .s-item__bids')
                        format_text = format_el.text.strip().lower()
                        if 'buy it now' in format_text or 'buy now' in format_text:
                            listing_type = 'Buy It Now'
                        elif 'bid' in format_text:
                            listing_type = 'Auction'
                        elif 'best offer' in format_text:
                            listing_type = 'Best Offer'
                    except:
                        pass
                    
                    # Get condition
                    condition = 'Unknown'
                    try:
                        cond_el = item.find_element(By.CSS_SELECTOR, '.SECONDARY_INFO')
                        condition = cond_el.text.strip()
                    except:
                        pass
                    
                    # Get location
                    location = 'Unknown'
                    try:
                        loc_el = item.find_element(By.CSS_SELECTOR, '.s-item__location')
                        location = loc_el.text.strip()
                    except:
                        pass
                    
                    # Get shipping
                    shipping = 'Unknown'
                    try:
                        ship_el = item.find_element(By.CSS_SELECTOR, '.s-item__shipping, .s-item__freeXDays')
                        shipping = ship_el.text.strip()
                    except:
                        pass
                    
                    # Get link
                    link = ''
                    try:
                        link_el = item.find_element(By.CSS_SELECTOR, '.s-item__link')
                        link = link_el.get_attribute('href')
                    except:
                        pass
                    
                    listings.append({
                        'title': title,
                        'price_usd': price,
                        'price_text': price_text,
                        'listing_type': listing_type,
                        'condition': condition,
                        'location': location,
                        'shipping': shipping,
                        'agent_type': agent_type,
                        'query': query,
                        'scrape_date': datetime.now().strftime('%Y-%m-%d'),
                        'url': link,
                    })
                    page_count += 1
                    
                except Exception as e:
                    continue
            
            print(f"  Page {page}: {page_count} listings found")
            
            if page_count == 0:
                break
            
            # Check if there's a next page
            try:
                next_btn = driver.find_element(By.CSS_SELECTOR, 'a.pagination__next')
                if not next_btn.is_enabled():
                    break
            except:
                break
            
            page += 1
            time.sleep(2)
            
        except Exception as e:
            print(f"  Error on page {page}: {e}")
            break
    
    return listings

def main():
    driver = setup_driver()
    all_listings = []
    
    queries = [
        ('desflurane vaporizer', 'Desflurane'),
        ('sevoflurane vaporizer', 'Sevoflurane'),
        ('isoflurane vaporizer', 'Isoflurane'),
    ]
    
    for query, agent_type in queries:
        print(f"\n=== Scraping: {query} ({agent_type}) ===")
        listings = scrape_ebay_listings(driver, query, agent_type)
        all_listings.extend(listings)
        print(f"  Total for {agent_type}: {len(listings)} listings")
        time.sleep(3)
    
    driver.quit()
    
    # Save to CSV
    outpath = '/home/ubuntu/vaporizer_research/data/ebay_asking_prices.csv'
    if all_listings:
        fieldnames = ['title', 'price_usd', 'price_text', 'listing_type', 'condition',
                      'location', 'shipping', 'agent_type', 'query', 'scrape_date', 'url']
        with open(outpath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(all_listings)
        print(f"\nSaved {len(all_listings)} listings to {outpath}")
    else:
        print("\nNo listings found!")
    
    # Summary
    print("\n=== Summary ===")
    for agent_type in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
        agent_listings = [l for l in all_listings if l['agent_type'] == agent_type]
        if agent_listings:
            prices = [l['price_usd'] for l in agent_listings]
            print(f"{agent_type}: n={len(prices)}, median=${sorted(prices)[len(prices)//2]:.0f}, "
                  f"mean=${sum(prices)/len(prices):.0f}, range=${min(prices):.0f}-${max(prices):.0f}")
        else:
            print(f"{agent_type}: No listings found")

if __name__ == '__main__':
    main()
