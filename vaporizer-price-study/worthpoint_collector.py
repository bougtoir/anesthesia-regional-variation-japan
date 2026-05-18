"""
WorthPoint Data Collection Script for Anaesthetic Vaporizer Research
====================================================================
Run this script on your local PC after logging into WorthPoint in your browser.

Prerequisites:
    pip install selenium pandas webdriver-manager

Usage:
    1. Open Chrome and log into WorthPoint (www.worthpoint.com) manually
    2. Close Chrome
    3. Run this script: python worthpoint_collector.py
    4. The script will open a new Chrome window using your existing profile
    5. Data will be saved to worthpoint_data.csv in the same directory

Note: This script uses your local Chrome profile to bypass bot detection.
      You must be logged into WorthPoint before running.
"""

import time
import csv
import re
import os
from datetime import datetime

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, NoSuchElementException
except ImportError:
    print("Please install required packages:")
    print("  pip install selenium")
    exit(1)

try:
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("Please install webdriver-manager:")
    print("  pip install webdriver-manager")
    ChromeDriverManager = None


# ==========================================
# Configuration
# ==========================================
SEARCH_QUERIES = [
    "desflurane vaporizer",
    "desflurane tec 6",
    "desflurane d-vapor",
    "sevoflurane vaporizer",
    "sevoflurane vapor 2000",
    "sevoflurane tec 7",
    "isoflurane vaporizer",
    "isoflurane vapor 2000",
    "isoflurane tec 5",
    "isoflurane isotec",
]

OUTPUT_FILE = "worthpoint_data.csv"
MAX_PAGES_PER_QUERY = 10  # Max pages to scrape per search query
DELAY_BETWEEN_PAGES = 3   # Seconds between page loads (be polite)
DELAY_BETWEEN_ITEMS = 1   # Seconds between item detail pages


def get_chrome_profile_path():
    """Get the default Chrome profile path based on OS."""
    import platform
    system = platform.system()
    home = os.path.expanduser("~")

    if system == "Windows":
        return os.path.join(home, "AppData", "Local", "Google", "Chrome", "User Data")
    elif system == "Darwin":  # macOS
        return os.path.join(home, "Library", "Application Support", "Google", "Chrome")
    else:  # Linux
        return os.path.join(home, ".config", "google-chrome")


def setup_driver():
    """Set up Chrome WebDriver with user's existing profile."""
    options = Options()

    # Use existing Chrome profile (keeps login session)
    profile_path = get_chrome_profile_path()
    if os.path.exists(profile_path):
        options.add_argument(f"--user-data-dir={profile_path}")
        options.add_argument("--profile-directory=Default")
        print(f"Using Chrome profile: {profile_path}")
    else:
        print("WARNING: Chrome profile not found. You may need to log in manually.")

    # Standard options
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    if ChromeDriverManager:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    else:
        driver = webdriver.Chrome(options=options)

    # Remove webdriver flag
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    return driver


def extract_price(text):
    """Extract numeric price from text like '$1,234.56'."""
    if not text:
        return None
    match = re.search(r'\$?([\d,]+\.?\d*)', text.replace(',', ''))
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None


def extract_date(text):
    """Extract date from various formats."""
    if not text:
        return None
    # Try common date patterns
    patterns = [
        (r'(\w+ \d{1,2}, \d{4})', '%B %d, %Y'),      # "January 15, 2024"
        (r'(\d{1,2}/\d{1,2}/\d{4})', '%m/%d/%Y'),      # "01/15/2024"
        (r'(\d{4}-\d{2}-\d{2})', '%Y-%m-%d'),           # "2024-01-15"
        (r'(\w+ \d{4})', '%B %Y'),                       # "January 2024"
    ]
    for pattern, fmt in patterns:
        match = re.search(pattern, text)
        if match:
            try:
                return datetime.strptime(match.group(1), fmt).strftime('%Y-%m-%d')
            except ValueError:
                continue
    return text  # Return raw text if no pattern matches


def classify_agent(title):
    """Classify vaporizer type from title."""
    title_lower = title.lower()
    if any(w in title_lower for w in ['desflurane', 'des ', 'tec 6', 'tec6', 'd-vapor', 'dvapor']):
        return 'Desflurane'
    elif any(w in title_lower for w in ['sevoflurane', 'sevo', 'tec 7', 'tec7', 'sevotec']):
        return 'Sevoflurane'
    elif any(w in title_lower for w in ['isoflurane', 'iso ', 'tec 5', 'tec5', 'isotec', 'tec 3', 'tec3']):
        return 'Isoflurane'
    return 'Unknown'


def scrape_search_results(driver, query):
    """Scrape search results for a given query."""
    results = []
    base_url = f"https://www.worthpoint.com/worthopedia?query={query.replace(' ', '+')}&sort=date_desc"

    for page in range(1, MAX_PAGES_PER_QUERY + 1):
        url = f"{base_url}&page={page}" if page > 1 else base_url
        print(f"  Page {page}: {url}")

        try:
            driver.get(url)
            time.sleep(DELAY_BETWEEN_PAGES)

            # Wait for results to load
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.search-results, div.worthopedia-results, div[class*='result']"))
            )
        except TimeoutException:
            print(f"  Timeout on page {page}, moving to next query")
            break

        # Try to find result items
        items = []
        for selector in [
            "div.search-result-item",
            "div[class*='SearchResult']",
            "div.worthopedia-item",
            "a[href*='/worthopedia/']",
            "div[class*='result'] a[href*='/worthopedia/']",
        ]:
            items = driver.find_elements(By.CSS_SELECTOR, selector)
            if items:
                break

        if not items:
            print(f"  No items found on page {page}, stopping")
            break

        print(f"  Found {len(items)} items")

        for item in items:
            try:
                # Extract basic info from search result
                title = ""
                price = ""
                date = ""
                link = ""

                # Try to get title
                for sel in ["h3", "h4", ".title", "[class*='title']", "a"]:
                    try:
                        title_elem = item.find_element(By.CSS_SELECTOR, sel)
                        title = title_elem.text.strip()
                        if title:
                            break
                    except NoSuchElementException:
                        continue

                # Try to get link
                try:
                    if item.tag_name == 'a':
                        link = item.get_attribute('href')
                    else:
                        link_elem = item.find_element(By.CSS_SELECTOR, "a[href*='/worthopedia/']")
                        link = link_elem.get_attribute('href')
                except NoSuchElementException:
                    pass

                # Try to get price
                for sel in [".price", "[class*='price']", "[class*='Price']", "span.amount"]:
                    try:
                        price_elem = item.find_element(By.CSS_SELECTOR, sel)
                        price = price_elem.text.strip()
                        if price:
                            break
                    except NoSuchElementException:
                        continue

                # Try to get date
                for sel in [".date", "[class*='date']", "[class*='Date']", "time", ".sold-date"]:
                    try:
                        date_elem = item.find_element(By.CSS_SELECTOR, sel)
                        date = date_elem.text.strip()
                        if date:
                            break
                    except NoSuchElementException:
                        continue

                if title:
                    agent_type = classify_agent(title)
                    results.append({
                        'date_sold': extract_date(date) or '',
                        'title': title,
                        'price_usd': extract_price(price) or '',
                        'condition': '',
                        'brand': '',
                        'model': '',
                        'agent_type': agent_type,
                        'source': 'WorthPoint',
                        'url': link,
                        'search_query': query,
                    })

            except Exception as e:
                print(f"  Error extracting item: {e}")
                continue

        # Check if there's a next page
        try:
            next_btn = driver.find_element(By.CSS_SELECTOR, "a[rel='next'], button[aria-label='Next'], .pagination .next")
            if not next_btn.is_enabled():
                break
        except NoSuchElementException:
            break

    return results


def get_item_details(driver, url):
    """Get detailed information from an individual WorthPoint listing page."""
    try:
        driver.get(url)
        time.sleep(DELAY_BETWEEN_ITEMS)

        details = {}

        # Try to get sold price
        for sel in [".sold-price", "[class*='sold'] [class*='price']", ".price-value"]:
            try:
                elem = driver.find_element(By.CSS_SELECTOR, sel)
                details['price'] = extract_price(elem.text)
                break
            except NoSuchElementException:
                continue

        # Try to get sold date
        for sel in [".sold-date", "[class*='sold'] [class*='date']", "time[datetime]"]:
            try:
                elem = driver.find_element(By.CSS_SELECTOR, sel)
                if elem.get_attribute('datetime'):
                    details['date'] = elem.get_attribute('datetime')[:10]
                else:
                    details['date'] = extract_date(elem.text)
                break
            except NoSuchElementException:
                continue

        # Try to get condition
        for sel in ["[class*='condition']", "[class*='Condition']"]:
            try:
                elem = driver.find_element(By.CSS_SELECTOR, sel)
                details['condition'] = elem.text.strip()
                break
            except NoSuchElementException:
                continue

        return details

    except Exception as e:
        print(f"  Error getting details: {e}")
        return {}


def main():
    print("=" * 60)
    print("WorthPoint Vaporizer Data Collection Script")
    print("=" * 60)
    print()
    print("IMPORTANT: Make sure you are logged into WorthPoint")
    print("in Chrome before running this script.")
    print()

    input("Press Enter to start...")

    print("\nStarting Chrome...")
    driver = setup_driver()

    all_results = []

    try:
        # First, verify we're logged in
        driver.get("https://www.worthpoint.com")
        time.sleep(3)

        # Check if we need to log in
        try:
            login_btn = driver.find_element(By.CSS_SELECTOR, "a[href*='login'], button[class*='login']")
            print("\nWARNING: You don't appear to be logged in.")
            print("Please log in manually in the browser window, then press Enter to continue.")
            input("Press Enter after logging in...")
        except NoSuchElementException:
            print("Logged in successfully!")

        # Scrape each search query
        for i, query in enumerate(SEARCH_QUERIES):
            print(f"\n[{i+1}/{len(SEARCH_QUERIES)}] Searching: '{query}'")
            results = scrape_search_results(driver, query)
            all_results.extend(results)
            print(f"  Collected {len(results)} results (total: {len(all_results)})")

        # Optionally get detailed info for items with missing data
        print(f"\nCollected {len(all_results)} total results")

        items_needing_details = [r for r in all_results if not r['price_usd'] and r['url']]
        if items_needing_details:
            print(f"\nFetching details for {len(items_needing_details)} items with missing prices...")
            for i, item in enumerate(items_needing_details[:50]):  # Limit to 50
                print(f"  [{i+1}/{min(len(items_needing_details), 50)}] {item['url'][:80]}...")
                details = get_item_details(driver, item['url'])
                if details.get('price'):
                    item['price_usd'] = details['price']
                if details.get('date'):
                    item['date_sold'] = details['date']
                if details.get('condition'):
                    item['condition'] = details['condition']

    except Exception as e:
        print(f"\nError: {e}")
    finally:
        driver.quit()

    # Remove duplicates based on URL
    seen_urls = set()
    unique_results = []
    for r in all_results:
        if r['url'] not in seen_urls:
            seen_urls.add(r['url'])
            unique_results.append(r)

    # Save to CSV
    if unique_results:
        fieldnames = ['date_sold', 'title', 'price_usd', 'condition', 'brand', 'model',
                       'agent_type', 'source', 'url', 'search_query']

        with open(OUTPUT_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(unique_results)

        print(f"\nSaved {len(unique_results)} unique results to {OUTPUT_FILE}")
        print(f"\nBreakdown by agent type:")
        for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane', 'Unknown']:
            count = sum(1 for r in unique_results if r['agent_type'] == agent)
            if count > 0:
                print(f"  {agent}: {count}")
    else:
        print("\nNo results collected.")

    print("\nDone! Please send the generated CSV file for analysis.")


if __name__ == '__main__':
    main()
