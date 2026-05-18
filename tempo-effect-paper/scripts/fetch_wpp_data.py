"""
Fetch UN WPP 2024 data for OECD countries + China + DRC
Indicators: Total population, TFR, Life expectancy, ASFR (5-year), MAC
"""
import requests
import json
import time
import os

base_url = "https://population.un.org/dataportalapi/api/v1"

# Country IDs
countries = {
    "Australia": 36, "Austria": 40, "Belgium": 56, "Canada": 124, "Chile": 152,
    "China": 156, "Colombia": 170, "Costa Rica": 188, "Czechia": 203,
    "DRC": 180, "Denmark": 208, "Estonia": 233, "Finland": 246, "France": 250,
    "Germany": 276, "Greece": 300, "Hungary": 348, "Iceland": 352,
    "Ireland": 372, "Israel": 376, "Italy": 380, "Japan": 392,
    "Latvia": 428, "Lithuania": 440, "Luxembourg": 442, "Mexico": 484,
    "Netherlands": 528, "New Zealand": 554, "Norway": 578, "Poland": 616,
    "Portugal": 620, "Republic of Korea": 410, "Slovakia": 703,
    "Slovenia": 705, "Spain": 724, "Sweden": 752, "Switzerland": 756,
    "Türkiye": 792, "United Kingdom": 826, "United States": 840
}

loc_ids = ",".join(str(v) for v in countries.values())

# Indicators to fetch
indicators = {
    "TPopulation": 49,   # Total population by sex
    "TFR5": 19,          # Total fertility rate (5-year)
    "E0": 61,            # Life expectancy at birth
    "MAC5": 18,          # Mean age at childbearing (5-year)
    "ASFR5": 17,         # Age-specific fertility rates (5-year)
}

os.makedirs("/home/ubuntu/wpp_data", exist_ok=True)

def fetch_indicator(ind_name, ind_id, extra_params=None):
    """Fetch all pages of data for an indicator"""
    all_data = []
    page = 1
    params = {
        "locationIds": loc_ids,
        "indicatorIds": str(ind_id),
        "startYear": 1950,
        "endYear": 2100,
        "pageSize": 100,
        "page": page,
    }
    if extra_params:
        params.update(extra_params)
    
    retries = 0
    while True:
        params["page"] = page
        try:
            resp = requests.get(f"{base_url}/data/indicators/{ind_id}", params=params, timeout=30)
            if resp.status_code != 200:
                print(f"  Error {resp.status_code} on page {page}: {resp.text[:200]}")
                break
            result = resp.json()
            items = result.get("data", [])
            if not items:
                break
            all_data.extend(items)
            retries = 0
            total_pages = result.get("pages", 1)
            print(f"  {ind_name}: page {page}/{total_pages}, got {len(items)} records (total so far: {len(all_data)})")
            if page >= total_pages:
                break
            page += 1
            time.sleep(0.3)
        except Exception as e:
            retries += 1
            print(f"  Error on page {page} (attempt {retries}): {e}")
            if retries >= 5:
                print(f"  Max retries reached on page {page}, skipping.")
                break
            time.sleep(2)
            continue
    
    return all_data

# Fetch each indicator
for ind_name, ind_id in indicators.items():
    print(f"\nFetching {ind_name} (ID={ind_id})...")
    extra = {}
    if ind_name in ["TPopulation", "E0"]:
        extra["sexId"] = "3"  # Both sexes
    data = fetch_indicator(ind_name, ind_id, extra)
    
    outfile = f"/home/ubuntu/wpp_data/{ind_name}.json"
    with open(outfile, "w") as f:
        json.dump(data, f)
    print(f"  Saved {len(data)} records to {outfile}")

print("\nDone fetching all indicators!")
