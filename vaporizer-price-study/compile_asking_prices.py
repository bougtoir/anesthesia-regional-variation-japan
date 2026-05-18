#!/usr/bin/env python3
"""
DEPRECATED: This script only compiles desflurane listings.
Use compile_all_asking_prices.py instead, which includes all three agents
(desflurane, sevoflurane, isoflurane) in a single dataset.

Original purpose:
Compile eBay asking price data from browser-extracted listings.
Data collected on 2026-03-27 from eBay active listings.
Prices are in JPY, converted to USD at 1 USD = 149.5 JPY (approximate rate).
"""
import csv
import os

# Exchange rate JPY to USD (approximate as of March 2026)
JPY_TO_USD = 1 / 149.5

# Data extracted from eBay browser on 2026-03-27
# Format: (title, price_jpy, condition, agent)
# Excluding: batteries, accessories, non-vaporizer items

desflurane_listings = [
    ("Maquet Desflurane Vaporizer, SAF-T-FIL", 36841, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Drager Anesthesia D-Vapor Desflurane Vaporizer M35500", 55934, "Used", "desflurane"),
    ("Drager Medical/Baxter M35500-18 D-Vapor Desflurane Vaporizer", 25568, "Used", "desflurane"),
    ("Datex Ohmeda Desflurane Aladin A-VDES Cassette Vaporizer", 11506, "Used", "desflurane"),
    ("DRAGER D Vapor Desflurane Vaporizer", 34588, "Used", "desflurane"),
    ("Desflurane D-Vapor 2000 (100V-240V)", 199763, "Refurbished", "desflurane"),
    ("DRAEGER DIVA Desflurane Vaporizer", 46053, "Used", "desflurane"),
    ("Drager D-Vapor Desflurane Vaporizer SW V3.0", 57532, "Used", "desflurane"),
    ("Drager D-Vapor Desflurane Vaporizer SW V2.03", 39953, "Used", "desflurane"),
    ("Desflurane Vaporizer Drager DIVA DM", 82896, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Desflurane Vaporizer", 71915, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer X1107-9601-000", 27743, "Used", "desflurane"),
    ("DATEX-OHMEDA TEC 6 PLUS DESFLURANE ANESTHESIA VAPORIZER", 47783, "Used", "desflurane"),
    ("Drager M35500-20 D-Vapor S2000 Desflurane Vaporizer", 79106, "Used", "desflurane"),
    ("DATEX OHMEDA TEC 6 DESFLURANE ANESTHESIA VAPORIZER", 12146, "Used", "desflurane"),
    ("Datex Ohmeda Desflurane Aladin A-VDES Cassette Vaporizer", 17500, "Used", "desflurane"),
    ("Drager D-Vapor Desflurane", 138467, "Used", "desflurane"),
    ("Baxter Drager D-Vapor Desflurane Vaporizer M35500-20", 53137, "Used", "desflurane"),
    ("Drager D-Vapor 3000 Desflurane", 84998, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 44018, "Used", "desflurane"),
    ("Baxter Drager D-Vapor Desflurane Vaporizer M35500-18", 41750, "Used", "desflurane"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE VAPORIZER", 22373, "Used", "desflurane"),
    ("GE Datex Ohmeda Tec 6 Anesthesia Vaporizer Desflurane", 24290, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Draeger Medical D-Vapor Desflurane Vaporizer", 16141, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Anesthesia Vaporizer", 63924, "Used", "desflurane"),
    ("Datex Ohmeda Tec 6 Desflurane Anaesthesia Vaporizer EMPTY", 42827, "Used", "desflurane"),
    ("DATEX OHMEDA 1100-9026-000 ALADIN 2 DESFLURANE CASSETTE VAPORIZER", 38227, "Used", "desflurane"),
    ("Siemens Maquet Desflurane Vaporizer", 44211, "Used", "desflurane"),
    ("Datex Ohmeda Tec 6 Plus Anes Vaporizer (Desflurane)", 22365, "Used", "desflurane"),
    ("DATEX-OHMEDA TEC 6 PLUS DESFLURANE ANESTHESIA VAPORIZER", 44619, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer REF 1107-9601-000", 17251, "Used", "desflurane"),
    ("DATEX-OHMEDA TEC 6 PLUS DESFLURANE ANESTHESIA VAPORIZER", 44619, "Used", "desflurane"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE VAPORIZER", 22373, "Used", "desflurane"),
    ("DATEX-OHMEDA TEC 6 PLUS DESFLURANE ANESTHESIA VAPORIZER", 44619, "Used", "desflurane"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE VAPORIZER", 22373, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Desflurane Vaporizer ref:X1107-9001-000", 22947, "Used", "desflurane"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE VAPORIZER", 22373, "Used", "desflurane"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE VAPORIZER", 25442, "Used", "desflurane"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE VAPORIZER", 22373, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Desflurane Vaporizer ref:X1107-9001-000", 21504, "Used", "desflurane"),
    ("Datex Ohmeda Tec 6 Vaporizer Desflurane", 28766, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Desflurane Vaporizer ref:X1107-9001-000", 32473, "Used", "desflurane"),
    ("Datex-Ohmeda (GE) Tec 6 Desflurane Vaporizer ref:X1107-9001-000", 36080, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer 30 Day Warranty", 39953, "Used", "desflurane"),
    ("Draeger Medical D-Vapor 3000 Desflurane Vaporizer", 55294, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Draeger Medical D-Vapor 3000 Desflurane Vaporizer", 55294, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used", "desflurane"),
]
# Note: Excluded battery listings (Battery 450mAh BATT/110451 items) as they are accessories, not vaporizers

# Now write all data to CSV
output_path = "/home/ubuntu/vaporizer_research/data/ebay_asking_prices.csv"
with open(output_path, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['title', 'price_jpy', 'price_usd', 'condition', 'agent', 'date_collected'])
    
    for title, price_jpy, condition, agent in desflurane_listings:
        price_usd = round(price_jpy * JPY_TO_USD, 2)
        writer.writerow([title, price_jpy, price_usd, condition, agent, '2026-03-27'])

print(f"Desflurane: {len(desflurane_listings)} listings written")
print(f"File saved to {output_path}")
