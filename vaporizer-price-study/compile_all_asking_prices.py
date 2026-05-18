#!/usr/bin/env python3
"""
Compile eBay asking price data from browser-extracted listings for all three agents.
Data collected on 2026-03-27 from eBay active listings.
Prices are in JPY, converted to USD at 1 USD = 149.5 JPY.
"""
import csv

JPY_TO_USD = 1 / 149.5

desflurane = [
    ("Maquet Desflurane Vaporizer SAF-T-FIL", 36841, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Drager D-Vapor Desflurane Vaporizer M35500", 55934, "Used"),
    ("Drager Medical Baxter D-Vapor Desflurane", 25568, "Used"),
    ("Datex Ohmeda Desflurane Aladin A-VDES Cassette", 11506, "Used"),
    ("DRAGER D Vapor Desflurane Vaporizer", 34588, "Used"),
    ("Desflurane D-Vapor 2000", 199763, "Refurbished"),
    ("DRAEGER DIVA Desflurane Vaporizer", 46053, "Used"),
    ("Drager D-Vapor Desflurane Vaporizer SW V3.0", 57532, "Used"),
    ("Drager D-Vapor Desflurane Vaporizer SW V2.03", 39953, "Used"),
    ("Desflurane Vaporizer Drager DIVA DM", 82896, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Desflurane Vaporizer", 71915, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane X1107-9601", 27743, "Used"),
    ("DATEX-OHMEDA TEC 6 PLUS DESFLURANE VAPORIZER", 47783, "Used"),
    ("Drager M35500-20 D-Vapor S2000 Desflurane", 79106, "Used"),
    ("DATEX OHMEDA TEC 6 DESFLURANE VAPORIZER", 12146, "Used"),
    ("Datex Ohmeda Desflurane Aladin A-VDES Cassette", 17500, "Used"),
    ("Drager D-Vapor Desflurane", 138467, "Used"),
    ("Baxter Drager D-Vapor Desflurane M35500-20", 53137, "Used"),
    ("Drager D-Vapor 3000 Desflurane", 84998, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 44018, "Used"),
    ("Baxter Drager D-Vapor Desflurane M35500-18", 41750, "Used"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE", 22373, "Used"),
    ("GE Datex Ohmeda Tec 6 Vaporizer Desflurane", 24290, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Draeger Medical D-Vapor Desflurane Vaporizer", 16141, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 63924, "Used"),
    ("Datex Ohmeda Tec 6 Desflurane Vaporizer EMPTY", 42827, "Used"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE", 38227, "Used"),
    ("Siemens Maquet Desflurane Vaporizer", 44211, "Used"),
    ("Datex Ohmeda Tec 6 Plus Vaporizer Desflurane", 22365, "Used"),
    ("DATEX-OHMEDA TEC 6 PLUS DESFLURANE VAPORIZER", 44619, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane REF 1107-9601", 17251, "Used"),
    ("DATEX-OHMEDA TEC 6 PLUS DESFLURANE VAPORIZER", 44619, "Used"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE", 22373, "Used"),
    ("DATEX-OHMEDA TEC 6 PLUS DESFLURANE VAPORIZER", 44619, "Used"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE", 22373, "Used"),
    ("Datex-Ohmeda Tec 6 Desflurane Vaporizer X1107-9001", 22947, "Used"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE", 22373, "Used"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE", 25442, "Used"),
    ("DATEX OHMEDA ALADIN 2 DESFLURANE CASSETTE", 22373, "Used"),
    ("Datex-Ohmeda Tec 6 Desflurane Vaporizer X1107-9001", 21504, "Used"),
    ("Datex Ohmeda Tec 6 Vaporizer Desflurane", 28766, "Used"),
    ("Datex-Ohmeda Tec 6 Desflurane Vaporizer X1107-9001", 32473, "Used"),
    ("Datex-Ohmeda GE Tec 6 Desflurane Vaporizer", 36080, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane 30 Day Warranty", 39953, "Used"),
    ("Draeger Medical D-Vapor 3000 Desflurane", 55294, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Draeger Medical D-Vapor 3000 Desflurane", 55294, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
    ("Datex-Ohmeda Tec 6 Plus Desflurane Vaporizer", 4475, "Used"),
]

sevoflurane = [
    ("Vaportec Sevoflurane Vaporiser", 66317, "Used"),
    ("Drager Vapor 2000 Sevoflurane M35170", 159009, "Used"),
    ("Baxter Drager Vapor 2000 Sevoflurane M35054-09", 190971, "Used"),
    ("Abbott Sevo Sevoflurane Blease Vaporiser", 89574, "Parts only"),
    ("Sevoflurane Vapor 2000", 296448, "Refurbished"),
    ("Somni Drager Sevoflurane 19.1 Vaporizer", 191770, "New-Open box"),
    ("Datex-Ohmeda Aladin 2 Sevoflurane Vaporizer", 119738, "Used"),
    ("GE DATEX-OHMEDA Sevo Tec 7 Sevoflurane Vaporizer", 115133, "Used"),
    ("Drager Vapor 19.1 Sevoflurane Vaporizer Fixed Mount", 159011, "Refurbished"),
    ("Drager M35170 Vapor 2000 Sevoflurane Vaporizer", 190973, "Used"),
    ("BAXTER GE Tec 850 Sevoflurane Vaporizer", 254214, "Used"),
    ("Baxter M35054-05 Sevoflurane Cartridge Vapor 2000", 225732, "Used"),
    ("Abbott Sevo Sevoflurane Blease Vaporiser", 132054, "Used"),
    ("Drager 19.1 Sevoflurane Anesthesia Vaporizer", 111867, "Used"),
    ("TEC 7 Sevoflurane Vaporizer", 263687, "Refurbished"),
    ("Datex Ohmeda Sevotec 5 Sevoflurane Key Fill", 263687, "Refurbished"),
    ("Datex Ohmeda Sevotec 5 Sevoflurane Funnel Fill", 190235, "Refurbished"),
    ("Drager Vapor 2000 Sevoflurane REF M35170", 230266, "Used"),
    ("Drager Vapor 2000 Sevoflurane REF M35170", 276319, "Used"),
    ("DATEX-OHMEDA SEVOTEC 5 SEVOFLURANE VAPORIZER", 377732, "Refurbished"),
    ("Draeger Medical Vapor 2000 Sevoflurane", 59290, "Used"),
]

isoflurane = [
    ("Draeger Medical Vapor 2000 Isoflurane", 31802, "Parts only"),
    ("Datex Ohmeda Isotec 4 Isoflurane Vaporizer", 55932, "Parts only"),
    ("Draeger Medical Vapor 19.3 Isoflurane", 90932, "Used"),
    ("Ohio Isoflurane Vapor Vaporizer Veterinary", 83101, "Used"),
    ("NorVap Jupiter Isoflurane Vaporizer", 143685, "New-Open box"),
    ("GE Datex Ohmeda TEC 7 Isoflurane Vaporizer", 79897, "Used"),
    ("Hoyer Blease Iso Isoflurane Vaporiser", 92106, "Used"),
    ("Isoflurane Vapor 19.1 Anesthetic Vaporizer", 39153, "Used"),
    ("PENLON Sigma Delta Vaporizer ISOFLURANE", 79544, "Used"),
    ("Norvap Luna Tec-3 Isoflurane Vaporizer Vet", 190235, "New"),
    ("Datex Ohmeda GE Aladin Isoflurane a-viso", 67919, "Parts only"),
    ("Datex Ohmeda GE Aladin Isoflurane a-viso", 54335, "Parts only"),
    ("Drager Vapor 2000 Isoflurane", 55934, "Parts only"),
    ("Datex Ohmeda Aladin 2 Isoflurane Cassette", 119738, "Used"),
    ("Draeger Medical Vapor 19.3 Isoflurane", 90932, "Used"),
    ("Draeger Medical Vapor 19.3 Isoflurane", 90932, "Used"),
    ("Datex Ohmeda Isoflurane Tec 7 Vaporizer", 57647, "Used"),
    ("Spacelabs BleaseDatum Isoflurane Vaporizer", 79905, "Parts only"),
    ("Datex-Ohmeda Tec 850 Isoflurane Vaporizer", 197046, "Used"),
    ("Drager 19.3 Vapor Isofluran", 197073, "Used"),
    ("Datex-Ohmeda Aladin A-VISO Isoflurane Cassette", 55294, "Used"),
    ("KeyMed Vapamasta 6 Isoflurane Vaporizer", 29245, "Used"),
    ("Penlon Sigma Delta Isoflurane Vaporiser", 143059, "Used"),
    ("Datex-Ohmeda Tec 850 Isoflurane Vaporizer", 197046, "Used"),
    ("Datex-Ohmeda Tec 850 Isoflurane Vaporizer", 197046, "Used"),
    ("Blease Datum Isoflurane Vaporizer", 121050, "Used"),
    ("Datex-Ohmeda Tec 5 Isoflurane Vaporizer", 69198, "Used"),
    ("Draeger Medical Vapor 19.3 Isoflurane", 90932, "Used"),
    ("Drager 19.1 Vapor Isoflurane Vaporizer", 33400, "Used"),
    ("Ohmeda Isotec 5 Isoflurane Vaporizer", 31962, "Used"),
    ("Draeger Medical Vapor 2000 Isoflurane", 45546, "Used"),
    ("Drager Isoflurane Vapor 19.2", 91891, "Used"),
    ("Dragger Isoflurane vaporizer", 40752, "Used"),
    ("Drager Isoflurane Vaporizer", 47943, "Used"),
    ("DRAGER ISOFLURANE VAPOR 19.1 VAPORIZER", 127848, "Used"),
    ("Datex Ohmeda Isotec 5 Key Fill", 95886, "Refurbished"),
    ("Drager Vapor Isoflurane 19.3 Wechselsystem", 191397, "Used"),
    ("Datex Ohmeda Isotec 4 Isofluran UNTESTED", 27632, "Parts only"),
    ("VetEquip COMPAC5 Isoflurane Vaporizer", 263317, "Used"),
]

output_path = "/home/ubuntu/vaporizer_research/data/ebay_asking_prices.csv"
with open(output_path, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['title', 'price_jpy', 'price_usd', 'condition', 'agent', 'date_collected'])
    for agent_name, listings in [('desflurane', desflurane), ('sevoflurane', sevoflurane), ('isoflurane', isoflurane)]:
        for title, price_jpy, condition in listings:
            price_usd = round(price_jpy * JPY_TO_USD, 2)
            writer.writerow([title, price_jpy, price_usd, condition, agent_name, '2026-03-27'])

print(f"Desflurane: {len(desflurane)} listings")
print(f"Sevoflurane: {len(sevoflurane)} listings")
print(f"Isoflurane: {len(isoflurane)} listings")
print(f"Total: {len(desflurane) + len(sevoflurane) + len(isoflurane)} listings")
print(f"Saved to {output_path}")
