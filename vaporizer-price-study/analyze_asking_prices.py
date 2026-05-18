#!/usr/bin/env python3
"""
Analyze eBay asking prices (active listings) as supplementary analysis.
Compares asking prices across three agents using Kruskal-Wallis test,
and calculates asking vs sold price spreads.
"""
import pandas as pd
import numpy as np
from scipy import stats as sp_stats
import json

# Load asking prices
asking = pd.read_csv('/home/ubuntu/vaporizer_research/data/ebay_asking_prices.csv')
# Load sold prices
sold = pd.read_csv('/home/ubuntu/vaporizer_research/data/combined_cleaned.csv')
sold['date_sold'] = pd.to_datetime(sold['date_sold'])

print("=" * 60)
print("SUPPLEMENTARY ANALYSIS: eBay Active Listing Prices")
print("=" * 60)
print(f"\nAsking prices collected: {asking['date_collected'].iloc[0]}")
print(f"Total active listings: {len(asking)}")
print()

# Summary by agent
print("--- Asking Price Summary by Agent ---")
results = {}
for agent in ['desflurane', 'sevoflurane', 'isoflurane']:
    sub = asking[asking['agent'] == agent]
    prices = sub['price_usd']
    agent_cap = agent.capitalize()
    results[agent_cap] = {
        'n': len(sub),
        'mean': prices.mean(),
        'sd': prices.std(),
        'median': prices.median(),
        'q25': prices.quantile(0.25),
        'q75': prices.quantile(0.75),
        'min': prices.min(),
        'max': prices.max(),
    }
    print(f"\n{agent_cap} (n={len(sub)}):")
    print(f"  Mean (SD): ${prices.mean():.0f} (${prices.std():.0f})")
    print(f"  Median (IQR): ${prices.median():.0f} (${prices.quantile(0.25):.0f}-${prices.quantile(0.75):.0f})")
    print(f"  Range: ${prices.min():.0f}-${prices.max():.0f}")

# Kruskal-Wallis test
print("\n--- Kruskal-Wallis Test (Asking Prices) ---")
des_asking = asking[asking['agent'] == 'desflurane']['price_usd']
sevo_asking = asking[asking['agent'] == 'sevoflurane']['price_usd']
iso_asking = asking[asking['agent'] == 'isoflurane']['price_usd']

kw_stat, kw_p = sp_stats.kruskal(des_asking, sevo_asking, iso_asking)
print(f"H = {kw_stat:.2f}, P = {kw_p:.6f}")

# Pairwise Mann-Whitney U tests (post-hoc)
print("\n--- Post-hoc Pairwise Mann-Whitney U Tests ---")
pairs = [
    ('Desflurane', 'Sevoflurane', des_asking, sevo_asking),
    ('Desflurane', 'Isoflurane', des_asking, iso_asking),
    ('Sevoflurane', 'Isoflurane', sevo_asking, iso_asking),
]
pairwise_results = {}
for name1, name2, g1, g2 in pairs:
    u_stat, u_p = sp_stats.mannwhitneyu(g1, g2, alternative='two-sided')
    pairwise_results[f'{name1} vs {name2}'] = {'U': u_stat, 'P': u_p}
    print(f"  {name1} vs {name2}: U={u_stat:.0f}, P={u_p:.6f}")

# Asking vs Sold price spread
print("\n--- Asking vs Recent Sold Price Spread ---")
# Use last 6 months of sold data for "recent" comparison
recent_cutoff = pd.Timestamp('2025-10-01')
spread_results = {}
for agent_cap in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    agent_lower = agent_cap.lower()
    recent_sold = sold[(sold['agent_type'] == agent_cap) & (sold['date_sold'] >= recent_cutoff)]['price_usd']
    ask_prices = asking[asking['agent'] == agent_lower]['price_usd']
    
    ask_median = ask_prices.median()
    if len(recent_sold) > 0:
        sold_median = recent_sold.median()
        spread = ask_median - sold_median
        spread_pct = (spread / sold_median) * 100
    else:
        sold_median = sold[sold['agent_type'] == agent_cap]['price_usd'].median()
        spread = ask_median - sold_median
        spread_pct = (spread / sold_median) * 100
    
    spread_results[agent_cap] = {
        'asking_median': ask_median,
        'sold_median': sold_median,
        'spread': spread,
        'spread_pct': spread_pct,
        'recent_sold_n': len(recent_sold) if len(recent_sold) > 0 else len(sold[sold['agent_type'] == agent_cap]),
    }
    print(f"\n{agent_cap}:")
    print(f"  Asking median: ${ask_median:.0f}")
    print(f"  Recent sold median: ${sold_median:.0f}")
    print(f"  Spread: ${spread:.0f} ({spread_pct:+.0f}%)")

# Save results
all_results = {
    'asking_summary': results,
    'kruskal_wallis': {'H': kw_stat, 'P': kw_p},
    'pairwise': {k: {'U': float(v['U']), 'P': float(v['P'])} for k, v in pairwise_results.items()},
    'spread': {k: {kk: float(vv) for kk, vv in v.items()} for k, v in spread_results.items()},
}

with open('/home/ubuntu/vaporizer_research/data/asking_price_analysis.json', 'w') as f:
    json.dump(all_results, f, indent=2)

print("\n\nResults saved to data/asking_price_analysis.json")

# Print key findings for paper integration
print("\n" + "=" * 60)
print("KEY FINDINGS FOR PAPER")
print("=" * 60)
print(f"\nAsking prices (n={len(asking)}): desflurane n={len(des_asking)}, "
      f"sevoflurane n={len(sevo_asking)}, isoflurane n={len(iso_asking)}")
print(f"\nDesflurane asking median: ${results['Desflurane']['median']:.0f} "
      f"vs Sevoflurane: ${results['Sevoflurane']['median']:.0f} "
      f"vs Isoflurane: ${results['Isoflurane']['median']:.0f}")
print(f"\nKruskal-Wallis: H={kw_stat:.2f}, P={'<0.001' if kw_p < 0.001 else f'{kw_p:.3f}'}")
print(f"\nDesflurane asking-sold spread: {spread_results['Desflurane']['spread_pct']:+.0f}%")
print(f"Sevoflurane asking-sold spread: {spread_results['Sevoflurane']['spread_pct']:+.0f}%")
print(f"Isoflurane asking-sold spread: {spread_results['Isoflurane']['spread_pct']:+.0f}%")
