import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns
from scipy import stats
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Set publication-quality style
plt.rcParams.update({
    'font.size': 11,
    'font.family': 'sans-serif',
    'axes.labelsize': 12,
    'axes.titlesize': 13,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'legend.fontsize': 10,
    'figure.dpi': 300,
    'savefig.dpi': 300,
    'savefig.bbox': 'tight',
})
sns.set_style("whitegrid")

# ==========================================
# KEY REGULATORY TIMELINE
# ==========================================
# Multiple inflection points for EU desflurane regulation
regulatory_events = {
    'commission_proposal': {
        'date': pd.Timestamp('2022-04-05'),
        'label': 'EC Proposal',
        'description': 'European Commission publishes revised F-gas Regulation proposal',
        'color': '#95A5A6',
        'linestyle': ':',
    },
    'parliament_vote': {
        'date': pd.Timestamp('2023-03-30'),
        'label': 'EP Vote',
        'description': 'European Parliament approves F-gas proposal in plenary',
        'color': '#F39C12',
        'linestyle': '-.',
    },
    'provisional_agreement': {
        'date': pd.Timestamp('2023-10-05'),
        'label': 'Trilogue\nAgreement',
        'description': 'Council and Parliament reach provisional agreement',
        'color': '#E67E22',
        'linestyle': '-.',
    },
    'formal_adoption': {
        'date': pd.Timestamp('2024-02-07'),
        'label': 'Regulation\nAdopted',
        'description': 'Regulation (EU) 2024/573 formally adopted',
        'color': '#E74C3C',
        'linestyle': '--',
    },
    'enters_force': {
        'date': pd.Timestamp('2024-03-11'),
        'label': 'Enters Force',
        'description': 'Regulation enters into force (20 days after OJ publication)',
        'color': '#C0392B',
        'linestyle': '--',
    },
    'desflurane_ban': {
        'date': pd.Timestamp('2026-01-01'),
        'label': 'Desflurane\nBan',
        'description': 'Desflurane prohibition takes effect under Art. 11(1) / Annex IV',
        'color': '#000000',
        'linestyle': '--',
    },
}

# Additional context events
context_events = {
    'scotland_ban': {
        'date': pd.Timestamp('2023-03-03'),
        'label': 'Scotland\nBan',
        'description': 'NHS Scotland stops purchasing desflurane',
    },
    'nhs_england_announce': {
        'date': pd.Timestamp('2023-01-13'),
        'label': 'NHS England\nAnnouncement',
        'description': 'NHS England announces desflurane decommissioning by 2024',
    },
}

# Primary regulatory dates for analysis
reg_date = pd.Timestamp('2026-01-01')           # Desflurane ban effective date
adoption_date = pd.Timestamp('2024-02-07')       # Regulation formally adopted
proposal_date = pd.Timestamp('2022-04-05')       # EC proposal published
agreement_date = pd.Timestamp('2023-10-05')      # Provisional trilogue agreement

# ==========================================
# 1. LOAD AND CLEAN eBay (Terapeak) DATA
# ==========================================

data_dir = '/home/ubuntu/vaporizer_research/data'

def load_terapeak(filepath):
    """Load Terapeak CSV and standardize columns."""
    df = pd.read_csv(filepath)
    df['date_sold'] = pd.to_datetime(df['sold_date'], format='mixed', dayfirst=False, errors='coerce')
    df['price_usd'] = pd.to_numeric(df['price_usd'], errors='coerce')
    df['qty_sold'] = pd.to_numeric(df['qty_sold'], errors='coerce').fillna(1).astype(int)
    df = df.dropna(subset=['date_sold', 'price_usd'])
    return df

# Load all three Terapeak datasets
des_raw = load_terapeak(f'{data_dir}/terapeak_desflurane.csv')
sevo_raw = load_terapeak(f'{data_dir}/terapeak_sevoflurane.csv')
iso_raw = load_terapeak(f'{data_dir}/terapeak_isoflurane.csv')

print("RAW DATA LOADED (eBay Terapeak 3-Year):")
print(f"  Desflurane: {len(des_raw)} listings")
print(f"  Sevoflurane: {len(sevo_raw)} listings")
print(f"  Isoflurane: {len(iso_raw)} listings")

# ==========================================
# 1b. FILTER TO VAPORIZERS ONLY
# ==========================================

des = des_raw.copy()
des['agent_type'] = 'Desflurane'

if 'item_category' in sevo_raw.columns:
    sevo = sevo_raw[sevo_raw['item_category'] == 'vaporizer'].copy()
else:
    sevo = sevo_raw.copy()
sevo['agent_type'] = 'Sevoflurane'

if 'item_category' in iso_raw.columns:
    iso = iso_raw[iso_raw['item_category'] == 'vaporizer'].copy()
else:
    iso = iso_raw.copy()
iso['agent_type'] = 'Isoflurane'

# Remove accessories/non-vaporizer items by keyword (safety net)
exclude_keywords = ['key fill', 'keyed filler', 'bottle adapter', 'easy-fil',
                    'adapter', 'filler cap', 'pour fill', 'anti-spill',
                    'fill adapter', 'lot of 10']
for df in [des, sevo, iso]:
    mask = ~df['title'].str.lower().str.contains('|'.join(exclude_keywords), na=False)
    df.drop(df[~mask].index, inplace=True)

# Remove veterinary anesthesia systems/machines (keep standalone vaporizers)
vet_system_keywords = ['anesthesia system', 'anesthesia center', 'with cart',
                       'rolling stand', 'w/ cart', 'rodent anesthesia',
                       'compac5', 'rc2 rodent', 'v-10 isoflurane laboratory',
                       'anesthesia workstation', 'apollo anesthesia']
for df in [iso, sevo]:
    mask = ~df['title'].str.lower().str.contains('|'.join(vet_system_keywords), na=False)
    df.drop(df[~mask].index, inplace=True)

# Remove lot listings
for df in [des, sevo, iso]:
    mask = ~df['title'].str.lower().str.contains(r'^lot of \d', na=False, regex=True)
    df.drop(df[~mask].index, inplace=True)

print(f"\nAFTER FILTERING (vaporizers only):")
print(f"  Desflurane: {len(des)} listings")
print(f"  Sevoflurane: {len(sevo)} listings")
print(f"  Isoflurane: {len(iso)} listings")

# Add source column
des['source'] = 'eBay (Terapeak)'
sevo['source'] = 'eBay (Terapeak)'
iso['source'] = 'eBay (Terapeak)'

# Combine all data (eBay Terapeak only - single source to avoid cross-listing duplication)
all_data = pd.concat([des, sevo, iso], ignore_index=True)

# Define multi-period classification based on regulatory timeline
def classify_period(date):
    if date < proposal_date:
        return '1_Pre-proposal'
    elif date < agreement_date:
        return '2_Post-proposal'
    elif date < adoption_date:
        return '3_Post-agreement'
    elif date < reg_date:
        return '4_Post-adoption'
    else:
        return '5_Post-ban'

all_data['period_detailed'] = all_data['date_sold'].apply(classify_period)

# Simple pre/post regulation flag (for backward compatibility)
all_data['period'] = np.where(all_data['date_sold'] < reg_date, 'Pre-regulation', 'Post-regulation')

# Pre/post adoption flag
all_data['period_adoption'] = np.where(all_data['date_sold'] < adoption_date, 'Pre-adoption', 'Post-adoption')

print("\n" + "=" * 60)
print("DATA SUMMARY (eBay Terapeak 3-Year Data)")
print("=" * 60)
print(f"Total listings: {len(all_data)}")
print(f"Date range: {all_data['date_sold'].min().strftime('%Y-%m-%d')} to {all_data['date_sold'].max().strftime('%Y-%m-%d')}")
for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    subset = all_data[all_data['agent_type'] == agent]
    print(f"\n{agent}:")
    print(f"  Total listings: {len(subset)}")
    if len(subset) > 0:
        print(f"  Date range: {subset['date_sold'].min().strftime('%Y-%m-%d')} to {subset['date_sold'].max().strftime('%Y-%m-%d')}")
        print(f"  Price range: ${subset['price_usd'].min():.2f} - ${subset['price_usd'].max():.2f}")
        print(f"  Mean price: ${subset['price_usd'].mean():.2f}")
        print(f"  Median price: ${subset['price_usd'].median():.2f}")
        for period in ['Pre-regulation', 'Post-regulation']:
            p_subset = subset[subset['period'] == period]
            if len(p_subset) > 0:
                print(f"  {period}: n={len(p_subset)}, mean=${p_subset['price_usd'].mean():.2f}, median=${p_subset['price_usd'].median():.2f}")
        print(f"  Detailed periods:")
        for period in sorted(subset['period_detailed'].unique()):
            p_subset = subset[subset['period_detailed'] == period]
            label = period.split('_', 1)[1]
            print(f"    {label}: n={len(p_subset)}, mean=${p_subset['price_usd'].mean():.2f}, median=${p_subset['price_usd'].median():.2f}")

# ==========================================
# 2. STATISTICAL ANALYSIS
# ==========================================

print("\n" + "=" * 60)
print("STATISTICAL ANALYSIS")
print("=" * 60)

results = {}
for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    subset = all_data[all_data['agent_type'] == agent]
    pre = subset[subset['period'] == 'Pre-regulation']['price_usd']
    post = subset[subset['period'] == 'Post-regulation']['price_usd']

    results[agent] = {
        'pre_n': len(pre),
        'post_n': len(post),
        'pre_mean': pre.mean() if len(pre) > 0 else np.nan,
        'post_mean': post.mean() if len(post) > 0 else np.nan,
        'pre_median': pre.median() if len(pre) > 0 else np.nan,
        'post_median': post.median() if len(post) > 0 else np.nan,
        'pre_sd': pre.std() if len(pre) > 0 else np.nan,
        'post_sd': post.std() if len(post) > 0 else np.nan,
    }

    # Mann-Whitney U test
    if len(pre) >= 2 and len(post) >= 2:
        u_stat, u_pval = stats.mannwhitneyu(pre, post, alternative='two-sided')
        results[agent]['u_stat'] = u_stat
        results[agent]['u_pval'] = u_pval

        # Welch's t-test
        t_stat, t_pval = stats.ttest_ind(pre, post, equal_var=False)
        results[agent]['t_stat'] = t_stat
        results[agent]['t_pval'] = t_pval

        # Effect size (Cohen's d)
        pooled_sd = np.sqrt((pre.var() + post.var()) / 2)
        if pooled_sd > 0:
            cohens_d = (post.mean() - pre.mean()) / pooled_sd
            results[agent]['cohens_d'] = cohens_d

        # Percent change
        if pre.mean() > 0:
            pct_change = ((post.mean() - pre.mean()) / pre.mean()) * 100
            results[agent]['pct_change'] = pct_change

        print(f"\n{agent}:")
        print(f"  Pre-regulation:  n={len(pre)}, mean=${pre.mean():.2f} +/- ${pre.std():.2f}, median=${pre.median():.2f}")
        print(f"  Post-regulation: n={len(post)}, mean=${post.mean():.2f} +/- ${post.std():.2f}, median=${post.median():.2f}")
        print(f"  Change: {results[agent].get('pct_change', 0):.1f}%")
        print(f"  Mann-Whitney U: U={u_stat:.1f}, p={u_pval:.4f}")
        print(f"  Welch's t-test: t={t_stat:.3f}, p={t_pval:.4f}")
        print(f"  Cohen's d: {results[agent].get('cohens_d', 0):.3f}")
    else:
        print(f"\n{agent}: Insufficient data for pre/post comparison (pre={len(pre)}, post={len(post)})")

# Multi-period Kruskal-Wallis test
print("\n" + "-" * 40)
print("MULTI-PERIOD ANALYSIS (Kruskal-Wallis)")
print("-" * 40)
for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    subset = all_data[all_data['agent_type'] == agent]
    groups = []
    group_labels = []
    for period in sorted(subset['period_detailed'].unique()):
        group = subset[subset['period_detailed'] == period]['price_usd'].values
        if len(group) >= 2:
            groups.append(group)
            group_labels.append(period.split('_', 1)[1])
    if len(groups) >= 2:
        h_stat, h_pval = stats.kruskal(*groups)
        print(f"\n{agent}: H={h_stat:.2f}, p={h_pval:.4f} ({', '.join(group_labels)})")
    else:
        print(f"\n{agent}: Not enough groups for Kruskal-Wallis test")

# ==========================================
# 3. FIGURES
# ==========================================

figdir = '/home/ubuntu/vaporizer_research/figures/'

colors = {'Desflurane': '#E74C3C', 'Sevoflurane': '#2E86C1', 'Isoflurane': '#27AE60'}
markers = {'Desflurane': 'o', 'Sevoflurane': 's', 'Isoflurane': '^'}


def add_regulatory_events(ax, events_to_show='main'):
    """Add regulatory event vertical lines to an axis."""
    if events_to_show == 'main':
        keys = ['commission_proposal', 'formal_adoption', 'desflurane_ban']
    elif events_to_show == 'all':
        keys = list(regulatory_events.keys())
    elif events_to_show == 'key':
        keys = ['commission_proposal', 'parliament_vote', 'formal_adoption', 'desflurane_ban']
    else:
        keys = events_to_show

    for key in keys:
        if key not in regulatory_events:
            continue
        event = regulatory_events[key]
        ax.axvline(x=event['date'], color=event['color'],
                   linestyle=event['linestyle'], linewidth=1.5, alpha=0.7)


def add_regulatory_legend(ax, events_to_show='main'):
    """Add a legend for regulatory events."""
    if events_to_show == 'main':
        keys = ['commission_proposal', 'formal_adoption', 'desflurane_ban']
    elif events_to_show == 'all':
        keys = list(regulatory_events.keys())
    elif events_to_show == 'key':
        keys = ['commission_proposal', 'parliament_vote', 'formal_adoption', 'desflurane_ban']
    else:
        keys = events_to_show

    from matplotlib.lines import Line2D
    legend_elements = []
    for key in keys:
        if key not in regulatory_events:
            continue
        event = regulatory_events[key]
        legend_elements.append(
            Line2D([0], [0], color=event['color'], linestyle=event['linestyle'],
                   linewidth=1.5, label=f"{event['label'].replace(chr(10), ' ')} ({event['date'].strftime('%b %Y')})")
        )
    return legend_elements


# --- Figure 1: Price time series with multiple inflection points ---
fig, axes = plt.subplots(3, 1, figsize=(12, 14), sharex=True)

for idx, agent in enumerate(['Desflurane', 'Sevoflurane', 'Isoflurane']):
    ax = axes[idx]
    subset = all_data[all_data['agent_type'] == agent].sort_values('date_sold')

    ax.scatter(subset['date_sold'], subset['price_usd'],
               c=colors[agent], marker=markers[agent], alpha=0.7, s=60,
               edgecolors='white', linewidth=0.5, label=f'{agent} vaporizer')

    # Add LOWESS trend line
    if len(subset) >= 5:
        try:
            from statsmodels.nonparametric.smoothers_lowess import lowess
            numeric_dates = (subset['date_sold'] - subset['date_sold'].min()).dt.days.values
            smoothed = lowess(subset['price_usd'].values, numeric_dates, frac=0.3)
            smooth_dates = subset['date_sold'].min() + pd.to_timedelta(smoothed[:, 0], unit='D')
            ax.plot(smooth_dates, smoothed[:, 1], color=colors[agent], linewidth=2, alpha=0.8)
        except Exception as e:
            print(f"  LOWESS failed for {agent}: {e}")

    # Add regulatory event lines
    add_regulatory_events(ax, events_to_show='key')

    ax.set_ylabel('Price (USD)')
    ax.set_title(f'{agent} Vaporizer Prices', fontweight='bold')
    ax.legend(loc='upper left')
    ax.grid(True, alpha=0.3)

# Add regulatory event labels on top subplot
ax0 = axes[0]
ylim = ax0.get_ylim()
for key in ['commission_proposal', 'parliament_vote', 'formal_adoption', 'desflurane_ban']:
    event = regulatory_events[key]
    ax0.text(event['date'], ylim[1] * 0.98, f"  {event['label']}",
             fontsize=7.5, va='top', ha='left', style='italic', color=event['color'],
             fontweight='bold')

axes[-1].set_xlabel('Date')
axes[-1].xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
axes[-1].xaxis.set_major_locator(mdates.MonthLocator(interval=3))
plt.xticks(rotation=45)

# Add regulatory events legend at bottom - preserve existing agent legend
reg_legend = add_regulatory_legend(axes[-1], 'key')
# Get existing legend (agent label) and re-add it after adding regulatory legend
existing_legend = axes[-1].get_legend()
leg2 = axes[-1].legend(handles=reg_legend, loc='lower left', fontsize=8,
                        framealpha=0.9, title='Regulatory Timeline', title_fontsize=9)
axes[-1].add_artist(leg2)
# Re-add the agent legend
if existing_legend:
    axes[-1].add_artist(existing_legend)

plt.tight_layout()
plt.savefig(f'{figdir}fig1_price_timeseries.png', dpi=300, bbox_inches='tight')
plt.savefig(f'{figdir}fig1_price_timeseries.pdf', bbox_inches='tight')
plt.close()
print("\nFigure 1 saved: Price time series with regulatory milestones")

# --- Figure 2: Box plot comparison pre/post regulation ---
fig, ax = plt.subplots(figsize=(10, 6))

plot_data = all_data.copy()
plot_data['Category'] = plot_data['agent_type'] + '\n(' + plot_data['period'] + ')'

order = []
for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    for period in ['Pre-regulation', 'Post-regulation']:
        order.append(f'{agent}\n({period})')

pre_post_colors = {
    'Desflurane\n(Pre-regulation)': '#E74C3C',
    'Desflurane\n(Post-regulation)': '#C0392B',
    'Sevoflurane\n(Pre-regulation)': '#5DADE2',
    'Sevoflurane\n(Post-regulation)': '#2E86C1',
    'Isoflurane\n(Pre-regulation)': '#58D68D',
    'Isoflurane\n(Post-regulation)': '#27AE60',
}

# Filter to only categories that have data
available_cats = [cat for cat in order if cat in plot_data['Category'].values]

bp = sns.boxplot(data=plot_data, x='Category', y='price_usd', order=available_cats,
                 palette=[pre_post_colors.get(c, '#888888') for c in available_cats],
                 showfliers=True, flierprops={'marker': 'o', 'markersize': 4, 'alpha': 0.5})

# Add individual points
sns.stripplot(data=plot_data, x='Category', y='price_usd', order=available_cats,
              color='black', alpha=0.3, size=3, jitter=True)

ax.set_xlabel('')
ax.set_ylabel('Price (USD)', fontsize=12)
ax.set_title('Vaporizer Prices: Pre- vs Post-EU Desflurane Ban (January 2026)\n(Source: eBay Terapeak, 3-Year Data)', fontweight='bold')

# Add P value labels
for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    if agent in results and 'u_pval' in results[agent]:
        pval = results[agent]['u_pval']
        if pval < 0.001:
            sig = 'P<0.001'
        elif pval < 0.01:
            sig = f'P={pval:.3f}'
        else:
            sig = f'P={pval:.2f}'

        pre_cat = f'{agent}\n(Pre-regulation)'
        post_cat = f'{agent}\n(Post-regulation)'
        if pre_cat in available_cats and post_cat in available_cats:
            pre_idx = available_cats.index(pre_cat)
            post_idx = available_cats.index(post_cat)
            y_max = plot_data[plot_data['Category'].isin([pre_cat, post_cat])]['price_usd'].max()
            ax.plot([pre_idx, pre_idx, post_idx, post_idx],
                    [y_max + 50, y_max + 80, y_max + 80, y_max + 50], 'k-', linewidth=1)
            ax.text((pre_idx + post_idx) / 2, y_max + 85, sig, ha='center', fontsize=10, fontweight='bold')

plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig(f'{figdir}fig2_boxplot_comparison.png', dpi=300, bbox_inches='tight')
plt.savefig(f'{figdir}fig2_boxplot_comparison.pdf', bbox_inches='tight')
plt.close()
print("Figure 2 saved: Box plot comparison")

# --- Figure 3: Monthly median prices trend with regulatory milestones ---
fig, ax = plt.subplots(figsize=(12, 6))

for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    subset = all_data[all_data['agent_type'] == agent].copy()
    subset['month'] = subset['date_sold'].dt.to_period('M').dt.to_timestamp()
    monthly = subset.groupby('month').agg(
        median_price=('price_usd', 'median'),
        count=('price_usd', 'count'),
        mean_price=('price_usd', 'mean')
    ).reset_index()

    monthly_filtered = monthly[monthly['count'] >= 1]

    ax.plot(monthly_filtered['month'], monthly_filtered['median_price'],
            color=colors[agent], marker=markers[agent], linewidth=2,
            markersize=8, label=f'{agent} (median)', alpha=0.9)

    # Add count annotations
    for _, row in monthly_filtered.iterrows():
        ax.annotate(f'n={int(row["count"])}',
                     xy=(row['month'], row['median_price']),
                     xytext=(0, 12), textcoords='offset points',
                     fontsize=7, ha='center', color=colors[agent], alpha=0.7)

# Add regulatory event lines with labels
add_regulatory_events(ax, events_to_show='key')

ylim = ax.get_ylim()
for key in ['commission_proposal', 'parliament_vote', 'formal_adoption', 'desflurane_ban']:
    event = regulatory_events[key]
    ax.text(event['date'], ylim[1] * 0.98, f"  {event['label']}",
            fontsize=8, va='top', ha='left', style='italic', color=event['color'],
            fontweight='bold')

ax.set_xlabel('Month', fontsize=12)
ax.set_ylabel('Median Price (USD)', fontsize=12)
ax.set_title('Monthly Median Prices of Anaesthetic Vaporizers\n(eBay Terapeak, 3-Year Sold Data) with EU Regulatory Milestones', fontweight='bold')
ax.legend(loc='upper left', framealpha=0.9)
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f'{figdir}fig3_monthly_median.png', dpi=300, bbox_inches='tight')
plt.savefig(f'{figdir}fig3_monthly_median.pdf', bbox_inches='tight')
plt.close()
print("Figure 3 saved: Monthly median prices with regulatory milestones")

# --- Figure 4: Price distribution histograms ---
fig, axes = plt.subplots(1, 3, figsize=(14, 5))

for idx, agent in enumerate(['Desflurane', 'Sevoflurane', 'Isoflurane']):
    ax = axes[idx]
    subset = all_data[all_data['agent_type'] == agent]

    pre_data = subset[subset['period'] == 'Pre-regulation']['price_usd']
    post_data = subset[subset['period'] == 'Post-regulation']['price_usd']

    bins = np.linspace(0, max(subset['price_usd'].max(), 100), 20)

    if len(pre_data) > 0:
        ax.hist(pre_data, bins=bins, alpha=0.5, color=colors[agent],
                label=f'Pre-ban (n={len(pre_data)})', edgecolor='white')
    if len(post_data) > 0:
        ax.hist(post_data, bins=bins, alpha=0.7, color=colors[agent],
                label=f'Post-ban (n={len(post_data)})', edgecolor='black', linewidth=0.5,
                hatch='///')

    ax.set_xlabel('Price (USD)')
    ax.set_ylabel('Count')
    ax.set_title(f'{agent}', fontweight='bold')
    ax.legend(fontsize=9)

plt.suptitle('Price Distribution: Pre- vs Post-Desflurane Ban (Jan 2026)\n(Source: eBay Terapeak)', fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(f'{figdir}fig4_histograms.png', dpi=300, bbox_inches='tight')
plt.savefig(f'{figdir}fig4_histograms.pdf', bbox_inches='tight')
plt.close()
print("Figure 4 saved: Price distribution histograms")


# --- Figure 5: NEW - Regulatory Timeline Overview ---
fig, ax = plt.subplots(figsize=(14, 7))

# Plot all data points
for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    subset = all_data[all_data['agent_type'] == agent].sort_values('date_sold')
    ax.scatter(subset['date_sold'], subset['price_usd'],
               c=colors[agent], marker=markers[agent], alpha=0.6, s=50,
               edgecolors='white', linewidth=0.5, label=f'{agent}')

# Add shaded regions for regulatory phases
date_min = all_data['date_sold'].min() - pd.Timedelta(days=30)
date_max = all_data['date_sold'].max() + pd.Timedelta(days=30)

ax.axvspan(date_min, proposal_date, alpha=0.05, color='green')
ax.axvspan(proposal_date, adoption_date, alpha=0.08, color='orange')
ax.axvspan(adoption_date, reg_date, alpha=0.08, color='red')
ax.axvspan(reg_date, date_max, alpha=0.1, color='darkred')

# Add vertical lines for all key events
for key in ['commission_proposal', 'parliament_vote', 'provisional_agreement', 'formal_adoption', 'desflurane_ban']:
    event = regulatory_events[key]
    ax.axvline(x=event['date'], color=event['color'],
               linestyle=event['linestyle'], linewidth=2, alpha=0.8)

# Add period labels at top
ylim_max = all_data['price_usd'].max() * 1.15
ax.set_ylim(bottom=-50, top=ylim_max)

# Add labels for regulatory events at top
event_labels = [
    ('commission_proposal', 'top'),
    ('parliament_vote', 'top'),
    ('provisional_agreement', 'bottom'),
    ('formal_adoption', 'top'),
    ('desflurane_ban', 'top'),
]
for key, pos in event_labels:
    event = regulatory_events[key]
    y_pos = ylim_max * 0.98 if pos == 'top' else ylim_max * 0.88
    ax.annotate(event['label'],
                xy=(event['date'], y_pos),
                fontsize=8, va='top', ha='center',
                fontweight='bold', color=event['color'],
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor=event['color'], alpha=0.8))

# Phase labels at bottom
phase_positions = [
    (date_min, proposal_date, 'Pre-Proposal\nPhase', 'green'),
    (proposal_date, adoption_date, 'Legislative\nProcess', 'orange'),
    (adoption_date, reg_date, 'Pre-Implementation\nPhase', 'red'),
    (reg_date, date_max, 'Post-Ban\nPhase', 'darkred'),
]
for start, end, label, color in phase_positions:
    mid = start + (end - start) / 2
    ax.text(mid, -30, label, ha='center', va='top', fontsize=8,
            color=color, fontweight='bold', alpha=0.7)

ax.set_xlabel('Date', fontsize=12)
ax.set_ylabel('Price (USD)', fontsize=12)
ax.set_title('Anaesthetic Vaporizer Prices with EU Regulatory Timeline\n'
             'EC Proposal (Apr 2022) \u2192 EP Vote (Mar 2023) \u2192 Trilogue (Oct 2023) \u2192 '
             'Adoption (Feb 2024) \u2192 Ban Effective (Jan 2026)\n'
             '(Source: eBay Terapeak, 3-Year Sold Data)',
             fontweight='bold', fontsize=11)
ax.legend(loc='upper left', framealpha=0.9, fontsize=10)
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f'{figdir}fig5_regulatory_timeline.png', dpi=300, bbox_inches='tight')
plt.savefig(f'{figdir}fig5_regulatory_timeline.pdf', bbox_inches='tight')
plt.close()
print("Figure 5 saved: Regulatory timeline overview")


# --- Figure 6: Quarterly volume and price trends ---
fig, axes = plt.subplots(2, 1, figsize=(12, 10), sharex=True)

for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
    subset = all_data[all_data['agent_type'] == agent].copy()
    subset['quarter'] = subset['date_sold'].dt.to_period('Q').dt.to_timestamp()
    quarterly = subset.groupby('quarter').agg(
        median_price=('price_usd', 'median'),
        count=('price_usd', 'count'),
        mean_price=('price_usd', 'mean')
    ).reset_index()

    axes[0].plot(quarterly['quarter'], quarterly['median_price'],
                 color=colors[agent], marker=markers[agent], linewidth=2,
                 markersize=8, label=f'{agent}', alpha=0.9)
    offset_days = {'Desflurane': -20, 'Sevoflurane': 0, 'Isoflurane': 20}
    axes[1].bar(quarterly['quarter'] + pd.Timedelta(days=offset_days[agent]),
                quarterly['count'], width=35, color=colors[agent], alpha=0.7, label=f'{agent}')

add_regulatory_events(axes[0], events_to_show='key')
add_regulatory_events(axes[1], events_to_show='key')

axes[0].set_ylabel('Median Price (USD)', fontsize=12)
axes[0].set_title('Quarterly Median Price Trends', fontweight='bold')
axes[0].legend(loc='upper left', framealpha=0.9)
axes[0].grid(True, alpha=0.3)

axes[1].set_ylabel('Number of Sales', fontsize=12)
axes[1].set_xlabel('Quarter', fontsize=12)
axes[1].set_title('Quarterly Sales Volume', fontweight='bold')
axes[1].legend(loc='upper left', framealpha=0.9)
axes[1].grid(True, alpha=0.3)

axes[-1].xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
axes[-1].xaxis.set_major_locator(mdates.MonthLocator(interval=3))
plt.xticks(rotation=45)

plt.suptitle('Quarterly Price and Volume Trends\n(Source: eBay Terapeak, 3-Year Sold Data)', fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(f'{figdir}fig6_quarterly_trends.png', dpi=300, bbox_inches='tight')
plt.savefig(f'{figdir}fig6_quarterly_trends.pdf', bbox_inches='tight')
plt.close()
print("Figure 6 saved: Quarterly trends")


# ==========================================
# 4. SAVE CLEANED DATA AND STATISTICS
# ==========================================

all_data.to_csv(f'{data_dir}/combined_cleaned.csv', index=False)

stats_df = pd.DataFrame(results).T
stats_df.to_csv(f'{data_dir}/statistics_summary.csv')

# Save regulatory timeline as CSV for reference
timeline_data = []
for key, event in regulatory_events.items():
    timeline_data.append({
        'event_key': key,
        'date': event['date'].strftime('%Y-%m-%d'),
        'label': event['label'].replace('\n', ' '),
        'description': event['description'],
    })
timeline_df = pd.DataFrame(timeline_data)
timeline_df.to_csv(f'{data_dir}/regulatory_timeline.csv', index=False)

print("\n" + "=" * 60)
print("ALL DATA AND ANALYSIS FILES SAVED SUCCESSFULLY")
print("=" * 60)
print(f"\nTotal listings analyzed: {len(all_data)}")
print(f"  Desflurane: {len(all_data[all_data['agent_type']=='Desflurane'])}")
print(f"  Sevoflurane: {len(all_data[all_data['agent_type']=='Sevoflurane'])}")
print(f"  Isoflurane: {len(all_data[all_data['agent_type']=='Isoflurane'])}")
print(f"\nData source: eBay (Terapeak) - 3 years of sold listings")
print(f"Date range: {all_data['date_sold'].min().strftime('%Y-%m-%d')} to {all_data['date_sold'].max().strftime('%Y-%m-%d')}")
print(f"\nRegulatory timeline: {len(timeline_data)} events")
for key, event in regulatory_events.items():
    print(f"  {event['date'].strftime('%Y-%m-%d')}: {event['description']}")
