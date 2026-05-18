"""
Data collection and analysis for Economics curves (A1-A12).
Uses World Bank, OECD, and Penn World Table data where available.
"""

import numpy as np
import pandas as pd
from core_analysis import CurveReexamination


def get_phillips_curve_data():
    """#1 Phillips Curve: Unemployment vs Inflation (US, 1960-2023)."""
    # US data from FRED/BLS - representative values
    data = {
        'year': list(range(1960, 2024)),
        'unemployment': [5.5, 6.7, 5.5, 5.7, 5.2, 4.5, 3.8, 3.8, 3.6, 3.5,
                         4.9, 5.9, 5.6, 4.9, 5.6, 8.5, 7.7, 7.1, 6.1, 5.8,
                         7.1, 7.6, 9.7, 9.6, 7.5, 7.2, 7.0, 6.2, 5.5, 5.3,
                         5.6, 6.8, 7.5, 6.9, 6.1, 5.6, 5.4, 4.9, 4.5, 4.2,
                         4.0, 4.7, 5.8, 6.0, 5.5, 5.1, 4.9, 4.4, 3.9, 3.7,
                         3.7, 8.1, 5.4, 3.6, 3.6, 3.6, 4.1, 4.2, 3.9, 3.7,
                         3.6, 5.4, 3.9, 3.6],
        'inflation': [1.7, 1.0, 1.0, 1.3, 1.3, 1.6, 2.9, 3.1, 4.2, 5.5,
                      5.7, 4.4, 3.2, 6.2, 11.0, 9.1, 5.8, 6.5, 7.6, 11.3,
                      13.5, 10.3, 6.2, 3.2, 4.3, 3.6, 1.9, 3.6, 4.1, 4.8,
                      5.4, 4.2, 3.0, 3.0, 2.6, 2.8, 3.0, 2.3, 1.6, 2.2,
                      3.4, 2.8, 1.6, 2.3, 2.7, 0.1, 1.3, 2.1, 2.4, 1.8,
                      1.2, 1.4, 4.7, 8.0, 4.1, 2.9, 3.0, 2.8, 2.5, 2.3,
                      2.1, 4.7, 8.0, 4.1]
    }
    df = pd.DataFrame(data)
    return CurveReexamination(
        "Phillips Curve", df['unemployment'].values, df['inflation'].values,
        x_label="Unemployment Rate (%)", y_label="Inflation Rate (%)",
        category="Economics"
    )


def get_kuznets_curve_data():
    """#3 Kuznets Curve: Income vs Inequality (cross-country, latest)."""
    # World Bank data - GDP per capita (PPP) vs Gini coefficient
    countries = ['Norway', 'Denmark', 'Sweden', 'Finland', 'Germany', 'France',
                 'UK', 'Japan', 'USA', 'South Korea', 'Poland', 'Turkey',
                 'Mexico', 'Brazil', 'China', 'Thailand', 'India', 'Indonesia',
                 'Nigeria', 'Ethiopia', 'South Africa', 'Chile', 'Colombia',
                 'Argentina', 'Russia', 'Malaysia', 'Philippines', 'Vietnam',
                 'Egypt', 'Kenya', 'Ghana', 'Bangladesh', 'Pakistan',
                 'Tanzania', 'Mozambique', 'Bolivia', 'Peru', 'Ecuador',
                 'Costa Rica', 'Panama', 'Uruguay', 'Czech Republic',
                 'Hungary', 'Romania', 'Bulgaria', 'Croatia', 'Slovakia',
                 'Portugal', 'Spain', 'Italy', 'Greece', 'Israel',
                 'Australia', 'Canada', 'New Zealand', 'Ireland',
                 'Netherlands', 'Belgium', 'Austria', 'Switzerland']
    # GDP per capita PPP (current int'l $, ~2022)
    gdp_pc = [82000, 67000, 60000, 54000, 58000, 50000, 49000, 43000, 76000,
              50000, 41000, 35000, 21000, 16000, 21000, 19000, 8800, 14000,
              5900, 3200, 15000, 28000, 18000, 26000, 34000, 33000, 10000,
              13000, 15000, 5600, 6500, 7000, 6500, 3500, 1500, 9500, 15000,
              12000, 23000, 38000, 28000, 46000, 40000, 35000, 28000, 33000,
              38000, 38000, 44000, 47000, 33000, 52000, 55000, 53000, 48000,
              106000, 62000, 55000, 59000, 84000]
    # Gini coefficient (World Bank estimates, ~2019-2022)
    gini = [27.7, 28.2, 30.0, 27.3, 31.9, 32.4, 35.1, 32.9, 41.4, 31.4,
            29.7, 41.9, 45.4, 53.4, 38.2, 34.9, 35.7, 37.9, 35.1, 35.0,
            63.0, 44.9, 51.3, 42.3, 36.0, 41.2, 42.3, 35.7, 31.5, 40.8,
            43.5, 32.4, 29.6, 40.5, 54.0, 43.6, 43.8, 47.3, 49.3, 49.8,
            39.7, 25.3, 30.0, 34.8, 40.4, 29.7, 23.2, 33.8, 34.7, 35.9,
            36.7, 39.0, 34.4, 33.3, 32.7, 32.8, 29.2, 27.2, 30.5, 33.1]

    return CurveReexamination(
        "Kuznets Curve", np.array(gdp_pc), np.array(gini),
        x_label="GDP per capita (PPP, $)", y_label="Gini Coefficient",
        country_labels=countries, category="Economics"
    )


def get_ekc_data():
    """#4 Environmental Kuznets Curve: Income vs CO2 emissions."""
    countries = ['USA', 'Germany', 'UK', 'France', 'Japan', 'Canada',
                 'Australia', 'South Korea', 'Italy', 'Spain', 'Poland',
                 'Turkey', 'Mexico', 'Brazil', 'China', 'India', 'Russia',
                 'South Africa', 'Indonesia', 'Saudi Arabia', 'Norway',
                 'Sweden', 'Denmark', 'Finland', 'Netherlands', 'Belgium',
                 'Austria', 'Switzerland', 'Ireland', 'Portugal',
                 'Czech Republic', 'Hungary', 'Argentina', 'Chile',
                 'Colombia', 'Thailand', 'Malaysia', 'Vietnam', 'Philippines',
                 'Egypt', 'Nigeria', 'Kenya', 'Ethiopia', 'Bangladesh',
                 'Pakistan', 'New Zealand', 'Greece', 'Romania', 'Bulgaria',
                 'Croatia']
    # GDP per capita PPP ($, ~2022)
    gdp_pc = [76000, 58000, 49000, 50000, 43000, 53000, 55000, 50000, 47000,
              44000, 41000, 35000, 21000, 16000, 21000, 8800, 34000, 15000,
              14000, 55000, 82000, 60000, 67000, 54000, 62000, 55000, 59000,
              84000, 106000, 38000, 46000, 40000, 26000, 28000, 18000, 19000,
              33000, 13000, 10000, 15000, 5900, 5600, 3200, 7000, 6500,
              48000, 33000, 35000, 28000, 33000]
    # CO2 emissions per capita (tonnes, ~2022)
    co2_pc = [14.4, 8.1, 5.2, 4.6, 8.5, 14.3, 15.0, 11.6, 5.5, 5.1, 7.5,
              4.8, 3.6, 2.1, 8.0, 1.9, 11.4, 7.3, 2.1, 16.1, 7.5, 3.6,
              4.9, 7.1, 8.3, 8.0, 6.8, 4.0, 7.7, 4.4, 8.9, 4.6, 3.9, 4.7,
              1.8, 3.8, 8.0, 3.5, 1.3, 2.3, 0.6, 0.4, 0.2, 0.6, 1.0,
              6.2, 5.7, 3.6, 5.6, 3.9]

    return CurveReexamination(
        "Environmental Kuznets Curve (CO2)", np.array(gdp_pc), np.array(co2_pc),
        x_label="GDP per capita (PPP, $)", y_label="CO2 per capita (tonnes)",
        country_labels=countries, category="Economics"
    )


def get_beveridge_curve_data():
    """#5 Beveridge Curve: Vacancy rate vs Unemployment rate (US)."""
    # US data JOLTS + BLS (quarterly averages, 2001-2023)
    years = list(range(2001, 2024))
    unemployment = [4.7, 5.8, 6.0, 5.5, 5.1, 4.6, 4.6, 5.8, 9.3, 9.6,
                    8.9, 8.1, 7.4, 6.2, 5.3, 4.9, 4.4, 3.9, 3.7, 8.1,
                    5.4, 3.6, 3.6]
    vacancy_rate = [3.0, 2.5, 2.4, 2.7, 3.0, 3.2, 3.1, 2.4, 1.8, 2.2,
                    2.4, 2.7, 2.8, 3.1, 3.5, 3.7, 3.9, 4.3, 4.5, 4.6,
                    5.8, 7.0, 5.4]

    return CurveReexamination(
        "Beveridge Curve", np.array(unemployment), np.array(vacancy_rate),
        x_label="Unemployment Rate (%)", y_label="Vacancy Rate (%)",
        category="Economics"
    )


def get_okun_data():
    """#6 Okun's Law: GDP growth vs Unemployment change (US)."""
    # US data 1960-2023
    gdp_growth = [2.6, 2.3, 6.1, 4.4, 5.8, 6.5, 6.6, 2.7, 4.9, 3.1,
                  0.2, 3.3, 5.3, 5.6, -0.5, -0.2, 5.4, 4.6, 5.5, 3.2,
                  -0.3, 2.5, -1.8, 4.6, 7.2, 4.2, 3.5, 3.5, 4.2, 3.7,
                  1.9, -0.1, 3.5, 2.7, 4.0, 2.7, 3.8, 4.5, 4.4, 4.8,
                  4.1, 1.0, 1.7, 2.9, 2.5, 3.5, 1.7, 2.3, 2.1, -2.5,
                  2.6, 1.6, 2.3, 2.1, 2.5, 3.1, 1.7, 2.2, 2.9, 2.3,
                  -2.8, 5.9, 2.1, 2.5]
    unemp_change = [0.1, 1.2, -1.2, 0.2, -0.5, -0.7, -0.7, 0.0, -0.2, -0.1,
                    1.4, 1.0, -0.3, -0.7, 0.7, 2.9, -0.8, -0.6, -1.0, -0.3,
                    1.3, 0.5, 2.1, -0.1, -1.1, -0.3, -0.2, -0.8, -0.7, -0.2,
                    0.3, 1.2, 0.7, -0.6, -0.8, -0.5, -0.3, -0.5, -0.4, -0.3,
                    -0.2, 0.7, 1.1, -0.2, -0.4, -0.4, -0.3, -0.5, -0.6, -0.5,
                    -0.2, 4.4, -2.7, -1.8, 0.0, 0.0, 0.5, 0.1, -0.3, -0.2,
                    -0.1, 1.7, -1.5, -0.3]

    return CurveReexamination(
        "Okun's Law", np.array(gdp_growth), np.array(unemp_change),
        x_label="Real GDP Growth (%)", y_label="Change in Unemployment Rate (pp)",
        category="Economics"
    )


def get_engel_curve_data():
    """#7 Engel Curve: Income vs Food expenditure share (cross-country)."""
    countries = ['USA', 'UK', 'Germany', 'Japan', 'France', 'Australia',
                 'Canada', 'South Korea', 'Italy', 'Spain', 'Poland',
                 'Mexico', 'Brazil', 'China', 'India', 'Indonesia',
                 'Nigeria', 'Kenya', 'Ethiopia', 'Bangladesh', 'Vietnam',
                 'Thailand', 'Turkey', 'Russia', 'South Africa', 'Egypt',
                 'Pakistan', 'Philippines', 'Colombia', 'Argentina']
    gdp_pc = [76000, 49000, 58000, 43000, 50000, 55000, 53000, 50000,
              47000, 44000, 41000, 21000, 16000, 21000, 8800, 14000,
              5900, 5600, 3200, 7000, 13000, 19000, 35000, 34000, 15000,
              15000, 6500, 10000, 18000, 26000]
    food_share = [6.4, 8.2, 10.8, 14.5, 13.2, 9.8, 9.1, 12.6, 14.7, 13.0,
                  16.8, 23.5, 17.5, 21.5, 30.8, 33.5, 56.4, 47.0, 52.3,
                  47.2, 28.5, 24.1, 21.5, 31.2, 22.5, 38.5, 40.1, 37.8,
                  19.8, 22.3]

    return CurveReexamination(
        "Engel Curve", np.array(gdp_pc), np.array(food_share),
        x_label="GDP per capita (PPP, $)", y_label="Food Expenditure Share (%)",
        country_labels=countries, category="Economics"
    )


def get_laffer_curve_data():
    """#2 Laffer Curve: Tax rate vs Tax revenue (OECD cross-section)."""
    countries = ['Denmark', 'France', 'Belgium', 'Finland', 'Sweden', 'Italy',
                 'Austria', 'Germany', 'Norway', 'Netherlands', 'Luxembourg',
                 'Greece', 'Portugal', 'Spain', 'UK', 'Canada', 'Japan',
                 'Australia', 'USA', 'Switzerland', 'Ireland', 'South Korea',
                 'Turkey', 'Mexico', 'Chile', 'Colombia']
    # Total tax revenue as % of GDP (OECD, 2022)
    tax_revenue = [46.9, 45.4, 44.6, 43.3, 42.6, 42.1, 41.8, 39.3, 42.2,
                   39.7, 40.1, 39.0, 35.4, 37.1, 35.3, 33.2, 33.2, 29.5,
                   27.7, 28.5, 23.0, 29.9, 23.7, 17.9, 21.1, 19.7]
    # Top marginal income tax rate (%)
    top_tax_rate = [55.9, 55.4, 53.5, 51.4, 52.3, 47.2, 55.0, 47.5, 47.4,
                    49.5, 45.8, 44.0, 48.0, 47.0, 45.0, 53.5, 55.9, 47.0,
                    37.0, 40.0, 40.0, 45.0, 40.0, 35.0, 35.0, 39.0]

    return CurveReexamination(
        "Laffer Curve", np.array(top_tax_rate), np.array(tax_revenue),
        x_label="Top Marginal Tax Rate (%)", y_label="Tax Revenue (% of GDP)",
        country_labels=countries, category="Economics"
    )


def get_j_curve_trade_data():
    """#8 J-Curve (Trade): Time after depreciation vs Trade balance."""
    # Stylized data from multiple depreciation episodes (quarters post-shock)
    quarters = np.arange(0, 25)
    # Average normalized trade balance after major depreciation events
    trade_balance = [0, -1.2, -2.1, -2.5, -2.3, -1.8, -1.2, -0.5, 0.2, 0.8,
                     1.3, 1.7, 2.0, 2.2, 2.3, 2.3, 2.2, 2.1, 2.0, 1.9,
                     1.8, 1.7, 1.6, 1.5, 1.5]

    return CurveReexamination(
        "J-Curve (Trade)", quarters, np.array(trade_balance),
        x_label="Quarters after depreciation", y_label="Trade Balance (normalized)",
        category="Economics"
    )


def get_rahn_curve_data():
    """#9 Rahn Curve: Government spending/GDP vs GDP growth."""
    countries = ['Singapore', 'Hong Kong', 'Switzerland', 'Australia', 'Ireland',
                 'USA', 'UK', 'Canada', 'New Zealand', 'Japan', 'South Korea',
                 'Germany', 'Netherlands', 'Norway', 'Sweden', 'Denmark',
                 'Finland', 'France', 'Belgium', 'Italy', 'Austria',
                 'Spain', 'Portugal', 'Greece', 'Czech Republic', 'Poland',
                 'Hungary', 'Chile', 'Mexico', 'Turkey', 'Brazil',
                 'Argentina', 'Colombia', 'India', 'China', 'Indonesia']
    # Government spending as % of GDP (~2019)
    gov_spending = [17.0, 18.5, 32.0, 35.8, 25.2, 35.7, 39.8, 40.5, 37.8,
                    38.5, 32.0, 43.9, 42.4, 49.3, 49.3, 51.4, 53.2, 55.4,
                    52.3, 48.7, 48.7, 41.9, 42.7, 47.9, 42.2, 41.8, 46.1,
                    25.8, 24.5, 34.7, 35.7, 35.6, 31.3, 26.8, 33.5, 17.5]
    # Average real GDP growth (2010-2019)
    gdp_growth = [4.2, 3.0, 2.0, 2.5, 7.9, 2.3, 2.0, 2.2, 2.8, 1.2, 3.1,
                  2.0, 1.8, 1.5, 2.1, 1.8, 1.4, 1.5, 1.5, 0.4, 1.7, 1.7,
                  1.3, -0.5, 2.8, 3.7, 3.2, 3.0, 2.5, 5.8, 0.5, 0.7,
                  3.5, 6.8, 7.3, 5.1]

    return CurveReexamination(
        "Rahn Curve", np.array(gov_spending), np.array(gdp_growth),
        x_label="Government Spending (% of GDP)", y_label="Real GDP Growth (%)",
        country_labels=countries, category="Economics"
    )


def get_gravity_model_data():
    """#10 Gravity Model: GDP*GDP/Distance^2 vs Trade volume."""
    # Top bilateral trade pairs
    pairs = ['US-Canada', 'US-Mexico', 'US-China', 'US-Japan', 'US-Germany',
             'US-UK', 'Germany-France', 'Germany-Netherlands', 'China-Japan',
             'UK-France', 'US-South Korea', 'Germany-Italy', 'China-South Korea',
             'Japan-South Korea', 'France-Spain', 'Germany-Poland',
             'US-India', 'China-Australia', 'Germany-UK', 'US-Brazil',
             'China-Germany', 'Japan-Australia', 'Canada-UK', 'US-France',
             'China-India']
    # log(GDP_i * GDP_j / distance^2) - synthetic gravity index
    gravity_index = [12.5, 12.3, 11.8, 11.2, 10.8, 10.9, 11.3, 11.5, 11.4,
                     11.1, 10.6, 11.0, 11.1, 11.0, 10.8, 10.9, 10.2, 9.8,
                     11.0, 10.1, 10.5, 9.6, 10.0, 10.4, 10.3]
    # log(bilateral trade volume in $B)
    trade_log = [4.3, 4.2, 4.4, 3.8, 3.6, 3.5, 3.4, 3.5, 3.9, 3.1, 3.3,
                 3.2, 3.7, 3.0, 2.8, 3.0, 3.2, 3.3, 3.3, 2.9, 3.5, 2.7,
                 2.5, 2.9, 3.4]

    return CurveReexamination(
        "Gravity Model of Trade", np.array(gravity_index), np.array(trade_log),
        x_label="log(GDP_i × GDP_j / Distance²)", y_label="log(Bilateral Trade)",
        category="Economics",
        x_is_logged=True
    )


def get_great_gatsby_data():
    """#11 Great Gatsby Curve: Inequality vs Intergenerational mobility."""
    countries = ['Denmark', 'Norway', 'Finland', 'Sweden', 'Canada',
                 'Australia', 'Germany', 'Japan', 'New Zealand', 'France',
                 'Spain', 'Switzerland', 'Italy', 'UK', 'USA',
                 'Pakistan', 'China', 'Brazil', 'Chile', 'Argentina',
                 'South Africa', 'Peru', 'Singapore']
    # Gini coefficient (income inequality)
    gini = [25.0, 26.0, 26.0, 27.0, 32.0, 34.0, 32.0, 33.0, 33.0, 33.0,
            35.0, 33.0, 36.0, 35.0, 41.0, 30.0, 42.0, 53.0, 44.0, 42.0,
            63.0, 43.0, 42.0]
    # Intergenerational earnings elasticity (higher = less mobility)
    ige = [0.15, 0.17, 0.18, 0.27, 0.19, 0.26, 0.32, 0.34, 0.25, 0.41,
           0.32, 0.46, 0.48, 0.50, 0.47, 0.46, 0.60, 0.58, 0.52, 0.49,
           0.69, 0.60, 0.45]

    return CurveReexamination(
        "Great Gatsby Curve", np.array(gini), np.array(ige),
        x_label="Gini Coefficient", y_label="Intergenerational Earnings Elasticity",
        country_labels=countries, category="Economics"
    )


def get_balassa_samuelson_data():
    """#12 Balassa-Samuelson: Productivity vs Real exchange rate."""
    countries = ['USA', 'Germany', 'Japan', 'UK', 'France', 'Canada',
                 'Australia', 'South Korea', 'Italy', 'Spain', 'Poland',
                 'Mexico', 'Brazil', 'China', 'India', 'Indonesia',
                 'Turkey', 'Russia', 'South Africa', 'Thailand',
                 'Malaysia', 'Philippines', 'Vietnam', 'Egypt',
                 'Nigeria', 'Bangladesh', 'Pakistan', 'Colombia',
                 'Chile', 'Argentina']
    # Relative productivity (GDP per worker, US=100)
    productivity = [100, 88, 72, 80, 82, 85, 87, 75, 75, 68, 55, 35, 28,
                    30, 15, 22, 45, 42, 25, 28, 40, 16, 18, 20, 10, 12,
                    11, 25, 38, 35]
    # Price level (PPP conversion factor / market exchange rate, US=1.0)
    price_level = [1.00, 0.92, 0.95, 0.95, 0.93, 0.88, 0.92, 0.73, 0.85,
                   0.78, 0.55, 0.50, 0.45, 0.52, 0.30, 0.35, 0.38, 0.40,
                   0.40, 0.38, 0.42, 0.32, 0.28, 0.25, 0.22, 0.25, 0.22,
                   0.42, 0.52, 0.38]

    return CurveReexamination(
        "Balassa-Samuelson Effect", np.array(productivity), np.array(price_level),
        x_label="Relative Labor Productivity (US=100)",
        y_label="Relative Price Level (US=1.0)",
        country_labels=countries, category="Economics"
    )


def run_economics_analysis():
    """Run all economics curve analyses."""
    curves = [
        get_phillips_curve_data(),      # 1
        get_laffer_curve_data(),         # 2
        get_kuznets_curve_data(),        # 3
        get_ekc_data(),                  # 4
        get_beveridge_curve_data(),      # 5
        get_okun_data(),                 # 6
        get_engel_curve_data(),          # 7
        get_j_curve_trade_data(),        # 8
        get_rahn_curve_data(),           # 9
        get_gravity_model_data(),        # 10
        get_great_gatsby_data(),         # 11
        get_balassa_samuelson_data(),    # 12
    ]

    results = []
    for curve in curves:
        print(f"  Analyzing: {curve.name} (N={curve.n})...")
        r = curve.run_full_analysis()
        results.append(r)
    return results


if __name__ == "__main__":
    results = run_economics_analysis()
    for r in results:
        print(f"{r['name']}: {r['verdict']['verdict']} (p_full={r['f_test']['p_value']:.4f})")
