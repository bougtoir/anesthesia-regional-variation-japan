"""
Data collection and analysis for Environmental Science / Ecology curves (D29-D34).
"""

import numpy as np
import pandas as pd
from core_analysis import CurveReexamination


def get_species_area_data():
    """#29 Species-Area Curve: Area vs Species richness (S=cA^z)."""
    # Island biogeography data (Caribbean islands + others)
    islands = ['Redonda', 'Saba', 'Montserrat', 'Nevis', 'Barbuda',
               'St. Lucia', 'Martinique', 'Dominica', 'Guadeloupe',
               'Trinidad', 'Jamaica', 'Puerto Rico', 'Hispaniola', 'Cuba',
               'Madagascar', 'New Guinea', 'Borneo', 'Sumatra',
               'Java', 'Sri Lanka', 'Taiwan', 'Iceland', 'Ireland',
               'UK', 'New Zealand', 'Philippines', 'Sulawesi', 'Luzon']
    # Area (km²)
    area = np.array([2.6, 13, 102, 93, 161, 617, 1128, 751, 1628, 5128,
                     10991, 8870, 76192, 109884, 587041, 786363, 743330,
                     473481, 129000, 65610, 36193, 103000, 84421, 242495,
                     268021, 300000, 180681, 104688])
    # Number of bird species (representative)
    species = np.array([5, 12, 40, 35, 45, 68, 95, 80, 110, 150, 200,
                        175, 230, 280, 258, 578, 420, 465, 340, 233,
                        155, 72, 135, 220, 168, 530, 280, 190])

    return CurveReexamination(
        "Species-Area Curve", np.log10(area), np.log10(species),
        x_label="log₁₀(Area, km²)", y_label="log₁₀(Species Count)",
        country_labels=islands, category="Environmental Science",
        x_is_logged=True
    )


def get_hubbert_peak_data():
    """#30 Hubbert Peak Oil: Year vs US oil production."""
    # US crude oil production (million barrels/day), 1930-2023
    years = np.array([1930, 1935, 1940, 1945, 1950, 1955, 1960, 1965, 1970,
                      1975, 1980, 1985, 1990, 1995, 2000, 2005, 2010, 2015,
                      2018, 2019, 2020, 2021, 2022, 2023])
    production = np.array([2.5, 2.7, 3.7, 4.7, 5.4, 6.8, 7.0, 7.8, 9.6,
                           8.4, 8.6, 8.9, 7.4, 6.6, 5.8, 5.2, 5.5, 9.4,
                           11.0, 12.3, 11.3, 11.2, 11.9, 12.9])

    return CurveReexamination(
        "Hubbert Peak Oil (US)", years, production,
        x_label="Year", y_label="Oil Production (million bbl/day)",
        category="Environmental Science"
    )


def get_keeling_curve_data():
    """#31 Keeling Curve: Year vs CO2 concentration."""
    # Mauna Loa CO2 data (annual averages)
    years = np.array([1960, 1965, 1970, 1975, 1980, 1985, 1990, 1995, 2000,
                      2005, 2010, 2015, 2016, 2017, 2018, 2019, 2020, 2021,
                      2022, 2023])
    co2 = np.array([316.9, 320.0, 325.7, 331.1, 338.8, 346.1, 354.4, 360.8,
                    369.5, 379.8, 389.9, 400.8, 404.2, 406.5, 408.5, 411.4,
                    414.2, 416.4, 418.6, 421.1])

    return CurveReexamination(
        "Keeling Curve (CO2)", years, co2,
        x_label="Year", y_label="CO₂ Concentration (ppm)",
        category="Environmental Science"
    )


def get_hanpp_data():
    """#32 HANPP: Economic development vs Land use intensity."""
    countries = ['USA', 'Germany', 'UK', 'France', 'Japan', 'Canada',
                 'Australia', 'South Korea', 'Brazil', 'China', 'India',
                 'Indonesia', 'Mexico', 'Turkey', 'Russia', 'South Africa',
                 'Nigeria', 'Kenya', 'Ethiopia', 'Thailand', 'Argentina',
                 'Poland', 'Spain', 'Italy', 'Vietnam']
    gdp_pc = [76000, 58000, 49000, 50000, 43000, 53000, 55000, 50000,
              16000, 21000, 8800, 14000, 21000, 35000, 34000, 15000,
              5900, 5600, 3200, 19000, 26000, 41000, 44000, 47000, 13000]
    # HANPP as % of net primary productivity
    hanpp = [25, 50, 65, 55, 35, 12, 8, 40, 18, 45, 55, 30, 20, 35,
             10, 22, 40, 45, 50, 35, 30, 55, 45, 50, 40]

    return CurveReexamination(
        "HANPP vs Development", np.array(gdp_pc), np.array(hanpp),
        x_label="GDP per capita (PPP, $)", y_label="HANPP (% of NPP)",
        country_labels=countries, category="Environmental Science"
    )


def get_jevons_paradox_data():
    """#33 Jevons Paradox: Energy efficiency vs Total energy consumption."""
    # US data: Energy intensity (BTU per $ GDP) vs Total energy consumption
    years = np.array([1970, 1975, 1980, 1985, 1990, 1995, 2000, 2005,
                      2010, 2015, 2020, 2022])
    # Energy intensity (thousand BTU per 2012$)
    intensity = np.array([14.0, 13.5, 12.5, 10.8, 10.0, 9.3, 8.5, 8.0,
                          7.2, 6.2, 5.5, 5.2])
    # Total primary energy consumption (quadrillion BTU)
    total_energy = np.array([67.8, 72.0, 78.1, 76.4, 84.7, 91.2, 98.8, 100.3,
                             97.5, 97.7, 92.9, 100.4])

    return CurveReexamination(
        "Jevons Paradox (Energy)", intensity, total_energy,
        x_label="Energy Intensity (kBTU/$GDP)", y_label="Total Energy (quad BTU)",
        category="Environmental Science"
    )


def get_forest_transition_data():
    """#34 Forest Transition Curve: Development vs Forest area."""
    countries = ['Ethiopia', 'Nigeria', 'Tanzania', 'Kenya', 'Ghana',
                 'Indonesia', 'Brazil', 'India', 'Vietnam', 'China',
                 'Thailand', 'Mexico', 'Turkey', 'Poland', 'South Korea',
                 'Chile', 'Spain', 'Italy', 'France', 'Germany', 'Japan',
                 'UK', 'USA', 'Canada', 'Sweden', 'Finland', 'Costa Rica',
                 'Cuba', 'New Zealand', 'Ireland']
    gdp_pc = [3200, 5900, 3500, 5600, 6500, 14000, 16000, 8800, 13000,
              21000, 19000, 21000, 35000, 41000, 50000, 28000, 44000,
              47000, 50000, 58000, 43000, 49000, 76000, 53000, 60000,
              54000, 23000, 12000, 48000, 106000]
    # Forest area as % of land area (~2020)
    forest_pct = [15.4, 7.7, 52.0, 7.8, 21.3, 49.1, 59.4, 24.3, 42.0,
                  23.3, 31.6, 34.0, 28.4, 30.9, 63.7, 24.5, 36.8, 32.4,
                  31.4, 32.7, 68.4, 13.1, 33.9, 38.7, 68.9, 73.7, 54.0,
                  30.1, 37.8, 11.0]

    return CurveReexamination(
        "Forest Transition Curve", np.array(gdp_pc), np.array(forest_pct),
        x_label="GDP per capita (PPP, $)", y_label="Forest Area (% of land)",
        country_labels=countries, category="Environmental Science"
    )


def run_environment_analysis():
    """Run all environmental science curve analyses."""
    curves = [
        get_species_area_data(),        # 29
        get_hubbert_peak_data(),        # 30
        get_keeling_curve_data(),       # 31
        get_hanpp_data(),               # 32
        get_jevons_paradox_data(),      # 33
        get_forest_transition_data(),   # 34
    ]

    results = []
    for curve in curves:
        print(f"  Analyzing: {curve.name} (N={curve.n})...")
        r = curve.run_full_analysis()
        results.append(r)
    return results


if __name__ == "__main__":
    results = run_environment_analysis()
    for r in results:
        print(f"{r['name']}: {r['verdict']['verdict']} (p_full={r['f_test']['p_value']:.4f})")
