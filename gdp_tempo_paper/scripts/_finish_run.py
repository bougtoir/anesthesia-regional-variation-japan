"""Helper: finish the pipeline by running only analysis 5 + figures using
cached CSVs from the first run."""
import os, json, sys
import pandas as pd
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run_paper_analyses import (DATA, FIG, prepare_countries,
                                run_gamma_price, make_figures)

countries = prepare_countries()
fair = pd.read_csv(os.path.join(DATA, "fair_eval.csv"))
gamma = run_gamma_price(countries, fair)
gamma.to_csv(os.path.join(DATA, "gamma_price.csv"), index=False)

oos = pd.read_csv(os.path.join(DATA, "oos.csv"))
boot = pd.read_csv(os.path.join(DATA, "bootstrap_ci.csv"))
make_figures(fair, oos, boot, gamma)
print("Done.")
