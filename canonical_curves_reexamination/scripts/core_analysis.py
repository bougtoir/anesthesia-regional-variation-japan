"""
Core statistical framework for re-examining canonical curves.

Methods:
- Nested F-test (linear vs quadratic/log/exponential)
- AIC/BIC comparison
- Leave-One-Out Cross-Validation (LOOCV) RMSE
- Cook's distance sensitivity analysis
- Outlier exclusion and re-testing

Onishi T. 2026.
"""

import numpy as np
import pandas as pd
from scipy import stats
from sklearn.model_selection import LeaveOneOut
from sklearn.metrics import mean_squared_error
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import OLSInfluence
import warnings
warnings.filterwarnings('ignore')


class CurveReexamination:
    """Framework for re-examining a single canonical curve."""

    def __init__(self, name, x, y, x_label="x", y_label="y",
                 country_labels=None, category="", x_is_logged=False):
        self.name = name
        self.x = np.array(x, dtype=float)
        self.y = np.array(y, dtype=float)
        self.x_label = x_label
        self.y_label = y_label
        self.country_labels = country_labels
        self.category = category
        self.x_is_logged = x_is_logged
        self.n = len(x)
        self.results = {}

    def fit_linear(self, x=None, y=None):
        """Fit y = a + b*x."""
        if x is None:
            x, y = self.x, self.y
        X = sm.add_constant(x)
        model = sm.OLS(y, X).fit()
        return model

    def fit_quadratic(self, x=None, y=None):
        """Fit y = a + b*x + c*x^2."""
        if x is None:
            x, y = self.x, self.y
        X = np.column_stack([np.ones(len(x)), x, x**2])
        model = sm.OLS(y, X).fit()
        return model

    def fit_log(self, x=None, y=None):
        """Fit y = a + b*log(x)."""
        if x is None:
            x, y = self.x, self.y
        mask = x > 0
        x_log = np.log(x[mask])
        X = sm.add_constant(x_log)
        model = sm.OLS(y[mask], X).fit()
        return model

    def fit_exponential(self, x=None, y=None):
        """Fit log(y) = a + b*x (exponential relationship)."""
        if x is None:
            x, y = self.x, self.y
        mask = y > 0
        y_log = np.log(y[mask])
        X = sm.add_constant(x[mask])
        model = sm.OLS(y_log, X).fit()
        return model

    def nested_f_test(self, x=None, y=None):
        """
        F-test: linear (restricted) vs quadratic (unrestricted).
        H0: quadratic term is not significant.
        """
        if x is None:
            x, y = self.x, self.y

        model_lin = self.fit_linear(x, y)
        model_quad = self.fit_quadratic(x, y)

        rss_r = model_lin.ssr  # restricted (linear)
        rss_u = model_quad.ssr  # unrestricted (quadratic)
        df_r = model_lin.df_resid
        df_u = model_quad.df_resid
        p = df_r - df_u  # number of additional parameters (1 for quadratic)

        if rss_u == 0:
            return np.inf, 0.0

        f_stat = ((rss_r - rss_u) / p) / (rss_u / df_u)
        p_value = 1 - stats.f.cdf(f_stat, p, df_u)

        return f_stat, p_value

    def compute_aic_bic(self, x=None, y=None):
        """Compute AIC and BIC for linear, quadratic, log models."""
        if x is None:
            x, y = self.x, self.y

        results = {}

        # Linear
        model_lin = self.fit_linear(x, y)
        results['linear'] = {'aic': model_lin.aic, 'bic': model_lin.bic,
                             'r2': model_lin.rsquared, 'r2_adj': model_lin.rsquared_adj}

        # Quadratic
        model_quad = self.fit_quadratic(x, y)
        results['quadratic'] = {'aic': model_quad.aic, 'bic': model_quad.bic,
                                'r2': model_quad.rsquared, 'r2_adj': model_quad.rsquared_adj}

        # Log (if x > 0 and not already log-transformed)
        if np.all(x > 0) and not self.x_is_logged:
            model_log = self.fit_log(x, y)
            results['log'] = {'aic': model_log.aic, 'bic': model_log.bic,
                              'r2': model_log.rsquared, 'r2_adj': model_log.rsquared_adj}

        return results

    def loocv_rmse(self, x=None, y=None):
        """LOOCV RMSE for linear and quadratic models."""
        if x is None:
            x, y = self.x, self.y

        loo = LeaveOneOut()
        errors_lin = []
        errors_quad = []

        for train_idx, test_idx in loo.split(x):
            x_train, x_test = x[train_idx], x[test_idx]
            y_train, y_test = y[train_idx], y[test_idx]

            # Linear
            X_train = sm.add_constant(x_train, has_constant='add')
            X_test = np.array([[1.0, x_test[0]]])
            model_lin = sm.OLS(y_train, X_train).fit()
            pred_lin = model_lin.predict(X_test)
            errors_lin.append((y_test[0] - pred_lin[0])**2)

            # Quadratic
            X_train_q = np.column_stack([np.ones(len(x_train)), x_train, x_train**2])
            X_test_q = np.array([[1.0, x_test[0], x_test[0]**2]])
            model_quad = sm.OLS(y_train, X_train_q).fit()
            pred_quad = model_quad.predict(X_test_q)
            errors_quad.append((y_test[0] - pred_quad[0])**2)

        rmse_lin = np.sqrt(np.mean(errors_lin))
        rmse_quad = np.sqrt(np.mean(errors_quad))

        return rmse_lin, rmse_quad

    def cooks_distance(self, x=None, y=None):
        """Compute Cook's distance for linear model."""
        if x is None:
            x, y = self.x, self.y

        X = sm.add_constant(x)
        model = sm.OLS(y, X).fit()
        influence = OLSInfluence(model)
        cooks_d = influence.cooks_distance[0]
        return cooks_d

    def sensitivity_analysis(self, top_k=3):
        """
        Remove top-k influential points (by Cook's distance) and re-test.
        Returns F-test p-value with and without outliers.
        """
        cooks_d = self.cooks_distance()
        top_indices = np.argsort(cooks_d)[-top_k:]

        # Full data
        f_full, p_full = self.nested_f_test()

        # Without outliers
        mask = np.ones(self.n, dtype=bool)
        mask[top_indices] = False
        x_clean = self.x[mask]
        y_clean = self.y[mask]

        if len(x_clean) < 4:
            return {
                'f_full': f_full, 'p_full': p_full,
                'f_clean': np.nan, 'p_clean': np.nan,
                'removed_indices': top_indices,
                'removed_labels': None,
                'n_full': self.n, 'n_clean': len(x_clean)
            }

        f_clean, p_clean = self.nested_f_test(x_clean, y_clean)

        removed_labels = None
        if self.country_labels is not None:
            labels = np.array(self.country_labels)
            removed_labels = labels[top_indices].tolist()

        return {
            'f_full': f_full, 'p_full': p_full,
            'f_clean': f_clean, 'p_clean': p_clean,
            'removed_indices': top_indices,
            'removed_labels': removed_labels,
            'n_full': self.n, 'n_clean': len(x_clean)
        }

    def run_full_analysis(self, top_k=3):
        """Run all analyses and return comprehensive results."""
        results = {
            'name': self.name,
            'category': self.category,
            'n': self.n,
            'x_label': self.x_label,
            'y_label': self.y_label,
        }

        # F-test
        f_stat, p_value = self.nested_f_test()
        results['f_test'] = {'f_stat': f_stat, 'p_value': p_value}

        # AIC/BIC
        results['model_comparison'] = self.compute_aic_bic()

        # LOOCV
        rmse_lin, rmse_quad = self.loocv_rmse()
        results['loocv'] = {'rmse_linear': rmse_lin, 'rmse_quadratic': rmse_quad}

        # Cook's distance
        cooks_d = self.cooks_distance()
        results['cooks_distance'] = {
            'max': float(np.max(cooks_d)),
            'mean': float(np.mean(cooks_d)),
            'n_influential': int(np.sum(cooks_d > 4/self.n))
        }

        # Sensitivity
        sensitivity = self.sensitivity_analysis(top_k=top_k)
        results['sensitivity'] = sensitivity

        # Verdict
        results['verdict'] = self._determine_verdict(results)

        self.results = results
        return results

    def _determine_verdict(self, results):
        """Determine whether nonlinearity is robust."""
        p_full = results['f_test']['p_value']
        p_clean = results['sensitivity']['p_clean']
        rmse_lin = results['loocv']['rmse_linear']
        rmse_quad = results['loocv']['rmse_quadratic']

        # Check model comparison
        mc = results['model_comparison']
        best_aic = min(mc[m]['aic'] for m in mc)
        best_model_aic = [m for m in mc if mc[m]['aic'] == best_aic][0]

        best_bic = min(mc[m]['bic'] for m in mc)
        best_model_bic = [m for m in mc if mc[m]['bic'] == best_bic][0]

        if p_full > 0.05:
            verdict = "NOT_SIGNIFICANT"
            explanation = "Nonlinear term not significant even with full data (p={:.3f})".format(p_full)
        elif np.isnan(p_clean):
            verdict = "INSUFFICIENT_DATA"
            explanation = "Too few data points for sensitivity analysis"
        elif p_clean > 0.05:
            verdict = "OUTLIER_DEPENDENT"
            explanation = ("Nonlinearity significant with full data (p={:.3f}) "
                          "but NOT after removing top influential points (p={:.3f})").format(p_full, p_clean)
        elif rmse_quad >= rmse_lin:
            verdict = "OVERFITTING"
            explanation = ("Nonlinearity significant but LOOCV shows quadratic overfits "
                          "(RMSE_lin={:.3f}, RMSE_quad={:.3f})").format(rmse_lin, rmse_quad)
        elif best_model_bic == 'linear':
            verdict = "BIC_PREFERS_LINEAR"
            explanation = "BIC prefers linear model despite significant F-test"
        else:
            verdict = "ROBUST_NONLINEAR"
            explanation = ("Nonlinearity is robust: significant F-test (p={:.3f}), "
                          "survives outlier removal (p={:.3f}), better LOOCV").format(p_full, p_clean)

        return {'verdict': verdict, 'explanation': explanation,
                'best_aic': best_model_aic, 'best_bic': best_model_bic}


def format_results_table(all_results):
    """Format all results into a summary DataFrame."""
    rows = []
    for r in all_results:
        row = {
            'Category': r['category'],
            'Curve': r['name'],
            'N': r['n'],
            'F-stat': r['f_test']['f_stat'],
            'p (full)': r['f_test']['p_value'],
            'p (clean)': r['sensitivity']['p_clean'],
            'AIC best': r['verdict']['best_aic'],
            'BIC best': r['verdict']['best_bic'],
            'RMSE_lin': r['loocv']['rmse_linear'],
            'RMSE_quad': r['loocv']['rmse_quadratic'],
            'Cook max': r['cooks_distance']['max'],
            'N influential': r['cooks_distance']['n_influential'],
            'Verdict': r['verdict']['verdict'],
        }
        rows.append(row)
    return pd.DataFrame(rows)
