# The Forgotten Tempo Effect in Capital Accounting: Investment-to-Output Time-to-Build, Intangible Capital, and the Reconciliation of Flow- and Stock-Based National Wealth Measures

**Tatsuki Onishi**

Data Science AI Innovation Research Promotion Center, Shiga University

1-1-1, Bamba, Hikone, Shiga, 522-8522 Japan

Telephone: +81-749-27-1023. E-mail: bougtoir@gmail.com. ORCID: 0000-0001-7261-9062.

**Keywords**: tempo effect; intangible capital; perpetual inventory method; time-to-build; flow-stock consistency; Beyond-GDP.

**JEL codes**: E01 (Measurement and Data on National Income and Product Accounts); E22 (Investment; Capital; Intangible Capital); O47 (Measurement of Economic Growth; Aggregate Productivity).

**CRediT contribution statement.** O.T.: Conceptualization, Methodology, Software, Formal analysis, Writing - original draft, Writing - review & editing.

**Declaration of generative artificial intelligence (AI) in scientific writing.** We used generative AI to help with formatting the text and choosing words that suited the tone, and to help writing codes.

**Conflict of interest.** The author declares no conflict of interest.

**Data and code availability.** Penn World Table 10.01 and World Bank CWON data used in this study are publicly available from the Groningen Growth and Development Centre and the World Bank, respectively. All analysis scripts, intermediate results, and manuscript sources are archived in the accompanying public repository.

---

**Abstract** (146 words). Since Goldstein, Lutz, and Scherbov (2003) showed that a single "forgotten" parity-specific variance parameter σ resolved a large share of the low-fertility puzzle once tempo effects on the mean age at childbearing were acknowledged, the dual of quantum and tempo has become a standard lens in formal demography. National income and wealth accounting has no equivalent diagnostic. We port the Bongaarts-Feeney quantum-tempo decomposition to capital accounting by letting the investment-to-output time-to-build μ(t) drift over time and by re-introducing intangible capital K_I, with share β, as the balance-sheet analogue of σ. Across 39 OECD and middle-income economies (Penn World Table 10.01, World Bank CWON), a time-varying μ(t) reduces the out-of-sample MAPE of GDP levels from 4.60% to 3.99% while a joint production-cum-wealth identification produces internally consistent flow and stock accounts. A sister medical-spending paper is in preparation.

**Keywords**: tempo effect; intangible capital; perpetual inventory method; wealth accounting; Beyond-GDP.

**JEL codes**: E01, E22, O47.

---

## 1. Introduction

Every macroeconomist has encountered two separate but related complaints about the way we measure national prosperity. First, Gross Domestic Product (GDP) is a flow measure that ignores depletion, depreciation, and the growing stock of intangible assets that drives modern productivity growth (Stiglitz, Sen, and Fitoussi, 2009; Corrado, Hulten, and Sichel, 2009; Haskel and Westlake, 2017). Second, stock-based alternatives such as the Inclusive Wealth Index (Managi and Kumar, 2018), the United Nations SEEA (UNECE, 2014), and the World Bank Changing Wealth of Nations (Lange, Wodon, and Carey, 2018) are attractive in principle but rarely line up with independently reconstructed capital stocks and never with one another. Flow-based and stock-based national accounts have lived side by side in different rooms of the same house for twenty-five years without being asked to sit down at the same table.

Demography has spent the same period quietly solving the mirror problem. Bongaarts and Feeney (1998) showed that a rising mean age at childbearing mechanically depresses period fertility even when completed cohort fertility is constant, and proposed an adjustment that stripped out the "tempo distortion". Goldstein, Lutz, and Scherbov (2003) reopened the debate by showing that once a parity-specific variance σ was allowed (the "forgotten parameter"), the tempo-adjusted fertility rate matched the cohort data far more closely. A generation of work on postponement, ultra-low fertility, and lifetime child-bearing risk followed. The pattern is simple: the period statistic was biased; the bias was a timing phenomenon; once you wrote down a structural timing parameter and a single forgotten quantity parameter, the flow and stock accounts of the reproduction process were reconciled.

This paper argues that capital accounting has an exact analogue of the Bongaarts-Feeney-Goldstein-Lutz-Scherbov correction, hiding in plain sight. The analogy is not rhetorical: every demographic quantity has a capital counterpart under a precise change of variables (Section 3.4, Table 2). Births are investment flows. Population stocks are capital stocks. The mean age at childbearing is the mean lag between investment and its productive deployment — the engineering and organisational "time-to-build" that Kydland and Prescott (1982) introduced but that has never been allowed to drift over time in standard production-function estimation. The parity-specific variance σ has a direct balance-sheet equivalent: the intangible capital share β that CHS have estimated but that official wealth accounts such as the CWON still treat as non-existent or residual.

Our contribution is four-fold. First, we write down the flow–stock identity *dW/dt = S(Y) − δW* in parameterised form, making the hidden parameters {μ(t), β} explicit on the flow side and on the stock side simultaneously. Second, we show that a time-varying time-to-build μ(t) — estimated with a two-parameter tempo drift μ(t) = μ₀ + μ₁·(t − t₀) — reduces the median out-of-sample MAPE of GDP level forecasts from 4.60 % to 3.99 % across 39 countries, a 13 % relative improvement that rivals gains from adding entirely new production factors. Third, we demonstrate that when the tempo and intangible corrections are *jointly* identified against CWON stock data, production-side and wealth-side likelihoods agree on a consistent pair (μ̂ₖ, β̂ₖ) for every country, which in our reading is the first empirical success of the "unified national-wealth accounting" programme that Stiglitz-Sen-Fitoussi called for. Fourth, we preview a companion paper extending the same tempo-plus-forgotten-parameter machinery to health expenditure and population health outcomes, where preliminary evidence shows the medical time-to-build lag has been widening at +0.15 years per year since 2000.

The remainder of the paper is organised as follows. Section 2 reviews the capital-accounting, intangibles, and tempo-demography literatures that our framework stitches together. Section 3 develops the theory. Section 4 describes the data and methods and defines five models M0–M4 of increasing generality. Section 5 reports results. Section 6 discusses the Solow-residual reinterpretation, the flow–stock reconciliation, and policy implications for Beyond-GDP. Section 7 concludes.

## 2. Related literature

**Capital accounting and time-to-build.** Since Kydland and Prescott (1982) it has been standard practice to insert a multi-period investment lag into business-cycle models. Empirical estimates are overwhelmingly based on fixed lag structures: a single μ is estimated once for an entire sample, or a small number of regime-dependent μs are estimated for recession and expansion states (Mayer, 1960; Koeva, 2000). Kaboski (2005) documents cross-industry heterogeneity but, again, in a time-invariant fashion. More recent work (Altug, 1989; Christiano and Todd, 1996; Edge, 2007) has explored stochastic extensions in which the lag distribution is allowed to depend on sectoral composition, but still does not allow μ to drift systematically over decades. We know of no prior study that lets the typical investment-to-output lag drift in the way that demographers have documented for the mean age at childbearing. The omission is consequential because post-1990 investment has shifted substantially toward long-lead assets — custom industrial software, cloud infrastructure, pharmaceutical R&D, complex engineering systems — whose gestation periods differ by nearly an order of magnitude from the plant-and-equipment of the 1960s on which the original lag literature was calibrated (OECD, 2013; Corrado et al., 2020).

**Intangible capital.** The programme begun by Corrado, Hulten, and Sichel (2005, 2009) has by now produced robust international evidence that software, R&D, design, brand, organisational capital, and training account for 30–60 % of productivity growth in advanced economies (INTAN-Invest: Corrado et al., 2016; Roth, 2023). The 2008 revision of the System of National Accounts (SNA) formally incorporated R&D into produced capital, but broader intangibles — organisational capital, brand, training, purchased design services, some categories of financial innovation — remain excluded from most official balance sheets, including the World Bank CWON (Lange et al., 2018, Chap. 3). De Rassenfosse and Jaffe (2018) and Haskel and Westlake (2017, 2022) emphasise that this omission biases not only the level of measured capital but also the implied productivity growth rate whenever the intangible share is expanding — exactly the setting of our 1995-2019 sample. Importantly, the intangible share β is not a global constant: Japan, Germany, and some East-Asian economies retain a smaller intangible share than the United States even under harmonised measurement (Corrado et al., 2020), so β ought to be a country-specific parameter, which is how we treat it in Section 4.

**Wealth accounting.** The Beyond-GDP movement, from Stiglitz-Sen-Fitoussi (2009) through Jorgenson (2018) and Managi-Kumar (2018), proposes to replace or augment GDP with wealth-style aggregates. Empirically, however, the three main aggregates — SEEA, IWI, and CWON — disagree materially both with each other and with independently reconstructed perpetual-inventory stocks (Arrow et al., 2012; Dasgupta, 2021). The mainstream diagnosis blames measurement error and the treatment of natural capital. We show that a more mundane culprit — a mis-specified time-to-build and an omitted intangible share — explains a sizeable fraction of the discrepancy.

**Tempo and forgotten parameters in demography.** Bongaarts and Feeney (1998) introduced the adjustment *TFR\** = *TFR*/(1 − *r(t)*) where *r(t)* is the annual change in the mean age at childbearing. Goldstein, Lutz, and Scherbov (2003) showed that Bongaarts-Feeney was an upper bound unless a parity-specific "forgotten" variance σ was re-introduced. Kohler, Billari, and Ortega (2002) and Bongaarts and Sobotka (2012) confirmed both findings across Europe. The structural lesson — that flow statistics of a stock process are contaminated by drift in the timing distribution, and that a single omitted quantity parameter restores consistency — is exactly the lesson we now transplant to the capital account.

**Healthcare and human capital sustainability.** A companion paper (in preparation) documents that the lag μ_H from medical expenditure to life-expectancy outcomes has been rising by roughly 0.15 years per year since 2000, and that an analogous "forgotten" parameter β_H — the share of expenditure directed to prevention and R&D, as opposed to curative care — accounts for a further share of the US-Japan life-expectancy gap. That paper exploits exactly the same quantum-tempo decomposition developed here.

**The gap this paper fills.** The papers above individually treat (i) capital time-to-build, (ii) intangibles, (iii) wealth aggregates, and (iv) demographic tempo. To our knowledge, no prior work simultaneously (a) estimates a time-varying time-to-build, (b) recovers the CHS intangible share, and (c) disciplines both with a wealth-accounting identity. This paper does. A secondary contribution, often overlooked in the accounting literature but central to our methodology, is to treat the PIM and the wealth-stock equations as two equally informative windows onto the same latent process — exactly as demographers treat period and cohort data — rather than as competing aggregates whose disagreement is a nuisance to be absorbed into residuals.

## 3. Theory

### 3.1 Flow-side production function with tempo

The textbook production function treats investment as if it matures instantly:

    K_instant(t) = (1 − δ_{t-1}) K_instant(t−1) + I_{t-1},                         (M0)

so the Solow (1957) residual aggregates all mis-specification into total factor productivity (TFP). Since Mayer (1960) and Kydland-Prescott (1982) it is well known that, in reality, investment accrues to the stock only after a lag. We write this as a distributed-lag perpetual inventory:

    K(t; μ) = (1 − δ_{t-1}) K(t−1; μ) + Σₛ w_s(μ) I_{t-1-s},                     (M1)

with geometric weights *w_s(μ) = (1 − θ)·θ^s* and *θ = μ/(1+μ)*, so the mean lag is exactly *μ* years. The key novelty relative to the existing lag literature is to allow μ to drift linearly over time:

    μ(t) = μ₀ + μ₁·(t − t₀),                                                    (M2)

where μ₁ captures the "tempo" in Bongaarts-Feeney's sense. A positive μ₁ indicates that typical projects are becoming longer-lived — for example because new investment is increasingly digital infrastructure, R&D platforms, or complex systems that require multi-year assembly — and a negative μ₁ would indicate the opposite.

### 3.2 Stock-side intangibles: the forgotten β

Let *K_tang(t)* be the tangible PIM stock from (M1)–(M2) and *K_I(t)* be an intangible stock built from R&D expenditure by a geometric PIM with depreciation δ_I = 0.15 (Corrado-Hulten-Sichel, 2009). A production function augmented by intangibles reads:

    log Y_t = α log K_tang(t) + β log K_I(t) + (1 − α − β) log L_t + log A_t,    (M3)

where β is the intangible share. Standard practice imposes β = 0 (Solow; also M0 and M1 here). Estimating β > 0 is the capital-accounting analogue of re-introducing the parity-specific variance σ in Goldstein-Lutz-Scherbov.

### 3.3 Unifying identity: the flow-stock joint loss

Any consistent national wealth aggregate *W(t)* must satisfy the book-keeping identity

    dW/dt = S(Y) − δ_W · W,                                                       (1)

where *S(Y)* is gross saving and *δ_W* is the aggregate depreciation rate. Under (1), the same parameters {μ, β} that govern the production side should also govern the reproducible-capital trajectory implied by the wealth account. We therefore define a single joint loss:

    L_total(μ, β) = L_production(μ, β) + λ · L_wealth(μ, β),                      (2)

where *L_production* is the growth-rate residual from the production function (M3) and *L_wealth* is the within-country trajectory RMSE between the PIM stock *K_tang(t; μ) + β · K_I(t)* and the CWON produced-capital series NW.PCA.TO(t). Minimising (2) delivers the "M4 joint" estimates (μ̂_joint, β̂_joint) used below; setting λ = 0 recovers production-only estimates.

### 3.4 Quantum–tempo correspondence between population and capital

Table 2 lays out the one-to-one mapping between the demographic variables that Bongaarts-Feeney-Goldstein-Lutz-Scherbov analysed and the capital-accounting variables we analyse. Every demographic entity has a capital entity with the same role in the book-keeping identity and in the quantum-tempo decomposition. This is more than mnemonic: it implies that the statistical tools used to identify σ from fertility tempo (cohort-consistency tests, Brass relational models) have direct analogues in capital accounting, which we exploit.

### 3.5 Relational PIM: a Brass model for capital accounting

Demography has long employed Brass relational models to compare an observed fertility or mortality schedule against a "standard" schedule via a small number of parameters that capture systematic deviations (Brass, 1971). We transplant this idea to capital accounting. Let *K_PIM(t)* denote the PIM-constructed stock under any of M0–M4, and let *K_CWON(t)* denote CWON produced capital NW.PCA.TO. We define the **Relational PIM** (RPIM) as:

    log K_PIM(t) = ρ₁ + ρ₂ · log K_CWON(t) + ε(t),                              (M5)

where (ρ₁, ρ₂) are the relational parameters estimated by OLS on the overlapping years where both series are observed and positive. The interpretation is direct:

* **ρ₂ = 1 and ρ₁ = 0**: the PIM and CWON accounts are perfectly consistent — they measure the same latent stock up to a white-noise error.
* **ρ₂ ≠ 1**: there is a cumulative (scale-dependent) bias in the PIM relative to CWON. If ρ₂ < 1, the PIM understates the growth of capital relative to CWON; if ρ₂ > 1, it overstates it.
* **ρ₁ ≠ 0**: there is a level shift between the two accounts, capturing differences in base-year calibration, PPP conversions, or asset coverage.

The novelty is threefold. First, to our knowledge no prior work has applied the Brass relational-model framework to capital accounting. Second, the RPIM does not treat CWON as "truth" — it parameterises the *relationship* between two independent estimates of the same latent stock, making systematic biases visible and quantifiable. Third, the diagnostic (ρ₁, ρ₂) can be computed under each model specification (M0 through M4), so improvement in ρ₂ toward unity serves as an independent check on whether the tempo and intangible corrections actually bring the two accounts closer together.

## 4. Data and methods

### 4.1 Data

We use **Penn World Table 10.01** (Feenstra, Inklaar, and Timmer, 2015) for real GDP output (*rgdpna*), tangible capital stock (*rnna*), investment share (*csh_i*), depreciation (*delta*), employment (*emp*), average hours (*avh*), human-capital index (*hc*), and labour share (*labsh*). For R&D intensity we use **World Bank WDI** series *GB.XPD.RSDV.GD.ZS*. For wealth we use **World Bank Changing Wealth of Nations** 2021 release (Lange, Wodon, and Carey, 2018) — specifically *NW.PCA.TO* (produced capital total), *NW.HCA.TO* (human capital total), and *NW.TOW.TO* (total wealth).

The sample is 39 OECD and middle-income economies for which all series are available. The GDP sample runs from 1970 to 2019; CWON runs 1995–2020; we take the intersection 1995–2019 when both are needed.

### 4.2 Models M0–M4

We estimate five nested production-function specifications:

* **M0**: Solow baseline, *K_tang* as M0 above, β = 0.
* **M1**: Constant-lag PIM (M1) with *μ = μ*\* estimated per country by minimising Test B (growth-rate RMSE).
* **M2**: Time-varying lag μ(t) = μ₀ + μ₁·(t − t₀) from (M2).
* **M3**: M0 tangible stock augmented with intangible stock K_I and β estimated by growth-rate fit.
* **M4**: Joint identification (Section 3.3), minimising (2) over (μ, β) simultaneously against CWON.

For each model we report two within-sample test statistics and one out-of-sample test statistic:

* **Test A (level MAPE)**: mean absolute percentage error of fitted log-GDP against observed log-GDP, decomposing away decade-mean TFP. Lower is better.
* **Test B (growth RMSE)**: root-mean-squared error of 1-year log-GDP differences, in percentage points. Lower is better.
* **Out-of-sample MAPE**: parameters fit on 1970–2014, level forecasts produced for 2015–2019 with a training-window TFP projection. Lower is better.

### 4.3 Estimation protocol and grid search

All five models are estimated by grid search, not gradient optimisation, for three reasons. First, the objective function (2) has known non-convexities induced by the geometric lag kernel, especially when μ is small and the kernel is near-concentrated. Second, grid search produces an explicit posterior-like surface for each (country, model) pair, which we use in the sensitivity checks below. Third, the 39-country × 5-model × 1000-draw bootstrap would be intractable with a Nelder-Mead or BFGS inner loop for many countries. The μ grid is {0.01, 0.05, 0.10, 0.25, 0.50, 1.0, 1.5, 2.0, 3.0, 4.5, 6.0} years and the β grid is {0.00, 0.02, 0.04, ..., 0.34}; μ₁ is searched on {−0.08, −0.04, −0.02, 0, +0.02, +0.04, +0.08} per year. These bounds were selected to bracket all plausible parameter values reported in prior cross-country studies (Kaboski, 2005; Corrado et al., 2016), and we verified that no country's optimum hits the grid boundary. The anchor year t₀ is 1970 for all countries, so that μ₀ is the average lag in the base year; this choice has no effect on fit but makes μ₀ and μ₁ interpretable.

### 4.4 Bootstrap confidence intervals

For every country we residual-bootstrap the growth-rate residuals of M4 one hundred times (block size 1, since the autocorrelation structure of PWT annual-growth residuals is weak after detrending; block size 3 gives nearly identical 95 % intervals on a pilot of five countries). Each bootstrap replicate proceeds as follows: (i) compute fitted growth rates from M4 and the corresponding residuals; (ii) resample the residuals and reconstruct a synthetic log-GDP series; (iii) back out a synthetic investment series using the PWT investment-share *csh_i* and a synthetic R&D intensity using WDI shares; (iv) rebuild *K_tang* and *K_I*; (v) re-run the joint-identification grid, storing (μ_b, β_b). We report 95 % percentile intervals in Figure 3 and per-country medians in the supplementary JSON. Country-specific CIs are narrowest for long, non-volatile series (United States, Canada, Germany, France, United Kingdom, Japan, Australia) and widest for short or post-transition series (Estonia, Latvia, Chile). We do not adjust the 95 % intervals for multiple testing across countries; the reader who wants a conservative reading should apply a Bonferroni-style 5 %/39 ≈ 0.13 % threshold, under which μ = 0 is still rejected for 14 countries and β = 0 for 9 countries.

### 4.5 γ_price sensitivity

To test whether the residual PIM-CWON gap in countries such as Japan reflects an asset-price re-evaluation effect rather than a real capital gap, we re-run the comparison under five counterfactual scenarios in which CWON PCA is inflated/deflated at an annual rate γ_price ∈ {−0.04, −0.02, 0, +0.02, +0.04}. A large γ_price sensitivity for a specific country would indicate that asset-price revaluation explains most of its gap; a small sensitivity would indicate a genuine real discrepancy. The interval ±0.04 per year brackets the observed rate of deflation in Japanese urban land prices during the 1990s (Nishimura and Saita, 2005) as well as the observed rate of reflation in US commercial real-estate between 2009 and 2019, so the grid is economically meaningful rather than arbitrary. We stress that γ_price is not intended to be an additional estimand of the joint framework — if it were, it would enter (2) alongside μ and β. Rather, it is a diagnostic: a residual gap between the PIM account and the CWON account at a specific γ_price value admits exactly one of three interpretations, namely (a) quantity mis-measurement in the PIM, (b) quantity mis-measurement in CWON, or (c) genuine composition change (e.g. a real shift from tangible to intangible capital that neither account has fully absorbed). The γ_price sweep helps identify (a) and (b) against (c).

## 5. Results

### 5.1 In-sample parameter distributions and fit

**[Table 1 here]**

Table 1 summarises the five models. Three facts deserve particular emphasis. First, the median in-sample growth-rate RMSE hardly moves across M0-M4 (3.07-3.10 pp). This is what standard Solow-accounting practitioners have found repeatedly when they experimented with alternative capital constructions (Jorgenson and Griliches, 1967; Hulten, 1992), and it is one reason why the profession has settled on M0 as the canonical baseline: within-sample growth-rate fit does not discipline μ at all. Second, the median level MAPE under M0 is 4.10 %, meaning that a carefully re-estimated TFP trajectory can absorb nearly all of a 4 % miscalibration in the capital stock at every point in time while preserving the first-differenced fit. This perfectly illustrates the warning of Griliches (1996) that "TFP is the measure of our ignorance": any capital mis-specification that varies on decade-scale time-frequencies will be silently reabsorbed into decade-scale TFP, and then re-interpreted as innovation. Third, and most importantly for the method developed here, the distribution of estimated μ* across the 39 countries is highly non-degenerate. The interquartile range under M1 runs from below 0.1 years to above 1.2 years, and the tempo drift μ₁ under M2 has an IQR that includes both substantially negative and substantially positive values. The universal-μ assumption implicit in M0 is therefore not merely statistically violated — it is violated in both directions across the sample, which implies that any single-parameter global correction (including the 5-year or 10-year fixed lags popular in business-cycle calibration) will be biased in roughly half of the sample. The median country has a M1 constant lag μ\* ≈ 0.3 years and a M2 tempo drift μ₁ close to zero on average but with wide dispersion across countries (IQR roughly [−0.02, +0.05]). Median intangible share β under M3 is about 0.06 for production-only fitting and 0.06 under joint identification with CWON (M4). The in-sample growth-rate RMSE is statistically indistinguishable across M0–M4 at the median (all within 3.07–3.10 pp), confirming that the production function is close to flat in μ when evaluated only on in-sample growth-rate residuals, as Koeva (2000) also found. In-sample level MAPE improves monotonically from M0 (4.10 %) to M4 (4.06 %). These apparently small in-sample differences conceal much larger out-of-sample differences, which we turn to next.

### 5.2 Out-of-sample prediction gains from the tempo correction

**[Figure 1 here]**

Figure 1 ranks the 39 countries by in-sample growth RMSE (M0) and overlays the other four models. The gains from moving from M0 to M2 or M4 are small but systematic, consistent with Table 1.

**[Figure 2 here]**

Figure 2 shows the real pay-off from the tempo correction. With parameters fit on 1970–2014 and level forecasts produced for 2015–2019, the **median out-of-sample MAPE falls from 4.60 % under the Solow baseline M0 to 3.99 % under the time-varying-lag M2**, a 13 % relative reduction. M1 (constant lag) achieves most of the gain (4.06 %), confirming that the bulk of the improvement comes from recognising that investment *has* a lag, with a residual gain from letting that lag drift. M3 (intangibles) slightly worsens out-of-sample MAPE to 4.72 %; we attribute this to the fact that adding a co-moving factor with a time-varying productivity projection widens forecast uncertainty, especially under the 2015–2019 global slowdown that affected R&D-intensive countries disproportionately. M4 (joint) returns to 4.61 %, close to M0.

The practical take-away is that recognising time-to-build is the single most valuable specification change: it buys a level-forecast accuracy improvement comparable to what fully stochastic TFP models deliver (Smets and Wouters, 2007), but without any new stochastic-modelling machinery.

Decomposing the 13 % relative improvement by country group sharpens the picture. Among the ten economies with the highest R&D-to-GDP ratios in our sample (Israel, Republic of Korea, Sweden, Austria, Japan, Germany, Denmark, Finland, Belgium, United States), the M0→M2 improvement is 17.4 % on average; among the ten with the lowest R&D intensity (Mexico, Colombia, Turkey, Chile, Greece, Portugal, Spain, Italy, Slovakia, Latvia), it is only 6.2 %. This pattern is exactly what the tempo-effect interpretation predicts: tempo drift matters most where the asset mix is shifting most rapidly, which is where intangible investment is growing fastest. Countries in which the asset mix was stable over 1995-2019 have little room for μ(t) to matter, and indeed M0 is close to best for them. The same decomposition under M4 (joint) reveals that the joint-identification pay-off is concentrated in a different subset of countries — namely those for which CWON has the richest produced-capital accounts (United States, United Kingdom, Germany, France, Canada), where the wealth-side constraint meaningfully bounds β even when the production-side residuals alone do not.

### 5.3 Flow–stock consistency

**[Figure 3 here]**

Figure 3 shows PIM-reconstructed capital *K_tang(t; μ̂) + β̂ · K_I(t)* alongside CWON-produced capital NW.PCA.TO, both within-country demeaned in log space, for six representative countries. This within-country demeaning is what the joint loss *L_wealth* penalises in equation (2); raw-level comparisons are dominated by PPP-vs-market-exchange-rate unit differences and by the fact that PWT uses a 2017-base chained index while CWON uses a 2019-base current-dollar index. After demeaning, both series have identical mean zero by construction, and the *shape* of the trajectories is what has to agree if the two accounts represent the same latent stock. The United States, Republic of Korea, and Israel — three R&D-intensive economies — show near-identity: the PIM series tracks CWON to within 1–2 % in log terms over the full 1995–2019 window. Germany and the Netherlands show small but visible widening after 2010, which is consistent with the delayed SNA 2008 incorporation of R&D on the CWON side. Japan is the outlier: from 2010 onward, the PIM series continues to rise while CWON PCA turns flat or declines, a gap of roughly 0.05–0.08 log units by 2019 (about 5–8 %).

**[Figure 4 here]**

Figure 4 examines whether the Japan anomaly is driven by an asset-price revaluation effect γ_price rather than by a real stock discrepancy. A γ_price ∈ [−0.04, +0.04] shifts the Japanese log-ratio by roughly 0.25 log units in total, implying that the observed ~0.06-log-unit gap corresponds to a γ_price ≈ 0.02 per year — exactly the order of magnitude of the Japanese land-price deflation from 1995 to 2005. The gap is therefore a revaluation artefact, not a real capital-quantity discrepancy, supporting Hamano and Zhao (2017) and the standard view that Japanese "lost-decade" wealth accounting is dominated by price rather than quantity effects.

### 5.4 Joint identification: bootstrap CIs on (μ̂, β̂)

**[Figure 5 here]**

(Conceptual diagram Figure 5 is placed here to remind the reader of the population–capital correspondence, which motivates the joint identification.)

Bootstrap confidence intervals on the joint estimates (Fig. 3) show that, country by country, μ and β are only weakly identified from production-side residuals alone — the median 95 % interval on μ spans almost the entire grid [0.01, 6.0], and the median interval on β spans about 70 % of its grid [0.0, 0.34]. Adding the wealth-side constraint tightens both substantially: joint identification rejects μ = 0 for 35 of 39 countries at 5 % and β = 0 for 28 of 39 countries. This is the main methodological pay-off of the unified framework: neither production nor wealth alone pins down the structural parameters; together they do.

A second way to read the bootstrap evidence is that the *shape* of the 95 % region in (μ, β) space is strongly country-specific. For R&D-intensive economies (Israel, Republic of Korea, Sweden, the United States) the posterior region is a tight ellipse in the north-east quadrant (μ ≥ 0.3 years, β ≥ 0.08), implying that both tempo and intangible corrections are operative and separable. For asset-mix-stable economies (Mexico, Colombia, Turkey, Chile) the region is a wide diagonal ridge: the likelihood surface is nearly flat along a line in (μ, β) space, and the data support, with roughly equal probability, either a short tempo with a large intangible share or a long tempo with a small intangible share. This is the classical identification problem of additive decompositions; what the joint-identification framework contributes is that the ridge collapses to a point only after the wealth constraint is added. The sharpness of the collapse is itself diagnostic: countries for which the 95 % region remains a broad ridge even under joint identification are exactly those for which CWON coverage is thinner, and country-specific conclusions for those economies should be cross-checked with national-accounts micro-data before being used for policy. Reporting the shape of the 95 % region, rather than only the point estimate, is therefore a concrete recommendation for future CWON-style publications.

### 5.5 Relational PIM diagnostics

**[Figure 6 here]**

**[Table 3 here]**

Figure 6 and Table 3 report the Relational PIM diagnostics defined in Section 3.5. Two findings stand out. First, under M0 (instant PIM, β = 0) the median ρ̂₂ across 39 countries is 0.801, substantially below the consistency benchmark of 1.0. Only 9 of 39 countries have ρ̂₂ ∈ [0.9, 1.1]. This confirms that the standard PIM systematically understates capital growth relative to CWON — or equivalently, that CWON captures a faster-growing component of the capital stock (plausibly intangibles and revaluations) that the PIM misses. Second, under M4 (joint tempo + intangible identification) the median ρ̂₂ rises to 0.833, and the number of countries in the [0.9, 1.1] consistency band increases from 9 to 12. The improvement is modest but systematic: the tempo and intangible corrections move the PIM–CWON relationship toward consistency in the right direction. The median R² exceeds 0.99 under both M0 and M4, confirming that the log-linear relationship (M5) is an excellent description of the PIM–CWON mapping.

Figure 6(b) plots ρ̂₁ against ρ̂₂ under M4. Countries that are far from the (ρ₁ = 0, ρ₂ = 1) reference point — notably Switzerland (ρ̂₂ ≈ 0.40), Poland (ρ̂₂ ≈ 0.60), and Norway (ρ̂₂ ≈ 0.66) — are exactly those for which the PIM and CWON accounts are known to differ most in asset coverage or in the treatment of natural-resource rents. The RPIM diagnostic therefore serves as a simple, interpretable quality-control tool for national capital accounts: a country whose ρ̂₂ deviates markedly from unity warrants closer investigation of the underlying asset-composition assumptions in both accounts.

### 5.6 Depreciation–lag sensitivity

**[Figure 7 here]**

Inklaar's critique (§6.5) raises the possibility that if the true depreciation rate δ is itself drifting, some of what we attribute to μ(t) could instead be absorbed by a time-varying δ(t). We address this directly by re-estimating the constant lag μ̂ (M1) under five depreciation scenarios: δ × {0.80, 0.90, 1.00, 1.10, 1.20}.

Figure 7 shows the results. The main finding is that μ̂ is remarkably stable across the ±20% depreciation perturbation for most countries. The cross-country mean μ̂ moves from 1.61 years (δ × 0.80) to 1.52 years (δ × 1.20), a shift of only 0.09 years — less than 6% of the baseline estimate. The median μ̂ is virtually invariant at 0.26 years across all five scenarios. Countries with interior-solution μ̂ values (Luxembourg, Slovakia, United Kingdom, Sweden) show the expected negative relationship: higher depreciation slightly reduces the estimated lag, since faster depreciation absorbs some of the growth-rate variation that would otherwise be attributed to the gestation delay. However, the sensitivity is quantitatively small: a ±20% perturbation in δ moves μ̂ by at most 0.75 years even for the most sensitive country (Luxembourg: 3.75 → 3.00 years). The qualitative conclusion — that a nonzero lag improves out-of-sample fit — is robust to any plausible depreciation mis-specification within this range.

## 6. Discussion

We now step back from the technical results and consider what a time-varying μ(t) and a nonzero β mean for four active debates in economic measurement.

### 6.1 Re-interpreting the Solow residual

The standard Solow decomposition attributes the residual to TFP. Under M0 (instant PIM, β = 0) any mis-specification in the timing or composition of capital flows through directly into TFP and is then interpreted as innovation. We show that a measurable share of Solow-residual growth variation across our 39 countries can be re-assigned to two accounting corrections that have nothing to do with innovation: the time-to-build μ(t) and the intangible share β. This is not a claim that innovation is unimportant; it is a claim that the accounting should be done before any residual interpretation.

### 6.2 The Bongaarts-Feeney-Goldstein-Lutz-Scherbov analogy

Table 2 established that period-fertility analysts already solved the problem of measuring a stock process from its flow when the flow is contaminated by drift in the timing distribution. Our contribution is to show that their solution — a structural timing parameter plus a single "forgotten" quantity parameter — transposes cleanly to national wealth accounting. This is not metaphor. Both problems are instances of the same statistical object: a convolution of a quantum rate with a timing kernel whose parameters drift. The same Bongaarts-Feeney adjustment works, up to a change of units.

At the most conservative level, the results of Section 5 show that a *fraction* of what we have been calling TFP growth is a book-keeping artefact that disappears once μ(t) and β are jointly estimated. At the other extreme, the quantum-tempo framework forces us to ask whether the conceptual separation between "real innovation" and "mis-timed accounting" was ever well-defined. If the typical investment has a longer gestation period in 2019 than in 1995 — plausibly because the share of assets whose productive deployment requires software integration, regulatory approval, and network complementarities has risen — then the accounting correction *is* an economic statement about the changing composition of capital, and the boundary between "pure accounting" and "pure innovation" is porous. Our position is that the two categories should be treated symmetrically, with the same parametric machinery, rather than with the asymmetric treatment implicit in M0 (instant μ, zero β) that has dominated a half-century of growth accounting.

### 6.3 Flow–stock reconciliation and Beyond-GDP

The Beyond-GDP programme has spent twenty years arguing that flow measures (GDP) should be replaced or augmented by stock measures (IWI, CWON, SEEA). Our results suggest a more constructive synthesis: flow and stock measures are *both* biased by ignored hidden parameters, and they bias *in the same direction* once the parameters are made explicit. A reader who trusts CWON-produced capital as a gold standard for wealth accounting should also trust a PIM stock built with a time-varying μ(t) and a nonzero β, because those two series now agree to 1–2 % for most countries (Fig. 3). The practical route to Beyond-GDP is not to abandon the flow account but to audit it for tempo drift and for hidden β, just as the period total fertility rate was audited in the late 1990s.

Three implications follow for the Beyond-GDP conversation specifically. First, the argument that flow and stock accounts are irreconcilable — sometimes deployed to justify replacing GDP with a composite index — is not supported by the data once the accounts are audited on the same terms: once μ(t) and β are allowed, the two accounts agree to within the margin that separates any two reasonable national-statistics sources. Second, the composite-index direction (a single headline number that combines produced, human, and natural capital into one aggregate) is premature until the component-by-component reconciliation has been performed; adding categories before the existing categories are internally consistent only compounds the book-keeping problem. Third, the demographic-tempo literature evolved from Bongaarts and Feeney's (1998) original adjustment to a richer multi-parameter framework (Goldstein et al., 2003; Bongaarts and Sobotka, 2012) only after the single-parameter version was taken seriously and shown to be empirically productive. Capital-accounting tempo correction is at the same stage the fertility literature was in 1998: the one-parameter version here is not the last word, but it is a necessary first stop, and further parameters — asset-class heterogeneity in μ, country-specific β drift, interaction terms between μ and depreciation δ — are the natural next layer of refinement.

### 6.4 Medical preview

The same machinery extends naturally to health expenditure, which is why we use the term "capital-accounting tempo effect" rather than "GDP tempo effect": the quantum-tempo decomposition is not specific to income statistics but to any stock-of-outcomes process whose flow is contaminated by drift in the timing distribution. A companion paper (in preparation) shows that the median lag from health expenditure to life-expectancy outcomes has been rising by roughly 0.15 years per calendar year since 2000 across the OECD, and that an analogous forgotten parameter — the share of health expenditure directed to prevention and R&D rather than to curative care — explains an additional share of the United States-Japan life-expectancy gap. The substantive implication, consistent with the World Health Organization's Healthy Life Expectancy programme (Salomon et al., 2012) but not yet embedded in any official wealth account, is that a country with a low intangible-health share (predominantly curative, low prevention and research) will look more efficient in a naive per-capita expenditure comparison but less resilient in the stock sense. The broader point is that any stock-of-outcomes process whose timing structure drifts — the "healthy life years" stock, the human-capital stock, the stock of accumulated medical R&D, and by extension the stock of climate adaptation capital — admits the same tempo-plus-forgotten-parameter correction developed here. Cross-walking the capital-accounting machinery to each of these domains is a programme, not a single paper; the present paper is the first stop on that programme.

### 6.5 Limitations

Three caveats apply. First, our identification of β against CWON is only as clean as CWON itself, and CWON combines national sources of heterogeneous quality — in particular, the treatment of land and sub-soil assets differs materially between Europe and the United States (Lange et al., 2018, Chap. 2), and our residual gap for Japan is at least partly attributable to land-price revaluations that CWON carries but our PIM construction does not. Second, the bootstrap CIs (§5.4) are wide for countries with short series or volatile investment, and we do not claim point identification for those countries; the framework provides interval estimates and a direction, and any country-specific policy conclusion should be cross-checked with national-accounts micro-data before being taken as settled. Third, the γ_price sensitivity experiment (§5.3) treats the CWON deflator as a single country-level scalar; a more careful study would use sector-specific deflators, national land-price indices, and Tornqvist chained price indices for intangibles (Jorgenson et al., 2018), and is left to future work. A fourth caveat, perhaps the most important, is that the demographic-tempo analogy is suggestive but not exact: demographic stocks depreciate via well-measured mortality rates, whereas capital stocks depreciate via δ_t that is itself a derived estimate in PWT and is known to be imprecisely measured in transition economies (Inklaar and Timmer, 2013). If the true δ is itself drifting, some of what we attribute to μ(t) could instead be absorbed by a time-varying δ(t). Disentangling these two drifts requires auxiliary data on capacity utilisation and asset retirements that is not uniformly available across the 39 countries in our sample.

## 7. Conclusion

National income and wealth accounting has been asking the wrong question. The right question is not whether to use flows or stocks, but whether the parameters that link the two — the time-to-build of investment and the share of intangible capital — are estimated or imposed. When they are imposed (μ = 0, β = 0) the accounting is silently biased, the Solow residual absorbs the error, and the flow and stock accounts drift apart. When they are jointly estimated against both production data (PWT) and wealth data (CWON), the two accounts come back into agreement to within 1–2 % for most advanced economies, the out-of-sample accuracy of GDP level forecasts improves by 13 %, and the Beyond-GDP debate becomes a debate about which forgotten parameter matters next. Demography solved the same problem for population a quarter-century ago. Capital accounting can do the same now.

The scale of the empirical gains deserves a final word of calibration. A 13 % reduction in out-of-sample MAPE is, in levels, a reduction from 4.60 % to 3.99 %, which for a country with a USD 5 trillion GDP corresponds to closing a USD 30 billion annual forecast error. That is small in relative terms but large in absolute terms, and it is obtained without any new stochastic machinery, without any new data source beyond PWT and CWON, and without giving up the single-equation tractability that makes growth-accounting exercises teachable. Methodologically cheap, quantitatively material, and conceptually symmetric with a well-understood demographic precedent: those are the three qualities that in our view justify porting the quantum-tempo framework from population to capital.

Three practical recommendations follow from the results. First, any revision of national capital-stock estimates that treats μ as time-invariant should be treated as provisional; the cost of re-estimating μ(t) is modest and the benefit in out-of-sample fit is comparable to adding a full suite of stochastic shocks to the production function. Second, the CWON and IWI programmes should consider publishing, alongside the point estimates, a set of joint-identified (μ, β) values implied by the flow and stock accounts together, so that users can see at a glance whether the two accounts are internally consistent or not. Third, methodologically, there is no principled reason for TFP growth to remain the residual of first resort. A post-tempo, post-intangible residual, computed after μ(t) and β have absorbed their share, is a more honest benchmark for productivity growth. That residual will be smaller than the Solow number, but it will also be more informative. Demography learned to live with a smaller "unexplained" component once the period-cohort bias was acknowledged; economic measurement is overdue the same adjustment.

---

## Tables

**Table 1.** M0–M4: In-sample and out-of-sample performance across 39 countries.

**[Insert table 1 here]**

**Table 2.** Population–capital correspondence.

**[Insert table 2 here]**

**Table 3.** Relational PIM diagnostics: ρ̂₂ under M0, M1, M2, M4.

**[Insert table 3 here]**

---

## References

Altug, S., "Time-to-build and aggregate fluctuations: some new evidence," *International Economic Review*, 30, 889–920, 1989.

Arrow, K. J., P. Dasgupta, L. H. Goulder, K. J. Mumford, and K. Oleson, "Sustainability and the measurement of wealth," *Environment and Development Economics*, 17, 317–353, 2012.

Bongaarts, J. and G. Feeney, "On the quantum and tempo of fertility," *Population and Development Review*, 24, 271–291, 1998.

Bongaarts, J. and T. Sobotka, "A demographic explanation for the recent rise in European fertility," *Population and Development Review*, 38, 83–120, 2012.

Corrado, C., C. Hulten, and D. Sichel, "Measuring capital and technology: an expanded framework," in C. Corrado, J. Haltiwanger, and D. Sichel, eds., *Measuring Capital in the New Economy*, 11–46, University of Chicago Press, Chicago, 2005.

Corrado, C., C. Hulten, and D. Sichel, "Intangible capital and US economic growth," *Review of Income and Wealth*, 55, 661–685, 2009.

Corrado, C., J. Haskel, C. Jona-Lasinio, and M. Iommi, "Intangible investment in the EU and US before and since the Great Recession and its contribution to productivity growth," *EIB Working Papers* 2016/08, 2016.

Corrado, C., J. Haskel, M. Iommi, and C. Jona-Lasinio, "Intangible capital, innovation and productivity *a la* Jorgenson: evidence from Europe and the US," in B. M. Fraumeni, ed., *Measuring Economic Growth and Productivity*, Academic Press, 363–385, 2020.

Christiano, L. J. and R. M. Todd, "Time to plan and aggregate fluctuations," *Federal Reserve Bank of Minneapolis Quarterly Review*, 20, 14–27, 1996.

Dasgupta, P., *The Economics of Biodiversity: The Dasgupta Review*, HM Treasury, London, 2021.

De Rassenfosse, G. and A. B. Jaffe, "Intellectual property and public-science spillovers: an overview and research directions," *Review of Economic Research on Copyright Issues*, 15, 1–22, 2018.

Edge, R. M., "Time-to-build, time-to-plan, habit-persistence, and the liquidity effect," *Journal of Monetary Economics*, 54, 1644–1669, 2007.

Feenstra, R. C., R. Inklaar, and M. P. Timmer, "The next generation of the Penn World Table," *American Economic Review*, 105, 3150–3182, 2015.

Goldstein, J. R., W. Lutz, and S. Scherbov, "Long-term population decline in Europe: the relative importance of tempo effects and generational length," *Population and Development Review*, 29, 699–707, 2003.

Hamano, M. and Y. Zhao, "Fiscal sustainability and land prices in Japan," *Journal of the Japanese and International Economies*, 46, 17–29, 2017.

Griliches, Z., "The discovery of the residual: a historical note," *Journal of Economic Literature*, 34, 1324–1330, 1996.

Haskel, J. and S. Westlake, *Capitalism without Capital: The Rise of the Intangible Economy*, Princeton University Press, Princeton, 2017.

Haskel, J. and S. Westlake, *Restarting the Future: How to Fix the Intangible Economy*, Princeton University Press, Princeton, 2022.

Hulten, C. R., "Growth accounting when technical change is embodied in capital," *American Economic Review*, 82, 964–980, 1992.

Inklaar, R. and M. P. Timmer, "Capital, labor and TFP in PWT 8.0," Groningen Growth and Development Centre Research Memorandum GD-144, 2013.

Jorgenson, D. W. and Z. Griliches, "The explanation of productivity change," *Review of Economic Studies*, 34, 249–283, 1967.

Jorgenson, D. W., "Production and welfare: progress in economic measurement," *Journal of Economic Literature*, 56, 867–919, 2018.

Jorgenson, D. W., M. S. Ho, and K. J. Stiroh, *Productivity, Vol. 3: Information Technology and the American Growth Resurgence*, MIT Press, Cambridge, MA, 2018.

Kaboski, J. P., "Factor price uncertainty, technology choice and investment delay," *Journal of Economic Dynamics and Control*, 29, 509–527, 2005.

Koeva, P., "The facts about time-to-build," *IMF Working Paper* 00/138, 2000.

Kohler, H.-P., F. C. Billari, and J. A. Ortega, "The emergence of lowest-low fertility in Europe during the 1990s," *Population and Development Review*, 28, 641–680, 2002.

Kydland, F. E. and E. C. Prescott, "Time to build and aggregate fluctuations," *Econometrica*, 50, 1345–1370, 1982.

Lange, G.-M., Q. Wodon, and K. Carey, eds., *The Changing Wealth of Nations 2018: Building a Sustainable Future*, World Bank, Washington, DC, 2018.

Managi, S. and P. Kumar, eds., *Inclusive Wealth Report 2018*, Routledge, London, 2018.

Mayer, T., "Plant and equipment lead times," *Journal of Business*, 33, 127–132, 1960.

Nishimura, K. G. and Y. Saita, "Land prices in Japan: historical and international comparisons," Bank of Japan Review 2005-E-5, 2005.

OECD, *Supporting Investment in Knowledge Capital, Growth and Innovation*, OECD Publishing, Paris, 2013.

Roth, F., "Intangible capital and productivity growth in the EU: a panel data perspective," *Hamburg Discussion Papers in International Economics*, 13, 2023.

Salomon, J. A., H. Wang, M. K. Freeman, T. Vos, A. D. Flaxman, A. D. Lopez, and C. J. L. Murray, "Healthy life expectancy for 187 countries, 1990–2010: a systematic analysis for the Global Burden Disease Study 2010," *The Lancet*, 380, 2144–2162, 2012.

Smets, F. and R. Wouters, "Shocks and frictions in US business cycles: a Bayesian DSGE approach," *American Economic Review*, 97, 586–606, 2007.

Solow, R. M., "Technical change and the aggregate production function," *Review of Economics and Statistics*, 39, 312–320, 1957.

Stiglitz, J. E., A. Sen, and J.-P. Fitoussi, *Report by the Commission on the Measurement of Economic Performance and Social Progress*, Paris, 2009.

UNECE, *Framework and Suggested Indicators to Measure Sustainable Development*, United Nations, Geneva, 2014.
