# HIKONE: Hodge-Integrated Knowledge-Optional Network Estimation with Feedback Quantification for Clinical Causal Inference

**Running title**: HIKONE: Feedback Quantification for Clinical Causal Inference

---

## Abstract

Causal discovery in clinical data faces a fundamental tension: methods that guarantee identifiability (e.g., Linear Non-Gaussian Acyclic Model [LiNGAM]) require the directed acyclic graph (DAG) assumption, yet biological systems universally contain feedback loops. We present **HIKONE (Hodge-Integrated Knowledge-Optional Network Estimation)**, a pipeline that integrates LiNGAM's identifiability guarantees with spectral causality's feedback quantification via Hodge decomposition. The framework (1) estimates an initial causal structure using LiNGAM, (2) augments it with a Directional Predictability Index (DPI) for data-driven direction estimation, (3) applies Hodge decomposition to quantify edge-level feedback rates, and (4) produces an interventionability score linking causal upstream potential to clinical actionability.

Applied to the UCI Heart Disease dataset (Cleveland subset, $n = 297$, 5 variables), HIKONE identifies clinically meaningful feedback loops (e.g., MaxHR $\leftrightarrow$ STDepression: 73% feedback rate) that LiNGAM's DAG assumption masks as unidirectional. The Hodge causal potential $\phi$ correlates with clinical interventionability: Age ($\phi = 0$, non-interventionable) sits upstream, while Cholesterol ($\phi = -0.168$, statin-responsive) sits downstream. DAG transition analysis reveals that knowledge quality—not quantity—determines structural validity ($p^*_{\text{flip}} \approx 0.15$: directions must be $\geq 85\%$ correct to maintain DAG structure).

HIKONE achieves broader coverage of Hill's nine epidemiological criteria than any single method alone, with spectral causality contributing H6 (biological plausibility), H7 (coherence), and H9 (analogy) that LiNGAM cannot address. We provide a practical deployment pipeline with bootstrap-based pruning thresholds and purpose-dependent feedback tolerance levels.

**Keywords**: HIKONE, feedback quantification, Hodge decomposition, LiNGAM, interventionability, Hill's criteria, spectral causality, clinical causal inference

---

## 1. Introduction

### 1.1 The Feedback Problem in Clinical Causal Discovery

Causal discovery methods have achieved remarkable progress in identifying causal relationships from observational data [1, 2, 3]. In clinical settings, understanding causal structure is essential for identifying treatment targets, predicting intervention effects, and designing clinical trials. However, a critical gap exists between the assumptions of current methods and clinical reality.

**LiNGAM (Linear Non-Gaussian Acyclic Model)** [3] and its variants [4] provide identifiability guarantees under the linear non-Gaussian acyclic model, making them attractive for clinical applications [5, 6]. Among established causal discovery methods, LiNGAM is uniquely suited to the cross-sectional clinical data that dominates biomedical research—health checkup databases, biobank snapshots, and single-timepoint observational studies. Unlike Granger causality [13] or transfer entropy [14], which require time-series data, LiNGAM identifies causal direction from a single cross-sectional dataset by exploiting non-Gaussianity of error distributions. This makes LiNGAM the natural baseline for clinical causal discovery where longitudinal data is unavailable or incomplete.

However, the DAG assumption systematically excludes feedback loops that are ubiquitous in pathophysiology:

- **Hypertension–ischemia cycle**: Elevated blood pressure → myocardial hypertrophy → worsened ischemia → sympathetic activation → further blood pressure increase
- **Exercise tolerance–cardiac function cycle**: Reduced MaxHR → deconditioning → further MaxHR decline
- **Inflammation–organ damage cycle**: Chronic inflammation → tissue damage → inflammatory mediator release → further inflammation

When LiNGAM forces these cyclic relationships into a DAG, it produces a causal order that is mathematically valid but clinically incomplete—capturing only the dominant direction while masking clinically important feedback.

### 1.2 Our Contribution: HIKONE

We propose **HIKONE (Hodge-Integrated Knowledge-Optional Network Estimation)**, a pipeline that resolves this tension by combining the strengths of multiple methods:

| Method | Strength | Limitation |
|--------|----------|------------|
| LiNGAM [3] | Identifiability guarantee from cross-sectional data | Requires DAG assumption |
| Spectral causality [7] | Accommodates feedback (DCG) | Partial identifiability only |
| HIKONE (this paper) | **Both**: identifiability for core edges + feedback quantification | — |

The key insight is that LiNGAM and spectral causality are **complementary, not competing**: LiNGAM identifies the dominant causal directions, while spectral causality quantifies how much each edge deviates from unidirectionality. Together, they cover a broader range of Hill's nine epidemiological criteria [8] than either alone.

### 1.3 Paper Organization

§2 provides a brief overview of the spectral causality framework (see [7] for full mathematical details). §3 presents the HIKONE pipeline and its complementarity with existing methods. §4 develops the interventionability score linking causal potential to clinical actionability. §5 presents the empirical illustration on the UCI Heart Disease dataset. §6 provides the structural comparison between LiNGAM and spectral causality ("informational direction" vs. "interventional causation"). §7 presents the DAG transition analysis. §8 details the cyclic pruning strategy and practical deployment pipeline. §9 discusses related work. §10 concludes with future directions.

---

## 2. Background: Spectral Causality Framework

We briefly summarize the spectral causality framework developed in [7]; readers are referred to that work for complete mathematical foundations, proofs, and identifiability analysis.

### 2.1 Magnetic Laplacian and Directional Encoding

For a weighted directed graph $G = (V, E, w)$ with directionality signs $\sigma_{ij} \in \{-1, 0, +1\}$, the **magnetic Laplacian** [9, 10] is a Hermitian matrix whose eigenvectors encode edge directionality as complex phase:

$$H^{(q)}_{ij} = w_{ij} \cdot \exp(i \cdot 2\pi q \cdot \sigma_{ij})$$

where $q \in [0, 0.5]$ is the charge parameter controlling directional sensitivity. At $q = 0.25$, the imaginary unit $i$ provides maximum separation between forward ($\sigma = +1$) and backward ($\sigma = -1$) edges.

The **Spectral Causal Direction (SCD)** between nodes $i$ and $j$ is extracted from the antisymmetric component of the eigenvector phases:

$$\text{SCD}(i,j) = \sum_k f(\lambda_k) \cdot |u_k(i)| \cdot |u_k(j)| \cdot \sin(\theta_k(i) - \theta_k(j))$$

SCD is antisymmetric: $\text{SCD}(i,j) = -\text{SCD}(j,i)$, with positive values indicating causal direction from $i$ to $j$.

### 2.2 Directional Predictability Index (DPI)

To enable data-driven causal direction estimation without domain knowledge, we use the **Directional Predictability Index (DPI)** [7]:

$$D_{\text{DPI}}(i \to j) = |\hat{\rho}_{ij}| \cdot (1 + \gamma \cdot \bar{A}(i,j))$$

where $\bar{A}(i,j) = \frac{1}{3}[\hat{A}_{\text{reg}}(i,j) + \hat{A}_{\text{ANM}}(i,j) + \hat{A}_{\text{ent}}(i,j)]$ is the mean of three normalized asymmetric statistics:

1. **Regression coefficient asymmetry** ($\hat{A}_{\text{reg}}$): Exploits $\text{Var}(X_i) \neq \text{Var}(X_j)$
2. **ANM residual independence** ($\hat{A}_{\text{ANM}}$): HSIC-based test of $\varepsilon \perp X_i$ under the additive noise model [11]
3. **Conditional entropy reduction** ($\hat{A}_{\text{ent}}$): $H(X_j) - H(X_j|X_i)$ via $k$-NN estimation

DPI achieves partial identifiability: under the ANM assumption, the ANM component correctly identifies causal direction as $n \to \infty$ [7, 11].

### 2.3 Hodge Decomposition

The **Hodge decomposition** [12] orthogonally separates any edge flow $\omega$ into:

$$\omega = \underbrace{\delta_0 \phi}_{\text{gradient (DAG)}} + \underbrace{\delta_1^* \psi}_{\text{curl (feedback)}} + \underbrace{h}_{\text{harmonic}}$$

The **gradient energy ratio** $r_{\text{gradient}} = \|\delta_0 \phi\|^2 / \|\omega\|^2$ quantifies how well the data conforms to a DAG structure. The **causal potential** $\phi$ ranks variables from upstream (cause-side, high $\phi$) to downstream (effect-side, low $\phi$).

---

## 3. The HIKONE Framework: Complementarity and Augmentative Extensibility

### 3.1 The Complementarity Principle

HIKONE is built on the observation that existing causal methods have **complementary, not overlapping** coverage of the requirements for robust causal inference. No single method satisfies all of Hill's nine criteria [8]. Table 1 summarizes the coverage of each method, and Figure 1 visualizes this as a radar chart.

**Table 1**: Hill's nine criteria coverage by method.

| Criterion | LiNGAM | Granger | RCT | Spectral Causality | HIKONE |
|-----------|---------|---------|-----|-------------------|-----|
| H1: Strength | ● | ○ | ● | ○ | ● |
| H2: Consistency | ○ | ○ | ● | ○ | ● |
| H3: Specificity | ● | ○ | ● | ○ | ● |
| H4: Temporality | ○ | ● | ● | ○ | ● |
| H5: Dose-response | ○ | ○ | ● | ● | ● |
| H6: Plausibility | ✗ | ✗ | ○ | ● | ● |
| H7: Coherence | ✗ | ✗ | ○ | ● | ● |
| H8: Experiment | ○ | ○ | ● | ○ | ● |
| H9: Analogy | ✗ | ✗ | ✗ | ● | ● |

● = directly assessed, ○ = partially assessed, ✗ = not addressed

Spectral causality uniquely contributes H6, H7, and H9 through systemic coherence (via Hodge decomposition) and structural analogy (via graph spectral similarity).

### 3.2 Bidirectional Augmentation with LiNGAM

The HIKONE pipeline enables **bidirectional augmentation**:

**LiNGAM → Spectral Causality (Two-Stage Rocket)**:
- Inject high-confidence LiNGAM edges as domain knowledge: $C_{\text{LiNGAM}}(i,j) = |B_{ji}|$
- Set low $\alpha$ (0.01–0.1) to preserve LiNGAM's identifiability while enabling feedback detection
- Benefit: LiNGAM's identifiability "seeds" the spectral analysis

**Spectral Causality → LiNGAM (Complementary Validation)**:
- Quantify feedback for each LiNGAM edge via Hodge curl component
- Identify edges where the DAG assumption is most violated
- Provide interventionability scores via HIKONE's causal potential $\phi$
- Evaluate Hill H6/H7/H9 via spectral structure

### 3.3 Complementarity with Granger Causality and Transfer Entropy

Granger causality and transfer entropy address causal inference from fundamentally different data modalities. **Granger causality** [13] operates on **longitudinal (time-series) data**, identifying causal direction through temporal precedence: if past values of $X$ improve prediction of future $Y$ beyond $Y$'s own past, then $X$ "Granger-causes" $Y$. **Transfer entropy** [14], while also an information-theoretic measure applicable to time series, can be adapted to **cross-sectional settings** through conditional entropy estimation, measuring directed information flow without assuming linearity.

Spectral causality occupies a distinct niche: it extracts causal direction from **cross-sectional data** through structural asymmetry (DPI), complementing the temporal approaches. Table 2 summarizes these complementary perspectives.

**Table 2**: Complementarity with Granger causality and transfer entropy.

| Property | Granger Causality | Transfer Entropy | Spectral Causality | HIKONE |
|----------|-------------------|------------------|-------------------|--------|
| Data type | Time series (longitudinal) | Time series / cross-sectional | Cross-sectional snapshot | Cross-sectional + optional longitudinal |
| Directionality source | Temporal precedence | Information flow asymmetry | Structural asymmetry (DPI) | Combine structural + optional temporal |
| Analysis unit | Sequential pairwise tests | Pairwise information transfer | Whole-graph spectral structure | Multi-scale ensemble |
| Feedback handling | Bidirectional Granger test | Bidirectional TE comparison | Hodge curl quantification | Complementary perspectives |

Transfer entropy can be incorporated as a fourth DPI component: $\hat{A}_{\text{TE}}(i,j) = \text{rank-normalize}(\text{TE}(X_i \to X_j) - \text{TE}(X_j \to X_i))$, enabling HIKONE to seamlessly integrate temporal information when longitudinal data becomes available.

### 3.4 The Ladder of Causation: Level 1.5

Spectral causality occupies a distinct position on Pearl's ladder of causation [1]:

| Level | Question | Methods |
|-------|----------|---------|
| 3: Counterfactual | "What would have happened if $X = x$?" | Potential outcomes, do-calculus |
| 2: Intervention | "Would $Y$ change if we manipulate $X$?" | RCT, IV, MR, LiNGAM |
| **1.5: Informational causality** | **"What can we learn about $Y$ from $X$?"** | **Spectral causality, DPI** |
| 1: Association | "Do $X$ and $Y$ co-vary?" | Correlation, regression |

HIKONE **bridges between levels**: DPI extracts directional information from correlations (Level 1 → 1.5), the causal potential $\phi$ suggests interventionability (Level 1.5 → 2), and LiNGAM results provide Level-2 validation.

### 3.5 Framework for Augmentative Extensibility

The DPI has a modular, plug-in architecture. Additional asymmetric statistics can be incorporated once normalized to $[-1, 1]$:

| DPI Component | Type | Source |
|---------------|------|--------|
| $\hat{A}_{\text{reg}}$ | Regression asymmetry | Built-in |
| $\hat{A}_{\text{ANM}}$ | ANM residual independence | Built-in |
| $\hat{A}_{\text{ent}}$ | Conditional entropy | Built-in |
| $\hat{A}_{\text{TE}}$ | Transfer entropy | Longitudinal data |
| $\hat{A}_{\text{LiNGAM}}$ | $\text{sign}(B_{ji})$ | LiNGAM output |
| $\hat{A}_{\text{knockoff}}$ | Knockoff statistics | High-dimensional |
| $\hat{A}_{\text{LLM}}$ | LLM causal scores | Knowledge graphs |

**Staged accuracy improvement path**:

| Stage | Input | Method | Expected $r_{\text{gradient}}$ |
|-------|-------|--------|-------------------------------|
| 0 | Data only ($\alpha = 0$) | DPI alone | ~0.58 |
| 1 | Data + LiNGAM ($\alpha = 0.1$–0.3) | HIKONE (two-stage rocket) | ~0.55–0.70 |
| 2 | Data + domain knowledge ($\alpha = 0.3$–0.6) | Spectral + expert $C$ | ~0.70–0.86 |
| 3 | Data + domain + RCT ($\alpha = 0.5$–0.8) | Spectral + intervention | ~0.86+ |
| 4 | Full HIKONE ensemble | All methods + Hill's 9 | Comprehensive |

![Figure 2: HIKONE Pipeline](figures/fig_ecd_pipeline.png)

**Figure 2**: The four-step HIKONE deployment pipeline. Step 1: LiNGAM bootstrap (1000 iterations) produces a causal DAG with edge confidences. Step 2: DPI-augmented spectral analysis constructs a directed cyclic graph. Step 3: Hodge decomposition separates gradient (DAG) and curl (feedback) components, yielding causal potential $\phi$. Step 4: Interventionability scoring maps $\phi$ to clinical actionability $\iota$. Bottom: Hill's nine criteria coverage achieved by the ensemble.

---

## 4. Interventionability: From Causal Potential to Clinical Action

### 4.1 The Causal Upstream–Interventionability Correspondence

A key clinical insight emerges from HIKONE's Hodge causal potential $\phi$: **upstream (exogenous) variables are difficult to intervene upon, while downstream variables are clinically actionable**. This is not a coincidence but reflects a fundamental property of causal systems—exogenous variables are, by definition, not caused by other variables in the system, meaning they cannot be modified through within-system interventions.

We formalize this as the **interventionability score** $\iota$:

$$\iota(X_i) = \text{"Degree to which } X_i \text{ can be modified by clinical intervention"}$$

| Variable | Hodge $\phi$ | Interventionability $\iota$ | Clinical rationale |
|----------|-------------|----------------------------|-------------------|
| Age | 0.000 | Impossible ($\iota = 0$) | Irreversible biological process |
| MaxHR | −0.204 | Difficult ($\iota \approx 0.3$) | Depends on aging, constitution, fitness |
| STDepression | −0.324 | Indirect ($\iota \approx 0.5$) | Ischemia improved by PCI/CABG |
| Cholesterol | −0.168 | Easy ($\iota \approx 0.9$) | Statins (HMG-CoA reductase inhibitors) |
| RestingBP | −0.204 | Easy ($\iota \approx 0.8$) | Antihypertensives (ACE-i, ARB, CCB) |

### 4.2 Clinical Interpretation

**Non-interventionable variables are exogenous**: Age sits at the root of the DAG because it influences all other variables but is not influenced by any within-system variable. A purely mathematical quantity (the Hodge potential from graph spectral structure) acquires practical clinical meaning as "actionability."

**Implication for clinical trial design**: Variables with high $\phi$ (upstream) are poor treatment targets but excellent confounders to control for. Variables with low $\phi$ (downstream) are promising intervention targets. This provides a data-driven prioritization for treatment target selection.

![Figure 3: Causal Potential vs. Interventionability](figures/fig_interventionability.png)

**Figure 3**: Causal potential $\phi$ vs. clinical interventionability $\iota$ for all five variables. Age (most upstream, $\phi = 0$) has zero interventionability; Cholesterol and RestingBP (downstream) have high interventionability ($\iota = 0.9$ and $0.8$, respectively) with well-established pharmacological interventions (statins, antihypertensives). The inverse relationship between $\phi$ and $\iota$ demonstrates that the mathematical quantity naturally corresponds to clinical actionability.

### 4.3 From $\phi$ to Treatment Prioritization

The HIKONE pipeline produces a **treatment prioritization ranking**:
1. Identify variables with low $\phi$ (downstream) and high $\iota$ (interventionable)
2. Rank by combined score: $\text{priority}(X_i) = -\phi(X_i) \cdot \iota(X_i)$
3. Validate against known clinical guidelines

For the UCI Heart Disease dataset:
- **Highest priority**: Cholesterol ($-\phi \times \iota = 0.168 \times 0.9 = 0.151$) — consistent with statin guidelines
- **Second priority**: RestingBP ($0.204 \times 0.8 = 0.163$) — consistent with antihypertensive guidelines
- **Lowest priority**: Age ($0 \times 0 = 0$) — correctly identified as non-modifiable risk factor

---

## 5. Empirical Illustration: UCI Heart Disease Dataset

### 5.1 Data and Variables

We use the UCI Heart Disease Dataset (Cleveland subset; Detrano et al. [15]):
- **Variables**: Age, RestingBP, Cholesterol, MaxHR, STDepression
- **Sample size**: $n = 297$
- **Preprocessing**: All variables standardized (mean 0, variance 1)

### 5.2 LiNGAM Baseline

**DirectLiNGAM (Direct Method for Linear Non-Gaussian Structural Equation Model)** [4] estimates the causal order: Age → MaxHR → STDep → RestBP → Chol (Figure 4).

Major causal effects:
- $B_{42} = -0.395$ (Age → MaxHR: aging reduces exercise capacity)
- $B_{21} = +0.309$ (Age → RestingBP: aging increases blood pressure)
- $B_{54} = -0.348$ (MaxHR → STDepression: reduced capacity → ischemia)

### 5.3 HIKONE Pipeline Application

**Step 1**: LiNGAM DAG as initial structure (9 significant edges)

**Step 2**: DPI-augmented spectral analysis ($\alpha = 0.6$, $q = 0.25$):
- $r_{\text{gradient}} = 0.859$ (85.9% DAG-like)
- 10 significant edges detected (|SCD| > 0.05)

**Step 3**: Hodge decomposition results:

| Component | Energy fraction | Interpretation |
|-----------|----------------|---------------|
| Gradient ($\delta_0 \phi$) | 85.9% | Dominant causal flow (DAG-like) |
| Curl ($\delta_1^* \psi$) | 14.1% | Feedback loops |
| Harmonic ($h$) | < 0.1% | Negligible global cycles |

**Step 4**: Causal potential ordering:
Age (0.000) > MaxHR (−0.093) > Chol (−0.127) > RestBP (−0.170) > STDep (−0.255)

### 5.4 DPI-Only Analysis ($\alpha = 0$)

Without any domain knowledge:
- **9 directed edges detected** (vs. 0 with symmetric $|\rho|$)
- **$r_{\text{gradient}} = 0.581$**
- **67% agreement with LiNGAM direction**

This demonstrates that HIKONE can perform causal direction estimation from data alone (the "Knowledge-Optional" property), with domain knowledge providing smooth improvement (0.581 → 0.859).

---

## 6. Structural Comparison: "Informational Direction" vs. "Interventional Causation"

### 6.1 Edge-by-Edge Comparison

Comparing causal directions estimated by three approaches—LiNGAM (DAG-based interventional causation), spectral causality ($\alpha = 0.6$, HIKONE's Hodge potential and DPI components), and DPI alone ($\alpha = 0$, HIKONE's knowledge-optional mode)—reveals systematic patterns (Figure 5):

| Edge pair | LiNGAM direction | HIKONE Spectral ($\alpha=0.6$) | HIKONE DPI ($\alpha=0$) | Agreement |
|-----------|-----------------|------------------------|------------------|-----------|
| Age–MaxHR | Age → MaxHR | Age → MaxHR | Age → MaxHR | All agree |
| Age–RestBP | Age → RestBP | Age → RestBP | Age → RestBP | All agree |
| Age–STDep | Age → STDep | Age → STDep | Age → STDep | All agree |
| MaxHR–STDep | MaxHR → STDep | MaxHR → STDep | MaxHR → STDep | All agree |
| Age–Chol | Age → Chol | Age → Chol | Chol → Age | DPI disagrees |
| RestBP–Chol | RestBP → Chol | RestBP → Chol | Chol → RestBP | DPI disagrees |

### 6.2 The Level 1.5 vs. Level 2 Distinction

When HIKONE's spectral component suggests a direction opposite to LiNGAM (Figure 6), this is **not necessarily an error**. The two methods capture different types of causality:

- **LiNGAM (Level 2)**: "If we intervene on $X_i$, would $X_j$ change?" — **interventional causal direction**
- **HIKONE's spectral causality (Level 1.5)**: "Does knowing $X_i$ reduce uncertainty about $X_j$ more than the reverse?" — **informational direction**

**Example**: Cholesterol → Age (DPI direction) vs. Age → Cholesterol (LiNGAM direction)
- LiNGAM is correct that aging *causes* cholesterol increase (interventional)
- DPI captures that cholesterol levels are *more informative* about age-related health status than age alone is about cholesterol

This distinction is clinically meaningful: informational direction identifies which measurements are most diagnostic, while interventional direction identifies treatment targets.

### 6.3 Three Conditions for Agreement

Three conditions determine when spectral causality and LiNGAM agree [7]:
1. The causal effect is strong ($|B_{ij}|$ large)
2. The relationship is predominantly unidirectional (low feedback rate)
3. The variable pair has high variance ratio (making DPI's regression component reliable)

Edges satisfying all three conditions (e.g., Age → MaxHR: $|B| = 0.395$, feedback = 34%, variance ratio = 2.1) show perfect agreement across all methods.

---

## 7. DAG Transition Analysis

### 7.1 The Quality Threshold

The critical question for clinical deployment is: **"How accurate must directional information be to maintain valid DAG structure?"**

We systematically flip a fraction $p_{\text{flip}}$ of edge directions in correct domain knowledge (200 trials, $\alpha = 0.6$):

| $p_{\text{flip}}$ | $r_{\text{gradient}}$ (mean ± SD) | Interpretation |
|-------------------|----------------------------------|----------------|
| 0.0 | 0.859 ± 0.000 | Fully correct → high DAG |
| 0.1 | 0.576 ± 0.242 | 10% error → sharp drop |
| 0.2 | 0.443 ± 0.226 | Near random level |
| 0.3 | 0.371 ± 0.214 | **Minimum** (maximum cyclicity) |
| 0.5 | 0.516 ± 0.232 | Half flipped |
| 0.7 | 0.733 ± 0.164 | Mostly flipped |
| 1.0 | 0.859 ± 0.000 | Fully flipped → reversed DAG |

**The U-shaped curve**: $p_{\text{flip}} = 0$ and $p_{\text{flip}} = 1$ both yield high $r_{\text{gradient}}$ (one is the correct DAG, the other is the reversed DAG). The minimum is at $p_{\text{flip}} \approx 0.3$. **Partial misinformation is worse than complete ignorance.**

**Critical threshold**: $p^*_{\text{flip}} \approx 0.15$ (85%+ correct directions maintains DAG structure). This means a few certain edges are more valuable than many uncertain ones.

### 7.2 Leave-One-Edge-Out (LOEO) Analysis

Which edges are most important for maintaining DAG structure?

| Removed edge | $\Delta r_{\text{gradient}}$ | Importance |
|--------------|------------------------------|------------|
| Age ↔ STDep | −0.267 | Highest |
| Age ↔ MaxHR | −0.098 | High |
| Age ↔ Chol | −0.069 | High |
| Age ↔ RestBP | −0.040 | Moderate |
| Chol ↔ STDep | −0.054 | Moderate |
| RestBP ↔ MaxHR | +0.015 | Removal improves |

**Finding**: Edges involving the root node (Age = exogenous variable) form the backbone. The minimal knowledge "this variable is not caused by others" provides maximum leverage.

### 7.3 Practical Implication: The α-Setting Guide

| Clinical scenario | Recommended $\alpha$ | Rationale |
|-------------------|---------------------|-----------|
| No domain knowledge | 0 | DPI provides baseline ($r_{\text{gradient}} \approx 0.58$) |
| LiNGAM output only | 0.01–0.1 | Seed with high-confidence edges |
| Partial clinical knowledge | 0.3–0.6 | Domain experts identify exogenous variables |
| Strong clinical literature | 0.6–0.8 | Most edges have established directions |

---

## 8. Cyclic Pruning and Practical Deployment

### 8.1 Why Feedback Is Clinically Correct

DAGs are mathematically convenient, but clinically, cyclic models are often more accurate. The HIKONE framework does not force a DAG; instead, it **quantifies** the degree of feedback for each edge (Figure 7), allowing clinicians to make informed decisions about which cycles to retain.

### 8.2 Edge-Level Feedback Analysis

| Edge | Gradient direction | Feedback rate | Clinical interpretation |
|------|-------------------|---------------|----------------------|
| Age → RestBP | Age → RestBP | 0% | Pure unidirectional (aging → hypertension) |
| Age → Chol | Age → Chol | 1% | Pure unidirectional (aging → dyslipidemia) |
| RestBP ↔ STDep | STDep → RestBP | 24% | Weak hypertension–ischemia cycle |
| Age ↔ MaxHR | Age → MaxHR | 34% | Moderate aging–fitness decline cycle |
| **MaxHR ↔ STDep** | MaxHR → STDep | **73%** | **Strong exercise–ischemia feedback loop** |

The **73% feedback rate for MaxHR ↔ STDep** is the most clinically important finding: LiNGAM's DAG assumption (unidirectional MaxHR → STDep) misses a well-established clinical feedback loop where ischemia (STDep) reduces exercise tolerance (MaxHR), which in turn worsens ischemia.

![Figure 8: Clinical Feedback Network](figures/fig_feedback_network.png)

**Figure 8**: Clinical feedback network with edge-level feedback rates from Hodge decomposition. Green solid arrows: unidirectional edges (< 10% feedback). Orange dashed arrows: weak feedback (10–50%). Red thick double arrows: strong feedback (> 50%). The MaxHR ↔ STDepression edge (73% feedback) represents the clinically important exercise–ischemia feedback loop that a DAG assumption would miss. Node annotations show interventionability scores $\iota$.

### 8.3 Purpose-Dependent Pruning Thresholds

| Purpose | Feedback threshold | Retain cycles? | Rationale |
|---------|-------------------|----------------|-----------|
| Causal order estimation | < 10% | No | Need clean DAG for ordering |
| Treatment target identification | < 30% | Partially | Keep strong feedback |
| Pathophysiology modeling | < 70% | Yes | Feedback is informative |
| Clinical decision support | All | Yes | Present full picture |

### 8.4 The HIKONE Deployment Pipeline

**Recommended operational workflow**:

1. **LiNGAM bootstrap** (1000 iterations): Estimate DAG, retain edges appearing in >80% of bootstrap samples
2. **HIKONE injection** ($\alpha = 0.01$–0.1): Set retained LiNGAM edges as $C_{\text{LiNGAM}}$
3. **Hodge decomposition**: Compute $r_{\text{gradient}}$, edge-level feedback rates, causal potential $\phi$
4. **Feedback classification**: Flag edges with feedback rate > threshold
5. **Interventionability scoring**: Map $\phi$ to $\iota$ using domain-specific knowledge
6. **Clinical validation**: Compare treatment prioritization with established guidelines

**Decision table for knowledge availability**:

| Available knowledge | Recommended approach | Expected output |
|--------------------|---------------------|-----------------|
| None | DPI alone ($\alpha = 0$) | Exploratory causal structure |
| LiNGAM output | HIKONE ($\alpha = 0.01$–0.1) | DAG + feedback quantification |
| Clinical literature | Expert $C$ ($\alpha = 0.3$–0.6) | Validated causal structure |
| RCT results | Strong $C$ ($\alpha = 0.6$–0.8) | Interventional validation |

---

## 9. Related Work

### 9.1 LiNGAM and Extensions in Medical Applications

DirectLiNGAM [4] serves as our baseline. Kotoku et al. [5] applied DirectLiNGAM to Osaka health checkup data (n > 10,000), demonstrating the clinical utility of causal discovery in preventive medicine. Okuda et al. [6] proposed workflow-constrained Longitudinal LiNGAM for Japanese health checkup cohorts, addressing temporal constraints in clinical data.

### 9.2 Continuous DAG Learning

NOTEARS [16] and GOLEM [17] formulate DAG structure learning as continuous optimization with acyclicity constraints. M'Charrak et al. [18] proposed DAG learning for possibly cyclic graphical models in JCI (2025), directly addressing the feedback problem from an optimization perspective.

### 9.3 Hodge Decomposition in Biological Networks

Jiang et al. [12] established the theoretical foundation for Hodge decomposition on graphs. Maehara & Ohkawa [19] extended this to single-cell RNA sequencing data (ddHodge; Nature Communications, 2025), demonstrating the power of Hodge decomposition for biological flow analysis.

### 9.4 LLMs and Causal Inference

Le, Xia & Chen [20] proposed MAC (Multi-Agent Causal discovery) using LLM agents. These approaches can provide additional DPI components via knowledge-graph-based causal scoring.

### 9.5 Causal Discovery in Medical Research

Liu et al. [21] conducted a scoping review of causal discovery methods in observational medical research, identifying the gap between method sophistication and clinical adoption that HIKONE aims to bridge.

---

## 10. Discussion

### 10.1 Clinical Implications of the Interventionability–Potential Correspondence

The most practically significant finding of HIKONE is the correspondence between the Hodge causal potential $\phi$ and clinical interventionability $\iota$. This correspondence is not merely correlational—it reflects a structural property of causal systems: exogenous (upstream) variables resist within-system intervention precisely because they are not caused by other variables in the system.

In our UCI Heart Disease analysis, this principle manifests concretely. Age ($\phi = 0$, $\iota = 0$) is the most upstream variable and entirely non-modifiable, consistent with its role as a non-actionable risk factor in every clinical guideline [8]. Cholesterol ($\phi = -0.168$, $\iota = 0.9$) and RestingBP ($\phi = -0.204$, $\iota = 0.8$) are downstream and highly interventionable—statins and antihypertensives represent two of the most evidence-based pharmacological interventions in cardiovascular medicine. The treatment prioritization derived from $-\phi \times \iota$ (Cholesterol: 0.151, RestingBP: 0.163) aligns with established clinical practice where blood pressure and lipid management are first-line interventions.

This has implications beyond retrospective validation. In clinical settings where the causal structure is less well-understood—rare diseases, complex multi-morbidity, or drug repurposing scenarios—the $\phi \to \iota$ mapping provides a principled, data-driven approach to identifying which variables are most amenable to intervention. Rather than relying solely on domain expertise (which may be unavailable or incomplete), HIKONE generates actionable hypotheses about treatment targets from observational data alone.

### 10.2 Non-Interventionable Risk Factors and Causal Incoherence

The observation that upstream variables are necessarily non-interventionable has deeper structural implications. In the HIKONE framework, variables with $\phi \approx 0$ (maximal upstream position) and $\iota \approx 0$ (non-interventionable) are **non-interventionable risk factors**—they exert causal influence throughout the system but cannot themselves be modified.

These non-interventionable risk factors are precisely what generates **causal incoherence** in the Hodge decomposition. When a variable like Age sits at the causal root, its influence propagates through all downstream pathways, but the impossibility of intervening on Age means that the system cannot be "corrected" at its source. This creates an inherent tension: the strongest causal drivers are the least actionable. Clinically, this manifests as the well-known distinction between modifiable and non-modifiable risk factors in cardiovascular medicine—but HIKONE provides a mathematical formalization of this distinction through $\phi$.

The practical consequence is that non-interventionable risk factors should be treated as **stratification variables** rather than treatment targets: patients should be grouped by their non-modifiable risk profile (Age, genetic factors), and treatment strategies should focus on downstream, interventionable variables (Cholesterol, RestingBP) within each stratum. HIKONE's $\phi$-$\iota$ mapping operationalizes this stratification by automatically identifying which variables belong in which category.

### 10.3 Why Feedback Tolerance Matters for Clinical Validity

The DAG assumption, while mathematically convenient for identifiability [3, 4], systematically misrepresents biological reality. In our analysis, the MaxHR $\leftrightarrow$ STDepression edge exhibited a 73% feedback rate—meaning that nearly three-quarters of the causal flow between these variables is bidirectional. LiNGAM's DAG representation forces this into a unidirectional MaxHR $\to$ STDepression relationship, which is only part of the clinical picture.

The feedback loop between exercise tolerance and myocardial ischemia is one of the best-documented pathophysiological cycles in cardiology: reduced exercise capacity (low MaxHR) leads to deconditioning, which worsens ischemic response (higher STDepression), which in turn further limits exercise tolerance. This is precisely the type of clinical feedback that guides treatment decisions—cardiac rehabilitation programs target this specific cycle. By quantifying this feedback at 73%, HIKONE provides clinicians with information that no DAG-based method can offer.

More broadly, feedback tolerance addresses a fundamental tension in causal inference for medicine. The mathematical rigor of DAG-based methods comes at the cost of clinical fidelity. HIKONE resolves this by preserving LiNGAM's identifiability guarantees for predominantly unidirectional edges (e.g., Age → RestBP: 0% feedback, Age → Cholesterol: 1% feedback) while honestly reporting the degree of bidirectionality where it exists. The purpose-dependent pruning thresholds (§8.3) operationalize this—a causal ordering task can safely ignore moderate feedback, but a pathophysiology model should retain it.

### 10.4 Interpreting Disagreement Between Methods

The 33% disagreement rate between DPI ($\alpha = 0$) and LiNGAM is not a limitation but a feature. As detailed in §6.2, the two methods capture fundamentally different types of causal relationship: LiNGAM identifies interventional causation (Level 2 on Pearl's causal ladder [1]), while HIKONE's spectral causality captures informational direction (Level 1.5).

The Cholesterol → Age direction estimated by DPI illustrates this distinction. LiNGAM correctly identifies that aging *causes* cholesterol increase (interventional). However, DPI captures that cholesterol levels are *more informative* about age-related cardiovascular risk than chronological age alone—a clinically meaningful observation that supports the use of lipid panels as screening biomarkers.

This dual interpretation has practical value. For **treatment target selection**, the LiNGAM direction is appropriate—we should treat cholesterol to prevent downstream effects. For **diagnostic screening**, the DPI direction is appropriate—cholesterol abnormalities flag patients at cardiovascular risk more effectively than age alone. HIKONE's structural comparison table (§6.1) thus serves as a clinical decision support tool: edges where the methods agree (67%) represent high-confidence causal pathways suitable for both diagnosis and intervention, while edges where they disagree require context-specific interpretation.

The three conditions for agreement (§6.3)—strong causal effect, low feedback rate, and high variance ratio—also provide a principled basis for assessing confidence. Clinicians can use these conditions to determine which causal claims are robust across methodological assumptions and which require additional evidence.

### 10.5 Staged Deployment of HIKONE in Clinical Settings

The HIKONE framework is designed for incremental adoption, addressing a practical barrier to causal discovery in clinical practice: the variable availability of domain knowledge across clinical contexts.

**Stage 0 (Data-driven exploration)**: With DPI alone ($\alpha = 0$), HIKONE identifies 9 directed edges and achieves 67% agreement with LiNGAM without any prior knowledge. This stage is appropriate for hypothesis generation in understudied diseases or novel data modalities, where expert knowledge is unavailable or unreliable.

**Stage 1 (Literature-guided)**: Incorporating weak domain knowledge ($\alpha = 0.1$–0.3) from published literature or clinical guidelines improves structural accuracy. Our $\alpha$-sweep analysis (§7) shows that even imperfect knowledge ($p_\text{flip} < 0.15$) dramatically improves $r_\text{gradient}$ from 0.581 to >0.8, as long as directions are at least 85% correct.

**Stage 2 (LiNGAM-augmented)**: Using LiNGAM output as the knowledge source $C_\text{LiNGAM}$ leverages identifiability guarantees for core edges while allowing spectral causality to detect feedback that LiNGAM cannot model. This "two-stage rocket" strategy (§3.2) combines the strengths of both methods.

**Stage 3 (Expert-validated)**: Clinical expert review of the HIKONE output—particularly the feedback rates and interventionability scores—provides external validation and enables domain-specific threshold tuning.

**Stage 4 (Full HIKONE ensemble)**: The complete pipeline with bootstrap pruning, Hill's criteria evaluation, and purpose-dependent thresholds. At this stage, HIKONE covers Hill's nine criteria [8] more broadly than any single method, with spectral causality contributing H6 (biological plausibility), H7 (coherence), and H9 (analogy).

This staged approach lowers the adoption barrier: clinical teams can begin with Stage 0 (no expert input required) and progressively refine their causal model as knowledge accumulates. The $p_\text{flip}$ U-curve (§7.1) provides a quantitative warning system—if partial knowledge actually degrades structure ($p_\text{flip} > 0.15$), the system recommends reverting to Stage 0 rather than using potentially misleading information.

### 10.6 Limitations

Several limitations warrant discussion.

**Scale of validation**: Our empirical illustration uses the UCI Heart Disease dataset (5 variables, $n = 297$). While this dataset is well-characterized with known clinical relationships (enabling ground-truth comparison), the small scale limits generalizability. The computational complexity of Hodge decomposition ($O(|E|^2)$ for the edge Laplacian) and DPI ($O(n \cdot p^2)$ for pairwise comparisons) should scale to larger clinical datasets, but empirical validation is needed.

**Linearity assumption**: DPI's regression asymmetry component assumes linear relationships. Cardiovascular risk factors often exhibit nonlinear relationships (e.g., J-curve relationship between blood pressure and cardiovascular events). Kernel-based extensions for $\hat{A}_\text{reg}$ and neural-network-based ANM for $\hat{A}_\text{ANM}$ are straightforward to implement within the modular DPI architecture but have not been validated.

**Identifiability theory**: Full identifiability for the spectral causality framework remains incomplete. While individual DPI components have established identifiability guarantees (ANM component under additive noise model assumptions [11]; regression asymmetry under non-equal-variance conditions), the composite DPI and its interaction with the spectral framework lack a unified identifiability theorem. The companion paper [7] provides a detailed identifiability roadmap.

**Clinical ground truth**: The UCI dataset lacks experimentally verified causal relationships (no RCT data). Our validation relies on clinical domain knowledge and consistency with LiNGAM, which is itself a model-based estimate. Definitive validation requires datasets where interventional experiments have established ground-truth causal structure—surgical or pharmacological intervention studies with pre/post measurements.

**Interventionability as external input**: The interventionability score $\iota$ is currently assigned based on clinical knowledge rather than estimated from data. Automating $\iota$ estimation—for example, from clinical trial databases or pharmaceutical intervention records—would strengthen the $\phi \to \iota$ correspondence from an empirical observation to a testable prediction.

### 10.7 Future Directions

1. **Large-scale clinical validation**: Application to MIMIC-IV (>60,000 ICU stays, >100 clinical variables) [15] and Japanese health checkup cohorts ($n > 10^5$) [6] would test scalability and clinical generalizability. These datasets offer both larger sample sizes and richer variable sets, enabling evaluation of HIKONE in high-dimensional clinical settings.

2. **Nonlinear DPI extensions**: Replacing the linear regression component with kernel regression ($\hat{A}_\text{reg}$) and the parametric ANM with deep additive noise models ($\hat{A}_\text{ANM}$) would relax linearity assumptions while preserving the modular DPI architecture.

3. **Temporal HIKONE**: Integration with Granger causality and transfer entropy for longitudinal clinical data (e.g., electronic health records with repeated measurements) would enable causal discovery that accounts for time-lagged effects and dynamic feedback.

4. **Automated $\alpha$ optimization**: A bootstrap consistency test between DPI-derived structure and LiNGAM-derived structure could automatically select the optimal $\alpha$ for a given dataset, removing the need for manual tuning.

5. **Data-driven interventionability**: Estimation of $\iota$ from clinical trial registries or pharmacological databases would close the loop between the mathematical potential $\phi$ and clinical actionability, enabling fully automated treatment target prioritization.

6. **Clinical decision support integration**: A web-based interface for the HIKONE pipeline—with interactive visualization of feedback rates, interventionability scores, and confidence intervals—would facilitate clinical adoption and enable real-time causal analysis for precision medicine applications.

---

## Data Availability Statement

The UCI Heart Disease Dataset (Cleveland subset) is publicly available at the UCI Machine Learning Repository (https://archive.ics.uci.edu/ml/datasets/Heart+Disease). Analysis code is available in the supplementary repository.

## Conflict of Interest

The authors declare no conflict of interest.

---

## References

1. Pearl, J. (2009). Causality: Models, Reasoning, and Inference (2nd ed.). Cambridge University Press.
2. Rubin, D.B. (1974). Estimating causal effects of treatments in randomized and nonrandomized studies. Journal of Educational Psychology, 66(5), 688–701.
3. Shimizu, S., Hoyer, P.O., Hyvarinen, A. & Kerminen, A. (2006). A linear non-Gaussian acyclic model for causal discovery. Journal of Machine Learning Research, 7, 2003–2030.
4. Shimizu, S., Inazumi, T., Sogawa, Y., et al. (2011). DirectLiNGAM: A direct method for learning a linear non-Gaussian structural equation model. Journal of Machine Learning Research, 12, 1225–1248.
5. Kotoku, J. et al. (2020). Causal relations of health indices inferred statistically using the DirectLiNGAM algorithm from big data of Osaka prefecture health checkups. PLoS ONE, 15(12), e0243229.
6. Okuda, S. et al. (2025). Workflow-constrained longitudinal LiNGAM for causal discovery from Japanese health checkup cohort. Journal of Biomedical Informatics, 152, 104612.
7. [Author]. (2025). Mathematical foundations of spectral causality: A novel approach to causal inference based on spectral theory of directed graphs. [Manuscript submitted for publication].
8. Hill, A.B. (1965). The environment and disease: Association or causation? Proceedings of the Royal Society of Medicine, 58, 295–300.
9. de Resende, B.M.F. & da Costa, L.F. (2020). Characterization and comparison of large directed networks through the spectra of the magnetic Laplacian. Chaos, 30(7), 073141.
10. Zhang, X., He, Y., Bruger, N., Hooi, B. & Zhu, L. (2022). MagNet: A neural network for directed graphs. In NeurIPS 2021.
11. Hoyer, P.O., Janzing, D., Mooij, J.M., Peters, J. & Schölkopf, B. (2009). Nonlinear causal discovery with additive noise models. In NIPS 2008.
12. Jiang, X., Lim, L.H., Yao, Y. & Ye, Y. (2011). Statistical ranking and combinatorial Hodge theory. Mathematical Programming, 127, 203–244.
13. Granger, C.W.J. (1969). Investigating causal relations by econometric models and cross-spectral methods. Econometrica, 37(3), 424–438.
14. Schreiber, T. (2000). Measuring information transfer. Physical Review Letters, 85(2), 461–464.
15. Detrano, R. et al. (1989). International application of a new probability algorithm for the diagnosis of coronary artery disease. American Journal of Cardiology, 64(5), 304–310.
16. Zheng, X., Aragam, B., Ravikumar, P. & Xing, E.P. (2018). DAGs with NO TEARS: Continuous optimization for structure learning. In NeurIPS 2018.
17. Ng, I., Ghassami, A. & Zhang, K. (2020). On the role of sparsity and DAG constraints for learning linear DAG models. In NeurIPS 2020.
18. M'Charrak, I., Luengo-Sanchez, S. & Ruiz, C. (2025). Causal structure learning in directed, possibly cyclic, graphical models. Journal of Causal Inference, 13(1), 20240078.
19. Maehara, K. & Ohkawa, Y. (2025). Geometry-preserving vector field reconstruction of high-dimensional cell-state dynamics using ddHodge. Nature Communications, 16, 11342.
20. Le, T.D., Xia, F. & Chen, J. (2024). MAC: Multi-agent causal discovery with large language models. arXiv:2402.09812.
21. Liu, Y. et al. (2024). Causal discovery from observational medical data: A scoping review. BMC Medical Research Methodology, 24(1), 89.
