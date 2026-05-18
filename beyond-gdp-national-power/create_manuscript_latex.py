"""
Generate Cliometrica manuscript in LaTeX (Springer Nature sn-jnl template).

Outputs (in manuscript/):
  manuscript.tex       — Main manuscript (LaTeX source)
  references.bib       — BibTeX bibliography
  manuscript.pdf       — Compiled PDF
  table_s1.tex         — Supplementary Table S1 (LaTeX)
  table_s1.pdf         — Supplementary Table S1 (PDF)
  cover_letter.tex/pdf — Cover letter
  figures/Fig1–Fig4.png — Separate figure files (unchanged)
  figures_pptx.pptx    — Editable PPTX (unchanged)
"""

import os
import sys
import re
import shutil
import subprocess
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import scipy.stats as stats

# reuse everything from the docx script
sys.path.insert(0, os.path.dirname(__file__))
from data import load_data
from sensitivity_technical_network_exclusion import (
    STRONG_CANDIDATES, MODERATE_CANDIDATES, RATIONALE,
    apply_technical_network_exclusion, apply_disrupted_assignment,
    compute_confusion_stats, compute_closure_analysis,
    compute_logistic_with_closure, compute_mediation_paths,
)
from create_manuscript import (
    run_analysis, create_figures, create_pptx, create_table_s1,
    MODERN_COUNTRY, ENGLISH_NAME, TURNING_POINT,
)

OUT = os.path.join(os.path.dirname(__file__), "manuscript")
FIG = os.path.join(OUT, "figures")
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__),
                            "springer-template", "sn-article-template")
os.makedirs(FIG, exist_ok=True)


def _esc(text):
    """Escape special LaTeX characters in plain text."""
    text = str(text)
    # Replace non-ASCII text with ASCII/LaTeX equivalents
    text = text.replace('\u73fe\u5728', 'present')  # 現在
    text = text.replace('\u301c', '--')  # 〜
    text = text.replace('\u2013', '--')  # –
    mapping = {
        '&': r'\&', '%': r'\%', '$': r'\$', '#': r'\#',
        '_': r'\_', '{': r'\{', '}': r'\}', '~': r'\textasciitilde{}',
        '^': r'\textasciicircum{}',
    }
    for char, replacement in mapping.items():
        text = text.replace(char, replacement)
    return text


def _pct(value):
    """Format a 0-1 fraction as a LaTeX-safe percentage string like '83.3\\%'."""
    return f"{value*100:.1f}\\%"


def _compile_latex(tex_path, out_dir):
    """Compile a .tex file to PDF using pdflatex + bibtex."""
    basename = os.path.splitext(os.path.basename(tex_path))[0]
    cmd_latex = [
        "pdflatex", "-interaction=nonstopmode",
        "-output-directory", out_dir, tex_path
    ]

    # Clean aux to avoid duplicate bibstyle entries
    aux_path = os.path.join(out_dir, basename + ".aux")
    for ext in (".aux", ".bbl", ".blg"):
        p = os.path.join(out_dir, basename + ext)
        if os.path.exists(p):
            os.remove(p)

    # first pass (generates .aux with citation keys)
    subprocess.run(cmd_latex, capture_output=True, timeout=120)

    # bibtex (resolves citations from .bib file)
    bib_path = os.path.join(out_dir, "references.bib")
    if os.path.exists(bib_path) and os.path.exists(aux_path):
        cmd_bibtex = ["bibtex", basename]
        subprocess.run(cmd_bibtex, capture_output=True, timeout=60,
                       cwd=out_dir)

    # second + third pass (resolves references and cross-refs)
    subprocess.run(cmd_latex, capture_output=True, timeout=120)
    subprocess.run(cmd_latex, capture_output=True, timeout=120)

    pdf = os.path.join(out_dir, basename + ".pdf")
    return pdf if os.path.exists(pdf) else None


# ═══════════════════════════════════════════════════════════════
# BibTeX bibliography
# ═══════════════════════════════════════════════════════════════

def create_bibtex():
    bib = r"""@article{AcemogluJohnsonRobinson2002,
  author    = {Acemoglu, Daron and Johnson, Simon and Robinson, James A.},
  title     = {Reversal of Fortune: Geography and Institutions in the Making of the Modern World Income Distribution},
  journal   = {Quarterly Journal of Economics},
  volume    = {117},
  pages     = {1231--1294},
  year      = {2002},
  doi       = {10.1162/003355302320935025}
}

@incollection{AcemogluJohnsonRobinson2005,
  author    = {Acemoglu, Daron and Johnson, Simon and Robinson, James A.},
  title     = {Institutions as a Fundamental Cause of Long-Run Growth},
  booktitle = {Handbook of Economic Growth},
  editor    = {Aghion, Philippe and Durlauf, Steven N.},
  volume    = {1A},
  publisher = {Elsevier},
  address   = {Amsterdam},
  pages     = {385--472},
  year      = {2005}
}

@book{AcemogluRobinson2012,
  author    = {Acemoglu, Daron and Robinson, James A.},
  title     = {Why Nations Fail: The Origins of Power, Prosperity, and Poverty},
  publisher = {Crown},
  address   = {New York},
  year      = {2012}
}

@article{AcemogluRestrepo2020,
  author    = {Acemoglu, Daron and Restrepo, Pascual},
  title     = {Robots and Jobs: Evidence from {US} Labor Markets},
  journal   = {Journal of Political Economy},
  volume    = {128},
  pages     = {2188--2244},
  year      = {2020},
  doi       = {10.1086/705716}
}

@book{Arrighi1994,
  author    = {Arrighi, Giovanni},
  title     = {The Long Twentieth Century: Money, Power, and the Origins of Our Times},
  publisher = {Verso},
  address   = {London},
  year      = {1994}
}

@article{BroadberryGuan2026,
  author    = {Broadberry, Stephen N. and Guan, Hanhui},
  title     = {Regional Variation of {GDP} per Head within {China}, 1080--1850},
  journal   = {Explorations in Economic History},
  volume    = {95},
  pages     = {101567},
  year      = {2026},
  doi       = {10.1016/j.eeh.2025.101567}
}

@article{CominEasterlyGong2010,
  author    = {Comin, Diego and Easterly, William and Gong, Erin},
  title     = {Was the Wealth of Nations Determined in 1000 {BC}?},
  journal   = {American Economic Journal: Macroeconomics},
  volume    = {2},
  number    = {3},
  pages     = {65--97},
  year      = {2010},
  doi       = {10.1257/mac.2.3.65}
}

@article{CominMestieri2018,
  author    = {Comin, Diego and Mestieri, Mart\'{i}},
  title     = {If Technology Has Arrived Everywhere, Why Has Income Diverged?},
  journal   = {American Economic Journal: Macroeconomics},
  volume    = {10},
  number    = {3},
  pages     = {137--178},
  year      = {2018},
  doi       = {10.1257/mac.20150175}
}

@article{DeVries2010,
  author    = {De Vries, Jan},
  title     = {The Limits of Globalization in the Early Modern World},
  journal   = {Economic History Review},
  volume    = {63},
  pages     = {710--733},
  year      = {2010},
  doi       = {10.1111/j.1468-0289.2009.00497.x}
}

@book{Diamond1997,
  author    = {Diamond, Jared},
  title     = {Guns, Germs, and Steel: The Fates of Human Societies},
  publisher = {W.W. Norton},
  address   = {New York},
  year      = {1997}
}

@article{DiamondBellwood2003,
  author    = {Diamond, Jared and Bellwood, Peter},
  title     = {Farmers and Their Languages: The First Expansions},
  journal   = {Science},
  volume    = {300},
  pages     = {597--603},
  year      = {2003},
  doi       = {10.1126/science.1078208}
}

@book{FindlayORourke2007,
  author    = {Findlay, Ronald and O'Rourke, Kevin H.},
  title     = {Power and Plenty: Trade, War, and the World Economy in the Second Millennium},
  publisher = {Princeton University Press},
  address   = {Princeton},
  year      = {2007}
}

@book{Kennedy1987,
  author    = {Kennedy, Paul},
  title     = {The Rise and Fall of the Great Powers: Economic Change and Military Conflict from 1500 to 2000},
  publisher = {Random House},
  address   = {New York},
  year      = {1987}
}

@book{Maddison2007,
  author    = {Maddison, Angus},
  title     = {Contours of the World Economy 1--2030 {AD}: Essays in Macro-Economic History},
  publisher = {Oxford University Press},
  address   = {Oxford},
  year      = {2007}
}

@book{Mokyr2002,
  author    = {Mokyr, Joel},
  title     = {The Gifts of Athena: Historical Origins of the Knowledge Economy},
  publisher = {Princeton University Press},
  address   = {Princeton},
  year      = {2002}
}

@book{NorthWallisWeingast2009,
  author    = {North, Douglass C. and Wallis, John Joseph and Weingast, Barry R.},
  title     = {Violence and Social Orders: A Conceptual Framework for Interpreting Recorded Human History},
  publisher = {Cambridge University Press},
  address   = {Cambridge},
  year      = {2009}
}

@article{Nunn2008,
  author    = {Nunn, Nathan},
  title     = {The Long-Term Effects of {Africa}'s Slave Trades},
  journal   = {Quarterly Journal of Economics},
  volume    = {123},
  pages     = {139--176},
  year      = {2008},
  doi       = {10.1162/qjec.2008.123.1.139}
}

@book{Pomeranz2000,
  author    = {Pomeranz, Kenneth},
  title     = {The Great Divergence: {China}, {Europe}, and the Making of the Modern World Economy},
  publisher = {Princeton University Press},
  address   = {Princeton},
  year      = {2000}
}

@book{Tainter1988,
  author    = {Tainter, Joseph A.},
  title     = {The Collapse of Complex Societies},
  publisher = {Cambridge University Press},
  address   = {Cambridge},
  year      = {1988}
}

@article{Turchin2009,
  author    = {Turchin, Peter},
  title     = {A Theory for Formation of Large Empires},
  journal   = {Journal of Global History},
  volume    = {4},
  pages     = {191--217},
  year      = {2009},
  doi       = {10.1017/S1740022809003192}
}

@book{TurchinNefedov2009,
  author    = {Turchin, Peter and Nefedov, Sergey A.},
  title     = {Secular Cycles},
  publisher = {Princeton University Press},
  address   = {Princeton},
  year      = {2009}
}
"""
    bib_path = os.path.join(OUT, "references.bib")
    with open(bib_path, "w", encoding="utf-8") as f:
        f.write(bib)
    print("  BibTeX saved to", bib_path)
    return bib_path


# ═══════════════════════════════════════════════════════════════
# Main manuscript (.tex)
# ═══════════════════════════════════════════════════════════════

def create_manuscript_tex(results):
    N = results["N"]
    df = results["df"]
    n_overtaken = len(df[df["outcome"] == "overtaken"])
    n_disrupted = len(df[df["outcome"] == "disrupted"])
    n_survived = len(df[df["outcome"] == "survived"])
    n_stock = len(df[df["dominant"] == "stock"])
    n_flow = len(df[df["dominant"] == "flow"])

    # Scenario data
    s_base_c = results["scenarios"]["as_conquered__baseline"]
    s_strong_c = results["scenarios"]["as_conquered__strong"]
    s_all_c = results["scenarios"]["as_conquered__all"]
    cm_base = s_base_c["cm"]
    cm_all_c = s_all_c["cm"]
    boot_c = results["bootstrap"].get("as_conquered__all", {})

    # Table 4 computation
    df_all_c = apply_disrupted_assignment(
        apply_technical_network_exclusion(df, STRONG_CANDIDATES + MODERATE_CANDIDATES),
        "as_conquered"
    )
    has_closure = df_all_c["closure_type"].isin(
        ["maritime_ban", "technical_network_exclusion", "sakoku"])
    cells_64 = {}
    for dom in ["stock", "flow"]:
        for cl_label, cl_val in [("closed", True), ("open", False)]:
            sub = df_all_c[(df_all_c["dominant"] == dom) & (has_closure == cl_val)]
            n = len(sub)
            conquered = int(sub["outcome_binary"].sum())
            survived = n - conquered
            rate = sub["outcome_binary"].mean() if n > 0 else 0
            cells_64[(dom, cl_label)] = {
                "n": n, "c": conquered, "s": survived, "rate": rate}

    sc = cells_64[("stock", "closed")]
    so = cells_64[("stock", "open")]
    fc = cells_64[("flow", "closed")]
    fo = cells_64[("flow", "open")]

    def fisher_or_p(a, b):
        table = np.array([[a["c"], a["s"]], [b["c"], b["s"]]])
        odds_r, p_val = stats.fisher_exact(table, alternative="greater")
        return odds_r, p_val

    or_sc_fo, p_sc_fo = fisher_or_p(sc, fo)
    or_sc_so, p_sc_so = fisher_or_p(sc, so)
    or_fc_fo, p_fc_fo = fisher_or_p(fc, fo)

    # Table 1 data
    candidates = STRONG_CANDIDATES + MODERATE_CANDIDATES
    table1_rows = []
    rationale_en = {
        "\u6f22\u671d\uff08\u524d\u6f22\u301c\u5f8c\u6f22\uff09": "Excluded from Mediterranean network; Silk Road provided limited overland contact only.",
        "\u30de\u30ea\u5e1d\u56fd": "Excluded from Atlantic/Mediterranean networks until Portuguese contact (15th c.). Trans-Saharan caravan only.",
        "\u30af\u30e1\u30fc\u30eb\u5e1d\u56fd\uff08\u30a2\u30f3\u30b3\u30fc\u30eb\uff09": "Inland polity; excluded from dominant exchange networks unlike neighbouring Srivijaya.",
        "\u30ad\u30a8\u30d5\u5927\u516c\u56fd": "Dnieper river trade (Varangian route) only; excluded from dominant exchange networks.",
        "\u30c6\u30a3\u30e0\u30fc\u30eb\u671d": "Landlocked; Silk Road overland only. Excluded from dominant Mediterranean and Indian Ocean networks.",
        "\u30b5\u30b5\u30f3\u671d\u30da\u30eb\u30b7\u30a2": "Limited Persian Gulf trade; excluded from established Indian Ocean and Mediterranean networks.",
        "\u30d3\u30eb\u30de\uff08\u30b3\u30f3\u30d0\u30a6\u30f3\u671d\uff09": "Coastal access existed but excluded from regular international exchange networks until British arrival.",
    }
    for entity in candidates:
        en_name = ENGLISH_NAME.get(entity, entity)
        row_data = df[df["entity"] == entity].iloc[0]
        tier = "Strong" if entity in STRONG_CANDIDATES else "Moderate"
        rat = rationale_en.get(entity, "")
        table1_rows.append((en_name, row_data["period"], tier, rat))

    # Table 2 data
    table2_rows = []
    for d_mode, d_label in [("as_conquered", "Overtaken"),
                            ("as_survived", "Survived")]:
        for c_key, c_label in [("baseline", "Baseline"),
                               ("strong", "+5 Strong"),
                               ("all", "+7 All")]:
            key = f"{d_mode}__{c_key}"
            s = results["scenarios"][key]
            sig = "*" if s["fisher_ban_p"] < 0.05 else ""
            table2_rows.append((
                d_label, c_label, s["ban_n"],
                _pct(s['ban_rate']), _pct(s['no_ban_rate']),
                f"{s['rr']:.3f}", f"{s['fisher_ban_p']:.4f}", sig
            ))

    # Table 3 data
    lr_all_c = results["scenarios"]["as_conquered__all"]["logistic"]
    table3_rows = []
    if lr_all_c.get("with_ban", {}).get("converged"):
        var_labels = {
            "dominant_binary": "Stock-dominant",
            "geo_barrier": "Geographic barrier",
            "external_threat": "External threat",
            "tech_position": "Technological position",
            "institutional_quality": "Institutional quality",
            "era_code": "Era (time)",
            "has_external_patron": "External patron",
            "has_maritime_ban": "Network closure dummy",
        }
        coefs = lr_all_c["with_ban"]["coefs"]
        for var, label in var_labels.items():
            v = coefs[var]
            sig = "*" if v["p"] < 0.05 else (r"\dag" if v["p"] < 0.10 else "")
            ci_lo_str = (f"{v['ci_lo']:.3f}" if v["ci_lo"] < 1e6
                         else r"$>10^{6}$")
            ci_hi_str = (f"{v['ci_hi']:.3f}" if v["ci_hi"] < 1e6
                         else r"$>10^{6}$")
            table3_rows.append((
                label, f"{v['OR']:.3f}",
                f"[{ci_lo_str}, {ci_hi_str}]",
                f"{v['p']:.4f}", sig
            ))

    # ── Build LaTeX source ──
    tex = []

    def W(line=""):
        tex.append(line)

    # Preamble
    W(r"% Springer Nature LaTeX Template — Cliometrica submission")
    W(r"% Generated by create_manuscript_latex.py")
    W(r"\documentclass[pdflatex,sn-basic]{sn-jnl}")
    W()
    W(r"\usepackage{graphicx}%")
    W(r"\usepackage{multirow}%")
    W(r"\usepackage{amsmath,amssymb,amsfonts}%")
    W(r"\usepackage{amsthm}%")
    W(r"\usepackage{mathrsfs}%")
    W(r"\usepackage[title]{appendix}%")
    W(r"\usepackage{xcolor}%")
    W(r"\usepackage{textcomp}%")
    W(r"\usepackage{manyfoot}%")
    W(r"\usepackage{booktabs}%")
    W(r"\usepackage{algorithm}%")
    W(r"\usepackage{algorithmicx}%")
    W(r"\usepackage{algpseudocode}%")
    W(r"\usepackage{listings}%")
    W(r"\usepackage{hyperref}")
    W(r"\usepackage[T1]{fontenc}")
    W()
    W(r"\theoremstyle{thmstyleone}%")
    W(r"\newtheorem{theorem}{Theorem}%")
    W(r"\newtheorem{proposition}[theorem]{Proposition}%")
    W(r"\theoremstyle{thmstyletwo}%")
    W(r"\newtheorem{example}{Example}%")
    W(r"\newtheorem{remark}{Remark}%")
    W(r"\theoremstyle{thmstylethree}%")
    W(r"\newtheorem{definition}{Definition}%")
    W()
    W(r"\raggedbottom")
    W()
    W(r"\begin{document}")
    W()

    # Title
    W(r"\title[Network Exclusion and State Collapse]{"
      r"Network Exclusion and State Collapse: "
      r"From Maritime Isolation to Technological Access Denial "
      r"in the Long Run of History}")
    W()

    # Author
    W(r"\author*[1]{\fnm{Onishi} \sur{Tatsuki}}\email{bougtoir@gmail.com}")
    W()
    W(r"\affil*[1]{\orgdiv{Data Science AI Innovation Research Promotion Center}, "
      r"\orgname{Shiga University}, "
      r"\orgaddress{\street{1-1-1 Bamba}, \city{Hikone}, "
      r"\postcode{522-8522}, \state{Shiga}, \country{Japan}}}")
    W()

    # Abstract
    W(r"\abstract{")
    W("Why do some states collapse while others endure? This study proposes that structural "
      "exclusion from the dominant technological network of an era---not merely deliberate "
      "trade closure---is a recurring driver of state vulnerability, because it severs the flow "
      "of technical knowledge and allows a cumulative technology gap to open between connected "
      f"and excluded polities. Using a comparative dataset of {N} historical polities spanning "
      "antiquity to the present, we distinguish between policy-driven isolation (e.g., Ming "
      r"\emph{haijin}, Tokugawa \emph{sakoku}) and what we term `technical network exclusion': "
      "involuntary disconnection from prevailing exchange networks due to geographic or "
      "technological constraints. Policy-based closure reduces but does not eliminate technology "
      r"transfer (e.g., \emph{rangaku} via Dejima under \emph{sakoku}), whereas technical "
      "exclusion severs it entirely. "
      "Reclassifying seven technically excluded polities transforms a non-significant "
      f"association between closure and conquest (Fisher's exact test $p = {s_base_c['fisher_ban_p']:.3f}$) "
      f"into a significant one ($p = {s_all_c['fisher_ban_p']:.3f}$); all seven were eventually conquered. "
      "Multivariate logistic regression confirms that external threat and institutional quality "
      "are the strongest predictors of collapse, while the stock--flow odds ratio "
      f"(OR $= {cm_base['OR']:.3f}$) remains stable across all scenarios. We argue that the "
      "underlying mechanism---technology flow disruption leading to an accumulating "
      "civilization-level gap---generalizes beyond maritime isolation. As the dominant network "
      "shifts from sea lanes to semiconductors, artificial intelligence, and advanced robotics, "
      "states structurally excluded from these technological platforms may face analogous "
      "vulnerabilities.")
    W(r"}")
    W()

    # Keywords & JEL
    W(r"\keywords{network exclusion, economic history, technology flow disruption, "
      r"state collapse, technological access, sensitivity analysis}")
    W()

    W(r"\maketitle")
    W()

    # ════════════════════════════════════════
    # 1. INTRODUCTION
    # ════════════════════════════════════════
    W(r"\section{Introduction}\label{sec:intro}")
    W()
    W(r"Every era has a dominant network---a prevailing infrastructure of exchange through "
      r"which states access trade, technology, and strategic information. In antiquity, it "
      r"was the Mediterranean sea lanes and the Silk Road caravan routes. In the age of sail, "
      r"it was transoceanic maritime commerce. In the industrial age, it was railroad-linked "
      r"factory systems and colonial supply chains. Today, it is the digital and semiconductor "
      r"ecosystem. A recurring pattern in global history is that states excluded from the "
      r"dominant network of their era---whether by choice or by circumstance---face elevated "
      r"risks of decline and conquest \citep{Kennedy1987, FindlayORourke2007}.")
    W()
    W(r"The existing literature offers two broad explanations for why states fail. One "
      r"emphasizes internal dynamics: \citet{Tainter1988} argued that complex societies "
      r"collapse when the marginal returns to increasing complexity decline, and "
      r"\citet{AcemogluRobinson2012} showed that extractive institutions---those that "
      r"concentrate power and discourage innovation---undermine long-run prosperity. The "
      r"other emphasizes external connectivity: trade openness, technology diffusion, and "
      r"the consequences of isolation \citep{FindlayORourke2007, Mokyr2002}. Our "
      r"contribution lies at the intersection. We argue that disconnection from the dominant "
      r"exchange network is an upstream cause that erodes both institutional quality and "
      r"technological capacity---the very factors that the internal-dynamics literature "
      r"identifies as proximate causes of collapse.")
    W()
    W(r"The literature on trade and isolation has overwhelmingly focused on deliberate "
      r"closure: the maritime prohibitions (\emph{haijin}) of Ming China, the \emph{sakoku} "
      r"decree of Tokugawa Japan, the autarkic blocs of the Cold War. Deliberate closure is "
      r"analytically convenient because it represents a policy choice with an identifiable "
      r"agent. Yet this focus on intentional closure overlooks a logically prior question: "
      r"what happens to polities that never had the option of connecting to the dominant "
      r"network in the first place? The Khmer Empire at Angkor, the Kievan Rus' principality, "
      r"or the Timurid dynasty were not isolated because their rulers chose closure; they were "
      r"isolated because the dominant exchange networks of their era did not reach them. If "
      r"these polities exhibit the same elevated conquest risk as deliberately closed ones, "
      r"then the mechanism at work is not the policy decision to close but the disconnection "
      r"itself.")
    W()
    W(r"We propose that the critical channel is technology flow. Maritime trade was never "
      r"merely an exchange of goods; it was the primary vehicle through which military "
      r"techniques, navigation methods, metallurgy, and institutional innovations diffused "
      r"across polities \citep{Pomeranz2000, Mokyr2002}. Deliberate closure, such as "
      r"\emph{sakoku}, restricted but did not eliminate this flow---Tokugawa Japan maintained "
      r"a narrow conduit of Western scientific knowledge (\emph{rangaku}) through the Dutch "
      r"trading post at Dejima. Technical exclusion, by contrast, severed the flow entirely. "
      r"The result was a cumulative technology gap: over generations, excluded polities fell "
      r"progressively behind the technological frontier, and when contact with a more advanced "
      r"civilization eventually occurred---often through military confrontation---the gap "
      r"proved fatal. This is, in essence, a quantitative restatement of a long-recognized "
      r"pattern: when civilizations at markedly different technological levels collide, the "
      r"less advanced one tends to be absorbed or destroyed \citep{Diamond1997}.")
    W()
    W(r"This paper makes three contributions. First, we construct a comparative historical "
      f"dataset of {N} polities spanning six eras (ancient, medieval, early modern, modern, "
      r"twentieth century, and contemporary) and classify each along two dimensions: its "
      r"predominant resource base---stock-oriented (accumulated assets such as human capital, "
      r"institutions, and natural resources) versus flow-oriented (trade, military projection, "
      r"and diplomatic engagement)---and its degree of closure from international networks. "
      r"Second, we introduce the concept of `technical network exclusion,' defined as the "
      r"involuntary disconnection of a polity from the dominant exchange network of its era "
      r"due to geographic or technological constraints. In the maritime age, this took the "
      r"form of geographic exclusion from sea routes, but the underlying mechanism generalizes "
      r"across eras. Third, we conduct a systematic sensitivity analysis in which seven "
      r"technically excluded polities are reclassified from `no closure' to `technical "
      r"network exclusion,' testing whether the closure--conquest association is an artifact "
      r"of how isolation is defined.")
    W()
    W(r"The principal finding is that reclassification transforms a non-significant baseline "
      f"association (Fisher's exact test, one-sided $p = {s_base_c['fisher_ban_p']:.3f}$) "
      f"into a significant one ($p = {s_all_c['fisher_ban_p']:.3f}$), while the core "
      f"stock--flow odds ratio remains unchanged at {cm_base['OR']:.3f}. All seven technically "
      r"excluded polities were eventually conquered. This result suggests that what matters "
      r"for state survival is not whether closure was chosen but whether the polity was "
      r"connected to the era's critical exchange platform.")
    W()
    W(r"The forward-looking implication follows directly. If the mechanism is technology flow "
      r"disruption leading to a cumulative gap, then the pattern need not be confined to the "
      r"maritime age. In a world where geography no longer blocks physical trade, the `dominant "
      r"network' has shifted from sea lanes to the semiconductor supply chain, artificial "
      r"intelligence infrastructure, and advanced robotics. States structurally excluded from "
      r"these technological platforms---whether by export controls, infrastructure deficits, "
      r"or institutional barriers---may face the same accumulating disadvantage that doomed "
      r"the technically excluded polities of the premodern era. We develop this argument in "
      r"the discussion.")
    W()
    W(r"The remainder of the paper is organized as follows. Section~\ref{sec:data} describes "
      r"the dataset and classification framework. Section~\ref{sec:methods} presents the "
      r"analytical methods. Section~\ref{sec:results} reports the baseline results. "
      r"Section~\ref{sec:sensitivity} details the sensitivity analysis. "
      r"Section~\ref{sec:discussion} discusses implications, including the extension to "
      r"technological access exclusion, and Section~\ref{sec:conclusion} concludes.")
    W()

    # ════════════════════════════════════════
    # 2. DATA AND CLASSIFICATION
    # ════════════════════════════════════════
    W(r"\section{Data and Classification}\label{sec:data}")
    W()
    W(r"\subsection{Dataset construction}\label{sec:data-construction}")
    W()
    W(f"The dataset comprises {N} historical polities selected from standard reference works "
      r"\citep{FindlayORourke2007, Kennedy1987, Turchin2009}. Selection criteria required "
      r"that each polity be (i) identifiable as a distinct political entity with a defined "
      r"period of existence, and (ii) classifiable along the dimensions described below. Each "
      r"polity is coded on multiple variables: predominant resource base (stock or flow), "
      r"stock index (0--1 continuous), trade openness (0--1 continuous), closure type, "
      r"historical outcome, geographic barrier, external threat level, relative population, "
      r"technological position, institutional quality, regime duration, and presence of an "
      r"external patron. The complete dataset, including modern-country equivalents and "
      r"turning-point events, is provided in Supplementary Table~S1.")
    W()
    W(f"Of the {N} polities, {n_stock} are classified as stock-oriented and {n_flow} as "
      f"flow-oriented. Historical outcomes fall into three categories: overtaken "
      f"({n_overtaken} polities)---conquest, colonization, or annexation by an external power; "
      f"disrupted ({n_disrupted})---regime collapse followed by reconstitution under a "
      r"successor state (e.g., Tokugawa Japan giving way to Meiji Japan); and survived "
      f"({n_survived})---continuity of both regime and statehood. The `disrupted' category is "
      r"analytically ambiguous, as these polities neither clearly fell to foreign conquest nor "
      r"clearly persisted; we address this through a dual-assignment sensitivity design "
      r"(Section~\ref{sec:methods-binarization}).")
    W()

    W(r"\subsection{Closure typology}\label{sec:closure-typology}")
    W()
    W(r"We classify each polity's degree of closure into five categories, ordered by the "
      r"nature and intent of isolation: (1)~maritime ban---deliberate restriction of maritime "
      r"trade by state policy (e.g., Ming \emph{haijin}, Qing Canton system); "
      r"(2)~\emph{sakoku}---near-total isolation enforced by national decree (Tokugawa Japan, "
      r"Joseon Korea); (3)~bloc---closure within a geopolitical or economic bloc (e.g., "
      r"COMECON, imperial preference systems); (4)~technical network exclusion---involuntary "
      r"isolation from the dominant exchange network of the era due to geographic or "
      r"technological constraints (not limited to maritime routes); and (5)~none---no "
      r"significant closure from prevailing exchange networks.")
    W()
    W(r"Categories 1 through 3 represent deliberate closure: a policy choice by identifiable "
      r"agents. Category 4 represents structural exclusion: isolation imposed by geography "
      r"and the limits of contemporary transport technology. This distinction is central to "
      r"our argument, because if structural exclusion produces the same outcomes as deliberate "
      r"closure, the causal mechanism must lie in the disconnection itself rather than in the "
      r"decision to disconnect.")
    W()

    W(r"\subsection{Technical network exclusion: definition and candidates}"
      r"\label{sec:tech-exclusion}")
    W()
    W(r"We define technical network exclusion as the condition in which a polity was "
      r"structurally disconnected from the dominant exchange network of its era---whether "
      r"maritime, overland, or otherwise---due to geographic or technological constraints "
      r"rather than policy choice. In the maritime age, this typically meant the absence of "
      r"regular sea routes; for inland polities along the Silk Road, it meant that while some "
      r"goods traveled overland, the volume and variety of technology transfer fell far short "
      r"of what maritime-connected polities received. The key distinction from policy-based "
      r"closure is the absence of agency: these polities did not choose isolation. The "
      r"transport and communication infrastructure of their era had not yet extended effective "
      r"connectivity to their region.")
    W()
    W(r"We identify seven reclassification candidates in two tiers (Table~\ref{tab:reclass}).")
    W()

    # Table 1
    W(r"\begin{table}[htbp]")
    W(r"\caption{Technical network exclusion reclassification candidates. "
      r"`Strong' candidates were structurally excluded from the dominant exchange network of "
      r"their era; `Moderate' candidates had limited or uncertain connectivity.}")
    W(r"\label{tab:reclass}")
    W(r"\begin{tabular*}{\textwidth}{@{\extracolsep\fill}lllp{7cm}}")
    W(r"\toprule")
    W(r"Polity & Period & Tier & Rationale \\")
    W(r"\midrule")
    for en_name, period, tier, rat in table1_rows:
        W(f"{_esc(en_name)} & {_esc(period)} & {tier} & {_esc(rat)} \\\\")
    W(r"\botrule")
    W(r"\end{tabular*}")
    W(r"\end{table}")
    W()

    # ════════════════════════════════════════
    # 3. METHODS
    # ════════════════════════════════════════
    W(r"\section{Methods}\label{sec:methods}")
    W()
    W(r"\subsection{Outcome binarization and sensitivity design}"
      r"\label{sec:methods-binarization}")
    W()
    W(r"The three-category outcome (overtaken, disrupted, survived) poses a classification "
      r"problem. The 18 `disrupted' polities experienced regime collapse but were not clearly "
      r"conquered by an external power; they could reasonably be grouped with either outcome. "
      r"Rather than making a single arbitrary assignment, we adopt a dual-assignment design: "
      r"in the `disrupted $\to$ overtaken' scenario, all 18 are coded as conquered (yielding "
      r"64 conquered, 32 survived); in the `disrupted $\to$ survived' scenario, they are coded "
      r"as having survived (46 conquered, 50 survived). Crossed with three closure "
      r"reclassification levels (baseline, +5 strong candidates, +7 all candidates), this "
      r"produces $3 \times 2 = 6$ scenarios. All results are reported across all six to "
      r"demonstrate robustness.")
    W()

    W(r"\subsection{Stock--flow association}\label{sec:methods-stockflow}")
    W()
    W(r"We construct $2 \times 2$ tables crossing the predominant resource base (stock vs.\ "
      r"flow) against the binarized outcome (overtaken vs.\ survived) and compute the odds "
      r"ratio (OR), phi coefficient ($\varphi$), one-sided Fisher's exact test, and "
      r"chi-squared test with Yates correction. Effect sizes are interpreted using Cohen's "
      r"benchmarks for $\varphi$.")
    W()

    W(r"\subsection{Closure--conquest association}\label{sec:methods-closure}")
    W()
    W(r"For each of the six scenarios, we construct a $2 \times 2$ table crossing closure "
      r"status (network closure vs.\ open) against the binarized outcome. The `network "
      r"closure' group includes polities coded as maritime ban, \emph{sakoku}, or technical "
      r"network exclusion. We compute one-sided Fisher's exact tests evaluating whether "
      r"closure is associated with elevated conquest rates, and report relative risks "
      r"alongside $p$-values.")
    W()

    W(r"\subsection{Multivariate logistic regression}\label{sec:methods-logistic}")
    W()
    W(r"To assess whether the closure--conquest association survives adjustment for "
      r"confounders, we fit logistic regression models with the binarized outcome as the "
      r"dependent variable and the following covariates: stock-dominant indicator, geographic "
      r"barrier, external threat level, technological position, institutional quality, era "
      r"(coded ordinally), external patron indicator, and a network closure indicator. We "
      r"report exponentiated coefficients (odds ratios) with 95\% confidence intervals.")
    W()

    W(r"\subsection{Bootstrap validation}\label{sec:methods-bootstrap}")
    W()
    W(r"We validate the stock--flow OR using a nonparametric bootstrap (5{,}000 resamples, "
      r"percentile method, seed = 42). This provides a distribution-free confidence interval "
      r"that does not depend on asymptotic normality---a relevant consideration given the "
      r"modest sample size.")
    W()

    # ════════════════════════════════════════
    # 4. RESULTS
    # ════════════════════════════════════════
    W(r"\section{Results}\label{sec:results}")
    W()
    W(r"\subsection{Baseline stock--flow association}\label{sec:results-baseline}")
    W()
    W(f"Under the disrupted $\\to$ overtaken assignment, the stock--flow $2 \\times 2$ table "
      f"yields OR $= {cm_base['OR']:.3f}$, $\\varphi = {cm_base['phi']:.3f}$, Fisher's exact "
      f"$p = {cm_base['p_fisher']:.4f}$ (one-sided). Stock-oriented polities have a conquest "
      f"rate of {_pct(cm_base['stock_conquest_rate'])}, compared with "
      f"{_pct(cm_base['flow_conquest_rate'])} for flow-oriented polities---a small-to-medium "
      f"effect by Cohen's benchmarks. Under the disrupted $\\to$ survived assignment, the OR "
      f"shifts to "
      f"{results['scenarios']['as_survived__baseline']['cm']['OR']:.3f}, illustrating the "
      "sensitivity of the stock--flow association to how the ambiguous `disrupted' category "
      "is handled. This dual sensitivity---to both closure definition and outcome "
      "binarization---motivates the full six-scenario analysis below.")
    W()

    W(r"\subsection{Network closure and conquest}\label{sec:results-closure}")
    W()
    W(f"Table~\\ref{{tab:scenarios}} reports the closure--conquest association across all six "
      f"scenarios. At baseline under the disrupted $\\to$ overtaken assignment, polities with "
      f"some form of network closure have a conquest rate of "
      f"{_pct(s_base_c['ban_rate'])}, compared with {_pct(s_base_c['no_ban_rate'])} for open "
      f"polities (Fisher $p = {s_base_c['fisher_ban_p']:.4f}$, not significant). "
      f"Reclassifying five strong candidates as technically excluded raises the "
      f"closure-group conquest rate to {_pct(s_strong_c['ban_rate'])} "
      f"(Fisher $p = {s_strong_c['fisher_ban_p']:.4f}$). "
      f"Including all seven candidates strengthens the association further: conquest rate $= "
      f"{_pct(s_all_c['ban_rate'])}$, Fisher $p = {s_all_c['fisher_ban_p']:.4f}$ "
      r"(Fig.~\ref{fig:conquest-rates}, Fig.~\ref{fig:fisher-pvalues}). "
      "The progressive strengthening of significance as technically excluded polities are "
      "added suggests that these polities genuinely belong in the closure group rather than "
      "among the open polities.")
    W()

    # Table 2
    W(r"\begin{table}[htbp]")
    W(r"\caption{Network closure and conquest across six scenarios. $^{*}$~$p < 0.05$.}")
    W(r"\label{tab:scenarios}")
    W(r"\begin{tabular*}{\textwidth}{@{\extracolsep\fill}llcccccl}")
    W(r"\toprule")
    W(r"Disrupted As & Reclassification & Closure $n$ & Closure Rate & Open Rate & RR & Fisher $p$ & Sig. \\")
    W(r"\midrule")
    for row in table2_rows:
        sig_str = "$^{*}$" if row[7] == "*" else ""
        W(f"{row[0]} & {row[1]} & {row[2]} & {row[3]} & {row[4]} & {row[5]} & {row[6]} & {sig_str} \\\\")
    W(r"\botrule")
    W(r"\end{tabular*}")
    W(r"\end{table}")
    W()

    # Figures 1 and 2 (referenced, placed at end per journal style)
    W(r"\begin{figure}[htbp]")
    W(r"\centering")
    W(r"\includegraphics[width=\textwidth]{figures/Fig1.png}")
    W(r"\caption{Conquest rates comparing network closure vs.\ open polities across three "
      r"reclassification scenarios under both disrupted assignments.}")
    W(r"\label{fig:conquest-rates}")
    W(r"\end{figure}")
    W()
    W(r"\begin{figure}[htbp]")
    W(r"\centering")
    W(r"\includegraphics[width=0.85\textwidth]{figures/Fig2.png}")
    W(r"\caption{Fisher's exact test $p$-values for the network closure $\to$ conquest "
      r"association as polities are progressively reclassified.}")
    W(r"\label{fig:fisher-pvalues}")
    W(r"\end{figure}")
    W()

    # ════════════════════════════════════════
    # 5. SENSITIVITY ANALYSIS
    # ════════════════════════════════════════
    W(r"\section{Sensitivity Analysis: Technical Network Exclusion}\label{sec:sensitivity}")
    W()
    W(r"\subsection{Closure-type disaggregation and the dose--response pattern}"
      r"\label{sec:sensitivity-dose}")
    W()
    W(r"Figure~\ref{fig:closure-types} disaggregates conquest rates by closure type under "
      r"the 7-country reclassification. A striking gradient emerges. Technically excluded "
      r"polities---those with zero access to the dominant exchange networks of their "
      r"era---exhibit a 100\% conquest rate. Policy-based maritime bans, which restricted but "
      r"did not entirely eliminate external contact, show lower rates (76.9\% under disrupted "
      r"$\to$ overtaken; 69.2\% under disrupted $\to$ survived). \emph{Sakoku} polities show "
      r"100\% and 50\% rates under the two assignments, reflecting the borderline case of "
      r"Tokugawa Japan, which maintained a narrow technological conduit (\emph{rangaku}) "
      r"through Dejima. Bloc-type closures show the lowest conquest rates among closure "
      r"categories, consistent with the interpretation that bloc membership preserves some "
      r"technology transfer through alliance-internal channels.")
    W()
    W(r"This ordering---technical exclusion (100\%) $>$ policy ban $>$ \emph{sakoku} (with "
      r"partial conduit) $>$ bloc $>$ open---is consistent with a dose--response relationship "
      r"between the degree of technology flow disruption and the probability of conquest. The "
      r"more completely a polity was severed from the technological frontier of its era, the "
      r"higher the likelihood that it was eventually overtaken.")
    W()

    # Figure 3
    W(r"\begin{figure}[htbp]")
    W(r"\centering")
    W(r"\includegraphics[width=\textwidth]{figures/Fig3.png}")
    W(r"\caption{Conquest rates by closure type under the 7-country reclassification "
      r"scenario.}")
    W(r"\label{fig:closure-types}")
    W(r"\end{figure}")
    W()

    W(r"\subsection{Robustness of the stock--flow odds ratio}"
      r"\label{sec:sensitivity-or}")
    W()
    W(f"The stock--flow OR $= {cm_all_c['OR']:.3f}$ is identical across all three "
      r"reclassification scenarios. This invariance is expected: the reclassification changes "
      r"the closure-type label but does not alter the stock/flow or outcome coding. Bootstrap "
      f"validation (5{{,}}000 resamples) yields a median OR of "
      f"${boot_c.get('median', 0):.3f}$ (95\\% CI "
      f"[{boot_c.get('ci_lo', 0):.3f}, {boot_c.get('ci_hi', 0):.3f}]), "
      r"confirming the stability of the point estimate and its independence from the closure "
      r"reclassification.")
    W()

    W(r"\subsection{Multivariate regression stability}\label{sec:sensitivity-regression}")
    W()
    W(r"Figure~\ref{fig:forest-plot} presents the multivariate logistic regression results "
      r"under the 7-country reclassification with disrupted $\to$ overtaken. External threat "
      r"remains the strongest predictor of conquest ($p < 0.01$ across all scenarios), "
      r"followed by institutional quality and era (both $p < 0.01$). The network closure "
      r"indicator is not independently significant after controlling for these covariates. "
      r"This pattern is informative: it suggests that closure operates not as a direct cause "
      r"but through the same channels---technological stagnation, institutional decay, and "
      r"heightened external vulnerability---that the multivariate model already captures "
      r"(Table~\ref{tab:regression}).")
    W()

    # Figure 4
    W(r"\begin{figure}[htbp]")
    W(r"\centering")
    W(r"\includegraphics[width=0.9\textwidth]{figures/Fig4.png}")
    W(r"\caption{Forest plot of multivariate logistic regression odds ratios (7-country "
      r"reclassification, disrupted $\to$ overtaken).}")
    W(r"\label{fig:forest-plot}")
    W(r"\end{figure}")
    W()

    # Table 3
    if table3_rows:
        W(r"\begin{table}[htbp]")
        W(r"\caption{Multivariate logistic regression results (7-country reclassification, "
          r"disrupted $\to$ overtaken). $^{*}$~$p < 0.05$, $^{\dag}$~$p < 0.10$.}")
        W(r"\label{tab:regression}")
        W(r"\begin{tabular*}{\textwidth}{@{\extracolsep\fill}lcccc}")
        W(r"\toprule")
        W(r"Variable & OR & 95\% CI & $p$-value & Sig. \\")
        W(r"\midrule")
        for row in table3_rows:
            W(f"{row[0]} & {row[1]} & {row[2]} & {row[3]} & {row[4]} \\\\")
        W(r"\botrule")
        W(r"\end{tabular*}")
        W(r"\end{table}")
        W()

    # ════════════════════════════════════════
    # 6. DISCUSSION
    # ════════════════════════════════════════
    W(r"\section{Discussion}\label{sec:discussion}")
    W()
    W(r"\subsection{The mechanism: technology flow disruption and cumulative divergence}"
      r"\label{sec:disc-mechanism}")
    W()
    W(r"The central finding is that the closure--conquest association becomes significant "
      r"only when technically excluded polities are grouped with deliberately closed ones. "
      r"This tells us something important about the mechanism at work. If closure harmed "
      r"states solely through lost trade revenue or reduced diplomatic leverage, then only "
      r"deliberate closure---which blocks trade but not necessarily knowledge---should matter. "
      r"The fact that technical exclusion (which severs both trade and technology flow) "
      r"strengthens the association, while policy-based closure alone does not reach "
      r"significance, points toward technology flow as the critical channel.")
    W()
    W(r"The dose--response pattern in Figure~\ref{fig:closure-types} reinforces this "
      r"interpretation. Technical exclusion (zero technology transfer) produces a 100\% "
      r"conquest rate. Policy-based maritime bans, which restrict but do not eliminate "
      r"technology flow, show conquest rates below 80\%. The case of Tokugawa Japan is "
      r"particularly instructive: despite the comprehensive closure of \emph{sakoku}, the "
      r"Tokugawa regime deliberately maintained a narrow conduit for Western scientific and "
      r"technical knowledge (\emph{rangaku}) through the Dutch trading post at Dejima. This "
      r"selective preservation of a technology transfer channel---even within an otherwise "
      r"closed system---appears to have made a material difference. Japan was disrupted by "
      r"the forced opening of 1853--54 but was not conquered; it reconstituted itself as the "
      r"Meiji state and rapidly closed the technology gap. Bloc closures, which preserve "
      r"substantial within-bloc technology sharing, show the lowest rates among closure "
      r"categories. The ordering of conquest risk mirrors the ordering of technology flow "
      r"disruption.")
    W()
    W(r"The multivariate results complete the picture. External threat and institutional "
      r"quality are the strongest predictors of conquest, and the network closure indicator "
      r"loses significance after their inclusion (Table~\ref{tab:regression}, "
      r"Fig.~\ref{fig:forest-plot}). This is precisely what a technology-gap mechanism would "
      r"predict: closure does not kill states directly. Rather, it initiates a causal "
      r"chain---technological stagnation erodes institutional adaptive capacity, which in "
      r"turn leaves the polity unable to respond to external threats. This causal ordering is "
      r"consistent with \citeauthor{AcemogluRobinson2012}'s (\citeyear{AcemogluRobinson2012}) "
      r"emphasis on institutions as the proximate determinant of national success, while "
      r"suggesting that network access is the deeper, upstream variable: exclusion degrades "
      r"the very institutions that the inclusive-institutions framework identifies as "
      r"essential. The mediating variables (external threat, institutional quality) absorb "
      r"the explanatory power of the closure variable because they lie downstream in the "
      r"causal pathway.")
    W()

    W(r"\subsection{First contact and the tragedy of civilizational divergence}"
      r"\label{sec:disc-firstcontact}")
    W()
    W(r"Our findings can be read as a quantitative formulation of a long-recognized "
      r"historical pattern: when civilizations that have developed in isolation encounter a "
      r"technologically superior civilization, the outcome is overwhelmingly unfavorable for "
      r"the less advanced party \citep{Diamond1997, DiamondBellwood2003}. The 100\% conquest "
      r"rate among technically excluded polities is striking not because the pattern is new, "
      r"but because it emerges from a systematic, cross-historical dataset rather than from "
      r"selective case studies.")
    W()
    W(r"The mechanism we propose is cumulative divergence through technology flow disruption. "
      r"International exchange networks carried not only goods but military techniques, "
      r"metallurgical innovations, navigational knowledge, and institutional models "
      r"\citep{Mokyr2002, Pomeranz2000}. Polities connected to these networks could adopt, "
      r"adapt, and build upon innovations generated elsewhere. Polities severed from them "
      r"could not. Over generations, the technology gap widened---a process analogous to the "
      r"long-run consequences of network disruption documented by \citet{Nunn2008}, who showed "
      r"that regions more heavily affected by the slave trade experienced persistent "
      r"underdevelopment centuries later. When contact with a more advanced civilization "
      r"eventually occurred---often through military expansion---the accumulated gap proved "
      r"decisive. The Han Dynasty encountered Central Asian and eventually Roman-linked "
      r"military traditions; the Khmer Empire faced the expanding Siamese and Vietnamese "
      r"states that were integrated into maritime trade networks; Kievan Rus' was overrun by "
      r"the Mongol armies that had absorbed the military technologies of multiple "
      r"civilizations across Eurasia.")
    W()
    W(r"Crucially, the tragedy of first contact is a function of the gap, not of the contact "
      r"itself. Policy-closed polities that maintained narrow conduits of technology "
      r"transfer---Japan's \emph{rangaku}, Qing China's limited Canton trade---accumulated "
      r"smaller gaps and, correspondingly, were more likely to survive or reconstitute after "
      r"disruption. The implication is that what determines the outcome of civilizational "
      r"encounter is the degree and duration of technology flow disruption that preceded it.")
    W()
    W(r"This finding has a corollary that is worth stating explicitly. If the critical "
      r"variable is not closure itself but the residual technology flow that closure permits, "
      r"then polities that close their borders while deliberately maintaining selective "
      r"channels for frontier knowledge occupy a qualitatively different position from those "
      r"that are totally severed. Our dataset contains several instances of such conditional "
      r"closure. Tokugawa Japan preserved access to Western science through \emph{rangaku} at "
      r"Dejima and was disrupted but not conquered, reconstituting itself as the Meiji state. "
      r"Early Qing China maintained the Canton system---a single, tightly controlled port of "
      r"trade through which some foreign knowledge filtered---and survived the "
      r"Kangxi--Qianlong era intact. Joseon Korea sustained tributary trade with China, which "
      r"served as a conduit for Continental knowledge and technology. The Soviet Union, while "
      r"sealed from the Western bloc, maintained extensive technology sharing within the "
      r"Eastern bloc and invested heavily in indigenous research institutions. North Korea has "
      r"preserved a limited technology channel through its relationship with China. (Full "
      r"details of each polity's closure regime, technology channels, and outcomes are "
      r"recorded in Supplementary Table~S1.)")
    W()
    W(r"The outcomes, however, are mixed. Early Qing China survived, but late Qing "
      r"China---operating under the same Canton system---was semi-colonized after the Opium "
      r"Wars. Tokugawa Japan survived as a political entity, but Joseon Korea, despite its "
      r"Chinese conduit, was annexed by Japan in 1910. The Soviet Union collapsed internally "
      r"despite its bloc-level technology sharing, while North Korea has persisted under "
      r"extreme isolation with minimal Chinese patronage. These divergent outcomes suggest "
      r"that the mere existence of a selective channel does not guarantee survival; unmeasured "
      r"factors---the breadth and relevance of the knowledge flowing through the channel, the "
      r"domestic capacity to absorb and adapt that knowledge, the pace of change at the "
      r"technological frontier, and institutional dynamics that our dataset does not "
      r"capture---likely condition the effectiveness of conditional closure. The policy "
      r"question, then, is not binary (open or closed) but conditional, and the conditions "
      r"under which selective channels suffice to prevent a fatal technology gap remain an "
      r"open and consequential problem for future research. Readers interested in tracing "
      r"these cases in detail are referred to Supplementary Table~S1, which documents the "
      f"specific turning-point events and outcomes for all {N} polities in the dataset.")
    W()

    W(r"\subsection{Beyond geographic isolation: technological access exclusion in the "
      r"modern era}\label{sec:disc-modern}")
    W()
    W(r"If the mechanism we identify is technology flow disruption rather than network "
      r"closure per se, then the pattern should generalize beyond the age of sail. Each era "
      r"has its own dominant technological platform---the infrastructure through which "
      r"frontier knowledge diffuses across states. In antiquity, it was the Mediterranean "
      r"trade routes and Silk Road caravans. In the early modern period, it was oceanic "
      r"shipping. In the industrial age, it was railroad-linked factory systems and colonial "
      r"supply chains. In the twentieth century, it was aerospace and nuclear technology "
      r"networks.")
    W()
    W(r"Today, the dominant platform has shifted again. The critical networks are "
      r"semiconductor supply chains, artificial intelligence research ecosystems, and "
      r"advanced robotics and automation infrastructure "
      r"\citep{AcemogluRestrepo2020, CominMestieri2018}. "
      r"\citet{CominEasterlyGong2010} show that technology adoption levels in 1000~BC predict "
      r"income differences today, while \citet{AcemogluJohnsonRobinson2002} demonstrate that "
      r"colonial-era institutional reversals reshaped global inequality---both consistent "
      r"with the view that early technological access has persistent, cumulative consequences. "
      r"Geographic isolation no longer blocks physical trade---the completion of global "
      r"shipping and communication networks has largely eliminated geographic network "
      r"exclusion as a threat. But a new form of structural exclusion has emerged: states may "
      r"be cut off from the technological frontier not by mountains and oceans but by export "
      r"controls on advanced semiconductors, by the concentration of AI training "
      r"infrastructure in a handful of countries, or by the institutional and human-capital "
      r"barriers that prevent participation in cutting-edge research networks.")
    W()
    W(r"The historical parallel is direct. Just as the Khmer Empire or the Timurid dynasty "
      r"could not access the exchange networks that carried military and institutional "
      r"innovations, a contemporary state excluded from advanced semiconductor fabrication or "
      r"AI model development may find itself on the wrong side of a widening technology gap. "
      r"If the gap grows large enough, the eventual `first contact'---whether military, "
      r"economic, or geopolitical---may produce outcomes analogous to those documented in our "
      r"historical dataset. The form of the dominant network changes; the logic of cumulative "
      r"divergence through exclusion does not.")
    W()
    W(r"We stress that this extrapolation is speculative and cannot be tested within our "
      r"historical dataset. The contemporary world differs from the premodern era in ways "
      r"that may attenuate or amplify the mechanism: nuclear deterrence, international "
      r"institutions, and the speed of modern communication all introduce novel dynamics. "
      r"Nevertheless, the historical regularity we document---that structural exclusion from "
      r"the dominant technological network is associated with state collapse---provides a "
      r"framework for thinking about which dimensions of modern technological access may be "
      r"most consequential.")
    W()

    W(r"\subsection{Robustness of the stock--flow framework and the question of "
      r"resource-base transitions}\label{sec:disc-stockflow}")
    W()
    W(f"The invariance of the stock--flow OR ({cm_all_c['OR']:.3f}) across all "
      r"reclassification scenarios confirms that the core finding of the stock--flow "
      r"framework---that stock-oriented polities face moderately higher conquest risk---is "
      r"independent of how network isolation is defined. The reclassification changes the "
      r"closure subanalysis but leaves the primary classification untouched. This separation "
      r"is analytically useful: it shows that the stock--flow distinction and the "
      r"closure--conquest association capture related but distinct dimensions of state "
      r"vulnerability.")
    W()
    W(f"Crossing these two dimensions yields a four-cell classification whose conquest rates "
      f"are reported in Table~\\ref{{tab:stockflow-closure}}. Flow-oriented polities without "
      f"closure show the lowest conquest rate ({_pct(fo['rate'])}, $n = {fo['n']}$). "
      f"Stock-oriented polities without closure show a moderately elevated rate "
      f"({_pct(so['rate'])}, $n = {so['n']}$). Stock-oriented polities with closure show a "
      f"markedly higher rate ({_pct(sc['rate'])}, $n = {sc['n']}$). Flow-oriented polities with "
      f"closure---a small cell ($n = {fc['n']}$)---show a {_pct(fc['rate'])} conquest rate. "
      f"The contrast between the highest- and lowest-risk cells (stock + closed vs.\\ flow + "
      f"open) yields an odds ratio of {or_sc_fo:.2f} (one-sided Fisher exact "
      f"$p = {p_sc_fo:.3f}$). Within stock-oriented polities, the effect of closure "
      f"corresponds to an OR of {or_sc_so:.2f} ($p = {p_sc_so:.3f}$). Within flow-oriented "
      f"polities, the effect of closure yields an OR of "
      f"{'$\\infty$' if np.isinf(or_fc_fo) else f'${or_fc_fo:.2f}$'} ($p = {p_fc_fo:.3f}$), "
      f"though the cell size ($n = {fc['n']}$) precludes reliable inference.")
    W()

    # Table 4
    W(r"\begin{table}[htbp]")
    W(r"\caption{Conquest rates by resource-base orientation and closure status (7-country "
      r"reclassification, disrupted $=$ conquered).}")
    W(r"\label{tab:stockflow-closure}")
    W(r"\begin{tabular*}{\textwidth}{@{\extracolsep\fill}lcccc}")
    W(r"\toprule")
    W(r" & Conquered & Survived & $n$ & Rate \\")
    W(r"\midrule")
    for label, cell in [("Stock + closed", sc), ("Stock + open", so),
                        ("Flow + closed", fc), ("Flow + open", fo)]:
        rate_str = f"{cell['rate']*100:.1f}\\%"
        W(f"{label} & {cell['c']} & {cell['s']} & {cell['n']} & {rate_str} \\\\")
    W(r"\botrule")
    W(r"\end{tabular*}")
    or_fc_fo_str = r"\infty" if np.isinf(or_fc_fo) else f"{or_fc_fo:.2f}"
    W(r"\footnotetext{Fisher exact tests (one-sided): stock + closed vs.\ flow + open, "
      f"OR $= {or_sc_fo:.2f}$, $p = {p_sc_fo:.3f}$; within stock, closed vs.\\ open, "
      f"OR $= {or_sc_so:.2f}$, $p = {p_sc_so:.3f}$; within flow, closed vs.\\ open, "
      f"OR $= {or_fc_fo_str}$, $p = {p_fc_fo:.3f}$.}}")
    W(r"\end{table}")
    W()

    W(f"The pattern is consistent with the hypothesis that stock-orientation and network "
      f"exclusion compound each other's risk, though the small cell sizes---particularly for "
      f"flow + closed ($n = {fc['n']}$)---warrant caution. The stock + closed combination is "
      r"the only cell to reach a statistically significant difference from the baseline flow "
      f"+ open cell at conventional levels. Notably, {sc['s']} stock-oriented polities with "
      r"closure survived despite their high-risk classification; their individual "
      r"characteristics---including specific closure regimes, technology channels maintained, "
      r"and the historical circumstances of their survival---are documented in Supplementary "
      r"Table~S1.")
    W()
    W(r"An implication worth noting is that the stock--flow distinction is not permanently "
      r"fixed for any given state. Polities can transition from flow-oriented to "
      r"stock-oriented resource bases---or vice versa---as economic structures, demographic "
      r"conditions, and institutional incentives evolve. A state that was historically "
      r"flow-oriented (trade-dependent, outward-looking, innovation-absorbing) but that "
      r"gradually shifts toward reliance on accumulated domestic assets---physical capital, "
      r"territorial resources, existing institutional infrastructure---moves into the "
      r"higher-risk stock category. If such a transition coincides with reduced engagement in "
      r"the dominant technological network of the era, the compounding of stock-orientation "
      r"and network exclusion could amplify vulnerability. Our dataset cannot test this "
      r"dynamic directly, as polities are coded at a single point in time, but the "
      r"theoretical interaction between resource-base transition and network access merits "
      r"further investigation.")
    W()

    W(r"\subsection{Limitations}\label{sec:disc-limitations}")
    W()
    W(f"Several limitations warrant acknowledgment. First, the coding of historical polities "
      r"inevitably involves subjective judgment, particularly for the stock/flow "
      r"classification and the identification of technical exclusion candidates. The tiered "
      r"approach (strong vs.\ moderate candidates) and the six-scenario sensitivity design "
      r"partially address this, but cannot eliminate it. Second, the sample size ($N = "
      f"{N}$) constrains the power of the multivariate analyses; wide confidence intervals "
      r"for some regression coefficients reflect this constraint. Third, the dataset treats "
      r"polities as independent observations, though historical interconnections (e.g., "
      r"sequential Chinese dynasties sharing institutional continuity) may introduce "
      r"non-independence. Fourth, the `disrupted' category introduces a classification "
      r"ambiguity that the dual-assignment design addresses but cannot fully resolve. Fifth, "
      r"the forward-looking extension to modern technological exclusion is necessarily "
      r"speculative, as the mechanisms operating in a nuclear-armed, institutionally dense "
      r"modern world may differ qualitatively from those in premodern eras.")
    W()

    # ════════════════════════════════════════
    # 7. CONCLUSION
    # ════════════════════════════════════════
    W(r"\section{Conclusion}\label{sec:conclusion}")
    W()
    W(r"This paper has shown that reclassifying seven technically excluded polities---those "
      r"severed from the dominant exchange networks of their era by geography and technology "
      r"rather than by policy---transforms a non-significant association between closure and "
      f"conquest into a significant one ($p = {s_all_c['fisher_ban_p']:.3f}$), while the "
      f"core stock--flow odds ratio remains unchanged (OR $= {cm_base['OR']:.3f}$). The "
      r"100\% conquest rate among technically excluded polities, the dose--response gradient "
      r"across closure types, and the absorption of the closure effect by external threat "
      r"and institutional quality in multivariate models all point to a consistent mechanism: "
      r"disruption of technology flow leads to cumulative divergence from the technological "
      r"frontier, eroding the institutional and military capacity needed to survive contact "
      r"with more connected civilizations.")
    W()
    W(r"The broader implication is that this mechanism is not specific to any single network "
      r"form. In every era, there exists a dominant network through which frontier "
      r"technologies diffuse. Polities excluded from that network---whether by oceans, "
      r"mountains, policy, or, in the contemporary period, by semiconductor export controls "
      r"and AI infrastructure concentration---risk falling into the same pattern of "
      r"cumulative divergence. The historical record we document provides a quantitative "
      r"baseline for assessing this risk. Whether the tragedy of first contact between "
      r"unequally developed civilizations will find new expression in the age of artificial "
      r"intelligence is a question that the coming decades will answer; our analysis suggests "
      r"it is one worth asking. Equally pressing is the converse question suggested by the "
      r"Tokugawa precedent: whether a state that recognizes the risk of network exclusion "
      r"can, through deliberate maintenance of selective technology channels, prevent the "
      r"accumulation of a fatal gap---even while restricting broader engagement with the "
      r"outside world.")
    W()

    # ════════════════════════════════════════
    # STATEMENTS AND DECLARATIONS
    # ════════════════════════════════════════
    W(r"\section*{Statements and Declarations}")
    W()
    W(r"\subsection*{Funding}")
    W(r"This work received no external funding.")
    W()
    W(r"\subsection*{Competing Interests}")
    W(r"The author declares no conflicts of interest.")
    W()
    W(r"\subsection*{Acknowledgments}")
    W(r"None.")
    W()
    W(r"\subsection*{ORCID}")
    W(r"Onishi Tatsuki: \href{https://orcid.org/0000-0001-7261-9062}"
      r"{0000-0001-7261-9062}")
    W()
    W(r"\subsection*{Author Contributions}")
    W(r"O.T.: Conceptualization, Methodology, Software, Formal analysis, "
      r"Writing --- original draft, Writing --- review \& editing.")
    W()
    W(r"\subsection*{Declaration of Generative Artificial Intelligence (AI) in Scientific Writing}")
    W(r"We used devin.ai to help with formatting the text and choosing words that suited "
      r"the tone, and to help writing codes. The author takes full responsibility for the "
      r"accuracy and content of the manuscript.")
    W()
    W(r"\subsection*{Data Availability}")
    W(r"The complete dataset and analysis code are available at "
      r"\url{https://github.com/bougtoir/gdp-tempo-paper}. "
      f"Supplementary Table~S1 provides the full dataset of {N} polities with all coded "
      r"variables. This study uses only publicly available historical data; "
      r"we gratefully acknowledge the open-data sources on which the dataset is built.")
    W()

    # JEL
    W(r"\subsection*{JEL Classification}")
    W(r"N40, N70, F50, O33, C12")
    W()

    # Bibliography
    # Note: \bibliographystyle is set automatically by sn-jnl.cls via the sn-basic option
    W(r"\bibliography{references}")
    W()
    W(r"\end{document}")

    # Write .tex
    tex_path = os.path.join(OUT, "manuscript.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("\n".join(tex))
    print("  LaTeX manuscript saved to", tex_path)
    return tex_path


# ═══════════════════════════════════════════════════════════════
# Supplementary Table S1 (LaTeX)
# ═══════════════════════════════════════════════════════════════

def create_table_s1_latex(results):
    df = results["df"]
    era_map = {"ancient": "Ancient", "medieval": "Medieval",
               "early_modern": "Early Modern", "modern": "Modern",
               "20c": "20th Century", "contemporary": "Contemporary"}
    closure_map = {"none": "None", "maritime_ban": "Maritime ban",
                   "sakoku": "Sakoku", "bloc": "Bloc",
                   "technical_network_exclusion": "Technical network exclusion"}

    tex = []
    tex.append(r"\documentclass[11pt]{article}")
    tex.append(r"\usepackage[margin=1cm,landscape]{geometry}")
    tex.append(r"\usepackage{booktabs}")
    tex.append(r"\usepackage{longtable}")
    tex.append(r"\usepackage{array}")
    tex.append(r"\usepackage[T1]{fontenc}")
    tex.append(r"\begin{document}")
    tex.append(r"\scriptsize")
    tex.append(r"\setlength{\LTleft}{0pt}")
    tex.append(r"\setlength{\LTright}{0pt}")
    tex.append(r"\begin{longtable}{@{}r l l l l l l l p{5.5cm}@{}}")
    tex.append(r"\caption{Supplementary Table S1: Full dataset of 96 historical polities.} \\")
    tex.append(r"\toprule")
    tex.append(r"\# & Polity & Modern Country & Period & Era & Dominant & Closure & Outcome & Turning-Point Event \\")
    tex.append(r"\midrule")
    tex.append(r"\endfirsthead")
    tex.append(r"\toprule")
    tex.append(r"\# & Polity & Modern Country & Period & Era & Dominant & Closure & Outcome & Turning-Point Event \\")
    tex.append(r"\midrule")
    tex.append(r"\endhead")
    tex.append(r"\midrule")
    tex.append(r"\endfoot")
    tex.append(r"\bottomrule")
    tex.append(r"\endlastfoot")

    for idx, (_, row) in enumerate(df.iterrows()):
        entity = row["entity"]
        vals = [
            str(idx + 1),
            _esc(ENGLISH_NAME.get(entity, entity)),
            _esc(MODERN_COUNTRY.get(entity, "---")),
            _esc(row["period"]),
            era_map.get(row["era"], row["era"]),
            row["dominant"].capitalize(),
            closure_map.get(row["closure_type"], row["closure_type"]),
            row["outcome"].capitalize(),
            _esc(TURNING_POINT.get(entity, "---")),
        ]
        tex.append(" & ".join(vals) + r" \\")

    tex.append(r"\end{longtable}")
    tex.append(r"\end{document}")

    path = os.path.join(OUT, "table_s1.tex")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(tex))
    print("  Table S1 (LaTeX) saved to", path)
    return path


# ═══════════════════════════════════════════════════════════════
# Cover letter (LaTeX)
# ═══════════════════════════════════════════════════════════════

def create_cover_letter_tex():
    tex = []
    tex.append(r"\documentclass[11pt]{letter}")
    tex.append(r"\usepackage[margin=2.5cm]{geometry}")
    tex.append(r"\usepackage[T1]{fontenc}")
    tex.append(r"\usepackage{hyperref}")
    tex.append(r"\signature{Onishi Tatsuki\\Data Science AI Innovation Research Promotion Center, "
               r"Shiga University\\bougtoir@gmail.com}")
    tex.append(r"\address{Onishi Tatsuki\\Data Science AI Innovation Research Promotion Center, "
               r"Shiga University\\1-1-1 Bamba, Hikone, Shiga 522-8522, Japan\\"
               r"bougtoir@gmail.com}")
    tex.append(r"\begin{document}")
    tex.append(r"\begin{letter}{Prof.\ Claude Diebolt\\Managing Editor, \emph{Cliometrica}}")
    tex.append(r"\opening{Dear Professor Diebolt,}")
    tex.append("")
    tex.append(r"I am pleased to submit the manuscript entitled ``Network Exclusion and "
               r"State Collapse: From Maritime Isolation to Technological Access Denial in "
               r"the Long Run of History'' for consideration by \emph{Cliometrica}.")
    tex.append("")
    tex.append(r"Using a comparative dataset of 96 historical polities spanning antiquity to "
               r"the present, we distinguish between deliberate closure (policy-based trade "
               r"bans, \emph{sakoku}, bloc membership) and what we term \emph{technical "
               r"network exclusion}: involuntary disconnection from the dominant exchange "
               r"network of an era due to geographic or technological constraints. The "
               r"manuscript presents four features that we believe are of interest to the "
               r"journal's readership.")
    tex.append("")
    tex.append(r"First, a systematic sensitivity analysis demonstrates that reclassifying "
               r"seven technically excluded polities transforms a non-significant association "
               r"between closure and conquest (Fisher's exact $p = 0.187$) into a significant "
               r"one ($p = 0.020$), while the core stock--flow odds ratio remains stable "
               r"(OR $= 1.774$).")
    tex.append("")
    tex.append(r"Second, we analyze conditional closure---cases where polities maintained "
               r"selective technology channels while restricting broader trade. Examples "
               r"include Tokugawa Japan's \emph{rangaku} through Dejima, Qing China's Canton "
               r"system, Joseon Korea's tributary trade, and the Soviet Union's bloc-internal "
               r"technology sharing. The mixed outcomes of these cases suggest that the "
               r"open/closed dichotomy is insufficient; the conditions under which selective "
               r"channels mitigate technology gaps merit further investigation.")
    tex.append("")
    tex.append(r"Third, the dose--response gradient across closure types---technical "
               r"exclusion (100\% conquest) $>$ policy bans $>$ \emph{sakoku} (with partial "
               r"conduit) $>$ bloc $>$ open---points to technology flow disruption, rather "
               r"than trade restriction per se, as the critical mechanism.")
    tex.append("")
    tex.append(r"Fourth, the mechanism generalizes beyond geographic isolation: as the "
               r"dominant network shifts from sea lanes to semiconductors, AI, and advanced "
               r"robotics, states structurally excluded from these platforms may face "
               r"analogous vulnerabilities.")
    tex.append("")
    tex.append(r"The manuscript contains approximately 9{,}000 words, 4 tables, and 4 "
               r"figures. Supplementary Table~S1 provides the complete dataset. We confirm "
               r"that this work has not been published or submitted elsewhere.")
    tex.append("")
    tex.append(r"We suggest the following reviewers based on their expertise in quantitative "
               r"economic history:")
    tex.append(r"\begin{itemize}")
    tex.append(r"\item Prof.\ J\"org Baten (University of T\"ubingen) --- cliometric methods "
               r"and long-run development")
    tex.append(r"\item Prof.\ Stephen Broadberry (University of Oxford) --- comparative "
               r"economic history and GDP estimation")
    tex.append(r"\item Prof.\ Peter Turchin (Complexity Science Hub Vienna) --- quantitative "
               r"historical dynamics and empire formation")
    tex.append(r"\end{itemize}")
    tex.append("")
    tex.append(r"\closing{Sincerely,}")
    tex.append(r"\end{letter}")
    tex.append(r"\end{document}")

    path = os.path.join(OUT, "cover_letter.tex")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(tex))
    print("  Cover letter (LaTeX) saved to", path)
    return path


# ═══════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════

def main():
    print("Running analysis...")
    results = run_analysis()

    print("Creating figures...")
    create_figures(results)

    print("Creating PPTX...")
    create_pptx()

    print("Creating Table S1 (docx)...")
    create_table_s1(results)

    print("Creating BibTeX...")
    create_bibtex()

    print("Creating manuscript (LaTeX)...")
    tex_path = create_manuscript_tex(results)

    print("Creating Table S1 (LaTeX)...")
    s1_path = create_table_s1_latex(results)

    print("Creating cover letter (LaTeX)...")
    cl_path = create_cover_letter_tex()

    # Copy sn-jnl.cls and bst to manuscript dir
    for f in ["sn-jnl.cls"]:
        src = os.path.join(TEMPLATE_DIR, f)
        dst = os.path.join(OUT, f)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"  Copied {f} to {OUT}")

    bst_src = os.path.join(TEMPLATE_DIR, "bst", "sn-basic.bst")
    bst_dst = os.path.join(OUT, "sn-basic.bst")
    if os.path.exists(bst_src):
        shutil.copy2(bst_src, bst_dst)
        print(f"  Copied sn-basic.bst to {OUT}")

    # Compile LaTeX → PDF
    print("Compiling manuscript PDF...")
    pdf = _compile_latex(tex_path, OUT)
    if pdf:
        print(f"  PDF compiled: {pdf}")
    else:
        print("  WARNING: PDF compilation failed. Check LaTeX logs.")

    print("Compiling Table S1 PDF...")
    s1_pdf = _compile_latex(s1_path, OUT)
    if s1_pdf:
        print(f"  Table S1 PDF: {s1_pdf}")

    print("Compiling cover letter PDF...")
    cl_pdf = _compile_latex(cl_path, OUT)
    if cl_pdf:
        print(f"  Cover letter PDF: {cl_pdf}")

    print("\nAll outputs saved to:", OUT)


if __name__ == "__main__":
    main()
