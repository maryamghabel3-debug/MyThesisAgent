# -*- coding: utf-8 -*-
"""Real-pipeline simulation for thesis chapters 4 and 5 (SIMULATED DATA ONLY).

Pipeline (deterministic; fixed SEED below):
  Phase 2: item-level Porsline-like raw CSV (consent/screening/D/A/S/R/duration)
  Phase 3: hierarchical screening -> cleaned CSV + scoring per final scoring keys
           (STAI 8 reversed, 20-80; YSQ 15 blocks, 75-450; Grable R01 reversed,
            R13 non-monotonic, 13-52; 12-item sensitivity variant)
  Phase 4: descriptives, reliability, assumptions, Pearson r, PROCESS-Model-4
           OLS mediation (5000 bootstrap), 13-vs-12 sensitivity re-run,
           saturated + constrained observed-variable path analysis (ML),
           stats.json + analysis_tables.json + path-diagram figure.

Scoring-key sources (read-only, side branches; frozen for this mission):
  feat/final-scoring-keys @ 18cd81c  -> final_scoring_keys.json (authoritative)
  feat/finalize-for-real-analysis @ bfcf99f -> spss_final_scoring.sps + checklist
Outputs (new files only):
  data/simulated/real_pipeline_v1/{raw_simulated_porsline_like.csv,
      cleaned_simulated_data.csv, stats.json, analysis_tables.json}
  output/drafts/chapter4_real_pipeline_fig1.png
"""
import json
import math
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, Rectangle
from scipy import stats as spstats
from numpy.linalg import slogdet

SEED = 20261010
N_RAW = 240
OUT_DIR = "data/simulated/real_pipeline_v1"
FIG_PATH = "output/drafts/chapter4_real_pipeline_fig1.png"
N_BOOT = 5000

# ---------------------------------------------------------------- design targets
# Latent correlation targets (attenuated slightly by sum-score reliability).
R_LAT = np.array([[1.00, 0.48, 0.28],
                  [0.48, 1.00, 0.32],
                  [0.28, 0.32, 1.00]])
FAST_CUTOFF_S = 120          # precedented cutoff: <120 s for 108 items = careless
INCOMPLETE_MISS_FRAC = 0.30  # >30% items missing -> incomplete

STAI_REV = [1, 6, 7, 10, 13, 14, 16, 19]          # 1-based, per final key
YSQ_BLOCKS = [(f"Schema_{k:02d}", list(range(5 * (k - 1) + 1, 5 * k + 1)))
              for k in range(1, 16)]
R13_MAP = {1: 4, 2: 1, 3: 2, 4: 3}                 # option -> score
R13_INV = {v: k for k, v in R13_MAP.items()}       # score -> option

A_COLS = [f"A{i:02d}" for i in range(1, 21)]
S_COLS = [f"S{i:02d}" for i in range(1, 76)]
R_COLS = [f"R{i:02d}" for i in range(1, 14)]
ITEM_COLS = A_COLS + S_COLS + R_COLS
DEMOG = ["D1", "D2", "D3", "D4", "D5", "D6", "D8", "D9", "D10"]  # no D7: real form has none
RAW_COLS = (["consent", "screening_age_18", "screening_real_trader",
             "screening_6_months"] + DEMOG + ITEM_COLS + ["duration_seconds"])

D2_OPTS = (["زن"] * 33 + ["مرد"] * 64 + ["ترجیح می‌دهم پاسخ ندهم"] * 3)
D3_OPTS = (["مجرد"] * 40 + ["متاهل"] * 52 + ["مطلقه/سایر"] * 8)
D4_OPTS = (["زیر دیپلم"] * 3 + ["دیپلم"] * 15 + ["فوق دیپلم"] * 10 +
           ["لیسانس"] * 42 + ["فوق لیسانس"] * 24 + ["دکتری"] * 6)
D5_OPTS = (["خانه‌دار"] * 8 + ["کارمند"] * 45 + ["مشاغل آزاد"] * 35 + ["سایر"] * 12)
D6_OPTS = (["بورس"] * 45 + ["رمزارز"] * 25 + ["فارکس"] * 15 +
           ["طلا و ارز"] * 8 + ["کالا"] * 3 + ["سایر"] * 4)
D9_OPTS = (["۱ الی ۱۰ معامله"] * 30 + ["۱۰ الی ۵۰"] * 45 + ["۵۰ به بالا"] * 25)
D10_OPTS = (["کمتر از ۵۰ میلیون تومان"] * 25 + ["۵۰ تا ۲۰۰ میلیون تومان"] * 30 +
            ["۲۰۰ تا ۵۰۰ میلیون تومان"] * 22 +
            ["۵۰۰ میلیون تا ۱ میلیارد تومان"] * 12 +
            ["بیش از ۱ میلیارد تومان"] * 8 +
            ["ترجیح می‌دهم پاسخ ندهم"] * 3)


def thresh_norm(z, cuts):
    """Map standard-ish continuous scores to 1..K via cut points."""
    return np.clip(np.digitize(z, cuts) + 1, 1, len(cuts) + 1)


SCHEMA_FA = ["محرومیت هیجانی (ED)", "رهاشدگی/بی‌ثباتی (AB)",
             "بی‌اعتمادی/بدرفتاری (MA)", "انزوای اجتماعی (SI)",
             "نقص/شرم (DS)", "شکست (FA)", "وابستگی/بی‌کفایتی (DI)",
             "آسیب‌پذیری (VH)", "درهم‌تنیدگی (EM)", "اطاعت (SB)",
             "ایثار (SS)", "بازداری هیجانی (EI)",
             "معیارهای سرسختانه (US)", "استحقاق/بزرگ‌منشی (ET)",
             "خویشتن‌داری ناکافی (IS)"]

_FA_TR = str.maketrans("0123456789.-", "۰۱۲۳۴۵۶۷۸۹٫−")


def to_fa_digits(s):
    return str(s).translate(_FA_TR)


def f0(x):
    return to_fa_digits(f"{int(round(float(x))):,}")


def f2(x):
    return to_fa_digits(f"{float(x):,.2f}")


def f3(x):
    return to_fa_digits(f"{float(x):.3f}")


def pfa(p):
    p = float(p)
    return f"p < {f3(0.001)}" if p < 0.001 else f"p = {f3(p)}"


def ordered_levels(series, order):
    counts = series.value_counts()
    return [(lv, int(counts.get(lv, 0))) for lv in order
            if lv in counts.index]


def demo_rows(demo_stats, n):
    rows = []
    for var, fa_name, levels in [
            ("D2", "جنسیت", None), ("D3", "وضعیت تأهل", None),
            ("D4", "سطح تحصیلات", None), ("D5", "وضعیت اشتغال", None),
            ("D6", "بازار اصلی فعالیت", None),
            ("D9", "تعداد معاملات ۶ ماهه اخیر", None),
            ("D10", "سرمایه تقریبی", None)]:
        for lv, c in demo_stats[var]["levels"]:
            rows.append([fa_name, lv, f0(c), f2(100 * c / n)])
    for var, fa_name in [("D1", "سن (سال)"), ("D8", "سابقه (ماه)")]:
        for lab, c in demo_stats[var]["groups"]:
            rows.append([fa_name, lab, f0(c), f2(100 * c / n)])
    return rows


def main():
    rng = np.random.default_rng(SEED)
    os.makedirs(OUT_DIR, exist_ok=True)

    # ------------------------------------------------- Phase 2: latent traits
    # Every screened-in respondent (valid + corrupted) gets latent traits first.
    n_pass = N_RAW - 8 - 10  # minus no-consent / screening-fail rows
    L = np.linalg.cholesky(R_LAT)
    z = rng.standard_normal((n_pass, 3)) @ L.T
    zA, zS, zR = z[:, 0], z[:, 1], z[:, 2]
    zBlock = rng.standard_normal((n_pass, 15))

    # ---- STAI items (scored direction), 4-point, mean ~2.15/item
    lamA = rng.uniform(0.48, 0.62, 20)
    eA = rng.standard_normal((n_pass, 20))
    contA = lamA * zA[:, None] + np.sqrt(np.maximum(1 - lamA ** 2, 0.05)) * eA
    cutsA = [-0.55, 0.25, 1.05]                      # -> mean ~2.15
    scoredA = thresh_norm(contA, cutsA)
    emitA = scoredA.copy()
    for j in STAI_REV:
        emitA[:, j - 1] = 5 - scoredA[:, j - 1]

    # ---- YSQ items, 6-point, mean ~2.68/item; block loadings vary 0.32..0.48
    lamG, lamB_vec = 0.35, rng.uniform(0.44, 0.54, 15)
    schema_off = rng.uniform(-0.30, 0.30, 15)
    emitS = np.zeros((n_pass, 75), dtype=int)
    for k in range(15):
        lb = lamB_vec[k]
        sd_e = math.sqrt(max(1 - lamG ** 2 - lb ** 2, 0.05))
        for t in range(5):
            cont = (lamG * zS + lb * zBlock[:, k] +
                    sd_e * rng.standard_normal(n_pass) + schema_off[k])
            emitS[:, 5 * k + t] = thresh_norm(cont, [-0.80, -0.10, 0.50, 1.10, 1.70])

    # ---- Grable scored values, 4-point, mean ~2.40/item
    lamR = rng.uniform(0.42, 0.58, 13)
    eR = rng.standard_normal((n_pass, 13))
    contR = lamR * zR[:, None] + np.sqrt(np.maximum(1 - lamR ** 2, 0.05)) * eR
    scoredR = thresh_norm(contR, [-0.75, 0.05, 0.85])
    emitR = scoredR.copy()
    emitR[:, 0] = 5 - scoredR[:, 0]                   # R01 reversed option
    emitR[:, 12] = np.array([R13_INV[s] for s in scoredR[:, 12]])  # R13 option

    # ---- demographics (valid-distribution; corrupted rows reuse then dropped)
    age = np.clip(19 + rng.gamma(3.0, 6.0, n_pass), 19, 65).round().astype(int)
    exp = np.clip(6 + rng.gamma(2.0, 18.0, n_pass), 6, 240).round().astype(int)
    pick = lambda opts: np.array([opts[i] for i in
                                  rng.integers(0, len(opts), n_pass)])
    dem = {"D1": age, "D2": pick(D2_OPTS), "D3": pick(D3_OPTS),
           "D4": pick(D4_OPTS), "D5": pick(D5_OPTS), "D6": pick(D6_OPTS),
           "D8": exp, "D9": pick(D9_OPTS), "D10": pick(D10_OPTS)}
    dur = np.clip(rng.lognormal(np.log(960), 0.45, n_pass), 300, 3600).round().astype(int)

    data = {"consent": np.ones(n_pass, dtype=int),
            "screening_age_18": np.ones(n_pass, dtype=int),
            "screening_real_trader": np.ones(n_pass, dtype=int),
            "screening_6_months": np.ones(n_pass, dtype=int)}
    for c in DEMOG:
        data[c] = dem[c]
    for j, c in enumerate(A_COLS):
        data[c] = emitA[:, j]
    for j, c in enumerate(S_COLS):
        data[c] = emitS[:, j]
    for j, c in enumerate(R_COLS):
        data[c] = emitR[:, j]
    data["duration_seconds"] = dur
    df = pd.DataFrame(data)

    # ---- plant invalid patterns (deterministic index blocks)
    idx = rng.permutation(n_pass)
    i_inc, i_str, i_fst = idx[0:17], idx[17:28], idx[28:37]
    # incomplete: abandonment tail inside the long YSQ section (some inside STAI)
    for r in i_inc:
        cut = rng.integers(8, 70)          # item position where dropout starts
        flat = ITEM_COLS[cut:]
        df.loc[r, flat] = np.nan
        if rng.random() < 0.4:             # some quit before finishing demographics
            df.loc[r, rng.choice(DEMOG, size=int(rng.integers(1, 4)),
                                 replace=False).tolist()] = np.nan
        df.loc[r, "duration_seconds"] = np.nan
    # straight-lining: constant raw option across all 108 items
    for r in i_str:
        df.loc[r, ITEM_COLS] = int(rng.choice([1, 2, 2, 2, 3]))
    # too fast: full data, unrealistically short duration
    df.loc[i_fst, "duration_seconds"] = rng.integers(45, 116, len(i_fst))
    # duplicates: exact copies of 3 valid rows (fresh durations)
    keep = np.ones(n_pass, bool)
    keep[np.concatenate([i_inc, i_str, i_fst])] = False
    src = rng.choice(np.where(keep)[0], 3, replace=False)
    dup = df.loc[src].copy()
    dup["duration_seconds"] = rng.integers(500, 1500, 3)
    df = pd.concat([df, dup], ignore_index=True)

    # ---- no-consent + screening-fail rows (form ends early -> rest empty)
    head = pd.DataFrame([{c: np.nan for c in RAW_COLS} for _ in range(18)])
    head.loc[0:7, "consent"] = 0
    head.loc[8:17, "consent"] = 1
    head.loc[8:10, "screening_age_18"] = 0     # under 18
    head.loc[8:10, ["screening_real_trader", "screening_6_months"]] = 1
    head.loc[11:14, "screening_real_trader"] = 0  # not a real trader
    head.loc[11:14, ["screening_age_18", "screening_6_months"]] = 1
    head.loc[15:17, "screening_6_months"] = 0  # < 6 months experience
    head.loc[15:17, ["screening_age_18", "screening_real_trader"]] = 1
    raw = pd.concat([df[RAW_COLS], head], ignore_index=True)
    raw = raw.sample(frac=1, random_state=SEED).reset_index(drop=True)
    raw_path = f"{OUT_DIR}/raw_simulated_porsline_like.csv"
    raw.to_csv(raw_path, index=False)
    assert len(raw) == 222 + 3 + 18, len(raw)  # pass-rows + dups + early-exit = 243
    print("raw rows:", len(raw))

    # --------------------------------------- Phase 3a: hierarchical screening
    n0 = len(raw)
    m_consent = raw["consent"] == 1
    m_screen = (m_consent & (raw["screening_age_18"] == 1) &
                (raw["screening_real_trader"] == 1) &
                (raw["screening_6_months"] == 1))
    miss_frac = raw[ITEM_COLS].isna().mean(axis=1)
    m_complete = m_screen & (miss_frac <= INCOMPLETE_MISS_FRAC)
    sd_items = raw[ITEM_COLS].std(axis=1, skipna=True)
    m_var = m_complete & (sd_items > 0)
    m_time = m_var & (raw["duration_seconds"] >= FAST_CUTOFF_S)
    tmp = raw[m_time].copy()
    m_dup = tmp.duplicated(subset=ITEM_COLS, keep="first")
    valid_idx = tmp.index[~m_dup.values]
    # expect-zero consistency gates on the surviving sample
    cons = raw.loc[valid_idx]
    gate_age = int(((pd.to_numeric(cons["D1"], errors="coerce") < 18)).sum())
    gate_exp = int(((pd.to_numeric(cons["D8"], errors="coerce") < 6)).sum())
    gate_tr0 = int((cons["D9"] == "هیچی").sum())
    flow = {"raw": int(n0), "no_consent": int((~m_consent).sum()),
            "screen_fail": int((m_consent & ~m_screen).sum()),
            "incomplete": int((m_screen & ~m_complete).sum()),
            "straight_line": int((m_complete & ~m_var).sum()),
            "too_fast": int((m_var & ~m_time).sum()),
            "duplicate": int(m_dup.sum()), "valid": int(len(valid_idx)),
            "gate_age_u18": gate_age, "gate_exp_lt6": gate_exp,
            "gate_zero_trades": gate_tr0}
    print("flow:", flow)
    clean = raw.loc[valid_idx, RAW_COLS].reset_index(drop=True)
    assert clean[ITEM_COLS].isna().sum().sum() == 0
    assert gate_age == 0 and gate_exp == 0 and gate_tr0 == 0
    clean_path = f"{OUT_DIR}/cleaned_simulated_data.csv"
    clean.to_csv(clean_path, index=False)

    # --------------------------------------- Phase 3b: scoring (mirror of SPSS)
    X = clean[A_COLS].to_numpy(dtype=float)
    Xr = X.copy()
    for j in STAI_REV:
        Xr[:, j - 1] = 5 - X[:, j - 1]
    STAI = Xr.sum(axis=1)
    S = clean[S_COLS].to_numpy(dtype=float)
    schemas = {name: S[:, [i - 1 for i in items]].sum(axis=1)
               for name, items in YSQ_BLOCKS}
    YSQ = S.sum(axis=1)
    R = clean[R_COLS].to_numpy(dtype=float)
    Rsc = R.copy()
    Rsc[:, 0] = 5 - R[:, 0]
    Rsc[:, 12] = np.array([R13_MAP[int(v)] for v in R[:, 12]])
    RISK = Rsc.sum(axis=1)
    RISK12 = Rsc[:, :12].sum(axis=1)
    # strict range gates: fix-before-report (raise = stop, never report bad)
    assert STAI.min() >= 20 and STAI.max() <= 80, (STAI.min(), STAI.max())
    for name, v in schemas.items():
        assert v.min() >= 5 and v.max() <= 30, (name, v.min(), v.max())
    assert YSQ.min() >= 75 and YSQ.max() <= 450
    assert RISK.min() >= 13 and RISK.max() <= 52
    assert RISK12.min() >= 12 and RISK12.max() <= 48
    n = len(clean)

    # ------------------------------------------------- Phase 4: statistics
    def desc(v):
        return {"n": int(n), "mean": float(np.mean(v)),
                "sd": float(np.std(v, ddof=1)), "min": float(np.min(v)),
                "max": float(np.max(v)),
                "median": float(np.median(v)),
                "q1": float(np.percentile(v, 25)),
                "q3": float(np.percentile(v, 75)),
                "skew": float(spstats.skew(v)), "kurt": float(spstats.kurtosis(v))}

    def alpha(mat):
        k = mat.shape[1]
        return float(k / (k - 1) *
                     (1 - np.trace(np.cov(mat, rowvar=False, ddof=1)) /
                      np.var(mat.sum(axis=1), ddof=1)))

    def pearson(a, b):
        r, p = spstats.pearsonr(a, b)
        return {"r": float(r), "p": float(p), "n": int(n)}

    def ols(y, *xs):
        Xm = np.column_stack([np.ones(n)] + [np.asarray(v) for v in xs])
        beta, res, _, _ = np.linalg.lstsq(Xm, y, rcond=None)
        yh = Xm @ beta
        e = y - yh
        dof = n - Xm.shape[1]
        s2 = float(e @ e / dof)
        V = s2 * np.linalg.inv(Xm.T @ Xm)
        se = np.sqrt(np.diag(V))
        t = beta / se
        p = 2 * spstats.t.sf(np.abs(t), dof)
        r2 = float(1 - e @ e / ((y - y.mean()) @ (y - y.mean())))
        return {"b": [float(v) for v in beta], "se": [float(v) for v in se],
                "t": [float(v) for v in t], "p": [float(v) for v in p],
                "r2": r2, "dof": int(dof), "resid": e, "yhat": yh}

    Sfull = np.column_stack([STAI, YSQ, RISK])
    covS = np.cov(Sfull, rowvar=False, ddof=1)
    stats = {"n_flow": flow, "duration_valid": desc(clean["duration_seconds"]
                                                   .to_numpy(dtype=float)),
             "fast_range": [float(raw.loc[m_var & ~m_time,
                                          "duration_seconds"].min()),
                            float(raw.loc[m_var & ~m_time,
                                          "duration_seconds"].max())],
             "desc": {"STAI_Total": desc(STAI), "YSQ_Total": desc(YSQ),
                      "Risk_Total": desc(RISK), "Risk_Total_12": desc(RISK12)},
             "alpha": {"STAI": alpha(Xr),
                       "YSQ_total": alpha(S),
                       "subscales": {name: alpha(S[:, [i - 1 for i in items]])
                                     for name, items in YSQ_BLOCKS},
                       "Risk_13": alpha(Rsc),
                       "Risk_12": alpha(Rsc[:, :12]),
                       "r_13_12": pearson(RISK, RISK12)["r"]},
             "r": {"H2_X_Y": pearson(STAI, RISK),
                   "H3_X_M": pearson(STAI, YSQ),
                   "H4_M_Y": pearson(YSQ, RISK)}}

    # assumptions
    maha = np.array([float((row - Sfull.mean(axis=0)) @
                           np.linalg.inv(covS) @ (row - Sfull.mean(axis=0)))
                     for row in Sfull])
    maha_crit = float(spstats.chi2.ppf(0.999, 3))
    reg_full = ols(RISK, STAI, YSQ)
    e = reg_full["resid"]
    dw = float(np.sum(np.diff(e) ** 2) / np.sum(e ** 2))
    r_xm = stats["r"]["H3_X_M"]["r"]
    vif = float(1 / (1 - r_xm ** 2))
    stats["assumptions"] = {"mahalanobis_max": float(maha.max()),
                            "mahalanobis_crit_chi2_3_001": maha_crit,
                            "mahalanobis_n_exceed": int((maha > maha_crit).sum()),
                            "VIF_XM": vif, "tolerance_XM": float(1 / vif),
                            "durbin_watson_Y_XM": dw}

    # PROCESS-Model-4-equivalent OLS system
    reg_a = ols(YSQ, STAI)          # M ~ X : path a
    reg_c = ols(RISK, STAI)         # Y ~ X : total c
    a, b = reg_a["b"][1], reg_full["b"][2]
    cp, c = reg_full["b"][1], reg_c["b"][1]
    indirect = a * b
    brng = np.random.default_rng(SEED + 1)
    boots = np.empty(N_BOOT)
    for i in range(N_BOOT):
        bi = brng.integers(0, n, n)
        Xb, Mb, Yb = STAI[bi], YSQ[bi], RISK[bi]
        ab = np.polyfit(Xb, Mb, 1)[0]
        H = np.column_stack([np.ones(n), Xb, Mb])
        bb = np.linalg.lstsq(H, Yb, rcond=None)[0][2]
        boots[i] = ab * bb
    ci_lo, ci_hi = float(np.percentile(boots, 2.5)), float(np.percentile(boots, 97.5))
    z0 = float(spstats.norm.ppf(np.mean(boots < indirect)))
    bc_lo = float(np.percentile(boots, 100 * spstats.norm.cdf(2 * z0 - 1.96)))
    bc_hi = float(np.percentile(boots, 100 * spstats.norm.cdf(2 * z0 + 1.96)))
    stats["mediation"] = {
        "a": {k: reg_a[k][1] for k in ("b", "se", "t", "p")},
        "b": {k: reg_full[k][2] for k in ("b", "se", "t", "p")},
        "c_total": {k: reg_c[k][1] for k in ("b", "se", "t", "p")},
        "c_prime": {k: reg_full[k][1] for k in ("b", "se", "t", "p")},
        "R2_M_on_X": reg_a["r2"], "R2_Y_on_XM": reg_full["r2"],
        "R2_Y_on_X": reg_c["r2"], "indirect_ab": float(indirect),
        "boot_n": N_BOOT, "boot_CI95_percentile": [ci_lo, ci_hi],
        "boot_CI95_BC": [bc_lo, bc_hi]}

    # sensitivity: re-run the risk-side system with the 12-item score
    reg_c12 = ols(RISK12, STAI)
    reg_f12 = ols(RISK12, STAI, YSQ)
    a12, b12, cp12 = a, reg_f12["b"][2], reg_f12["b"][1]
    ind12 = a12 * b12
    boots12 = np.empty(N_BOOT)
    for i in range(N_BOOT):
        bi = brng.integers(0, n, n)
        Xb, Mb, Yb = STAI[bi], YSQ[bi], RISK12[bi]
        ab = np.polyfit(Xb, Mb, 1)[0]
        H = np.column_stack([np.ones(n), Xb, Mb])
        boots12[i] = ab * np.linalg.lstsq(H, Yb, rcond=None)[0][2]
    stats["sensitivity_12"] = {
        "r_X_Y12": pearson(STAI, RISK12), "r_M_Y12": pearson(YSQ, RISK12),
        "c_total12": reg_c12["b"][1], "c_prime12": cp12,
        "c_prime12_p": float(reg_f12["p"][1]),
        "indirect12": float(ind12),
        "boot_CI95_percentile12": [float(np.percentile(boots12, 2.5)),
                                     float(np.percentile(boots12, 97.5))],
        "R2_Y12_on_XM": reg_f12["r2"]}

    # observed-variable path analysis (ML). Saturated model (df=0): no fit test.
    sdx = stats["desc"]["STAI_Total"]["sd"]
    sdm = stats["desc"]["YSQ_Total"]["sd"]
    sdy = stats["desc"]["Risk_Total"]["sd"]
    stats["saturated_std"] = {"a_std": float(a * sdx / sdm),
                              "b_std": float(b * sdm / sdy),
                              "c_prime_std": float(cp * sdx / sdy),
                              "df": 0}
    # constrained full-mediation model (c' = 0): recursive ML = OLS per equation
    a_c = float(np.polyfit(STAI, YSQ, 1)[0])
    b_c = float(np.polyfit(YSQ, RISK, 1)[0])
    vx = float(np.var(STAI, ddof=1))
    psi_m = float(np.var(YSQ - np.polyval([a_c, np.polyfit(STAI, YSQ, 1)[1]],
                                          STAI), ddof=1))
    Mh = a_c * STAI + np.polyfit(STAI, YSQ, 1)[1]
    psi_y = float(np.var(RISK - (b_c * YSQ +
                                 np.polyfit(YSQ, RISK, 1)[1]), ddof=1))
    vm = a_c ** 2 * vx + psi_m
    Sig = np.array([[vx, a_c * vx, a_c * b_c * vx],
                    [a_c * vx, vm, b_c * vm],
                    [a_c * b_c * vx, b_c * vm, b_c ** 2 * vm + psi_y]])
    Sinv = np.linalg.inv(Sig)
    Fml = float(np.trace(covS @ Sinv) - slogdet(covS @ Sinv)[1] - 3)
    chi2, df_c = (n - 1) * Fml, 1
    Sig0 = np.diag(np.diag(covS))
    F0 = float(np.trace(covS @ np.linalg.inv(Sig0)) -
               slogdet(covS @ np.linalg.inv(Sig0))[1] - 3)
    chi0, df0 = (n - 1) * F0, 3
    cfi = float(1 - max(chi2 - df_c, 0) / max(chi0 - df0, 0))
    tli = float((chi0 / df0 - chi2 / df_c) / (chi0 / df0 - 1))
    rmsea = float(math.sqrt(max(chi2 - df_c, 0) / ((n - 1) * df_c)))
    R = np.corrcoef(Sfull, rowvar=False)
    D = np.sqrt(np.diag(Sig))
    Rs = Sig / np.outer(D, D)
    iu = np.triu_indices(3)
    srmr = float(math.sqrt(np.mean((R[iu] - Rs[iu]) ** 2)))
    stats["constrained_full_mediation"] = {
        "a": a_c, "b": b_c, "chi2": float(chi2), "df": df_c,
        "p": float(spstats.chi2.sf(chi2, df_c)), "CFI": cfi, "TLI": tli,
        "RMSEA": rmsea, "SRMR": float(srmr),
        "null_chi2": float(chi0), "null_df": df0}
    def grp(v, cuts, labels):
        return [(lab, int(((v >= lo) & (v < hi)).sum()))
                for lab, lo, hi in
                [(labels[i], cuts[i], cuts[i + 1]) for i in range(len(labels))]]

    age = pd.to_numeric(clean["D1"]).to_numpy(dtype=float)
    expm = pd.to_numeric(clean["D8"]).to_numpy(dtype=float)
    demo_stats = {
        "D1": {**desc(age), "groups": grp(age, [18, 30, 40, 50, 100],
                                          ["۱۸ تا ۲۹", "۳۰ تا ۳۹",
                                           "۴۰ تا ۴۹", "۵۰ به بالا"])},
        "D8": {**desc(expm), "groups": grp(expm, [6, 12, 36, 72, 10 ** 9],
                                           ["۶ تا ۱۱", "۱۲ تا ۳۵",
                                            "۳۶ تا ۷۱", "۷۲ به بالا"])},
        "D2": {"levels": ordered_levels(clean["D2"], ["زن", "مرد",
                        "ترجیح می‌دهم پاسخ ندهم"])},
        "D3": {"levels": ordered_levels(clean["D3"], ["مجرد", "متاهل",
                                                     "مطلقه/سایر"])},
        "D4": {"levels": ordered_levels(clean["D4"], ["زیر دیپلم", "دیپلم",
                        "فوق دیپلم", "لیسانس", "فوق لیسانس", "دکتری"])},
        "D5": {"levels": ordered_levels(clean["D5"], ["خانه‌دار", "کارمند",
                                                 "مشاغل آزاد", "سایر"])},
        "D6": {"levels": ordered_levels(clean["D6"], ["بورس", "رمزارز",
                        "فارکس", "طلا و ارز", "کالا", "سایر"])},
        "D9": {"levels": ordered_levels(clean["D9"], ["هیچی",
                        "۱ الی ۱۰ معامله", "۱۰ الی ۵۰", "۵۰ به بالا"])},
        "D10": {"levels": ordered_levels(clean["D10"],
                        ["کمتر از ۵۰ میلیون تومان",
                         "۵۰ تا ۲۰۰ میلیون تومان",
                         "۲۰۰ تا ۵۰۰ میلیون تومان",
                         "۵۰۰ میلیون تا ۱ میلیارد تومان",
                         "بیش از ۱ میلیارد تومان",
                         "ترجیح می‌دهم پاسخ ندهم"])}}
    stats["demographics"] = demo_stats

    with open(f"{OUT_DIR}/stats.json", "w") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    # --------------------------------- analysis_tables.json (chapter-4 payloads)
    def T(tid, title, columns, rows, note=""):
        return {"id": tid, "title": title, "columns": columns, "rows": rows,
                "note": note}

    def star(p):
        return "***" if p < 0.001 else ("**" if p < 0.01 else
                                       ("*" if p < 0.05 else ("†" if p < 0.10
                                                              else "")))

    r = stats["r"]
    med = stats["mediation"]
    dsc = stats["desc"]
    tables = [
        T("T4-1", "روند غربالگری و ریزش پاسخ‌ها",
          ["مرحله", "تعداد حذف‌شده", "باقی‌مانده"],
          [["خام اولیه", "—", f0(flow['raw'])],
           ["رضایت آگاهانه (بدون رضایت)", f0(flow['no_consent']),
            f0(flow['raw'] - flow['no_consent'])],
           ["غربالگری ورود (سن/معامله‌گر واقعی/سابقه)", f0(flow['screen_fail']),
            f0(flow['raw'] - flow['no_consent'] - flow['screen_fail'])],
           ["ناقص‌بودن (بیش از ۳۰٪ بی‌پاسخ)", f0(flow['incomplete']),
            f0(flow['raw'] - flow['no_consent'] - flow['screen_fail'] -
               flow['incomplete'])],
           ["الگوی ثابت پاسخ (واریانس صفر)", f0(flow['straight_line']),
            f0(flow['raw'] - flow['no_consent'] - flow['screen_fail'] -
               flow['incomplete'] - flow['straight_line'])],
           ["زمان نامتعارف کوتاه (کمتر از ۱۲۰ ثانیه)",
            f0(flow['too_fast']),
            f0(flow['raw'] - flow['no_consent'] - flow['screen_fail'] -
               flow['incomplete'] - flow['straight_line'] -
               flow['too_fast'])],
           ["تکراری (بردار پاسخ یکسان)", f0(flow['duplicate']),
            f0(flow['valid'])],
           ["نمونه نهایی معتبر", "—", f0(flow['valid'])]]),
        T("T4-2", "توزیع متغیرهای جمعیت‌شناختی و معاملاتی نمونه نهایی",
          ["متغیر", "سطح", "تعداد", "درصد"],
          demo_rows(demo_stats, n)),
        T("T4-3", "شاخص‌های توصیفی نمره‌های کل",
          ["متغیر", "میانگین", "انحراف معیار", "کمینه", "بیشینه",
           "چولگی", "کشیدگی", "دامنه نظری", "دامنه مشاهده‌شده"],
          [["اضطراب صفتی", f2(dsc['STAI_Total']['mean']),
            f2(dsc['STAI_Total']['sd']), f0(dsc['STAI_Total']['min']),
            f0(dsc['STAI_Total']['max']), f2(dsc['STAI_Total']['skew']),
            f2(dsc['STAI_Total']['kurt']), "۲۰ تا ۸۰",
            f"{f0(dsc['STAI_Total']['min'])} تا {f0(dsc['STAI_Total']['max'])}"],
           ["طرحواره‌های ناسازگار اولیه (نمره کل)", f2(dsc['YSQ_Total']['mean']),
            f2(dsc['YSQ_Total']['sd']), f0(dsc['YSQ_Total']['min']),
            f0(dsc['YSQ_Total']['max']), f2(dsc['YSQ_Total']['skew']),
            f2(dsc['YSQ_Total']['kurt']), "۷۵ تا ۴۵۰",
            f"{f0(dsc['YSQ_Total']['min'])} تا {f0(dsc['YSQ_Total']['max'])}"],
           ["تحمل ریسک مالی (۱۳ سؤالی)", f2(dsc['Risk_Total']['mean']),
            f2(dsc['Risk_Total']['sd']), f0(dsc['Risk_Total']['min']),
            f0(dsc['Risk_Total']['max']), f2(dsc['Risk_Total']['skew']),
            f2(dsc['Risk_Total']['kurt']), "۱۳ تا ۵۲",
            f"{f0(dsc['Risk_Total']['min'])} تا {f0(dsc['Risk_Total']['max'])}"]]),
        T("T4-4", "پایایی ابزارها و خرده‌مقیاس‌ها (آلفای کرونباخ)",
          ["مقیاس", "تعداد گویه", "آلفا"],
          [["اضطراب صفتی اسپیلبرگر", "۲۰", f2(stats['alpha']['STAI'])],
           ["فرم کوتاه یانگ (نمره کل)", "۷۵",
            f2(stats['alpha']['YSQ_total'])]]
          + [[SCHEMA_FA[i], "۵", f2(stats['alpha']['subscales'][name])]
             for i, (name, _) in enumerate(YSQ_BLOCKS)]
          + [["تحمل ریسک گرابل و لایتون (۱۳ سؤالی)", "۱۳",
              f2(stats['alpha']['Risk_13'])],
             ["تحمل ریسک (۱۲ سؤالی، بدون سؤال سیزدهم)", "۱۲",
              f2(stats['alpha']['Risk_12'])]],
          note="همبستگی نسخه ۱۳ و ۱۲ سؤالی تحمل ریسک: "
               f"{f2(stats['alpha']['r_13_12'])}."),
        T("T4-5", "ماتریس همبستگی پیرسون متغیرهای پژوهش",
          ["", "۱", "۲", "۳"],
          [["۱. اضطراب صفتی", "۱", "", ""],
           ["۲. طرحواره‌های ناسازگار اولیه",
            f"{f3(r['H3_X_M']['r'])}{star(r['H3_X_M']['p'])}", "۱", ""],
           ["۳. تحمل ریسک مالی",
            f"{f3(r['H2_X_Y']['r'])}{star(r['H2_X_Y']['p'])}",
            f"{f3(r['H4_M_Y']['r'])}{star(r['H4_M_Y']['p'])}", "۱"]],
          note=f"n = {n}. "
               "† p < ۰٫۱۰؛ * p < ۰٫۰۵؛ ** p < ۰٫۰۱؛ *** p < ۰٫۰۰۱."),
        T("T4-6", "نتایج تحلیل میانجی‌گری (الگوی شماره ۴، خودگردان‌سازی ۵۰۰۰ نمونه‌ای)",
          ["اثر", "ضریب", "خطای معیار", "آماره", "p", "فاصله اطمینان ۹۵٪"],
          [["مسیر a (اضطراب → طرحواره‌ها)", f3(med['a']['b']),
            f3(med['a']['se']), f3(med['a']['t']),
            pfa(med['a']['p']), ""],
           ["مسیر b (طرحواره‌ها → ریسک، با کنترل اضطراب)", f3(med['b']['b']),
            f3(med['b']['se']), f3(med['b']['t']), pfa(med['b']['p']), ""],
           ["اثر کل c (اضطراب → ریسک)", f3(med['c_total']['b']),
            f3(med['c_total']['se']), f3(med['c_total']['t']),
            pfa(med['c_total']['p']), ""],
           ["اثر مستقیم c′ (با کنترل میانجی)", f3(med['c_prime']['b']),
            f3(med['c_prime']['se']), f3(med['c_prime']['t']),
            pfa(med['c_prime']['p']), ""],
           ["اثر غیرمستقیم (a×b)",
            f3(med['indirect_ab']), "—", "—", "—",
            f"[{f3(med['boot_CI95_percentile'][0])}، "
            f"{f3(med['boot_CI95_percentile'][1])}]"]],
          note=f"R² مدل میانجی = {f3(med['R2_M_on_X'])}؛ "
               f"R² مدل پیامد = {f3(med['R2_Y_on_XM'])}؛ "
               f"R² اثر کل = {f3(med['R2_Y_on_X'])}. "
               "فاصله اطمینان اثر غیرمستقیم از نوع صدکی است."),
        T("T4-7", "شاخص‌های مدل مقید میانجی‌گری کامل (مکمل، df=۱)",
          ["شاخص", "مقدار"],
          [["خی‌دو", f2(stats['constrained_full_mediation']['chi2'])],
           ["درجه آزادی", "۱"],
           ["p", pfa(stats['constrained_full_mediation']['p'])],
           ["CFI", f3(stats['constrained_full_mediation']['CFI'])],
           ["TLI", f3(stats['constrained_full_mediation']['TLI'])],
           ["RMSEA", f3(stats['constrained_full_mediation']['RMSEA'])],
           ["SRMR", f3(stats['constrained_full_mediation']['SRMR'])]],
          note="مدل مقید است و تفسیر آن تکمیلی است؛ آزمون فرضیه اول بر "
               "مبنای فاصله اطمینان خودگردان‌سازی انجام شده است."),
    ]
    with open(f"{OUT_DIR}/analysis_tables.json", "w") as f:
        json.dump(tables, f, ensure_ascii=False, indent=2)

    # ------------------------------------------------------- figure (v2 style)
    try:
        from arabic_reshaper import reshape as _reshape
        from bidi.algorithm import get_display as _bidi

        def fa(s):
            return _bidi(_reshape(str(s)))
    except ImportError:
        def fa(s):
            return str(s)
    fig, ax = plt.subplots(figsize=(12, 7))
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 7)
    ax.axis("off")
    boxes = {"X": (0.3, 2.2, 3.0, 1.0), "M": (4.5, 4.9, 3.0, 1.0),
             "Y": (8.7, 2.2, 3.0, 1.0)}
    labels = {"X": "اضطراب صفتی", "M": "طرحواره‌های ناسازگار اولیه",
              "Y": "رفتارهای مالی پرخطر"}
    for k, (x, y, w, h) in boxes.items():
        ax.add_patch(Rectangle((x, y), w, h, fill=False, lw=2))
        ax.text(x + w / 2, y + h / 2, fa(labels[k]), ha="center",
                va="center", fontsize=13)
    arrows = [(3.3, 2.8, 4.6, 5.0, f"a = {med['a']['b']:.3f}{star(med['a']['p'])}",
               3.0, 4.1),
              (7.6, 5.0, 8.9, 2.8, f"b = {med['b']['b']:.3f}{star(med['b']['p'])}",
               8.4, 4.1),
              (3.3, 2.7, 8.7, 2.7,
               f"c' = {med['c_prime']['b']:.3f}{star(med['c_prime']['p'])}",
               6.0, 2.95)]
    for x1, y1, x2, y2, lab, tx, ty in arrows:
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2),
                                     arrowstyle="-|>", mutation_scale=18, lw=1.6))
        ax.text(tx, ty, lab, ha="center", va="center", fontsize=11)
    lo, hi = med["boot_CI95_percentile"]
    ax.text(6.0, 0.9,
            fa(f"CI95% [{lo:.3f}, {hi:.3f}] ، {med['indirect_ab']:.3f} = اثر غیرمستقیم"),
            ha="center", va="center", fontsize=12)
    fig.tight_layout()
    fig.savefig(FIG_PATH, dpi=150)
    print("wrote", FIG_PATH)

    # ------------------------------------------------------- scenario report
    g = {"H2 r in [.25,.36]": 0.25 <= r["H2_X_Y"]["r"] <= 0.36,
         "H3 r in [.42,.52]": 0.42 <= r["H3_X_M"]["r"] <= 0.52,
         "H4 r in [.25,.36]": 0.25 <= r["H4_M_Y"]["r"] <= 0.36,
         "c' sig or near (p<.10)": med["c_prime"]["p"] < 0.10,
         "indirect CI != 0": not (lo <= 0 <= hi),
         "alpha STAI>=.80": stats["alpha"]["STAI"] >= 0.80,
         "alpha YSQ>=.88": stats["alpha"]["YSQ_total"] >= 0.88,
         "alpha R13>=.72": stats["alpha"]["Risk_13"] >= 0.72,
         "subscales>=.60": min(stats["alpha"]["subscales"].values()) >= 0.60,
         "n valid 170..185": 170 <= flow["valid"] <= 185}
    print("scenario gates:", "ALL PASS" if all(g.values()) else
          [k for k, v in g.items() if not v])
    print(json.dumps({"r": r, "med_cp_p": med["c_prime"]["p"],
                      "indirect_CI": [lo, hi]}, ensure_ascii=False))


main()
