# -*- coding: utf-8 -*-
"""نسخهٔ واقع‌نمای شبیه‌سازی فصل چهارم — realism_v2 (مأموریت ۴۳-ب، فاز ۳).

تولید در سطح گویه با ویژگی‌های دادهٔ واقعی:
- بارهای گویه‌ای نامتساوی، گویه‌های معکوس STAI، علت‌مندی (acquiescence)،
  پاسخ‌های بی‌دقت پراکنده و چولگی جمعیت در اضطراب؛
- ۲۱۲ ردیف خام شامل ۴۱ ردیف نامعتبر با پرچم‌های قابل بازتولید
  (رضایت/ناقص/یکنواخت/سریع) — غربالگری از خودِ داده استخراج می‌شود؛
- زمان پاسخ لوگ‌نرمال با دنبالهٔ آهسته و رابطهٔ منطقی سابقه↔حجم معاملات.

ساختار نهفت: zS = .50·zA + خطا؛ zR = .30·zS + .10·zA + خطا (میانجی‌گری جزئی).
بذر ۷۷۸۸ یک‌بار تثبیت شده و برای دست‌کاری ضرایب جست‌وجو نشده است.

خروجی: data/simulated/realism_v2/simulated_traders_raw_v2.csv
"""

from pathlib import Path

import numpy as np
import pandas as pd

SEED = 7788
N_RAW = 212
N_CONSENT = 7      # انصراف از رضایت
N_MISSING = 14      # ناقص (بیش از ۳۰٪ گویه)
N_STRAIGHT = 11     # الگوی یکنواخت
N_FAST = 9          # زمان زیر ۱۲۰ ثانیه
# ۴۱ ردیف نامعتبر → ۱۷۱ ردیف معتبر (حذف ۱۹٫۳٪)

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / 'data' / 'simulated' / 'realism_v2'

K_A, K_S, K_R = 20, 75, 13
BETA_AS, BETA_SR, BETA_AR = 0.50, 0.22, 0.05  # مسیرهای نهفت استاندارد


def _rng():
    return np.random.default_rng(SEED)


def generate_raw():
    rng = _rng()
    n = N_RAW

    # ---- عوامل نهفت با چولگی واقع‌نما ----
    zA = rng.normal(0, 1, n)
    stress = rng.random(n) < 0.12          # زیرجمعیت پرتنش → دنبالهٔ راست
    zA = zA + stress * rng.uniform(0.55, 1.1, n)
    zA = (zA - zA.mean()) / zA.std()
    resS = rng.normal(0, 1, n)
    zS = BETA_AS * zA + np.sqrt(1 - BETA_AS ** 2) * resS
    rA_r = rng.normal(0, 1, n)
    R2 = np.array([BETA_AR, BETA_SR]) @ np.array([[1, BETA_AS], [BETA_AS, 1]]) @ np.array([BETA_AR, BETA_SR])
    zR = BETA_SR * zS + BETA_AR * zA + np.sqrt(max(1 - R2, 0.05)) * rA_r
    Z = np.column_stack([zA, zS, zR])

    # ---- علت‌مندی و خطای فردی ----
    acq = rng.normal(0.0, 0.22, n)

    # ---- تولید گویه‌ها: بار + علت‌مندی + خطا، برش با آستانه‌های تجمعی جهانی ----
    from scipy.stats import norm

    def _cut_probs(K, rng_j):
        """حاشیهٔ گزینه‌ای واقع‌نما (نامتوازن) برای هر مقیاس."""
        if K == 4:
            pi = np.array([0.22, 0.40, 0.24, 0.14])
        elif K == 6:
            pi = np.array([0.24, 0.30, 0.22, 0.13, 0.07, 0.04])
        else:
            pi = np.array([0.28, 0.31, 0.21, 0.12, 0.08])
        pi = np.clip(pi + rng_j.normal(0, 0.012, K), 0.02, None)
        pi /= pi.sum()
        cum = np.cumsum(pi)[:-1]
        return norm.ppf(cum)

    def gen_scale(idx, k, K, lam_lo, lam_hi, beta_g, dom_beta=None, dom_fac=None,
                  rev_mask=None, target_mean=None, seed_off=0):
        """یک مقیاس کامل: n×k گویه با کالیبراسیون سختی برای میانگین هدف."""
        r = np.random.default_rng(SEED + seed_off)
        z = Z[:, idx]
        lam = r.uniform(lam_lo, lam_hi, k)
        lin = np.empty((n, k))
        zdom = None
        if dom_beta is not None:
            dn = len(dom_beta)
            zdom = r.multivariate_normal(np.zeros(dn), np.eye(dn) - 0.04 * np.ones((dn, dn)), n)
        for jj in range(k):
            if zdom is not None:
                dj = jj // (k // dn)
                g = dom_beta[dj] * z + np.sqrt(dom_fac[dj]) * zdom[:, dj] \
                    + np.sqrt(max(0.02, 1 - dom_beta[dj] ** 2 - dom_fac[dj])) * r.normal(0, 1, n)
            else:
                g = lam[jj] * z + np.sqrt(max(0.02, 1 - lam[jj] ** 2)) * r.normal(0, 1, n)
            lin[:, jj] = g + 0.18 * acq
        rev = np.ones(k)
        if rev_mask is not None:
            rev = np.where(rev_mask == 1, -1.0, 1.0)
        lin = lin * rev
        # آستانه‌های تجمعی مشترک + سختی گویه (δ)؛ بای‌سکشن δ برای رسیدن به میانگین هدف
        delta0 = 0.0

        def items_at(delta):
            its = np.empty((n, k), dtype=int)
            for jj in range(k):
                pr = np.random.default_rng(SEED * 17 + seed_off * 101 + jj)
                cuts = _cut_probs(K, pr)
                delta = delta0 + delta + pr.uniform(-0.18, 0.18)
                its[:, jj] = np.digitize(lin[:, jj], cuts + delta * rev[jj] * -1) + 1
                its[:, jj] = np.clip(its[:, jj], 1, K)
            return its

        def score_mean(its):
            sc = its.astype(float)
            if rev_mask is not None:
                sc[:, rev_mask == 1] = (K + 1) - sc[:, rev_mask == 1]
            return sc.sum(1).mean(), sc

        lo, hi = -1.2, 1.2
        for _ in range(12):
            mid = (lo + hi) / 2
            m, _ = score_mean(items_at(mid))
            if m > target_mean:
                hi = mid
            else:
                lo = mid
        delta = (lo + hi) / 2
        items = items_at(delta)
        # بی‌دقتی پراکندهٔ ۳٪
        sp = np.random.default_rng(SEED * 5 + seed_off)
        sloppies = sp.random(n) < 0.03
        for ii in np.where(sloppies)[0]:
            js = sp.choice(k, size=int(sp.integers(1, 4)), replace=False)
            items[ii, js] = sp.integers(1, K + 1, len(js))
        return items

    revmask = np.zeros(K_A, dtype=int)
    revmask[::2] = 1                              # ۱۰ گویهٔ معکوس STAI
    a_items = gen_scale(0, K_A, 4, 0.56, 0.70, None, rev_mask=revmask,
                        target_mean=46.0, seed_off=1)
    s_items = gen_scale(1, K_S, 6, 0.62, 0.74, None,
                        dom_beta=[0.40, 0.36, 0.44, 0.34, 0.32],
                        dom_fac=[0.08, 0.15, 0.04, 0.12, 0.06],
                        target_mean=222.0, seed_off=2)
    r_items = gen_scale(2, K_R, 5, 0.50, 0.62, None, target_mean=30.0, seed_off=3)

    # مقیاس‌های هدف: α≈.87/.92/.80 — در تحلیل راستی‌آزمایی می‌شوند

    # ---- جمعیت‌شناختی نامتقارن ----
    age = np.clip(np.round(18 + rng.gamma(2.0, 8.2, n)), 18, 63).astype(int)
    gender = rng.choice(['Male', 'Female'], n, p=[0.63, 0.37])
    market = rng.choice(['Bourse', 'Crypto', 'Forex', 'Gold'], n, p=[0.44, 0.24, 0.18, 0.14])
    edu = rng.choice(['Diploma', 'Bachelor', 'Master', 'Other'], n, p=[0.19, 0.47, 0.26, 0.08])
    exp = np.clip(np.round(rng.gamma(2.1, 17.5, n) + 3), 3, 240).astype(int)
    trades = np.clip(np.round(rng.gamma(1.6, 3 + exp / 30.0)), 1, 90).astype(int)

    # ---- زمان پاسخ ----
    time = np.round(rng.lognormal(np.log(515), 0.5, n)).astype(int)
    time = np.clip(time, 125, 3600)

    rows = []
    for i in range(n):
        rec = {
            'ID': i + 1, 'Age': int(age[i]), 'Gender': gender[i], 'Market': market[i],
            'Education': edu[i], 'Experience': int(exp[i]), 'TradesPerMonth': int(trades[i]),
            'ResponseTimeSec': int(time[i]),
        }
        for j in range(K_A):
            rec[f'A{j + 1:02d}'] = int(a_items[i, j])
        for j in range(K_S):
            rec[f'S{j + 1:02d}'] = int(s_items[i, j])
        for j in range(K_R):
            rec[f'R{j + 1:02d}'] = int(r_items[i, j])
        rows.append(rec)
    df = pd.DataFrame(rows)

    # ---- آلودگی‌های هدفمند (ردیف‌های نامعتبر) ----
    perm = rng.permutation(n)
    c_i, m_i, s_i, f_i = perm[:N_CONSENT], perm[N_CONSENT:N_CONSENT + N_MISSING], \
        perm[N_CONSENT + N_MISSING:N_CONSENT + N_MISSING + N_STRAIGHT], \
        perm[-N_FAST:]
    df['ConsentOK'] = 1
    df.loc[c_i, 'ConsentOK'] = 0

    for i in m_i:  # ناقص: ۳۱–۵۸٪ گویه‌ها NaN
        frac = rng.uniform(0.31, 0.58)
        cols = df.columns[df.columns.str.match(r'^[ASR]\d\d$')].tolist()
        drop = rng.choice(cols, int(frac * len(cols)), replace=False)
        df.loc[i, drop] = np.nan
    df['MissingRate'] = 0.0
    item_cols = df.columns[df.columns.str.match(r'^[ASR]\d\d$')].tolist()
    df.loc[m_i, 'MissingRate'] = df.loc[m_i, item_cols].isna().mean(axis=1).round(3)

    for i in s_i:  # یکنواخت: همهٔ گویه‌های هر مقیاس یک عدد
        for pre, k in (('A', K_A), ('S', K_S), ('R', K_R)):
            df.loc[i, [f'{pre}{j + 1:02d}' for j in range(k)]] = int(rng.integers(2, 4))
    df['StraightLine'] = 0
    df.loc[s_i, 'StraightLine'] = 1

    df.loc[f_i, 'ResponseTimeSec'] = rng.integers(38, 118, len(f_i))
    df.loc[f_i, item_cols] = np.round(df.loc[f_i, item_cols].astype(float).values * 0 + 2)

    return df


def save():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    df = generate_raw()
    path = OUT_DIR / 'simulated_traders_raw_v2.csv'
    df.to_csv(path, index=False, encoding='utf-8-sig')
    print(f'دادهٔ خام v2 ذخیره شد: {path} | ردیف‌ها: {len(df)}')
    return df


if __name__ == '__main__':
    d = save()
    ok = (d['ConsentOK'] == 1) & (d['MissingRate'] <= 0.3) & (d['StraightLine'] == 0) & (d['ResponseTimeSec'] >= 120)
    print('ردیف معتبر پس از غربالگری:', int(ok.sum()))
