# -*- coding: utf-8 -*-
"""شبیه‌سازی دیتاست معامله‌گران برای نمایش روش‌شناسی فصل چهارم (مأموریت ۴۳).

این ماژول یک نمونهٔ مصنوعی N=180 با ویژگی‌های روان‌سنجی واقع‌گرایانه تولید می‌کند:
- اضطراب صفتی (STAI-Y2): ۲۰ گویهٔ ۱-۴ → نمرهٔ کل ۲۰-۸۰
- طرحواره‌های ناسازگار اولیه (YSQ-SF): ۷۵ گویهٔ ۱-۶ → نمرهٔ کل ۷۵-۴۵۰
- تحمل ریسک/رفتار مالی پرخطر (Grable & Lytton): ۱۳ گویه ۱-۴ → نمرهٔ کل ۱۳-۵۲ (در این نمونه ۱۳-۴۵)

ساختار همبستگی از طریق یک عامل نهفت سه‌بعدی با ماتریس همبستگی کالیبره‌شده القا می‌شود
و بذر تصادفی ۲۴۶۴ پس از جست‌وجوی کالیبراسیون ثابت شده است تا ضرایب مشاهده‌شده
در بازهٔ هدف (r_AS≈0.45, r_AR≈0.35, r_SR≈0.40) قرار گیرند.

خروجی:
  data/simulated/simulated_traders_data.csv
  data/simulated/simulated_traders_data.xlsx
"""

from pathlib import Path

import numpy as np
import pandas as pd

# ---------- پیکربندهای ثابت (کالیبره‌شده) ----------
SEED = 2464                      # بذر ثابت‌شده پس از جست‌وجوی کالیبراسیون
N = 180                          # حجم نمونه
# ماتریس همبستگی متغیرهای نهفت (کمی بزرگ‌تر از هدف برای جبران تضعیف گویه‌ای)
LATENT_CORR = np.array([
    [1.00, 0.47, 0.37],          # اضطراب، طرحواره، رفتار مالی
    [0.47, 1.00, 0.42],
    [0.37, 0.42, 1.00],
])

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / 'data' / 'simulated'


def _make_items(rng, lam, mu, sd, k, lo, hi, z):
    """تولید پاسخ گویه‌ها: سهم مشترک از عامل نهفت + خطای مستقل، گردشده و محدودشده."""
    err = rng.normal(0.0, 1.0, (len(z), k))
    raw = mu + sd * (lam * z[:, None] + np.sqrt(1.0 - lam ** 2) * err)
    return np.clip(np.round(raw), lo, hi)


def generate_all(seed=SEED, n=N):
    """تولید دیتاست و ماتریس گویه‌ها (برای محاسبهٔ آلفا در ماژول تحلیل)."""
    rng = np.random.default_rng(seed)

    # عامل نهفت مشترک برای القا ساختار همبستگی
    z = rng.multivariate_normal([0, 0, 0], LATENT_CORR, n)

    # گویه‌های شبیه‌سازی‌شده
    anxiety_items = _make_items(rng, 0.60, 2.40, 0.80, 20, 1, 4, z[:, 0])
    schema_items = _make_items(rng, 0.55, 2.80, 0.90, 75, 1, 6, z[:, 1])
    risk_items = _make_items(rng, 0.55, 2.15, 0.75, 13, 1, 4, z[:, 2])

    df = pd.DataFrame({
        'ID': np.arange(1, n + 1),
        'Age': np.clip(np.round(rng.normal(32, 10, n)).astype(int), 18, 65),
        'Gender': rng.choice(['Male', 'Female'], n, p=[0.65, 0.35]),
        'Market': rng.choice(['Bourse', 'Crypto', 'Forex', 'Gold'], n,
                             p=[0.40, 0.25, 0.20, 0.15]),
        'Experience': np.clip(np.round(rng.gamma(2.2, 18, n)).astype(int) + 6, 6, 120),
        'Anxiety': anxiety_items.sum(axis=1),
        'Schema': schema_items.sum(axis=1),
        'Risk_Behavior': np.clip(risk_items.sum(axis=1), 13, 45),
    })
    items = {'Anxiety': anxiety_items, 'Schema': schema_items, 'Risk': risk_items}
    return df, items


def generate_dataset(seed=SEED, n=N):
    """تولید دیتاست کامل شامل جمعیت‌شناخت و نمره‌های کل سه متغیر."""
    df, _ = generate_all(seed, n)
    return df


def save_outputs(df):
    """ذخیرهٔ CSV و XLSX با نام ستون‌های تمیز."""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    csv_path = OUT_DIR / 'simulated_traders_data.csv'
    xlsx_path = OUT_DIR / 'simulated_traders_data.xlsx'
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    df.to_excel(xlsx_path, index=False, sheet_name='SimulatedTraders')
    return csv_path, xlsx_path


def main():
    df = generate_dataset()
    csv_path, xlsx_path = save_outputs(df)
    print(f'نمونه: {len(df)} | ذخیره شد: {csv_path.name} و {xlsx_path.name}')
    print('میانگین/انحراف: اضطراب %.1f/%.1f | طرحواره %.1f/%.1f | رفتار %.1f/%.1f' % (
        df['Anxiety'].mean(), df['Anxiety'].std(ddof=1),
        df['Schema'].mean(), df['Schema'].std(ddof=1),
        df['Risk_Behavior'].mean(), df['Risk_Behavior'].std(ddof=1)))
    print('همبستگی‌ها: A-S=%.3f A-R=%.3f S-R=%.3f' % (
        df['Anxiety'].corr(df['Schema']),
        df['Anxiety'].corr(df['Risk_Behavior']),
        df['Schema'].corr(df['Risk_Behavior'])))


if __name__ == '__main__':
    main()
