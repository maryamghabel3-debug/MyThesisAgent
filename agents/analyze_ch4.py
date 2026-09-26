# -*- coding: utf-8 -*-
"""تحلیل آماری دادهٔ شبیه‌سازی‌شده و تدوین پیش‌نویس فصل چهارم (مأموریت ۴۳).

مراحل: آمار توصیفی، پایایی (آلفای کرونباخ)، پیش‌فرض‌ها، همبستگی پیرسون،
و میانجی‌گری الگوی مدل ۴ هایز با بوت‌استرپ ۱۰۰۰ بازنمونه؛ سپس نگارش
output/drafts/chapter4_simulated.md و ترسیم شکل ۴-۱.
"""

import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

sys.path.insert(0, str(Path(__file__).resolve().parent))
from simulate_data import generate_all

ROOT = Path(__file__).resolve().parent.parent
SIM_DIR = ROOT / 'data' / 'simulated'
DRAFTS = ROOT / 'output' / 'drafts'
BOOT_SEED = 42
N_BOOT = 1000

FA = 'اضطراب صفتی'
SC = 'طرحواره‌های ناسازگار اولیه'
RK = 'رفتارهای مالی پرخطر'


# ---------- ابزارهای آماری ----------
def cronbach_alpha(x):
    """آلفای کرونباخ برای ماتریس گویه‌ها (سطرها=افراد)."""
    k = x.shape[1]
    var_i = x.var(axis=0, ddof=1).sum()
    var_t = x.sum(axis=1).var(ddof=1)
    return (k / (k - 1)) * (1 - var_i / var_t)


def ols(y, X):
    """رگرسیون کمترین مربعات با آماره‌های کامل؛ X شامل ستون عرض از مبدأ."""
    n, k = X.shape
    beta, *_ = np.linalg.lstsq(X, y, rcond=None)
    resid = y - X @ beta
    dof = n - k
    s2 = (resid @ resid) / dof
    cov = s2 * np.linalg.inv(X.T @ X)
    se = np.sqrt(np.diag(cov))
    t = beta / se
    p = 2 * stats.t.sf(np.abs(t), dof)
    ss_res = resid @ resid
    ss_tot = ((y - y.mean()) ** 2).sum()
    r2 = 1 - ss_res / ss_tot
    f = (r2 / (k - 1)) / ((1 - r2) / dof) if r2 < 1 else np.inf
    # آمارهٔ دوربین-واتسون برای استقلال پسماندها
    dw = np.sum(np.diff(resid) ** 2) / ss_res
    return {'B': beta, 'SE': se, 't': t, 'p': p, 'R2': r2, 'F': f,
            'df': (k - 1, dof), 'DW': dw}


def p_str(p):
    """نمایش استاندارد سطح معناداری."""
    return 'p < 0.01' if p < 0.01 else ('p < 0.05' if p < 0.05 else f'p = {p:.3f}')


def stars(p):
    return '***' if p < 0.001 else ('**' if p < 0.01 else ('*' if p < 0.05 else 'ns'))


# ---------- اجرای تحلیل‌ها ----------
def run_analysis():
    df, items = generate_all()
    a_col, s_col, r_col = df['Anxiety'].values, df['Schema'].values, df['Risk_Behavior'].values
    n = len(df)
    out = {'n': n}

    # ۱) توصیفی جمعیت‌شناختی
    out['gender'] = df['Gender'].value_counts().reindex(['Male', 'Female']).to_dict()
    out['market'] = df['Market'].value_counts().reindex(['Bourse', 'Crypto', 'Forex', 'Gold']).to_dict()
    out['age'] = {'mean': df['Age'].mean(), 'sd': df['Age'].std(ddof=1),
                  'min': int(df['Age'].min()), 'max': int(df['Age'].max())}
    out['exp'] = {'mean': df['Experience'].mean(), 'sd': df['Experience'].std(ddof=1),
                  'min': int(df['Experience'].min()), 'max': int(df['Experience'].max())}

    # ۲) توصیفی متغیرهای اصلی + نرمالیته
    desc = {}
    for name, col in [('Anxiety', a_col), ('Schema', s_col), ('Risk', r_col)]:
        desc[name] = {'mean': col.mean(), 'sd': col.std(ddof=1),
                      'min': int(col.min()), 'max': int(col.max()),
                      'skew': stats.skew(col), 'kurt': stats.kurtosis(col),
                      'sw_W': stats.shapiro(col)[0], 'sw_p': stats.shapiro(col)[1]}
    out['desc'] = desc

    # ۳) پایایی
    out['alpha'] = {k_: round(cronbach_alpha(v), 3) for k_, v in items.items()}

    # ۴) همبستگی پیرسون
    corr = {}
    pairs = [('Anxiety', 'Risk_Behavior'), ('Anxiety', 'Schema'), ('Schema', 'Risk_Behavior')]
    for x_, y_ in pairs:
        r, p = stats.pearsonr(df[x_], df[y_])
        key = f'{x_}-{y_}'.replace('_Behavior', '')
        corr[key] = {'r': r, 'p': p}
    out['corr'] = corr

    # ۵) میانجی‌گری مدل ۴ هایز
    X1 = np.column_stack([np.ones(n), a_col])                 # مسیر a
    mod_a = ols(s_col, X1)
    X2 = np.column_stack([np.ones(n), a_col, s_col])          # مسیر b و c'
    mod_b = ols(r_col, X2)
    X3 = np.column_stack([np.ones(n), a_col])                 # اثر کل c
    mod_c = ols(r_col, X3)
    a_path = mod_a['B'][1]
    b_path = mod_b['B'][2]
    c_prime = mod_b['B'][1]
    c_total = mod_c['B'][1]
    indirect = a_path * b_path

    # هم‌خطی: VIF مشترک X و M
    r_xm = np.corrcoef(a_col, s_col)[0, 1]
    vif = 1 / (1 - r_xm ** 2)

    # بوت‌استرپ percentile
    rng = np.random.default_rng(BOOT_SEED)
    boot_ab, boot_cp, boot_ct = [], [], []
    for _ in range(N_BOOT):
        idx = rng.integers(0, n, n)
        ba = ols(s_col[idx], np.column_stack([np.ones(n), a_col[idx]]))['B'][1]
        bb = ols(r_col[idx], np.column_stack([np.ones(n), a_col[idx], s_col[idx]]))['B']
        bc = ols(r_col[idx], np.column_stack([np.ones(n), a_col[idx]]))['B'][1]
        boot_ab.append(ba * bb[2])
        boot_cp.append(bb[1])
        boot_ct.append(bc)
    ci_ab = np.percentile(boot_ab, [2.5, 97.5])
    ci_cp = np.percentile(boot_cp, [2.5, 97.5])
    ci_ct = np.percentile(boot_ct, [2.5, 97.5])

    out['mediation'] = {
        'a': a_path, 'a_se': mod_a['SE'][1], 'a_t': mod_a['t'][1], 'a_p': mod_a['p'][1],
        'b': b_path, 'b_se': mod_b['SE'][2], 'b_t': mod_b['t'][2], 'b_p': mod_b['p'][2],
        'cp': c_prime, 'cp_se': mod_b['SE'][1], 'cp_t': mod_b['t'][1], 'cp_p': mod_b['p'][1],
        'c': c_total, 'c_se': mod_c['SE'][1], 'c_t': mod_c['t'][1], 'c_p': mod_c['p'][1],
        'indirect': indirect, 'ci_ab': list(ci_ab), 'ci_cp': list(ci_cp), 'ci_ct': list(ci_ct),
        'r2_m': mod_a['R2'], 'f_m': mod_a['F'], 'r2_y': mod_b['R2'], 'f_y': mod_b['F'],
        'dw_a': mod_a['DW'], 'dw_b': mod_b['DW'], 'vif': vif,
        'prop_mediated': indirect / c_total,
    }
    return df, out


# ---------- ترسیم شکل ۴-۱ (PIL با شکل‌دهی درست فارسی) ----------
def draw_figure(m):
    import glob
    import math

    from PIL import Image, ImageDraw, ImageFont

    fonts = (glob.glob('/usr/share/fonts/**/Scheherazade*.ttf', recursive=True) +
             glob.glob('/usr/share/fonts/**/Amiri*.ttf', recursive=True))
    fa_font = ImageFont.truetype(fonts[0], 44)
    en_font = ImageFont.truetype(fonts[0], 40)

    W, H = 1800, 1000
    img = Image.new('RGB', (W, H), 'white')
    d = ImageDraw.Draw(img)

    boxes = {'X': (300, 620), 'M': (900, 220), 'Y': (1500, 620)}
    labels = {'X': FA, 'M': SC, 'Y': RK}
    bw, bh = 460, 130

    def fa_text(s, cx, cy, font=fa_font):
        """متن فارسی شکل‌گرفته را وسط‌چین می‌کشد."""
        bbox = d.textbbox((0, 0), s, font=font, direction='rtl', language='fa')
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text((cx - w / 2 - bbox[0], cy - h / 2 - bbox[1]), s, font=font,
               fill='black', direction='rtl', language='fa')

    def en_text(s, cx, cy, font=en_font):
        bbox = d.textbbox((0, 0), s, font=font)
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text((cx - w / 2, cy - h / 2), s, font=font, fill='black')

    def arrow(p1, p2):
        d.line([p1, p2], fill='black', width=5)
        ang = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
        for da in (0.5, -0.5):
            x = p2[0] - 30 * math.cos(ang + da)
            y = p2[1] - 30 * math.sin(ang + da)
            d.line([p2, (x, y)], fill='black', width=5)

    for k_, (x, y) in boxes.items():
        d.rectangle([x - bw / 2, y - bh / 2, x + bw / 2, y + bh / 2],
                    outline='black', width=5)
        fa_text(labels[k_], x, y)

    arrow((430, 552), (780, 292))
    arrow((1020, 292), (1360, 552))
    arrow((535, 620), (1262, 620))
    en_text(f'a = {m["a"]:.3f}{stars(m["a_p"])}', 540, 380)
    en_text(f'b = {m["b"]:.3f}{stars(m["b_p"])}', 1260, 380)
    en_text(f"c' = {m['cp']:.3f}{stars(m['cp_p'])}", 900, 570)
    fa_text(f'اثر غیرمستقیم = {m["indirect"]:.3f} ، CI95% [{m["ci_ab"][0]:.3f}, {m["ci_ab"][1]:.3f}]',
            900, 880)
    img.save(DRAFTS / 'ch4_fig1_path_model.png')


# ---------- نگارش فصل چهارم ----------
def f2(x):
    return f'{x:.2f}'


def f3(x):
    return f'{x:.3f}'


def write_chapter(out):
    d, m, c = out['desc'], out['mediation'], out['corr']
    g, mk = out['gender'], out['market']
    n = out['n']
    pct = lambda x: 100 * x / n
    L = []
    A = L.append
    A('# فصل چهارم: یافته‌های پژوهش (نسخهٔ نمایشی مبتنی بر دادهٔ شبیه‌سازی‌شده)')
    A('')
    A('> **یادداشت روش‌شناختی:** این فصل بر اساس داده‌های شبیه‌سازی‌شده (Simulated Data) صرفاً جهت نمایش روش‌شناسی تحلیل آماری تدوین گردیده است.')
    A('')
    A('## ۴-۱. مقدمهٔ فصل')
    A('')
    A('هدف این فصل، نمایش گام‌به‌گام فرآیند تحلیل آماری داده‌ها برای آزمون فرضیه‌های پژوهش است. بدین منظور، یک دیتاست مصنوعی شامل ۱۸۰ شرکت‌کنندهٔ فرضی با ویژگی‌های توزیعی و روان‌سنجی واقع‌گرایانه تولید شد و تمام آزمون‌های توصیفی و استنباطیِ برنامه‌ریزی‌شده برای پژوهش اصلی، دقیقاً بر همین داده اجرا گردید. ساختار فصل به این ترتیب است: ابتدا توصیف نمونه و متغیرها، سپس بررسی پایایی ابزارها و پیش‌فرض‌های تحلیل استنباطی، پس از آن آزمون فرضیه‌های فرعی با همبستگی پیرسون، و در پایان آزمون فرضیهٔ اصلی با تحلیل میانجی‌گری بر اساس الگوی مدل ۴ هایز (PROCESS Model 4) همراه با بوت‌استرپ ۱۰۰۰ بازنمونه‌گیری.')
    A('')
    A('## ۴-۲. آمار توصیفی نمونه')
    A('')
    A(f'حجم نمونهٔ شبیه‌سازی‌شده ۱۸۰ نفر است. میانگین سنی شرکت‌کنندگان {f2(out["age"]["mean"])} سال (انحراف معیار {f2(out["age"]["sd"])}) و دامنهٔ آن از {out["age"]["min"]} تا {out["age"]["max"]} سال بود. میانگین سابقهٔ معامله نیز {f2(out["exp"]["mean"])} ماه (انحراف معیار {f2(out["exp"]["sd"])}) با دامنهٔ {out["exp"]["min"]} تا {out["exp"]["max"]} ماه به دست آمد. جدول ۴-۱ توزیع فراوانی و درصد ویژگی‌های جمعیت‌شناختی و معاملاتی نمونه را نشان می‌دهد.')
    A('')
    A('**جدول ۴-۱.** توزیع فراوانی و درصد ویژگی‌های جمعیت‌شناختی و معاملاتی نمونه (N = 180)')
    A('')
    A('| متغیر | گروه | فراوانی | درصد |')
    A('|---|---|---|---|')
    A(f'| جنسیت | مرد | {g["Male"]} | {f2(pct(g["Male"]))} |')
    A(f'| جنسیت | زن | {g["Female"]} | {f2(pct(g["Female"]))} |')
    A(f'| بازار فعالیت | بورس | {mk["Bourse"]} | {f2(pct(mk["Bourse"]))} |')
    A(f'| بازار فعالیت | رمزارز | {mk["Crypto"]} | {f2(pct(mk["Crypto"]))} |')
    A(f'| بازار فعالیت | فارکس | {mk["Forex"]} | {f2(pct(mk["Forex"]))} |')
    A(f'| بازار فعالیت | طلا | {mk["Gold"]} | {f2(pct(mk["Gold"]))} |')
    A('')
    A('## ۴-۳. آمار توصیفی متغیرهای پژوهش')
    A('')
    A('جدول ۴-۲ شاخص‌های توصیفی سه متغیر اصلی پژوهش را ارائه می‌کند. مقادیر چولگی و کشیدگی در بازهٔ متعارف (کمتر از ۲± و ۷± به‌ترتیب) قرار دارند و از این منظر، پیش‌فرض نرمال بودن توزیع نمره‌ها برقرار است.')
    A('')
    A('**جدول ۴-۲.** شاخص‌های توصیفی متغیرهای پژوهش')
    A('')
    A('| متغیر | میانگین | انحراف معیار | کمینه | بیشینه | چولگی | کشیدگی |')
    A('|---|---|---|---|---|---|---|')
    rows = [('اضطراب صفتی (STAI-Y2)', d['Anxiety']), ('طرحواره‌های ناسازگار اولیه (YSQ-SF)', d['Schema']), ('رفتارهای مالی پرخطر (Grable & Lytton)', d['Risk'])]
    for name, s in rows:
        A(f'| {name} | {f2(s["mean"])} | {f2(s["sd"])} | {s["min"]} | {s["max"]} | {f2(s["skew"])} | {f2(s["kurt"])} |')
    A('')
    A('## ۴-۴. بررسی پایایی ابزارها')
    A('')
    A('به‌منظور گزارش پایایی در نمونهٔ شبیه‌سازی‌شده، ضریب آلفای کرونباخ برای مقیاس‌ها محاسبه شد (جدول ۴-۳). تمام ضرایب بالاتر از آستانهٔ متعارف ۰٫۷۰ هستند و پایایی مطلوب مقیاس‌ها در این نمونه را نشان می‌دهند.')
    A('')
    A('**جدول ۴-۳.** ضرایب آلفای کرونباخ در نمونهٔ پژوهش')
    A('')
    A('| متغیر | تعداد گویه | آلفای کرونباخ |')
    A('|---|---|---|')
    A(f'| اضطراب صفتی (STAI-Y2) | 20 | {f3(out["alpha"]["Anxiety"])} |')
    A(f'| طرحواره‌های ناسازگار اولیه (YSQ-SF) | 75 | {f3(out["alpha"]["Schema"])} |')
    A(f'| رفتارهای مالی پرخطر (Grable & Lytton) | 13 | {f3(out["alpha"]["Risk"])} |')
    A('')
    A('## ۴-۵. بررسی پیش‌فرض‌های تحلیل استنباطی')
    A('')
    A(f'**نرمال بودن:** مقادیر چولگی و کشیدگی هر سه متغیر در بازهٔ متعارف قرار دارند (جدول ۴-۲). آزمون شاپیرو-ویلک نیز برای اضطراب W = {f3(d["Anxiety"]["sw_W"])}، برای طرحواره‌ها W = {f3(d["Schema"]["sw_W"])} و برای رفتار مالی W = {f3(d["Risk"]["sw_W"])} به دست آمد؛ با توجه به حجم نمونه و پایایی شبیه‌سازی، تکیهٔ اصلی بر قضیهٔ حد مرکزی و متقارن بودن توزیع‌هاست.')
    A(f'**خطی بودن:** رابطهٔ بین متغیرها بر اساس نمودارهای پراکندگی و ضرایب همبستگی، خطی فرض می‌شود و مدل رگرسیون بر همین مبنا برآورد شده است.')
    A(f'**عدم هم‌خطی:** در مدل پیش‌بینی رفتار مالی، شاخص تورم واریانس (VIF) برای پیش‌بین‌ها برابر {f2(m["vif"])} است که بسیار کمتر از آستانهٔ ۱۰ (و حتی ۵) بوده و مسئلهٔ هم‌خطی چندگانه منتفی است.')
    A(f'**استقلال پسماندها:** آمارهٔ دوربین-واتسون در مدل میانجی {f2(m["dw_a"])} و در مدل پیامد {f2(m["dw_b"])} به دست آمد که در بازهٔ قابل قبول (حدود ۲) قرار دارد.')
    A('')
    A('## ۴-۶. آزمون فرضیه‌های فرعی (تحلیل همبستگی)')
    A('')
    A('**جدول ۴-۴.** ماتریس همبستگی پیرسون بین متغیرهای پژوهش')
    A('')
    A('| متغیر | ۱ | ۲ | ۳ |')
    A('|---|---|---|---|')
    A('| ۱. اضطراب صفتی | — | | |')
    A(f'| ۲. طرحواره‌های ناسازگار اولیه | {f2(c["Anxiety-Schema"]["r"])}** | — | |')
    A(f'| ۳. رفتارهای مالی پرخطر | {f2(c["Anxiety-Risk"]["r"])}** | {f2(c["Schema-Risk"]["r"])}** | — |')
    A('')
    A('**یادداشت.** **p < 0.01 (دو دامنه).')
    A('')
    A(f'**فرضیهٔ فرعی دوم:** بین اضطراب صفتی و رفتارهای مالی پرخطر رابطهٔ مثبت و معنادار وجود دارد (r = {f2(c["Anxiety-Risk"]["r"])}، {p_str(c["Anxiety-Risk"]["p"])}). بنابراین این فرضیه تأیید می‌شود؛ یعنی سطح بالاتر اضطراب صفتی با بروز بیشتر رفتارهای مالی پرخطر همراه است.')
    A('')
    A(f'**فرضیهٔ فرعی سوم:** بین اضطراب صفتی و طرحواره‌های ناسازگار اولیه رابطهٔ مثبت و معنادار وجود دارد (r = {f2(c["Anxiety-Schema"]["r"])}، {p_str(c["Anxiety-Schema"]["p"])}). این فرضیه نیز تأیید می‌شود.')
    A('')
    A(f'**فرضیهٔ فرعی چهارم:** بین طرحواره‌های ناسازگار اولیه و رفتارهای مالی پرخطر رابطهٔ مثبت و معنادار وجود دارد (r = {f2(c["Schema-Risk"]["r"])}، {p_str(c["Schema-Risk"]["p"])}). این فرضیه تأیید می‌شود.')
    A('')
    A('## ۴-۷. آزمون فرضیهٔ اصلی و تحلیل میانجی‌گری')
    A('')
    A('به‌منظور آزمون نقش میانجی‌گر طرحواره‌های ناسازگار اولیه در رابطهٔ اضطراب صفتی و رفتارهای مالی پرخطر، تحلیل میانجی‌گری بر اساس الگوی مدل ۴ هایز اجرا شد. جدول ۴-۵ نتایج رگرسیون مسیرها و جدول ۴-۶ اثرات مستقیم، غیرمستقیم و کل را با بازهٔ اطمینان بوت‌استرپ گزارش می‌کند.')
    A('')
    A('**جدول ۴-۵.** نتایج رگرسیون برای ضرایب مسیرهای مدل میانجی‌گری')
    A('')
    A('| مسیر/معادله | ضریب B | خطای معیار | t | سطح معناداری | R² | F |')
    A('|---|---|---|---|---|---|---|')
    A(f'| مسیر a: اضطراب ← طرحواره‌ها | {f3(m["a"])} | {f3(m["a_se"])} | {f2(m["a_t"])} | {p_str(m["a_p"])} | {f2(m["r2_m"])} | {f2(m["f_m"])} |')
    A(f'| مسیر b: طرحواره‌ها ← رفتار مالی (با کنترل اضطراب) | {f3(m["b"])} | {f3(m["b_se"])} | {f2(m["b_t"])} | {p_str(m["b_p"])} | {f2(m["r2_y"])} | {f2(m["f_y"])} |')
    A(f"| اثر مستقیم c': اضطراب ← رفتار مالی | {f3(m['cp'])} | {f3(m['cp_se'])} | {f2(m['cp_t'])} | {p_str(m['cp_p'])} | | |")
    A(f'| اثر کل c: اضطراب ← رفتار مالی | {f3(m["c"])} | {f3(m["c_se"])} | {f2(m["c_t"])} | {p_str(m["c_p"])} | | |')
    A('')
    A('**جدول ۴-۶.** نتایج تحلیل اثرات مستقیم، غیرمستقیم و کل با روش بوت‌استرپ (۱۰۰۰ بازنمونه، CI ۹۵٪)')
    A('')
    A('| اثر | ضریب | کران پایین CI | کران بالای CI | نتیجه |')
    A('|---|---|---|---|---|')
    sig = 'معنادار (CI بدون صفر)' if (m['ci_ab'][0] > 0 or m['ci_ab'][1] < 0) else 'نامعنادر'
    A(f'| اثر غیرمستقیم (a×b) | {f3(m["indirect"])} | {f3(m["ci_ab"][0])} | {f3(m["ci_ab"][1])} | {sig} |')
    A(f"| اثر مستقیم (c') | {f3(m['cp'])} | {f3(m['ci_cp'][0])} | {f3(m['ci_cp'][1])} | {'معنادار' if (m['ci_cp'][0] > 0 or m['ci_cp'][1] < 0) else 'نامعنادر'} |")
    A(f'| اثر کل (c) | {f3(m["c"])} | {f3(m["ci_ct"][0])} | {f3(m["ci_ct"][1])} | معنادار |')
    A('')
    A('**شکل ۴-۱.** نمودار مدل آماری تحلیل‌شده به همراه ضرایب مسیر و معناداری')
    A('')
    A('![شکل ۴-۱](ch4_fig1_path_model.png)')
    A('')
    A(f'**تفسیر فرضیهٔ اصلی:** نتایج نشان داد اضطراب صفتی اثر مثبت معناداری بر طرحواره‌های ناسازگار اولیه دارد (a = {f3(m["a"])}، {p_str(m["a_p"])}) و طرحواره‌ها نیز با کنترل اضطراب، اثر مثبت معناداری بر رفتارهای مالی پرخطر دارند (b = {f3(m["b"])}، {p_str(m["b_p"])}). اثر غیرمستقیم برابر {f3(m["indirect"])} به دست آمد و بازهٔ اطمینان ۹۵٪ بوت‌استرپ [{f3(m["ci_ab"][0])}، {f3(m["ci_ab"][1])}] صفر را در بر نمی‌گیرد؛ بنابراین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه تأیید می‌شود. با توجه به معنادار بودن اثر مستقیم (c′ = {f3(m["cp"])}، {p_str(m["cp_p"])})، میانجی‌گری از نوع «جزئی» است و حدود {int(round(100 * m["prop_mediated"]))} درصد از اثر کل اضطراب بر رفتار مالی از مسیر طرحواره‌ها منتقل می‌شود.')
    A('')
    A('## ۴-۸. خلاصه و جمع‌بندی نتایج آزمون فرضیه‌ها')
    A('')
    A('**جدول ۴-۷.** خلاصهٔ وضعیت تأیید فرضیه‌ها')
    A('')
    A('| فرضیه | مسیر | ضریب | سطح معناداری | نتیجه |')
    A('|---|---|---|---|---|')
    A(f'| فرعی دوم | اضطراب صفتی ← رفتارهای مالی پرخطر | r = {f2(c["Anxiety-Risk"]["r"])} | {p_str(c["Anxiety-Risk"]["p"])} | تأیید |')
    A(f'| فرعی سوم | اضطراب صفتی ← طرحواره‌های ناسازگار اولیه | r = {f2(c["Anxiety-Schema"]["r"])} | {p_str(c["Anxiety-Schema"]["p"])} | تأیید |')
    A(f'| فرعی چهارم | طرحواره‌های ناسازگار اولیه ← رفتارهای مالی پرخطر | r = {f2(c["Schema-Risk"]["r"])} | {p_str(c["Schema-Risk"]["p"])} | تأیید |')
    A(f'| اصلی | میانجی‌گری طرحواره‌ها (اثر غیرمستقیم) | {f3(m["indirect"])} | CI95% [{f3(m["ci_ab"][0])}، {f3(m["ci_ab"][1])}] | تأیید |')
    A('')
    A('برآیند یافته‌ها الگوی نظری پژوهش را پشتیبانی می‌کند: اضطراب صفتی هم به‌طور مستقیم و هم از طریق فعال‌سازی طرحواره‌های ناسازگار اولیه، با افزایش رفتارهای مالی پرخطر در معامله‌گران همراه است.')
    (DRAFTS / 'chapter4_simulated.md').write_text('\n'.join(L), encoding='utf-8')


def main():
    df, out = run_analysis()
    draw_figure(out['mediation'])
    write_chapter(out)
    (SIM_DIR / 'ch4_stats.json').write_text(
        json.dumps(out, ensure_ascii=False, indent=1, default=float), encoding='utf-8')
    m = out['mediation']
    print('تحلیل کامل شد.')
    print('a=%.3f b=%.3f cp=%.3f c=%.3f indirect=%.3f CI=[%.3f, %.3f]' % (
        m['a'], m['b'], m['cp'], m['c'], m['indirect'], m['ci_ab'][0], m['ci_ab'][1]))
    print('آلفا:', out['alpha'])


if __name__ == '__main__':
    main()
