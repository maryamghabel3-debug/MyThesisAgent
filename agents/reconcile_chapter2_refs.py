# فاز ۴ مأموریت ۲۱ — آشتی‌دادن نهایی منابع فصل دوم (تطبیق دوجهته)
# خروجی: ۱) فهرست منابع نهایی در انتهای chapter2_full.md  ۲) جدول آشتی در chapter2_final_reconciliation.md

import re
from pathlib import Path

ROOT = Path('/home/user/MyThesisAgent')
BASE = (ROOT / 'output/drafts/chapter2_cleaned_base.md').read_text(encoding='utf-8').split('\n')
FULL_PATH = ROOT / 'output/drafts/chapter2_full.md'
full = FULL_PATH.read_text(encoding='utf-8')

# فهرست منابع فارسی: سطرهای ۲۷۶ تا ۲۸۴؛ انگلیسی: سطرهای ۲۸۶ تا ۳۶۲ (با ادغام سطرهای ادامهٔ مدخل)
fa_entries = [l.strip() for l in BASE[276:285] if l.strip()]
assert len(fa_entries) == 9, f'منابع فارسی: {len(fa_entries)}'

en_entries = []
CONT = ('Journal of Information Systems', 'The effect of early maladaptive schemas')  # سطرهای ادامهٔ مدخل‌های سفیتیری و صدیقی
for l in BASE[286:363]:
    if not l.strip():
        continue
    if l.strip().startswith(CONT):
        en_entries[-1] = en_entries[-1] + ' ' + l.strip()
    else:
        en_entries.append(l.rstrip())
assert len(en_entries) == 75, f'مدخل‌های انگلیسی: {len(en_entries)}'

# ------------------------------------------------------------- اصلاحات فاز ۲ روی مدخل‌ها
def find_entry(prefix):
    cands = [e for e in en_entries if e.startswith(prefix)]
    assert len(cands) == 1, f'{prefix}: {len(cands)} مدخل'
    return cands[0]

HALVORSEN_OLD = find_entry('Halvorsen')
HALVORSEN_NEW = ('Halvorsen, M., Wang, C. E., Richter, J., Myrland, I., Pedersen, S. K., Eisemann, M., & Waterloo, K. (2009). '
                 'Early maladaptive schemas, temperament and character traits in clinically depressed and previously depressed subjects. '
                 'Clinical Psychology & Psychotherapy, 16(5), 394-407. https://doi.org/10.1002/cpp.618')
FENTON_TFD_OLD = find_entry("Fenton-O'Creevy, M., Soane")
FENTON_TFD_NEW = ("Fenton-O'Creevy, M., Soane, E., Nicholson, N., & Willman, P. (2011). "
                  'Thinking, feeling and deciding: The influence of emotions on the decision making and performance of traders. '
                  'Journal of Organizational Behavior, 32(8), 1044-1061. https://doi.org/10.1002/job.720')
OLOLO_OLD = find_entry('Ololo')
OLOLO_NEW = re.sub(r'345[–-]\s*367', '398-418', OLOLO_OLD)
assert OLOLO_NEW != OLOLO_OLD
replacements = {HALVORSEN_OLD: HALVORSEN_NEW, FENTON_TFD_OLD: FENTON_TFD_NEW, OLOLO_OLD: OLOLO_NEW}

# ------------------------------------------------------------- نگاشت استنادهای درون‌متنی به مدخل‌ها
CITED = {
    'امیدی': 'امیدی و همکاران (۱۳۹۲)', 'بهدیو': 'بهدیو و همکاران (۱۳۹۸)', 'ستایش': 'ستایش و همکاران (۱۳۹۹)',
    'طالبی': 'طالبی و همکاران (۱۳۹۸)', 'فلاح': 'فلاح و خدایی (۱۳۹۸)', 'فروغی': 'فروغی و تهرانی (۱۳۹۷)',
    'نوری': 'نوری و هاشمیان (۱۳۹۶)', 'یوسفی': 'یوسفی (۱۳۹۶)',
    'An, H.': 'آن و همکاران (۲۰۲۱)', 'Arena,': 'آرنا و همکاران (۲۰۲۲)', 'American Psychiatric': 'انجمن روانپزشکی آمریکا (۲۰۱۳)',
    'Arntz': 'آرنتز و جیکوب (۲۰۱۲)', 'Auer': 'اور و شوماخر (۲۰۲۳)', 'Baddeley': 'بدلی و همکاران (۲۰۲۱)',
    'Baker': 'بیکر و نوفسینگر (۲۰۱۰)', 'Barber': 'باربر و اودین (۲۰۰۰)', 'Barlow': 'بارلو (۲۰۰۲)',
    'Beck, A. T. (1976)': 'بک (۱۹۷۶)', 'Beck, A. T., & Clark': 'بک و کلارک (۱۹۹۷)', 'Bishop': 'بیشاپ (۲۰۰۷)',
    'Borkovec': 'بورکووک (۱۹۹۴)', 'Charles': 'چارلز و کاسیلینگام (۲۰۲۲)', 'De Martino': 'دی مارتینو و همکاران (۲۰۰۶)',
    'Dhingra': 'دینگرا و آگاروال (۲۰۲۵)', 'Estévez': 'استِوِز و همکاران (۲۰۲۴)', 'Eysenck': 'آیزنک و همکاران (۲۰۰۷)',
    'Fenton-O’Creevy': 'فنتون-اکریوی و همکاران (۲۰۱۲الف)', "Fenton-O'Creevy, M., Soane": 'فنتون-اکریوی و همکاران (۲۰۱۲ب)',
    'Gambetti': 'گامبتی و گیوسبرتی (۲۰۱۲)', 'Goetzmann': 'گوتزمن و کومار (۲۰۰۸)',
    'Grable, J. E., & Lytton': 'گرابل و لایتون (۱۹۹۹)', 'Grable, J. E. (2017)': 'گرابل (۲۰۱۷)',
    'Grall': 'گرال-برونک و همکاران (۲۰۲۵)', 'Grupe': 'گروپ و نیچکه (۲۰۱۳)', 'Halvorsen': 'هالورسن و همکاران (۲۰۰۹)',
    'Hartley': 'هارتلی و فلپس (۲۰۱۲)', 'Hayes': 'هایز (۲۰۱۸)', 'Hwang': 'هوانگ و سالمون (۲۰۰۴)',
    'Kahneman, D. (2011)': 'کانمن (۲۰۱۱)', 'Kahneman, D., & Tversky': 'کانمن و تورسکی (۱۹۷۹)',
    'Kaplan': 'کاپلان و سادوک (۲۰۱۵)', 'Kuhnen, C. M., & Chiao': 'کوهنن و چیائو (۲۰۰۹)',
    'Lauriola': 'لوریولا و لوین (۲۰۰۱)', 'Lindberg': 'لیندبرگ و همکاران (۲۰۱۱)',
    'Lo, A. W., & Repin': 'لو و رپین (۲۰۰۲)', 'Lo, A. W., Repin': 'لو و همکاران (۲۰۰۵)',
    'Mahato': 'ماهاتو و همکاران (۲۰۲۵)', 'Malkiel': 'ملکیل (۲۰۰۳)', 'Maner': 'مانر و اشمیت (۲۰۰۶)',
    'Mitte': 'میت (۲۰۰۷)', 'Nguyen': 'نوین و همکاران (۲۰۱۹)', 'OECD': 'سازمان همکاری و توسعه اقتصادی (۲۰۲۱)',
    'Ololo': 'اولولو و همکاران (۲۰۲۴)', 'Pak': 'پاک و محمود (۲۰۱۹)', 'Rehm': 'رایم و همکاران (۲۰۱۵)',
    'Safitri': 'سافیتری و همکاران (۲۰۲۵)', 'Saggino': 'ساژینو و همکاران (۲۰۱۸)',
    'Schmidt, N. B., Joiner': 'اشمیت و همکاران (۱۹۹۵)', 'Sedighi': 'صدیقی ارفعی و همکاران (۲۰۲۴)',
    'Shefrin': 'شفرین و استاتمن (۱۹۸۵)', 'Shiller': 'شیلر (۲۰۱۵)', 'Shorey': 'شوری و همکاران (۲۰۱۲)',
    'Spielberger, C. D. (1983)': 'اسپیلبرگر (۱۹۸۳)', 'Thaler': 'تیلر (۲۰۱۶)', 'Vasile': 'واسیل (۲۰۱۴)',
    'Vieira': 'ویرا و همکاران (۲۰۲۳)', 'Yalom': 'یالوم (۱۹۸۰)',
    'Young, J. E. (1998)': 'یانگ (۱۹۹۸)', 'Young, J. E., Klosko': 'یانگ و همکاران (۲۰۰۳)',
    'Zuckerman': 'زاکرمن (۱۹۹۹)',
}

REMOVE_8 = ('Barnes', 'Julian', 'Kuhnen, C. M., & Knutson', 'Lusardi', 'Odean, T. (1998)', 'Richards', 'Spada', 'Young, J. E. (2005)')

def match_key(entry):
    return next((k for k in CITED if entry.startswith(k)), None)

rows, en_final, uncited = [], [], []
for e in en_entries:
    k = match_key(e)
    if k:
        en_final.append(replacements.get(e, e))
    else:
        uncited.append(e)

fa_final = list(fa_entries)  # هر ۹ منبع فارسی حفظ می‌شوند (قانون تخطی‌ناپذیر)

# ------------------------------------------------------------- نوشتن فهرست در انتهای فصل
if '## منابع فصل دوم' in full:
    full = full[:full.index('## منابع فصل دوم')].rstrip()
ref_block = '\n\n'.join(
    ['## منابع فصل دوم (نسخهٔ کاری — در مونتاژ نهایی پایان‌نامه ادغام می‌شود)', '',
     '### الف) منابع فارسی', '', *[e.replace('\xa0', ' ') for e in fa_final], '',
     '### ب) منابع لاتین', '', *[e.replace('\xa0', ' ') for e in en_final]])
FULL_PATH.write_text(full.rstrip() + '\n\n' + ref_block + '\n', encoding='utf-8')

# ------------------------------------------------------------- جدول آشتی
def status_of(e, k):
    if k:
        s = '✅ استنادشده — مدخل موجود'
        if e is HALVORSEN_OLD: s += '؛ مدخل با نسخهٔ رسمی (مأموریت ۲۱/فاز ۲) جایگزین شد'
        if e is FENTON_TFD_OLD: s += '؛ مدخل با نسخهٔ رسمی جایگزین + تفکیک «ب»؛ مغایرت سال رسمی (2011) با استناد درون‌متنی (۲۰۱۲) گزارش شد'
        if e is OLOLO_OLD: s += '؛ صفحات به 398-418 (رکورد رسمی) اصلاح شد'
        if e.startswith('Vasile'): s = '⚠️ استنادشده — مدخل موجود؛ رکورد رسمی یافت نشد (پرچم باز، حذف نشد)'
        if e.startswith('Sedighi'): s = '⚠️ استنادشده — مدخل موجود (عنوان با رکورد DOI هم‌خوان است)؛ اما عنوانِ درون‌متنی با مدخل نمی‌خواند و سال‌ها ناهمخوان‌اند (پرچم باز)'
        return s, 'حفظ'
    tag = '🗑 نامزد حذف (قانون ۸ مورد)' if any(e.startswith(r) for r in REMOVE_8) else '— خارج از متن فصل'
    return f'❌ استنادنشده در فصل ۲ {tag}', 'حذف از فهرست فصل'

md = ['# جدول آشتی نهایی منابع فصل دوم (تطبیق دوجهته)', '',
      '> **مأموریت ۲۱ — فاز ۴** | تاریخ: ۲۰۲۶-۰۹-۱۹',
      '> **مأخذ تطبیق:** استنادهای استخراج‌شده از `output/drafts/chapter2_full.md` ↔ مدخل‌های فهرست منابع `chapter2_cleaned_base.md`',
      '> اصلاحات اعمال‌شده بر مدخل‌ها صرفاً بر اساس رکوردهای رسمی (فاز ۲) بوده و هیچ مدخل جدیدی ساخته نشده است.', '',
      '## ۱) منابع فارسی (۹ مدخل — همه حفظ می‌شوند)', '',
      '| استناد یکتا | مدخل موجود؟ | وضعیت نهایی |', '|---|---|---|']
for e in fa_entries:
    k = match_key(e)
    name = CITED[k] if k else e.split('،')[0] + ' (بدون استناد در فصل ۲)'
    st = '✅ استنادشده' if k else '❌ در متن فصل ۲ استناد نشده — به‌دلیل قانون «۹ منبع فارسی حفظ شوند» در فهرست می‌ماند (استناد آن در فصل ۳ است)'
    md.append(f'| {name} | بله | {st} |')
md += ['', '## ۲) منابع لاتین (۷۵ مدخل پایه)', '',
       '| استناد یکتا / مدخل | مدخل موجود؟ | وضعیت نهایی | اقدام |', '|---|---|---|---|']
for e in en_entries:
    k = match_key(e)
    label = CITED[k] if k else e[:80] + ('…' if len(e) > 80 else '')
    st, act = status_of(e, k)
    md.append(f"| {label} | {'بله' if k else 'بله'} | {st} | {act} |")
md += ['', '## ۳) خلاصه', '',
       f'- استناد یکتای درون‌متنی: ۷۰ کلید → ۷۰ اثر (فارسی ۸ + لاتین ۶۲).',
       f'- مدخل‌های لاتین پایه: ۷۵ → استنادشده: {len(en_final)}؛ حذف‌شده از فهرست فصل: {len(uncited)}.',
       '- حذف‌شدگان: ۸ نامزد `REMOVE_AT_FINAL_ASSEMBLY` (بارنز، جولیان، کوهنن و نوتسون، لوساردی و میچل، اودین ۱۹۹۸، ریچاردز، اسپادا، یانگ ۲۰۰۵) + ۵ مدخل بدون استناد در فصل ۲ (فاما ۱۹۷۰، کلاین ۲۰۱۵، اسپیلبرگر/گورچ/لوشن ۱۹۷۰، بودی/کین/مارکوس ۲۰۱۸، میشکین ۲۰۱۹ — دو مورد آخر متعلق به فصل ۱).',
       '- تفکیک الف/ب فنتون-اکریوی اعمال شد (الف = مقالهٔ تنظیم هیجان/HRV؛ ب = مقالهٔ چندوجهی بانک‌های سرمایه‌گذاری).',
       '- پرچم‌های باز: واسیل ۲۰۱۴ (رکورد یافت نشد)، صدیقی ارفعی (ناهمخوانی عنوان/سال درون‌متنی با مدخل و رکورد).',
       '- لیندبرگ ۲۰۱۱: مدخل موجود و استنادشده؛ در این مأموریت راستی‌آزمایی مستقل انجام نشد و پرچم بازی هم ندارد.',
       '- بشارت ۱۳۹۵: تنها منبع فارسی بدون استناد در فصل ۲ (استناد آن در بخش ابزار فصل ۳ است) — حفظ شد.']
(ROOT / 'data/processed/chapter2_final_reconciliation.md').write_text('\n'.join(md) + '\n', encoding='utf-8')
print('فهرست نهایی: فارسی', len(fa_final), '| لاتین', len(en_final), '| حذف‌شده', len(uncited))
for e in uncited: print('  حذف:', e[:80])
