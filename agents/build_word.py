#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
اسکریپت ساخت فایل Word فصل‌های ۱ تا ۳ پایان‌نامه طبق راهنمای واقعی دانشگاه
---------------------------------------------------------------------------
مأموریت ۳۹:
- قطعه‌بندی هر پاراگراف به اجراهای فارسی/لاتین: فارسی = B Nazanin روی ascii/hAnsi/cs
  با w:hint="cs" + <w:rtl/> + <w:lang w:bidi="fa-IR"/>؛ لاتین = Times New Roman بدون rtl
- حذف همهٔ صفات theme فونت از styles.xml (docDefaults و استایل‌ها)
- پاورقی لاتین نام نویسندگان خارجی در اولین استناد (املای لاتین فقط از فهرست منابع)
- شماره‌گذاری پاورقی در هر صفحه از ۱ (footnotePr/numRestart=eachPage) + اعداد فارسی (numFmt=hindi)
ورودی:  output/drafts/chapter1.md، chapter2_full.md، chapter3_full.md
خروجی:  output/final/thesis_ch1_to_ch3.docx + data/processed/citation_latin_map.md
"""

import re
from pathlib import Path
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ------------------------- تنظیمات قالب -------------------------
FA_FONT = 'B Nazanin'
EN_FONT = 'Times New Roman'
FA_SIZE = 12
EN_SIZE = 11
CHAPTER_KICKER_PT = 28
CHAPTER_TITLE_PT = 26
H2_PT = 14
H3_PT = 13
TABLE_FONT_PT = 11
LINE_SPACING_PT = 29
FOOTNOTE_PT = 10   # در راهنما صریح نیست؛ پیش‌فرض ثبت‌شده در گزارش

ROOT = Path(__file__).resolve().parents[1]
CH_SOURCES = [
    ('فصل اول', ROOT / 'output/drafts/chapter1.md'),
    ('فصل دوم', ROOT / 'output/drafts/chapter2_full.md'),
    ('فصل سوم', ROOT / 'output/drafts/chapter3_full.md'),
]
THESIS_TITLE = ('تعیین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه در رابطه بین '
                'اضطراب صفتی و رفتارهای مالی پرخطر در معامله‌گران بازارهای مالی')
OUT_PATH = ROOT / 'output/final/thesis_ch1_to_ch3.docx'
MAP_PATH = ROOT / 'data/processed/citation_latin_map.md'

FN_TOKEN = '\x01FN{}\x01'
FN_TOKEN_RE = re.compile(r'\x01FN(\d+)\x01')
FA_DIG = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')

# ------------------------- نقشهٔ استنادهای خارجی -------------------------
# املای لاتین در همهٔ موارد عیناً از مدخل فهرست منابع یکپارچه گرفته شده است.
AUTHOR_MAP = {
    ('گورچ و لوشن', '1970'): 'Gorsuch & Lushene',
    ('اسپیلبرگر، گورچ و لوشن', '1970'): 'Spielberger, Gorsuch, & Lushene',
    ('بک', '1976'): 'Beck',
    ('کانمن و تورسکی', '1979'): 'Kahneman & Tversky',
    ('یالوم', '1980'): 'Yalom',
    ('اسپیلبرگر', '1983'): 'Spielberger',
    ('شفرین و استاتمن', '1985'): 'Shefrin & Statman',
    ('بورکووک', '1994'): 'Borkovec',
    ('اشمیت و همکاران', '1995'): 'Schmidt et al.',
    ('بک و کلارک', '1997'): 'Beck & Clark',
    ('یانگ', '1998'): 'Young',
    ('دانیل، هیرشلایفر و سوبرامانیام', '1998'): 'Daniel, Hirshleifer, & Subrahmanyam',
    ('زاکرمن', '1999'): 'Zuckerman',
    ('گرابل و لایتون', '1999'): 'Grable & Lytton',
    ('باربر و اودین', '2000'): 'Barber & Odean',
    ('لوریولا و لوین', '2001'): 'Lauriola & Levin',
    ('بارلو', '2002'): 'Barlow',
    ('لو و رپین', '2002'): 'Lo & Repin',
    ('ملکیل', '2003'): 'Malkiel',
    ('یانگ و همکاران', '2003'): 'Young et al.',
    ('یانگ، کلوسکو و ویشار', '2003'): 'Young, Klosko, & Weishaar',
    ('هوانگ و سالمون', '2004'): 'Hwang & Salmon',
    ('لو و همکاران', '2005'): 'Lo et al.',
    ('دی مارتینو و همکاران', '2006'): 'De Martino et al.',
    ('مانر و اشمیت', '2006'): 'Maner & Schmidt',
    ('آیزنک و همکاران', '2007'): 'Eysenck et al.',
    ('بیشاپ', '2007'): 'Bishop',
    ('میت', '2007'): 'Mitte',
    ('گوتزمن و کومار', '2008'): 'Goetzmann & Kumar',
    ('هالورسن و همکاران', '2009'): 'Halvorsen et al.',
    ('کوهنن و چیائو', '2009'): 'Kuhnen & Chiao',
    ('بیکر و نوفسینگر', '2010'): 'Baker & Nofsinger',
    ('فنتون-اکریوی و همکاران', '2011'): "Fenton-O'Creevy et al., 2011",
    ('لیندبرگ و همکاران', '2011'): 'Lindberg et al.',
    ('کانمن', '2011'): 'Kahneman',
    ('آرنتز و جیکوب', '2012'): 'Arntz & Jacob',
    ('شوری و همکاران', '2012'): 'Shorey et al.',
    ('فنتون-اکریوی و همکاران', '2012'): "Fenton-O'Creevy et al., 2012",
    ('هارتلی و فلپس', '2012'): 'Hartley & Phelps',
    ('گامبتی و گیوسبرتی', '2012'): 'Gambetti & Giusberti',
    ('انجمن روانپزشکی آمریکا', '2013'): 'American Psychiatric Association',
    ('گروپ و نیچکه', '2013'): 'Grupe & Nitschke',
    ('رایم و همکاران', '2015'): 'Rehm et al.',
    ('شیلر', '2015'): 'Shiller',
    ('کاپلان و سادوک', '2015'): 'Kaplan & Sadock',
    ('تیلر', '2016'): 'Thaler',
    ('کلاین', '2016'): 'Kline',
    ('گرابل', '2017'): 'Grable',
    ('بودی، کین و مارکوس', '2018'): 'Bodie, Kane, & Marcus',
    ('ساژینو و همکاران', '2018'): 'Saggino et al.',
    ('هایز', '2018'): 'Hayes',
    ('میشکین', '2019'): 'Mishkin',
    ('آن و همکاران', '2021'): 'An et al.',
    ('آن، لی و وو', '2021'): 'An, Li, & Wu',
    ('بدلی و همکاران', '2021'): 'Baddeley et al.',
    ('بدلی، آیزنک و اندرسون', '2021'): 'Baddeley, Eysenck, & Anderson',
    ('آرنا و همکاران', '2022'): 'Arena et al.',
    ('چارلز و کاسیلینگام', '2022'): 'Charles & Kasilingam',
    ('اور و شوماخر', '2023'): 'Auer & Schuhmacher',
    ('صدیقی ارفعی و همکاران', '2023'): 'Sedighi Arfaee et al.',
    ('ویرا و همکاران', '2023'): 'Vieira et al.',
    ('استِوِز و همکاران', '2024'): 'Estévez et al.',
    ('اولولو و همکاران', '2024'): 'Ololo et al.',
    ('سافیتری و همکاران', '2025'): 'Safitri et al.',
    ('ماهاتو و همکاران', '2025'): 'Mahato et al.',
    ('گرال-برونک و همکاران', '2025'): 'Grall-Bronnec et al.',
}
# نام‌های روایی بدون استناد که در اولین رخداد، پاورقیِ معادل لاتین می‌گیرند (مأموریت ۳۹-ب)
NARRATIVE_NAMES = [
    ('زیگموند فروید', 'Sigmund Freud'),
    ('آلبرت الیس', 'Albert Ellis'),
    ('رولو می', 'Rollo May'),
]


def chapter_title_from_h1(text):
    """عنوان رسمی فصل را از سرفصل `#` فایل مارک‌داون می‌گیرد."""
    for line in text.splitlines():
        if line.startswith('# ') and ':' in line:
            return line.split(':', 1)[1].strip()
    return ''


# ------------------------- ابزارهای XML -------------------------

_SECTPR_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr',
                 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
                 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
                 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings',
                 'sectPrChange']


def _insert_sectpr_element(section, tag, attrs=None, children=()):
    """درج عنصر در sectPr مطابق ترتیب طرحوارهٔ CT_SectPr."""
    sectPr = section._sectPr
    el = sectPr.makeelement(qn(tag), attrs or {})
    for ctag, cattrs in children:
        el.append(sectPr.makeelement(qn(ctag), cattrs))
    local = lambda e: e.tag.split('}')[-1]
    for child in sectPr:
        if _SECTPR_ORDER.index(local(child)) > _SECTPR_ORDER.index(local(el)):
            child.addprevious(el)
            return el
    sectPr.append(el)
    return el


def _set_paragraph_bidi(par, rtl=True):
    """افزودن w:bidi (یا val=0 برای LTR) به pPr."""
    pPr = par._p.get_or_add_pPr()
    old = pPr.find(qn('w:bidi'))
    if old is not None:
        pPr.remove(old)
    el = pPr.makeelement(qn('w:bidi'), {})
    el.set(qn('w:val'), '1' if rtl else '0')
    pPr.append(el)


def _set_section_rtl(section):
    if section._sectPr.find(qn('w:bidi')) is None:
        _insert_sectpr_element(section, 'w:bidi')


# ------------------------- قطعه‌بندی متن و اجراهای دوخطی -------------------------

def _char_class(ch):
    o = ord(ch)
    if (0x0600 <= o <= 0x06FF or 0x0750 <= o <= 0x077F or 0xFB50 <= o <= 0xFDFF
            or 0xFE70 <= o <= 0xFEFF or o == 0x200C):
        return 'F'
    if 'A' <= ch <= 'Z' or 'a' <= ch <= 'z':
        return 'L'
    return 'N'


def segment_text(text):
    """تقسیم متن به قطعه‌های ('F'|'L', رشته)؛ خنثی‌ها به قطعهٔ مجاور می‌چسبند
    مگر آن‌که بین دو قطعهٔ لاتین باشند."""
    raw = [_char_class(c) for c in text]
    kinds = []
    for i, c in enumerate(raw):
        if c != 'N':
            kinds.append(c)
            continue
        prev = next((raw[j] for j in range(i - 1, -1, -1) if raw[j] != 'N'), None)
        nxt = next((raw[j] for j in range(i + 1, len(raw)) if raw[j] != 'N'), None)
        if prev == 'L' and nxt == 'L':
            kinds.append('L')
        elif prev == 'F' or nxt == 'F':
            kinds.append('F')
        else:
            kinds.append(prev or nxt or 'F')
    segs = []
    for ch, k in zip(text, kinds):
        if segs and segs[-1][0] == k:
            segs[-1][1] += ch
        else:
            segs.append([k, ch])
    return [(k, s) for k, s in segs]


def _make_run(par, text, kind, pt, bold):
    """ساخت اجرا با rPr کامل: فارسی = B Nazanin + rtl + lang bidi؛ لاتین = TNR."""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    font = FA_FONT if kind == 'F' else EN_FONT
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:cs'), font)
    if kind == 'F':
        rFonts.set(qn('w:hint'), 'cs')
    rPr.append(rFonts)
    if kind == 'F':
        rPr.append(OxmlElement('w:rtl'))
        lang = OxmlElement('w:lang')
        lang.set(qn('w:bidi'), 'fa-IR')
        rPr.append(lang)
    else:
        lang = OxmlElement('w:lang')
        lang.set(qn('w:val'), 'en-US')
        rPr.append(lang)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(pt * 2)))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(int(pt * 2)))
    rPr.append(szCs)
    if bold:
        rPr.append(OxmlElement('w:b'))
        rPr.append(OxmlElement('w:bCs'))
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    par._p.append(r)


def add_script_runs(par, text, fa_pt=FA_SIZE, en_pt=EN_SIZE, bold=False):
    """افزودن متن با قطعه‌بندی فارسی/لاتین به پاراگراف."""
    for kind, seg in segment_text(text):
        _make_run(par, seg, kind, fa_pt if kind == 'F' else en_pt, bold)


def _add_footnote_ref_run(par, fn_id):
    """اجرای بالانویس ارجاع پاورقی (فارسی، B Nazanin)."""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    for at in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rFonts.set(qn(at), FA_FONT)
    rFonts.set(qn('w:hint'), 'cs')
    rPr.append(rFonts)
    rPr.append(OxmlElement('w:rtl'))
    va = OxmlElement('w:vertAlign')
    va.set(qn('w:val'), 'superscript')
    rPr.append(va)
    r.append(rPr)
    ref = OxmlElement('w:footnoteReference')
    ref.set(qn('w:id'), str(fn_id))
    r.append(ref)
    par._p.append(r)


def add_rich_text(par, text, fa_pt=FA_SIZE, en_pt=EN_SIZE, bold_all=False):
    """افزودن متن با بولد درون‌خطی (**…**) و توکن‌های پاورقی، قطعه‌بندی‌شده."""
    for seg in FN_TOKEN_RE.split(text):
        if seg == '':
            continue
        if seg.isdigit() and seg.isascii():
            _add_footnote_ref_run(par, int(seg))
            continue
        for sub in re.split(r'(\*\*.+?\*\*)', seg):
            if not sub:
                continue
            if sub.startswith('**') and sub.endswith('**') and len(sub) > 4:
                add_script_runs(par, sub[2:-2], fa_pt, en_pt, bold=True)
            else:
                add_script_runs(par, sub, fa_pt, en_pt, bold=bold_all)


def clean_styles(doc):
    """حذف صفات theme فونت از docDefaults و همهٔ استایل‌ها + فونت صریح."""
    styles_el = doc.styles.element
    for rFonts in styles_el.iter(qn('w:rFonts')):
        for at in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
            if rFonts.get(qn(at)) is not None:
                del rFonts.attrib[qn(at)]
    # حذف سراسری همهٔ صفات وابسته به تم از هر عنصر؛ مقادیر صریح باقی می‌مانند
    THEME_ATTRS = ('w:themeColor', 'w:themeShade', 'w:themeTint',
                   'w:themeFill', 'w:themeFillShade', 'w:themeFillTint')
    for el in styles_el.iter():
        for at in THEME_ATTRS:
            if el.get(qn(at)) is not None:
                del el.attrib[qn(at)]
        if rFonts.get(qn('w:ascii')) is None:
            rFonts.set(qn('w:ascii'), EN_FONT)
            rFonts.set(qn('w:hAnsi'), EN_FONT)
        if rFonts.get(qn('w:cs')) is None:
            rFonts.set(qn('w:cs'), FA_FONT)
    # lang پیش‌فرض: bidi = fa-IR
    for lang in styles_el.iter(qn('w:lang')):
        lang.set(qn('w:bidi'), 'fa-IR')
    # استایل‌های Heading: bCs برای بولد دوطرفه
    for name in ('Heading 1', 'Heading 2', 'Heading 3', 'Title', 'Subtitle'):
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        rPr = st.element.get_or_add_rPr()
        if rPr.find(qn('w:b')) is not None and rPr.find(qn('w:bCs')) is None:
            rPr.append(OxmlElement('w:bCs'))


def add_field_paragraph(doc, instr, placeholder):
    """پاراگراف حاوی فیلد Word (TOC) با متن جای‌دار."""
    par = doc.add_paragraph()
    _set_paragraph_bidi(par)
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for ftype in ('begin',):
        r = OxmlElement('w:r')
        el = OxmlElement('w:fldChar')
        el.set(qn('w:fldCharType'), ftype)
        r.append(el)
        par._p.append(r)
    r = OxmlElement('w:r')
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = instr
    r.append(it)
    par._p.append(r)
    r = OxmlElement('w:r')
    el = OxmlElement('w:fldChar')
    el.set(qn('w:fldCharType'), 'separate')
    r.append(el)
    par._p.append(r)
    add_script_runs(par, placeholder)
    r = OxmlElement('w:r')
    el = OxmlElement('w:fldChar')
    el.set(qn('w:fldCharType'), 'end')
    r.append(el)
    par._p.append(r)
    return par


def add_page_number_footer(section):
    """شمارهٔ صفحه: وسط، ۱ سانتی‌متر از لبهٔ پایین، اجرای فارسی B Nazanin."""
    footer = section.footer
    footer.is_linked_to_previous = False
    par = footer.paragraphs[0]
    par.text = ''
    _set_paragraph_bidi(par)
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fmt = par.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(LINE_SPACING_PT)
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), ' PAGE ')
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    for at in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rFonts.set(qn(at), FA_FONT)
    rFonts.set(qn('w:hint'), 'cs')
    rPr.append(rFonts)
    rPr.append(OxmlElement('w:rtl'))
    r.append(rPr)
    t = OxmlElement('w:t')
    t.text = '۱'
    r.append(t)
    fld.append(r)
    par._p.append(fld)


def clear_footer(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    footer.paragraphs[0].text = ''


# ------------------------- موتور پاورقی‌ها (اصطلاح + نویسنده) -------------------------

class FootnoteMaster:
    """شماره‌گذاری یکپارچهٔ پاورقی‌ها به ترتیب سند: معادل انگلیسی اصطلاحات
    و نام لاتین نویسندگان خارجی در اولین استناد."""

    ELIG_RE = re.compile(r'\(([^()\n]+)\)')
    # استناد درون‌پرانتز: «نام، سال» یا چندتایی با «؛»
    PAR_CITE = re.compile(r'\(([^()\n]*?(?:۱۹|۲۰)[۰-۹]{2}[^()\n]*)\)')
    # استناد روایی: «نام (سال)»
    NAR_CITE = re.compile(
        r'([\u0600-\u06FF][\u0600-\u06FF\s‌،\-]{0,80}?)\s*'
        r'\(((?:۱۹|۲۰)[۰-۹]{2})\)')

    def __init__(self):
        self.entries = []            # متن پاورقی‌ها به ترتیب ساخت
        self.term_keys = {}          # کلید اصطلاح -> (id, ch, line)
        self.author_used = set()     # کلیدهای نویسندهٔ پاورقی‌گرفته
        self.latin_emitted = set()   # متن‌های لاتین پاورقی‌شده (جلوگیری از پاورقی تکراری یک نام)
        self.name_used = set()       # نام‌های روایی پاورقی‌گرفته
        self.name_notes = 0          # شمار پاورقی نام‌های روایی
        self.removed_later = 0
        self.left_intact = set()
        self.not_found = {}          # استناد بدون تطبیق -> شمار
        self.author_notes = 0
        self.term_notes = 0

    # ---------- اصطلاح‌ها ----------
    @staticmethod
    def term_eligible(content):
        if not re.search(r'[A-Za-z]', content):
            return False
        if re.search(r'[0-9۰-۹٠-٩]', content):
            return False
        if any(ch in content for ch in '<>=×+%/'):
            return False
        return True

    @staticmethod
    def term_key(content):
        return re.sub(r'\s+', ' ', content.split(';')[0]).strip().lower()

    def scan_terms(self, chapters):
        for ch_idx, lines in enumerate(chapters):
            for ln_no, line in enumerate(lines):
                kind = _line_kind(line)
                if kind in ('table', 'fence', 'ref', 'blank', 'h'):
                    continue
                for m in self.ELIG_RE.finditer(line):
                    c = m.group(1)
                    if not self.term_eligible(c) or not _persian_before(line, m.start()):
                        continue
                    k = self.term_key(c)
                    if k not in self.term_keys:
                        self.entries.append(c.strip())
                        self.term_keys[k] = (len(self.entries), ch_idx, ln_no)
                        self.term_notes += 1

    # ---------- نویسندگان ----------
    def lookup_author(self, raw_name, year):
        name = raw_name.strip().strip('،,. ')
        for cand in (name,):
            if (cand, year) in AUTHOR_MAP:
                return cand
        words = name.split()
        for i in range(1, len(words)):
            cand = ' '.join(words[i:]).strip('،,. ')
            if (cand, year) in AUTHOR_MAP:
                return cand
        return None

    # ---------- تبدیل خط ----------
    def transform_line(self, line, ch_idx, ln_no):
        kind = _line_kind(line)
        if kind in ('table', 'fence', 'ref', 'blank'):
            return line
        out, last = [], 0
        # ترکیب هر سه الگو در یک گذر؛ ترتیب اولویت با موقعیت در خط تعیین می‌شود
        matches = []
        for m in self.PAR_CITE.finditer(line):
            matches.append(('par', m))
        for m in self.ELIG_RE.finditer(line):
            matches.append(('term', m))
        for m in self.NAR_CITE.finditer(line):
            matches.append(('nar', m))
        # نام‌های روایی: فقط اولین رخداد، با کنترل مرز (حرف فارسی یا نیم‌فاصله نچسبد)
        for fa, la in NARRATIVE_NAMES:
            if fa in self.name_used:
                continue
            for nm in re.finditer(re.escape(fa), line):
                nxt = line[nm.end()] if nm.end() < len(line) else ''
                if nxt and ('\u0600' <= nxt <= '\u06FF' or nxt == '\u200c'):
                    continue
                matches.append(('name', nm, fa, la))
                break
        matches.sort(key=lambda x: (x[1].start(), -len(x[1].group(0))))
        taken = []
        for item in matches:
            tag, m = item[0], item[1]
            if any(s <= m.start() < e or s < m.end() <= e for s, e in taken):
                continue
            taken.append((m.start(), m.end()))
            if tag == 'term':
                last = self._do_term(line, m, ch_idx, ln_no, kind, out, last)
            elif tag == 'par':
                last = self._do_par_cite(line, m, kind, out, last)
            elif tag == 'nar':
                last = self._do_nar_cite(line, m, kind, out, last)
            else:
                last = self._do_name(line, m, item[2], item[3], kind, out, last)
        out.append(line[last:])
        return ''.join(out)

    def _do_name(self, line, m, fa, latin, kind, out, last):
        """نام روایی: درج پاورقی معادل لاتین در اولین رخداد (فقط نام، بدون سال/عنوان)."""
        if kind not in ('p', 'bullet') or fa in self.name_used:
            return last
        self.name_used.add(fa)
        self.entries.append(latin)
        self.name_notes += 1
        fid = len(self.entries)
        out.append(line[last:m.end()])
        out.append(FN_TOKEN.format(fid))
        return m.end()

    def _do_term(self, line, m, ch_idx, ln_no, kind, out, last):
        content = m.group(1)
        if not self.term_eligible(content):
            return last
        key = self.term_key(content)
        reg = self.term_keys.get(key)
        if reg is not None:
            fid, fch, fln = reg
            if (fch, fln) == (ch_idx, ln_no) and kind in ('p', 'bullet'):
                out.append(line[last:m.start()])
                out.append(FN_TOKEN.format(fid))
            else:
                out.append(line[last:m.start()].rstrip(' '))
                self.removed_later += 1
            return m.end()
        first_word = re.sub(r'[^a-z]', '', content.split()[0].lower()) \
            if content.split() else ''
        if first_word and any(k == first_word or k.startswith(first_word + ' ')
                              for k in self.term_keys) \
                and _persian_or_paren_before(line, m.start()):
            out.append(line[last:m.start()].rstrip(' '))
            self.removed_later += 1
            return m.end()
        self.left_intact.add(content)
        return last

    def _do_par_cite(self, line, m, kind, out, last):
        if kind not in ('p', 'bullet'):
            return last
        inner = m.group(1)
        new_parts, changed = [], False
        insert_pos = None
        for part in re.split(r'؛', inner):
            pm = re.search(r'^(.*?)[،\s]+((?:۱۹|۲۰)[۰-۹]{2})$', part.strip())
            if not pm:
                new_parts.append(part)
                continue
            name, yr = pm.group(1).strip(), pm.group(2).translate(FA_DIG)
            if not re.search('[؀-ۿ]', name):
                new_parts.append(part)
                continue
            cand = self.lookup_author(name, yr)
            if cand is None:
                self.not_found[(name, yr)] = self.not_found.get((name, yr), 0) + 1
                new_parts.append(part)
                continue
            akey = (cand, yr)
            latin = AUTHOR_MAP[akey]
            # اگر همین نویسنده قبلاً پاورقی گرفته، یا دقیقاً همین متن لاتین تکراری است، رد شو
            if akey in self.author_used or _norm_apos(latin) in self.latin_emitted:
                new_parts.append(part)
                continue
            self.author_used.add(akey)
            self.latin_emitted.add(_norm_apos(latin))
            self.entries.append(latin)
            self.author_notes += 1
            fid = len(self.entries)
            new_parts.append(part.replace(name, name + FN_TOKEN.format(fid), 1))
            changed = True
        if changed:
            out.append(line[last:m.start() + 1])
            out.append('؛'.join(new_parts))
            return m.end() - 1
        return last

    def _do_nar_cite(self, line, m, kind, out, last):
        if kind not in ('p', 'bullet'):
            return last
        name, yr = m.group(1), m.group(2).translate(FA_DIG)
        cand = self.lookup_author(name, yr)
        if cand is None:
            self.not_found[(name.strip(), yr)] = \
                self.not_found.get((name.strip(), yr), 0) + 1
            return last
        akey = (cand, yr)
        latin = AUTHOR_MAP[akey]
        if akey in self.author_used or _norm_apos(latin) in self.latin_emitted:
            return last
        self.author_used.add(akey)
        self.latin_emitted.add(_norm_apos(latin))
        self.entries.append(latin)
        self.author_notes += 1
        fid = len(self.entries)
        out.append(line[last:m.end(1)])
        out.append(FN_TOKEN.format(fid))
        return m.end(1)


def _line_kind(line):
    s = line.strip()
    if not s:
        return 'blank'
    if s.startswith('|'):
        return 'table'
    if s.startswith('```'):
        return 'fence'
    if s.startswith('## منابع'):
        return 'ref'
    if s.startswith('#'):
        return 'h'
    if s.startswith('- '):
        return 'bullet'
    return 'p'


def _persian_before(line, pos):
    prev = line[:pos].rstrip()
    if not prev:
        return False
    ch = prev[-1]
    return ('؀' <= ch <= 'ۿ') or ch in '»\'"‌'


def _persian_or_paren_before(line, pos):
    prev = line[:pos].rstrip()
    if not prev:
        return False
    return prev[-1] in ')»\'"' or ('؀' <= prev[-1] <= 'ۿ')


def attach_footnotes_part(doc, master):
    """ساخت word/footnotes.xml و پیوند آن به بدنه."""
    if not master.entries:
        return
    body = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
            '<w:footnote w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
            '<w:footnote w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>']
    for i, text in enumerate(master.entries, start=1):
        is_latin = bool(re.match(r'[A-Za-z]', text))
        font = EN_FONT if is_latin else FA_FONT
        bidi = '<w:pPr><w:bidi w:val="0"/></w:pPr>' if is_latin else ''
        body.append(
            f'<w:footnote w:id="{i}"><w:p>{bidi}'
            '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{FOOTNOTE_PT * 2}"/><w:szCs w:val="{FOOTNOTE_PT * 2}"/></w:rPr>'
            f'<w:t xml:space="preserve"> {escape(text)}</w:t></w:r>'
            '</w:p></w:footnote>')
    body.append('</w:footnotes>')
    blob = ''.join(body).encode('utf-8')
    ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
    part = Part(PackURI('/word/footnotes.xml'), ct, blob, doc.part.package)
    doc.part.relate_to(part, RT.FOOTNOTES)


# ------------------------- تحلیل مارک‌داون -------------------------

def parse_md_body(text):
    """تبدیل بدنهٔ مارک‌داون به بلوک‌ها؛ فهرست منابع جدا می‌شود."""
    text = text.split('## منابع')[0]
    blocks = []
    lines = text.split('\n')
    i = 0
    fig_pending = False
    while i < len(lines):
        line = lines[i]
        if line.startswith('# '):
            i += 1
            continue
        if line.startswith('```'):
            i += 1
            while i < len(lines) and not lines[i].startswith('```'):
                i += 1
            i += 1
            fig_pending = True
            continue
        if line.startswith('### '):
            blocks.append(('h3', line[4:].strip()))
        elif line.startswith('## '):
            blocks.append(('h2', line[3:].strip()))
        elif line.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].startswith('|'):
                cells = [c.strip() for c in lines[i].strip('|').split('|')]
                if not all(re.fullmatch(r':?-{3,}:?', c) for c in cells):
                    rows.append(cells)
                i += 1
            blocks.append(('table', rows))
            continue
        elif line.startswith('- '):
            blocks.append(('bullet', line[2:].strip()))
        elif line.strip() == '':
            pass
        else:
            blocks.append(('p', line.strip()))
        if fig_pending and blocks and blocks[-1][0] == 'p' \
                and blocks[-1][1].startswith('شکل ۲-۱'):
            caption = blocks.pop()[1]
            blocks.append(('figure', caption))
            fig_pending = False
        i += 1
    out = []
    for idx, (k, payload) in enumerate(blocks):
        nxt = blocks[idx + 1][0] if idx + 1 < len(blocks) else ''
        if k in ('p', 'h3') and payload.startswith('جدول') and nxt == 'table':
            out.append(('caption', payload))
        else:
            out.append((k, payload))
    return out


def parse_references(text):
    if '## منابع' not in text:
        return []
    ref_text = text.split('## منابع', 1)[1]
    if '\n' in ref_text:
        ref_text = ref_text.split('\n', 1)[1]
    else:
        ref_text = ''
    entries = []
    for chunk in re.split(r'\n\s*\n', ref_text):
        chunk = chunk.strip()
        if not chunk or chunk.startswith('#') or len(chunk) < 30:
            continue
        entries.append(' '.join(chunk.split()))
    return entries


# ------------------------- رندر اجزا -------------------------

def setup_heading_styles(doc):
    for name, pt in (('Heading 2', H2_PT), ('Heading 3', H3_PT)):
        st = doc.styles[name]
        st.font.name = EN_FONT
        st.font.size = Pt(pt)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        st.paragraph_format.line_spacing = Pt(LINE_SPACING_PT)
        st.paragraph_format.space_before = Pt(12 if pt == H2_PT else 9)
        st.paragraph_format.space_after = Pt(3 if pt == H2_PT else 2)
        rPr = st.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:cs'), FA_FONT)
        rFonts.set(qn('w:ascii'), EN_FONT)
        rFonts.set(qn('w:hAnsi'), EN_FONT)
        pPr = st.element.get_or_add_pPr()
        if pPr.find(qn('w:bidi')) is None:
            bidi = OxmlElement('w:bidi')
            bidi.set(qn('w:val'), '1')
            pPr.append(bidi)


def render_blocks(doc, blocks):
    for kind, payload in blocks:
        if kind == 'h2':
            par = doc.add_paragraph(style='Heading 2')
            _set_paragraph_bidi(par)
            add_rich_text(par, payload, fa_pt=H2_PT, en_pt=H2_PT, bold_all=True)
        elif kind == 'h3':
            par = doc.add_paragraph(style='Heading 3')
            _set_paragraph_bidi(par)
            add_rich_text(par, payload, fa_pt=H3_PT, en_pt=H3_PT, bold_all=True)
        elif kind == 'caption':
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 9, 2)
            add_rich_text(par, payload, bold_all=True)
        elif kind == 'table':
            render_table(doc, payload)
        elif kind == 'figure':
            render_conceptual_figure(doc, payload)
        elif kind == 'bullet':
            par = base_paragraph(doc, space_after=0)
            par.paragraph_format.right_indent = Cm(0.75)
            add_rich_text(par, '• ' + payload)
        else:
            par = base_paragraph(doc, space_after=0)
            add_rich_text(par, payload)


def render_table(doc, rows):
    if not rows:
        return
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = 'Table Grid'
    tblPr = table._tbl.tblPr
    if tblPr.find(qn('w:bidiVisual')) is None:
        tblPr.append(tblPr.makeelement(qn('w:bidiVisual'), {}))
    for r, row in enumerate(rows):
        for c, cell_text in enumerate(row):
            if c >= len(rows[0]):
                continue
            cell = table.cell(r, c)
            cell.text = ''
            par = cell.paragraphs[0]
            _set_paragraph_bidi(par)
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER if r == 0 else WD_ALIGN_PARAGRAPH.JUSTIFY
            add_rich_text(par, cell_text, fa_pt=TABLE_FONT_PT, en_pt=10,
                          bold_all=(r == 0))


def render_conceptual_figure(doc, caption):
    table = doc.add_table(rows=2, cols=3)
    table.style = 'Table Grid'
    tblPr = table._tbl.tblPr
    if tblPr.find(qn('w:bidiVisual')) is None:
        tblPr.append(tblPr.makeelement(qn('w:bidiVisual'), {}))
    header = ['اضطراب صفتی (متغیر مستقل)',
              'طرحواره‌های ناسازگار اولیه (متغیر میانجی)',
              'رفتارهای مالی پرخطر (متغیر وابسته)']
    labels = ['مسیر a', 'مسیر غیرمستقیم (a×b)', 'مسیر b']
    for c, txt in enumerate(header):
        cell = table.cell(0, c)
        cell.text = ''
        par = cell.paragraphs[0]
        _set_paragraph_bidi(par)
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_rich_text(par, txt, fa_pt=TABLE_FONT_PT, en_pt=10, bold_all=True)
    for c, txt in enumerate(labels):
        cell = table.cell(1, c)
        cell.text = ''
        par = cell.paragraphs[0]
        _set_paragraph_bidi(par)
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_rich_text(par, txt, fa_pt=TABLE_FONT_PT, en_pt=10)
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 3, 6)
    add_rich_text(par, caption, bold_all=True)


def base_paragraph(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                   space_before=None, space_after=None, rtl=True):
    par = doc.add_paragraph()
    _set_paragraph_bidi(par, rtl)
    par.alignment = align
    fmt = par.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(LINE_SPACING_PT)
    if space_before is not None:
        fmt.space_before = Pt(space_before)
    if space_after is not None:
        fmt.space_after = Pt(space_after)
    return par


def setup_section(section, numbering='none', page_start=None, title_pg=False,
                  vcenter=False):
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.right_margin, section.left_margin = Cm(4), Cm(2.5)
    section.top_margin, section.bottom_margin = Cm(3), Cm(2.5)
    section.footer_distance = Cm(1)
    _set_section_rtl(section)
    for tag in ('w:pgNumType', 'w:titlePg', 'w:vAlign', 'w:footnotePr'):
        for el in section._sectPr.findall(qn(tag)):
            section._sectPr.remove(el)
    # شروع شمارهٔ پاورقی در هر صفحه از ۱ + اعداد فارسی (راهنما ص ۱۷)
    _insert_sectpr_element(
        section, 'w:footnotePr', None,
        (('w:numFmt', {qn('w:val'): 'hindi'}),
         ('w:numRestart', {qn('w:val'): 'eachPage'})))
    if vcenter:
        _insert_sectpr_element(section, 'w:vAlign', {qn('w:val'): 'center'})
    if numbering == 'none':
        clear_footer(section)
    else:
        num_attrs = {qn('w:numFmt'): 'hindi' if numbering == 'hindi' else 'upperRoman'}
        if page_start is not None:
            num_attrs[qn('w:start')] = str(page_start)
        _insert_sectpr_element(section, 'w:pgNumType', num_attrs)
        add_page_number_footer(section)
    if title_pg:
        _insert_sectpr_element(section, 'w:titlePg')
        fpf = section.first_page_footer
        fpf.is_linked_to_previous = False
        fpf.paragraphs[0].text = ''


def page_break(doc):
    par = doc.add_paragraph()
    run = par.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)


def centered(doc, text, size, bold=False, before=0, after=0):
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, before, after)
    add_rich_text(par, text, fa_pt=size, en_pt=size, bold_all=bold)
    return par


def right(doc, text, size, bold=False, before=0, after=0):
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT, before, after)
    add_rich_text(par, text, fa_pt=size, en_pt=size, bold_all=bold)
    return par


def build_front_matter(doc):
    setup_section(doc.sections[0], numbering='none', vcenter=True)
    centered(doc, 'بسم الله الرحمن الرحیم', 16, bold=True)

    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec, numbering='none')
    centered(doc, '[نام دانشگاه]', 10, bold=True, after=3)
    centered(doc, 'وابسته به جهاد دانشگاهی', 10, bold=True, after=18)
    centered(doc, 'پایان‌نامه کارشناسی ارشد دانشکدهٔ روان‌شناسی', 14, bold=True, after=6)
    centered(doc, 'گروه روان‌شناسی بالینی', 14, bold=True, after=18)
    centered(doc, THESIS_TITLE, 24, bold=True, before=12, after=18)
    right(doc, 'استاد راهنما:', 14, after=3)
    centered(doc, 'دکتر سعید وزیری یزدی', 14, bold=True, after=9)
    right(doc, 'نام دانشجو:', 14, after=3)
    centered(doc, '[نام دانشجو]', 14, bold=True, after=9)
    centered(doc, 'مهر ۱۴۰۵', 14, bold=True, before=12)
    page_break(doc)
    centered(doc, '[فرم صورت‌جلسه دفاع پس از برگزاری جلسه دفاع در این محل قرار می‌گیرد]',
             12, before=120)
    page_break(doc)
    centered(doc, '[فرم تعهدنامه اصالت اثر پس از امضا در این محل قرار می‌گیرد]',
             12, before=120)
    page_break(doc)
    right(doc, 'تقدیم به:', 14, bold=True, after=12)
    centered(doc, '[متن تقدیم توسط دانشجو تکمیل می‌شود]', 12, before=60)
    page_break(doc)
    right(doc, 'سپاسگزاری', 14, bold=True, after=12)
    centered(doc, '[متن سپاسگزاری توسط دانشجو تکمیل می‌شود]', 12, before=60)
    page_break(doc)
    right(doc, 'چکیده:', 14, bold=True, after=6)
    par = base_paragraph(doc, space_after=6)
    add_rich_text(par, '[چکیده پس از تکمیل فصل‌های چهارم و پنجم نوشته می‌شود]')
    par = base_paragraph(doc, space_after=0)
    add_rich_text(par, 'کلمات کلیدی: اضطراب صفتی، طرحواره‌های ناسازگار اولیه، '
                       'رفتارهای مالی پرخطر، معامله‌گران بازارهای مالی، '
                       'مدل‌سازی معادلات ساختاری', fa_pt=12, en_pt=11)


def build_lists_section(doc):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec, numbering='roman')
    centered(doc, 'فهرست مطالب', 14, bold=True, after=9)
    add_field_paragraph(doc, ' TOC \\o "1-3" \\h \\z \\u ',
                        '[فهرست مطالب در Word از مسیر References > Table of Contents '
                        'ساخته و به‌روزرسانی شود — با کلیک راست و Update Field]')
    page_break(doc)
    centered(doc, 'فهرست جدول‌ها', 14, bold=True, after=9)
    right(doc, 'جدول ۲-۱. حوزه‌ها و طرحواره‌های ناسازگار اولیه در مدل یانگ '
               '— [شمارهٔ صفحه پس از صفحه‌بندی نهایی در Word تکمیل شود]', 12, after=3)
    right(doc, 'جدول ۳-۱. خلاصهٔ ابزارهای پژوهش '
               '— [شمارهٔ صفحه پس از صفحه‌بندی نهایی در Word تکمیل شود]', 12, after=3)
    page_break(doc)
    centered(doc, 'فهرست شکل‌ها', 14, bold=True, after=9)
    right(doc, 'شکل ۲-۱. مدل مفهومی پژوهش '
               '— [شمارهٔ صفحه پس از صفحه‌بندی نهایی در Word تکمیل شود]', 12, after=3)
    page_break(doc)
    centered(doc, 'فهرست علائم', 14, bold=True, after=9)
    par = base_paragraph(doc)
    add_rich_text(par, '[در صورت وجود علائم و اختصارات، در این بخش تکمیل می‌شود]')


def fa_sort_key(text):
    order = 'آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی'

    def rank(ch):
        return order.find(ch) if ch in order else 99
    cleaned = re.sub(r'[\u200c\s]+', ' ', text).strip()
    return [rank(ch) for ch in cleaned if not ch.isspace()]


def build_unified_references(doc, ref_sets):
    def normalize(s):
        return re.sub(r'\s+', ' ', s.replace('\u200c', ' ').replace('\xa0', ' ')).strip()

    seen, fa_list, en_list, dup_count, flagged = {}, [], [], 0, []
    for refs in ref_sets:
        for entry in refs:
            key = normalize(entry).lower()
            if key in seen:
                dup_count += 1
                continue
            near = [k for k in seen if k[:45] == key[:45]]
            if near:
                flagged.append((seen[near[0]], entry))
            seen[key] = entry
            (en_list if re.match(r'[A-Za-z]', entry) else fa_list).append(entry)
    fa_list.sort(key=lambda e: (fa_sort_key(e.split('،')[0]), fa_sort_key(e)))
    en_list.sort(key=lambda e: e.lower())

    right(doc, 'منابع', 14, bold=True, before=12, after=9)
    right(doc, 'الف) منابع فارسی', 14, bold=True, before=9, after=3)
    for entry in fa_list:
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT, space_after=2)
        add_rich_text(par, entry, fa_pt=12, en_pt=11)
    right(doc, 'ب) منابع لاتین', 14, bold=True, before=9, after=3)
    for entry in en_list:
        par = doc.add_paragraph()
        _set_paragraph_bidi(par, rtl=False)
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt = par.paragraph_format
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        fmt.line_spacing = Pt(LINE_SPACING_PT)
        fmt.left_indent = Cm(1.27)
        fmt.first_line_indent = Cm(-1.27)
        fmt.space_after = Pt(2)
        add_script_runs(par, entry, fa_pt=12, en_pt=11)
    return len(fa_list), len(en_list), dup_count, flagged


def _norm_apos(s):
    """یکدست‌سازی انواع آپوستروف برای تطبیق."""
    return s.replace('\u2019', "'").replace('\u2018', "'")


def _ref_entry_lookup(latin, year, ref_index):
    """یافتن مدخل فهرست منابع متناظر با شکل پاورقی (نام خانوادگی + سال)."""
    base = _norm_apos(latin).split(' et al.')[0].split(' & ')[0]
    base = base.split(', ')[0].strip()
    for line in ref_index:
        if _norm_apos(line).startswith(base) and \
                re.search(r'\b' + re.escape(year) + r'\b', line):
            return line
    return '—'


def write_citation_map(master, ref_index):
    """گزارش نقشهٔ استنادهای خارجی به تفکیک وضعیت (مأموریت ۴۰)."""
    lines = ['# نقشهٔ استنادهای خارجی: شکل فارسی ← پاورقی لاتین (مأموریت ۴۰ و ۳۹-ب)',
             '',
             '> املای لاتین در همهٔ موارد عیناً از مدخل فهرست منابع یکپارچه گرفته شده است.',
             '> وضعیت‌ها: پاورقی شد = در اولین استناد درج شد؛ پاورقی تکراری نشد = استناد وجود دارد ولی پاورقی لاتین مشابه قبلاً آمده؛ در متن استناد نشد = مدخل فهرست بدون استناد.',
             '',
             '| شکل فارسی | سال | مدخل منبع | شکل لاتین پاورقی | وضعیت |',
             '|---|---|---|---|---|']
    for (name, year), latin in sorted(AUTHOR_MAP.items(), key=lambda kv: kv[0][1]):
        entry = _ref_entry_lookup(latin, year, ref_index)
        if (name, year) in master.author_used:
            status = 'پاورقی شد (MATCHED)'
        elif _norm_apos(latin) in master.latin_emitted:
            status = 'پاورقی تکراری نشد (پیش‌تر عیناً آمده)'
        else:
            status = 'در متن استناد نشد'
        lines.append(f'| {name} | {year} | {entry} | {latin} | {status} |')
    lines += ['', '## استنادهای بدون تطبیق (NOT_FOUND — پاورقی نگرفتند)', '']
    if master.not_found:
        lines += ['| شکل فارسی | سال |', '|---|---|']
        for (name, year), c in sorted(master.not_found.items()):
            lines.append(f'| {name} | {year} ×{c} |')
    else:
        lines.append('موردی یافت نشد.')
    lines += ['', '## نام‌های روایی بدون استناد (پاورقی معادل لاتین در اولین رخداد — مأموریت ۳۹-ب)', '']
    lines += ['| شکل فارسی در متن | املای لاتین پاورقی | در فهرست منابع؟ | وضعیت |',
              '|---|---|---|---|']
    for fa, la in NARRATIVE_NAMES:
        status = 'پاورقی شد (اولین رخداد)' if fa in master.name_used else 'رخدادی در متن نداشت'
        lines.append(f'| {fa} | {la} | خیر | {status} |')
    lines += ['', '## نام‌های از پیش لاتین (بدون پاورقی، فقط گزارش)', '',
              '- سازمان همکاری و توسعه اقتصادی (OECD) — استناد «(سازمان همکاری و توسعه اقتصادی، ۲۰۲۱)» چون نام سازمان است و معادل لاتین‌اش در متن به‌صورت اختصار OECD جا ندارد، پاورقی نگرفت.']
    MAP_PATH.write_text('\n'.join(lines) + '\n', encoding='utf-8')


def main():
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    raw = [(k, p, p.read_text(encoding='utf-8')) for k, p in CH_SOURCES]

    master = FootnoteMaster()
    master.scan_terms([t.split('## منابع')[0].split('\n') for _, _, t in raw])
    processed = {}
    for kicker, path, text in raw:
        ch_idx = [r[0] for r in raw].index(kicker)
        body, _, refs_part = text.partition('## منابع')
        new_lines = [master.transform_line(ln, ch_idx, i)
                     for i, ln in enumerate(body.split('\n'))]
        processed[kicker] = '\n'.join(new_lines) + ('## منابع' + refs_part)

    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = EN_FONT
    normal.font.size = Pt(EN_SIZE)
    clean_styles(doc)
    setup_heading_styles(doc)

    build_front_matter(doc)
    build_lists_section(doc)

    ref_sets = []
    first_chapter = True
    for kicker, path, text in raw:
        ref_sets.append(parse_references(text))
        blocks = parse_md_body(processed[kicker])
        title = chapter_title_from_h1(text)
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        setup_section(sec, numbering='hindi',
                      page_start=1 if first_chapter else None, title_pg=True)
        first_chapter = False
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 24, 9)
        pPr = par._p.get_or_add_pPr()
        olvl = OxmlElement('w:outlineLvl')
        olvl.set(qn('w:val'), '0')
        pPr.append(olvl)
        add_rich_text(par, kicker, fa_pt=CHAPTER_KICKER_PT, en_pt=CHAPTER_KICKER_PT,
                      bold_all=True)
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 6, 24)
        add_rich_text(par, title, fa_pt=CHAPTER_TITLE_PT, en_pt=CHAPTER_TITLE_PT,
                      bold_all=True)
        render_blocks(doc, blocks)

    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec, numbering='hindi')
    stats = build_unified_references(doc, ref_sets)

    # فهرست مسطح منابع لاتین برای ستون «مدخل منبع» در نقشهٔ استنادها
    ref_index = [e for refs in ref_sets for e in refs
                 if re.match(r'[A-Za-z]', e.strip())]
    attach_footnotes_part(doc, master)
    write_citation_map(master, ref_index)
    doc.save(OUT_PATH)

    fa_n, en_n, dups, flagged = stats
    print(f'فایل ساخته شد: {OUT_PATH}')
    print(f'منابع: {fa_n} فارسی + {en_n} لاتین | تکرار حذف‌شده: {dups}')
    print(f'پاورقی اصطلاح: {master.term_notes} | پاورقی نویسنده: {master.author_notes}'
          f' | پاورقی نام روایی: {master.name_notes}'
          f' | مجموع: {len(master.entries)} | تکراری پاک‌شده: {master.removed_later}')
    print('NOT_FOUND:', sorted(master.not_found.items()) or 'هیچ')
    for a, b in flagged:
        print('فلگ شباهت:', a[:60], '<->', b[:60])


if __name__ == '__main__':
    main()
