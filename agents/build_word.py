#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
اسکریپت ساخت فایل Word فصل‌های ۱ تا ۳ پایان‌نامه طبق قالب دانشگاه
--------------------------------------------------------------------
ورودی:  output/drafts/chapter1.md، chapter2_full.md، chapter3_full.md
خروجی:  output/final/thesis_ch1_to_ch3.docx

قالب (از templates/university_format_spec.md):
- کاغذ A4؛ حاشیه: راست ۴، چپ ۲٫۵، بالا ۳، پایین ۲٫۵ سانتی‌متر
- فونت فارسی: B Nazanin (روی rFonts.cs با سایز ۱۲)؛ فونت لاتین داخل متن: Times New Roman 11 (روی rFonts.ascii/hAnsi)
- فاصلهٔ خطوط: Exactly 29pt؛ تراز متن: Justify؛ جهت سند و پاراگراف‌ها: راست‌به‌چپ (bidi)
- «فصل اول/دوم/سوم»: نازنین سیاه ۲۸ وسط‌چین؛ عنوان فصل: سیاه ۲۶ وسط‌چین
- تیتر سطح ۲: سیاه ۱۴؛ تیتر سطح ۳: سیاه ۱۳ (سایزها در فایل دانشگاه صریح نبود — ثبت در گزارش)
- شمارهٔ صفحه: وسط پایین، فاصله ۱ سانتی‌متر از لبه، اعداد فارسی (numFmt="hindi")
- صفحهٔ نخست هر فصل شماره ندارد ولی به حساب می‌آید (با w:titlePg)
"""

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree

# ------------------------- تنظیمات قالب -------------------------
FA_FONT = 'B Nazanin'      # فونت فارسی (در rFonts.cs قرار می‌گیرد)
EN_FONT = 'Times New Roman'  # فونت لاتین داخل متن (در rFonts.ascii/hAnsi)
FA_SIZE = 12               # سایز متن فارسی
EN_SIZE = 11               # سایز متن لاتین
CHAPTER_KICKER_PT = 28     # «فصل اول» — سیاه ۲۸
CHAPTER_TITLE_PT = 26      # عنوان فصل — سیاه ۲۶
H2_PT = 14                 # تیتر سطح ۲ (در فایل دانشگاه صریح نبود؛ در گزارش ثبت شده)
H3_PT = 13                 # تیتر سطح ۳ (در فایل دانشگاه صریح نبود؛ در گزارش ثبت شده)
TABLE_FONT_PT = 11         # فونت داخل سلول‌های جدول
LINE_SPACING_PT = 29       # فاصلهٔ خطوط: Exactly 29pt

ROOT = Path(__file__).resolve().parents[1]
CH_SOURCES = [
    ('فصل اول', 'کلیات پژوهش', ROOT / 'output/drafts/chapter1.md'),
    ('فصل دوم', 'مبانی نظری و پیشینه پژوهش', ROOT / 'output/drafts/chapter2_full.md'),
    ('فصل سوم', 'روش‌شناسی پژوهش (بخش اول)', ROOT / 'output/drafts/chapter3_full.md'),
]
OUT_PATH = ROOT / 'output/final/thesis_ch1_to_ch3.docx'

# ------------------------- ابزارهای کمکی -------------------------

def _set_paragraph_bidi(par):
    """فعال‌سازی راست‌به‌چپ برای پاراگراف (افزودن w:bidi به pPr)."""
    pPr = par._p.get_or_add_pPr()
    if pPr.find(qn('w:bidi')) is None:
        el = pPr.makeelement(qn('w:bidi'), {})
        el.set(qn('w:val'), '1')
        pPr.append(el)


def _set_section_rtl(section):
    """راست‌به‌چپ کردن بخش (افزودن w:bidi به sectPr)."""
    sectPr = section._sectPr
    if sectPr.find(qn('w:bidi')) is None:
        el = sectPr.makeelement(qn('w:bidi'), {})
        grid = sectPr.find(qn('w:docGrid'))
        if grid is not None:
            grid.addprevious(el)
        else:
            sectPr.append(el)


def _insert_sectpr_element(section, tag, attrs=None):
    """درج عنصر در sectPr با رعایت ترتیب طرحواره (پیش از w:cols)."""
    sectPr = section._sectPr
    el = sectPr.makeelement(qn(tag), attrs or {})
    cols = sectPr.find(qn('w:cols'))
    if cols is not None:
        cols.addprevious(el)
    else:
        sectPr.append(el)
    return el


def style_run(run, fa_pt=FA_SIZE, en_pt=EN_SIZE, bold=False):
    """تنظیم فونت اجرا: لاتین با TNR و سایز ۱۱؛ فارسی با B Nazanin و سایز ۱۲."""
    run.font.name = EN_FONT                 # rFonts.ascii و rFonts.hAnsi
    run.font.size = Pt(en_pt)               # w:sz و w:szCs هر دو مقدار می‌گیرند
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:                       # در صورت نبودن، دستی ایجاد می‌شود
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:cs'), FA_FONT)         # فونت اسکریپت مختلط (فارسی)
    szCs = rPr.find(qn('w:szCs'))
    if szCs is not None:
        szCs.set(qn('w:val'), str(int(fa_pt * 2)))  # szCs برحسب نیم‌پوینت
    if bold:
        run.font.bold = True


def base_paragraph(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                   space_before=None, space_after=None):
    """ساخت پاراگراف پایه با جهت راست‌به‌چپ، تراز و فاصلهٔ خطوط ثابت ۲۹."""
    par = doc.add_paragraph()
    _set_paragraph_bidi(par)
    par.alignment = align
    fmt = par.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(LINE_SPACING_PT)
    if space_before is not None:
        fmt.space_before = Pt(space_before)
    if space_after is not None:
        fmt.space_after = Pt(space_after)
    return par


def add_rich_text(par, text, fa_pt=FA_SIZE, en_pt=EN_SIZE, bold_all=False):
    """افزودن متن با پشتیبانی از بولد درون‌خطی (**متن**) به پاراگراف."""
    for idx, seg in enumerate(re.split(r'(\*\*.+?\*\*)', text)):
        if not seg:
            continue
        if seg.startswith('**') and seg.endswith('**') and len(seg) > 4:
            style_run(par.add_run(seg[2:-2]), fa_pt, en_pt, bold=True)
        else:
            style_run(par.add_run(seg), fa_pt, en_pt, bold=bold_all)


def add_page_number_footer(section, with_number=True):
    """ساخت پاورقی با فیلد شمارهٔ صفحه (وسط، اعداد فارسی با تنظیم بخش)."""
    footer = section.footer
    footer.is_linked_to_previous = False
    par = footer.paragraphs[0]
    par.text = ''
    _set_paragraph_bidi(par)
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fmt = par.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(LINE_SPACING_PT)
    if with_number:
        fld = etree.SubElement(par._p, qn('w:fldSimple'))
        fld.set(qn('w:instr'), ' PAGE ')
        r = etree.SubElement(fld, qn('w:r'))
        rPr = etree.SubElement(r, qn('w:rPr'))
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
        rFonts.set(qn('w:ascii'), EN_FONT)
        rFonts.set(qn('w:hAnsi'), EN_FONT)
        rFonts.set(qn('w:cs'), FA_FONT)
        szCs = etree.SubElement(rPr, qn('w:szCs'))
        szCs.set(qn('w:val'), str(FA_SIZE * 2))
        t = etree.SubElement(r, qn('w:t'))
        t.text = '۱'  # متن پیش‌فرض تا زمان به‌روزرسانی فیلد در Word

# ------------------------- تحلیل مارک‌داون -------------------------

def parse_md_body(text):
    """تبدیل بدنهٔ مارک‌داون به فهرست بلوک‌ها؛ فهرست منابع جدا شده است."""
    text = text.split('## منابع')[0]          # حذف فهرست منابع فصل از بدنه
    blocks = []
    lines = text.split('\n')
    i = 0
    fig_pending = False                        # پرچم بلوک شکل (پایان بلوک ASCII)
    while i < len(lines):
        line = lines[i]
        if line.startswith('# '):             # تیتر فصل — جداگانه مدیریت می‌شود
            i += 1
            continue
        if line.startswith('```'):             # بلوک ASCII — در خروجی نمی‌آید
            i += 1
            while i < len(lines) and not lines[i].startswith('```'):
                i += 1
            i += 1
            fig_pending = True                 # شکل به‌صورت جدول ساخته می‌شود
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
        # اگر بلافاصله پس از بلوک شکل، عنوان «شکل ۲-۱» آمد: جدول جایگزین درج می‌شود
        if fig_pending and blocks and blocks[-1][0] == 'p' \
                and blocks[-1][1].startswith('شکل ۲-۱'):
            caption = blocks.pop()[1]
            blocks.append(('figure', caption))
            fig_pending = False
        i += 1
    # عنوان جدول (خط آغازشونده با «جدول») به‌عنوان کپشن علامت‌گذاری می‌شود
    out = []
    for idx, (kind, payload) in enumerate(blocks):
        if kind == 'p' and payload.startswith('جدول') and idx + 1 < len(blocks) \
                and blocks[idx + 1][0] == 'table':
            out.append(('caption', payload))
        elif kind == 'h3' and payload.startswith('جدول') and idx + 1 < len(blocks) \
                and blocks[idx + 1][0] == 'table':
            out.append(('caption', payload))
        else:
            out.append((kind, payload))
    return out


def parse_references(text):
    """استخراج مدخل‌های فهرست منابع یک فصل از متن مارک‌داون."""
    if '## منابع' not in text:
        return []
    ref_text = text.split('## منابع', 1)[1]
    # باقی‌ماندهٔ همان سطرِ سرفصل (مثلاً « فصل اول») مدخل منبع نیست؛ حذف می‌شود.
    if '\n' in ref_text:
        ref_text = ref_text.split('\n', 1)[1]
    else:
        ref_text = ''
    entries = []
    for chunk in re.split(r'\n\s*\n', ref_text):
        chunk = chunk.strip()
        if not chunk or chunk.startswith('#'):
            continue
        if len(chunk) < 30:          # خرده‌متن‌های کوتاه، مدخل منبع نیستند
            continue
        entries.append(' '.join(chunk.split()))
    return entries

# ------------------------- رندر اجزای سند -------------------------

def render_blocks(doc, blocks):
    """رندر بلوک‌های تحلیل‌شده در بدنهٔ سند."""
    for kind, payload in blocks:
        if kind == 'h2':
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT, space_before=12, space_after=3)
            add_rich_text(par, payload, fa_pt=H2_PT, en_pt=H2_PT, bold_all=True)
        elif kind == 'h3':
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT, space_before=9, space_after=2)
            add_rich_text(par, payload, fa_pt=H3_PT, en_pt=H3_PT, bold_all=True)
        elif kind == 'caption':
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, space_before=9, space_after=2)
            add_rich_text(par, payload, bold_all=True)
        elif kind == 'table':
            render_table(doc, payload)
        elif kind == 'figure':
            render_conceptual_figure(doc, payload)
        elif kind == 'bullet':
            par = base_paragraph(doc, space_after=0)
            par.paragraph_format.right_indent = Cm(0.75)
            add_rich_text(par, '• ' + payload)
        else:  # پاراگراف عادی
            par = base_paragraph(doc, space_after=0)
            add_rich_text(par, payload)


def render_table(doc, rows):
    """رندر جدول مارک‌داون به‌صورت جدول Word با کادر و جهت راست‌به‌چپ."""
    if not rows:
        return
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = 'Table Grid'
    tblPr = table._tbl.tblPr
    if tblPr.find(qn('w:bidiVisual')) is None:    # جهت راست‌به‌چپ جدول
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
    """شکل ۲-۱: جدول سه‌ستونهٔ متغیرها + کپشن زیر آن (بلوک جایگزین تصویر)."""
    table = doc.add_table(rows=2, cols=3)
    table.style = 'Table Grid'
    tblPr = table._tbl.tblPr
    if tblPr.find(qn('w:bidiVisual')) is None:
        tblPr.append(tblPr.makeelement(qn('w:bidiVisual'), {}))
    # ستون‌ها از راست به چپ: متغیر مستقل، میانجی، وابسته
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
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, space_before=3, space_after=6)
    add_rich_text(par, caption, bold_all=True)   # عنوان شکل: زیر شکل، وسط‌چین

# ------------------------- ساخت سند -------------------------

def setup_section(section, first=False, page_start=None):
    """تنظیم صفحه (A4 و حاشیه‌ها)، جهت بخش و شماره‌گذاری فارسی."""
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.right_margin, section.left_margin = Cm(4), Cm(2.5)
    section.top_margin, section.bottom_margin = Cm(3), Cm(2.5)
    section.footer_distance = Cm(1)              # فاصلهٔ شمارهٔ صفحه از لبهٔ پایین
    _set_section_rtl(section)
    num_attrs = {qn('w:numFmt'): 'hindi'}        # اعداد فارسی صفحه
    if page_start is not None:
        num_attrs[qn('w:start')] = str(page_start)
    _insert_sectpr_element(section, 'w:pgNumType', num_attrs)
    if first:
        # صفحهٔ نخست فصل: بدون نمایش شماره ولی به حساب می‌آید
        _insert_sectpr_element(section, 'w:titlePg')
        fpf = section.first_page_footer
        fpf.is_linked_to_previous = False
        fpf.paragraphs[0].text = ''
        add_page_number_footer(section, with_number=True)
    else:
        add_page_number_footer(section, with_number=False)


def build_title_page(doc):
    """صفحهٔ عنوان ساده با جانگهدارهای [نام دانشگاه] و [نام دانشجو]."""
    doc.sections[0].page_width, doc.sections[0].page_height = Cm(21), Cm(29.7)
    sec = doc.sections[0]
    sec.right_margin, sec.left_margin = Cm(4), Cm(2.5)
    sec.top_margin, sec.bottom_margin = Cm(3), Cm(2.5)
    _set_section_rtl(sec)

    def centered(text, size, bold=False, before=0, after=0):
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER,
                             space_before=before, space_after=after)
        add_rich_text(par, text, fa_pt=size, en_pt=size, bold_all=bold)
        return par

    for _ in range(4):
        base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER)      # فاصلهٔ ابتدای صفحه
    centered('نام دانشگاه: [نام دانشگاه]', 14, bold=True, after=12)
    centered('تعیین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه در رابطه بین '
             'اضطراب صفتی و رفتارهای مالی پرخطر در معامله‌گران بازارهای مالی',
             24, bold=True, before=24, after=18)
    centered('پایان‌نامه کارشناسی ارشد رشتهٔ روان‌شناسی بالینی', 14, bold=True, after=24)
    centered('استاد راهنما: دکتر سعید وزیری یزدی', 14, before=18, after=6)
    centered('نگارنده: [نام دانشجو]', 14, after=24)
    centered('سال: ۱۴۰۵', 14, before=12)


def fa_sort_key(text):
    """کلید مرتب‌سازی الفبایی فارسی بر اساس ترتیب الفبای فارسی."""
    order = 'آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی'
    def rank(ch):
        ch = ch.replace('ۀ', 'ه').replace('ی', 'ی')
        return order.find(ch) if ch in order else 99
    cleaned = re.sub(r'[\u200c\s]+', ' ', text).strip()
    return [rank(ch) for ch in cleaned if not ch.isspace()]


def build_unified_references(doc, ref_sets):
    """ادغام، حذف تکرار و مرتب‌سازی فهرست منابع واحد در انتهای سند."""
    def normalize(s):
        return re.sub(r'\s+', ' ', s.replace('\u200c', ' ').replace('\xa0', ' ')).strip()

    seen, fa_list, en_list, dup_count, flagged = {}, [], [], 0, []
    for refs in ref_sets:
        for entry in refs:
            key = normalize(entry)
            key_low = key.lower()
            if key_low in seen:
                dup_count += 1
                continue
            # فلگ موارد شبیه ولی نه یکسان (۴۵ نویسهٔ نخست یکسان)
            near = [k for k in seen if k[:45] == key_low[:45]]
            if near:
                flagged.append((seen[near[0]], entry))
            seen[key_low] = entry
            if re.match(r'[A-Za-z]', entry):
                en_list.append(entry)
            else:
                fa_list.append(entry)
    fa_list.sort(key=lambda e: (fa_sort_key(e.split('،')[0]), fa_sort_key(e)))
    en_list.sort(key=lambda e: e.lower())

    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=9)
    add_rich_text(par, 'منابع', fa_pt=14, en_pt=14, bold_all=True)
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT, space_before=9, space_after=3)
    add_rich_text(par, 'الف) منابع فارسی', fa_pt=14, en_pt=14, bold_all=True)
    for entry in fa_list:
        par = base_paragraph(doc, space_after=2)
        add_rich_text(par, entry, fa_pt=12, en_pt=11)
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT, space_before=9, space_after=3)
    add_rich_text(par, 'ب) منابع لاتین', fa_pt=14, en_pt=14, bold_all=True)
    for entry in en_list:
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.LEFT, space_after=2)
        run = par.add_run(entry)
        run.font.name = EN_FONT
        run.font.size = Pt(11)
    return len(fa_list), len(en_list), dup_count, flagged


def main():
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()

    # تنظیم سبک پایهٔ سند (فونت پیش‌فرض برای پاراگراف‌های بدون تنظیم صریح)
    normal = doc.styles['Normal']
    normal.font.name = EN_FONT
    normal.font.size = Pt(EN_SIZE)

    build_title_page(doc)

    ref_sets = []
    for kicker, title, path in CH_SOURCES:
        text = path.read_text(encoding='utf-8')
        ref_sets.append(parse_references(text))
        blocks = parse_md_body(text)

        section = doc.add_section(WD_SECTION.NEW_PAGE)   # هر فصل از صفحهٔ جدید
        first_chapter = (kicker == 'فصل اول')
        setup_section(section, first=True, page_start=1 if first_chapter else None)

        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, space_before=24, space_after=9)
        add_rich_text(par, kicker, fa_pt=CHAPTER_KICKER_PT, en_pt=CHAPTER_KICKER_PT,
                      bold_all=True)
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=24)
        add_rich_text(par, title, fa_pt=CHAPTER_TITLE_PT, en_pt=CHAPTER_TITLE_PT,
                      bold_all=True)
        render_blocks(doc, blocks)

    # بخش منابع واحد
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(section, first=False)
    stats = build_unified_references(doc, ref_sets)

    doc.save(OUT_PATH)
    fa_n, en_n, dups, flagged = stats
    print(f'فایل ساخته شد: {OUT_PATH}')
    print(f'منابع: {fa_n} فارسی + {en_n} لاتین | تکرار حذف‌شده: {dups}')
    for a, b in flagged:
        print('فلگ شباهت:', a[:60], '<->', b[:60])


if __name__ == '__main__':
    main()
