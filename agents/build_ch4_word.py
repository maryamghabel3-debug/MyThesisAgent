# -*- coding: utf-8 -*-
"""ساخت سند Word فصل چهارم (نسخهٔ نمایشی دادهٔ شبیه‌سازی‌شده) — مأموریت ۴۳.

خروجی: output/final/chapter4_simulated.docx
همان قالب فصل‌های ۱-۳: B Nazanin/Times New Roman، راست‌به‌چپ، تیترهای
راست‌چین با bidi معتبر، جداول APA سه‌خطی، شکل با عنوان زیرین وسط‌چین.
"""

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_word import (FA_SIZE, EN_SIZE, TABLE_FONT_PT, LINE_SPACING_PT,
                        add_rich_text, base_paragraph, clean_styles,
                        _reorder_ppr, _make_run, _set_paragraph_bidi,
                        _set_section_rtl, render_table, setup_heading_styles,
                        setup_section)

ROOT = Path(__file__).resolve().parents[1]
MD_PATH = ROOT / 'output' / 'drafts' / 'chapter4_simulated.md'
FIG_PATH = ROOT / 'output' / 'drafts' / 'ch4_fig1_path_model.png'
OUT_PATH = ROOT / 'output' / 'final' / 'chapter4_simulated.docx'

# ---------- جزیرۀ لاتین برای پرانتز/براکت‌های آماری (مأموریت ۴۳-ج) ----------
# هر گروه متوازنِ کاملاً لاتین (حرف/رقم/نمادِ لاتین، بدون نویسۀ فارسی) به‌همراہ
# هر دو پرانتز در یک رانِ چپ‌به‌راست با فونت Times New Roman نوشته می‌شود تا
# در پاراگراف راست‌به‌چپ، پرانتزِ باز و بسته از محتوا جدا نشوند. هیچ کاراکتر
# کنترلیِ جداره (LRM/RLM/isolate) افزوده نمی‌شود؛ راه‌حل کاملاً ساختاری است.
_GROUP_RE = re.compile(r'\([^()]*\)|\[[^\]]*\]')


def _is_fa(ch):
    o = ord(ch)
    return (0x0600 <= o <= 0x06FF or 0xFB50 <= o <= 0xFDFF
            or 0xFE70 <= o <= 0xFEFF or o == 0x200C)


def _latin_island(group):
    """True اگر محتوای گروه فقط لاتین/رقم/نماد باشد و دست‌کم یک حرف یا رقم داشته باشد."""
    inner = group[1:-1]
    if not inner or any(_is_fa(c) for c in inner):
        return False
    return any(c.isascii() and c.isalnum() for c in inner)


def add_text_islands(par, text, fa_pt=FA_SIZE, en_pt=EN_SIZE, bold=False):
    """مانند add_rich_text، جز آنکه گروه‌های لاتینِ داخل ()/[] یک‌جا در یک ران
    می‌نشینند (جزیرۀ چپ‌به‌راست) و بقیۀ متن قطعه‌بندی معمولی می‌شود."""
    pos = 0
    for m in _GROUP_RE.finditer(text):
        if not _latin_island(m.group(0)):
            continue
        if m.start() > pos:
            add_rich_text(par, text[pos:m.start()], fa_pt, en_pt, bold_all=bold)
        _make_run(par, m.group(0), 'L', en_pt, bold)
        pos = m.end()
    if pos < len(text):
        add_rich_text(par, text[pos:], fa_pt, en_pt, bold_all=bold)


def add_md_rich(par, text, fa_pt=FA_SIZE, en_pt=EN_SIZE):
    """متن مارک‌داون با بخش‌های **بولد** را با ران‌های فارسی/لاتین می‌افزاید.
    (لایۀ جزیرۀ لاتین روی add_rich_text — مأموریت ۴۳-ج)"""
    pos = 0
    for m in re.finditer(r'\*\*(.+?)\*\*', text):
        if m.start() > pos:
            add_text_islands(par, text[pos:m.start()], fa_pt, en_pt)
        add_text_islands(par, m.group(1), fa_pt, en_pt, bold=True)
        pos = m.end()
    if pos < len(text):
        add_text_islands(par, text[pos:], fa_pt, en_pt)


def apa_table_borders(table):
    """قالب APA: فقط خط بالای جدول، خط زیر سرستون و خط پایین جدول."""
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old)
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'bottom'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '12')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        borders.append(el)
    tblPr.append(borders)
    for cell in table.rows[0].cells:
        tcPr = cell._tc.get_or_add_tcPr()
        for old in tcPr.findall(qn('w:tcBorders')):
            tcPr.remove(old)
        cb = OxmlElement('w:tcBorders')
        b = OxmlElement('w:bottom')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '12')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        cb.append(b)
        tcPr.append(cb)


def render_apa_table(doc, rows):
    """جدول APA با bidi و بدون خطوط عمودی.
    (مأموریت ۴۳-ج: سلول‌های کاملاً لاتین/عددی به‌صورت پاراگراف چپ‌به‌راست
    رندر می‌شوند تا ستاره/براکت انتهایی جابه‌جا افت نکند؛ سلول‌های فارسی
    همان رفتار قبلی را دارند.)"""
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    # نکته: ران‌های ساده (نه add_rich_text)؛ ترکیب ران‌های دوزبانهٔ صریح با
    # جدول‌های بیش از دو سطر در LibreOffice بارگذاری را مختل می‌کند.
    for r, row in enumerate(rows):
        for c, cell_text in enumerate(row):
            cell = table.cell(r, c)
            cell.text = ''
            par = cell.paragraphs[0]
            txt = cell_text.strip()
            if any(_is_fa(ch) for ch in txt):
                _set_paragraph_bidi(par)
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = par.add_run(txt)
            run.bold = (r == 0)
            run.font.size = Pt(11)
    apa_table_borders(table)
    # bidiVisual پس از tblBorders الحاق شود (ترتیب تحمل‌شده توسط LibreOffice)
    if table._tbl.tblPr.find(qn('w:bidiVisual')) is None:
        table._tbl.tblPr.append(table._tbl.makeelement(qn('w:bidiVisual'), {}))


def build(md_path=MD_PATH, fig_path=FIG_PATH, out_path=OUT_PATH):
    doc = Document()
    clean_styles(doc)
    setup_heading_styles(doc)
    setup_section(doc.sections[0])
    _set_section_rtl(doc.sections[0])

    lines = Path(md_path).read_text(encoding='utf-8').splitlines()
    i = 0
    pending_fig_caption = None
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line:
            continue
        if line.startswith('# '):
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 6, 12)
            add_md_rich(par, '**' + line[2:] + '**', fa_pt=16, en_pt=15)
            continue
        if line.startswith('> '):
            par = base_paragraph(doc, space_after=6)
            add_md_rich(par, line[2:])
            continue
        if line.startswith('@subtitle '):
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 0, 10)
            add_md_rich(par, '**' + line[len('@subtitle '):] + '**', fa_pt=14, en_pt=13)
            continue
        if line.startswith('@pagebreak'):
            doc.add_page_break()
            continue
        if line.startswith('## '):
            par = doc.add_paragraph(style='Heading 2')
            _set_paragraph_bidi(par)
            par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _reorder_ppr(par._p.get_or_add_pPr())
            add_text_islands(par, line[3:], fa_pt=14, en_pt=14, bold=True)
            continue
        if line.startswith('|'):
            rows = []
            cur = line
            while cur.startswith('|'):
                cells = [c.strip() for c in cur.strip('|').split('|')]
                if not set(''.join(cells)) <= set('-: '):
                    rows.append(cells)
                if i >= len(lines):
                    break
                cur = lines[i].strip()
                if cur.startswith('|'):
                    i += 1
                # اگر سطر بعدی جدول نبود، همان سطر برای دورۀ بیرونی می‌ماند؛
                # افزودن i+=1 بدون این شرط، نخستین پاراگراف پس از جدول را می‌بلعید
                # (رفع اشکال مأموریت ۴۳-ج).
            render_apa_table(doc, rows)
            continue
        if line.startswith('!['):
            doc.add_picture(str(fig_path), width=Cm(14))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # کپشن می‌تواند از قبل در pending باشد یا در سطر بعد از تصویر بیاید
            if not pending_fig_caption and i < len(lines) and lines[i].strip().startswith('**شکل'):
                pending_fig_caption = lines[i].strip()
                i += 1
            if pending_fig_caption:
                par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 3, 6)
                add_md_rich(par, pending_fig_caption, fa_pt=TABLE_FONT_PT, en_pt=10)
                pending_fig_caption = None
            continue
        if line.startswith('**شکل'):
            pending_fig_caption = line
            continue
        if line.startswith('**جدول'):
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 9, 2)
            add_md_rich(par, line, fa_pt=TABLE_FONT_PT, en_pt=10)
            continue
        par = base_paragraph(doc, space_after=0)
        add_md_rich(par, line)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print('سند فصل چهارم ساخته شد:', out_path)


if __name__ == '__main__':
    build()
