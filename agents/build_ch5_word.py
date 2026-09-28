# -*- coding: utf-8 -*-
"""ساخت سند Word فصل پنجم آموزشی (مأموریت ۴۹).

ورودی: output/drafts/chapter5_simulated_demo.md
خروجی: output/final/chapter5_simulated_full_demo.docx
رجیستری اصطلاحات: data/processed/ch5_english_term_registry.md

قالب همان فصل‌های ۱ تا ۴ است: B Nazanin/Times New Roman، راست‌به‌چپ،
تیترهای راست‌چین، جداول APA سه‌خطی، عنوان جدول بالا، حاشیه‌ها و فاصلهٔ
خطوط مطابق قالب دانشگاه. تفاوت‌ها با سازندهٔ فصل ۴: پشتیبانی از تیتر
سطح ۳ (###)، بولت‌ها، و بخش منابع فارسی/لاتین در پایان فصل.
"""

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_word import (FA_SIZE, EN_SIZE, TABLE_FONT_PT, H2_PT, H3_PT,
                        base_paragraph, clean_styles, setup_heading_styles,
                        setup_section, _set_section_rtl, _set_paragraph_bidi,
                        _reorder_ppr, add_rich_text, add_script_runs,
                        attach_footnotes_part, right)
from build_ch4_word import (Ch4TermMaster, add_text_islands, add_md_rich,
                            render_apa_table)

ROOT = Path(__file__).resolve().parents[1]
MD_PATH = ROOT / 'output' / 'drafts' / 'chapter5_simulated_demo.md'
OUT_PATH = ROOT / 'output' / 'final' / 'chapter5_simulated_full_demo.docx'
CH5_REGISTRY_OUT = ROOT / 'data' / 'processed' / 'ch5_english_term_registry.md'


class Ch5TermMaster(Ch4TermMaster):
    """قانون پاورقی فصل ۵: اصطلاحات دارندهٔ پاورقی در فصل ۱-۳ فارسیِ تنها
    می‌مانند و اصطلاح کاملاً تازه فقط در نخستین کاربرد پاورقی می‌گیرد.
    متن فصل ۵ آموزشی فارسیِ خالص نوشته شده و انتظار می‌رود پاورقی تازه‌ای
    ساخته نشود؛ رجیستری جداگانه همین را مستند می‌کند."""

    def write_registry(self):
        L = ['# رجیستری الگوهای «فارسی (English)» در فصل ۵ (مأموریت ۴۹)', '',
             'مبنا: رجیستری سراسریِ فصل ۱-۳. دستهٔ الف = از پیش پاورقی‌گرفته → فقط حذف '
             'پرانتز؛ دستهٔ ب = جدید → پاورقی در نخستین کاربرد؛ دسته‌های ج (آماره/فرمول)، '
             'د (استناد) و هـ (جدول) دست‌نخورده. متن فصل ۵ فارسیِ خالص است و پرانتز لاتین '
             'تازه‌ای ندارد؛ بنابراین انتظار می‌رود شمار پاورقی‌های تازه صفر باشد.', '',
             '| اصطلاح فارسی | معادل انگلیسی | در فصل ۱-۳ بود؟ | اقدام | محل (سطر md) |',
             '|---|---|---|---|---|']
        for r in sorted(self.rows, key=lambda x: x['ln']):
            L.append(f"| {r['fa'] or '—'} | {r['en']} | {r['known']} | {r['act']} "
                     f"| سطر {r['ln']} |")
        L += ['', '## یادداشت‌ها', '',
              '- شمارۀ پاورقی در footnote partِ همین سند است؛ نمایشِ Word به‌ازای هر صفحه '
              'از ۱ شروع می‌شود (numRestart=eachPage، اعداد هندی) — هم‌راستا با فصل ۱-۴.',
              '- استنادهای درون‌متنی فصل ۵ همگی فارسی‌اند و مشمول قانون پاورقی نیستند.',
              '- فایل md منبع دست‌نخورده است؛ تبدیل در زمان ساخت انجام می‌شود.', '']
        CH5_REGISTRY_OUT.write_text('\n'.join(L), encoding='utf-8')


def render_heading(doc, text, level):
    """تیتر راست‌چین با bidi مستقیم (نگهبان: هرگز تراز پیش‌فرض چپ)."""
    style = 'Heading 2' if level == 2 else 'Heading 3'
    pt = H2_PT if level == 2 else H3_PT
    par = doc.add_paragraph(style=style)
    _set_paragraph_bidi(par)
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _reorder_ppr(par._p.get_or_add_pPr())
    add_text_islands(par, text, fa_pt=pt, en_pt=pt, bold=True)


def render_references(doc, ref_lines):
    """بخش منابع: فارسی راست‌چین، لاتین چپ‌چین LTR با تورفتگی آویخته."""
    for line in ref_lines:
        s = line.strip()
        if not s:
            continue
        if s.startswith('### '):
            right(doc, s[4:], 14, bold=True, before=9, after=3)
            continue
        if s.startswith('#'):
            continue
        if re.match(r'[A-Za-z]', s):
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.LEFT,
                                 space_after=2, rtl=False)
            fmt = par.paragraph_format
            fmt.left_indent = Cm(1.27)
            fmt.first_line_indent = Cm(-1.27)
            add_script_runs(par, s, fa_pt=12, en_pt=11)
        else:
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT, space_after=2)
            add_rich_text(par, s, fa_pt=12, en_pt=11)


def build(md_path=MD_PATH, out_path=OUT_PATH):
    doc = Document()
    clean_styles(doc)
    setup_heading_styles(doc)
    setup_section(doc.sections[0])
    _set_section_rtl(doc.sections[0])

    master = Ch5TermMaster()
    raw_lines = Path(md_path).read_text(encoding='utf-8').splitlines()
    # جداسازی بخش منابع پیش از تبدیل اصطلاحات: مدخل‌های لاتین (مانند
    # ‎(Eds.)‎ و ‎(4th ed.)‎) نباید دست‌خورده یا پاورقی شوند.
    try:
        ref_at = next(i for i, l in enumerate(raw_lines)
                      if l.strip() == '## منابع')
    except StopIteration:
        ref_at = len(raw_lines)
    raw_body, ref_lines = raw_lines[:ref_at], raw_lines[ref_at + 1:]
    master.inventory_skipped(raw_body)
    body_lines = master.transform(raw_body)
    master.write_registry()

    n_tables = 0
    i = 0
    while i < len(body_lines):
        line = body_lines[i].strip()
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
        if line.startswith('### '):
            render_heading(doc, line[4:], 3)
            continue
        if line.startswith('## '):
            render_heading(doc, line[3:], 2)
            continue
        if line.startswith('|'):
            rows = []
            cur = line
            while cur.startswith('|'):
                cells = [c.strip() for c in cur.strip('|').split('|')]
                if not set(''.join(cells)) <= set('-: '):
                    rows.append(cells)
                if i >= len(body_lines):
                    break
                cur = body_lines[i].strip()
                if cur.startswith('|'):
                    i += 1
            render_apa_table(doc, rows)
            n_tables += 1
            continue
        if line.startswith('**جدول'):
            par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, 9, 2)
            add_md_rich(par, line, fa_pt=TABLE_FONT_PT, en_pt=10)
            continue
        if line.startswith('- '):
            par = base_paragraph(doc, space_after=0)
            par.paragraph_format.right_indent = Cm(0.75)
            add_md_rich(par, '• ' + line[2:])
            continue
        par = base_paragraph(doc, space_after=0)
        add_md_rich(par, line)

    if ref_lines:
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT, 12, 9)
        add_md_rich(par, '**منابع**', fa_pt=14, en_pt=14)
        render_references(doc, ref_lines)

    attach_footnotes_part(doc, master)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f'سند فصل پنجم ساخته شد: {out_path} | جدول‌ها: {n_tables} | '
          f'پاورقی جدید (ب): {master.footnoted_new} | حذف تکراری از ۱-۳ (الف): '
          f'{master.removed_known} | حذف کاربرد بعدی (ب): {master.removed_later}')


if __name__ == '__main__':
    build()
