#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
اسکریپت ساخت فایل Word فصل‌های ۱ تا  پایان‌نامه طبق راهنمای واقعی دانشگاه
---------------------------------------------------------------------------
ورودی:  output/drafts/chapter1.md، chapter2_full.md، chapter3_full.md
خروجی:  output/final/thesis_ch1_to_ch3.docx

مبنای قالب: templates/university_format_spec.md + templates/thesis_structure_order.md
(راستی‌آزمایی‌شده با templates/university_template.pdf)

ویژگی‌های این بازسازی (مأموریت ۳۸):
- صفحات مقدماتی: بسم‌الله، عنوان فارسی، فرم دفاع، تعهدنامه، تقدیم، سپاسگزاری، چکیده
- فهرست مطالب (فیلد TOC قابل به‌روزرسانی)، فهرست جدول‌ها/شکل‌ها/علائم
- شماره‌گذاری: مقدماتی بدون شماره؛ فهرست‌ها با اعداد رومی؛ متن اصلی پیوسته از ۱ با اعداد فارسی
- صفحهٔ نخست هر فصل بدون نمایش شماره ولی به حساب (w:titlePg)
- پاورقی خودکار معادل انگلیسی اصطلاحات (footnotes.xml با دستکاری XML)
- منابع لاتین: LTR، چپ‌چین، TNR 11، تورفتگی آویزان APA
- تیترهای درون‌فصل: Heading واقعی Word (برای TOC)، راست‌چین و RTL
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
FA_FONT = 'B Nazanin'        # فونت فارسی (روی rFonts.cs)
EN_FONT = 'Times New Roman'  # فونت لاتین (روی rFonts.ascii/hAnsi)
FA_SIZE = 12
EN_SIZE = 11
CHAPTER_KICKER_PT = 28       # «فصل اول» — سیاه ۲۸ وسط‌چین
CHAPTER_TITLE_PT = 26        # عنوان فصل — سیاه ۲۶ وسط‌چین
H2_PT = 14                   # در فایل دانشگاه صریح نبود؛ ثبت در گزارش
H3_PT = 13                   # در فایل دانشگاه صریح نبود؛ ثبت در گزارش
TABLE_FONT_PT = 11
LINE_SPACING_PT = 29         # فاصلهٔ خطوط Exactly 29pt
FOOTNOTE_PT = 10             # اندازهٔ متن پاورقی (در راهنما صریح نیست؛ ثبت در گزارش)

ROOT = Path(__file__).resolve().parents[1]
CH_SOURCES = [
    ('فصل اول', ROOT / 'output/drafts/chapter1.md'),
    ('فصل دوم', ROOT / 'output/drafts/chapter2_full.md'),
    ('فصل سوم', ROOT / 'output/drafts/chapter3_full.md'),
]
THESIS_TITLE = ('تعیین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه در رابطه بین '
                'اضطراب صفتی و رفتارهای مالی پرخطر در معامله‌گران بازارهای مالی')
OUT_PATH = ROOT / 'output/final/thesis_ch1_to_ch3.docx'

# توکن درونی نشانهٔ پاورقی هنگام رندر
FN_TOKEN = '\x01FN{}\x01'
FN_TOKEN_RE = re.compile(r'\x01FN(\d+)\x01')
ELIG_RE = re.compile(r'\(([^()\n]+)\)')


def chapter_title_from_h1(text):
    """عنوان رسمی فصل را از سرفصل `#` فایل مارک‌داون می‌گیرد (بعد از «:»)."""
    for line in text.splitlines():
        if line.startswith('# ') and ':' in line:
            return line.split(':', 1)[1].strip()
    return ''


# ------------------------- ابزارهای کمکی XML/قالب -------------------------

def _set_paragraph_bidi(par, rtl=True):
    """افزودن w:bidi (یا w:bidi val=0 برای چپ‌به‌راست) به pPr پاراگراف."""
    pPr = par._p.get_or_add_pPr()
    old = pPr.find(qn('w:bidi'))
    if old is not None:
        pPr.remove(old)
    el = pPr.makeelement(qn('w:bidi'), {})
    if rtl:
        el.set(qn('w:val'), '1')
    else:
        el.set(qn('w:val'), '0')
    pPr.append(el)


def _set_section_rtl(section):
    """راست‌به‌چپ کردن بخش (افزودن w:bidi به sectPr با ترتیب طرحواره)."""
    if section._sectPr.find(qn('w:bidi')) is None:
        _insert_sectpr_element(section, 'w:bidi')


# ترتیب مجاز فرزندهای w:sectPr طبق طرحوارهٔ OOXML
_SECTPR_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr',
                 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
                 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
                 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings',
                 'sectPrChange']


def _insert_sectpr_element(section, tag, attrs=None):
    """درج عنصر در sectPr دقیقاً مطابق ترتیب طرحوارهٔ CT_SectPr."""
    sectPr = section._sectPr
    el = sectPr.makeelement(qn(tag), attrs or {})
    local = lambda e: e.tag.split('}')[-1]
    for child in sectPr:
        if _SECTPR_ORDER.index(local(child)) > _SECTPR_ORDER.index(local(el)):
            child.addprevious(el)
            return el
    sectPr.append(el)
    return el


def style_run(run, fa_pt=FA_SIZE, en_pt=EN_SIZE, bold=False):
    """تنظیم فونت اجرا: لاتین TNR و فارسی B Nazanin با سایزهای جداگانه."""
    run.font.name = EN_FONT
    run.font.size = Pt(en_pt)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:cs'), FA_FONT)
    szCs = rPr.find(qn('w:szCs'))
    if szCs is not None:
        szCs.set(qn('w:val'), str(int(fa_pt * 2)))
    if bold:
        run.font.bold = True


def base_paragraph(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                   space_before=None, space_after=None, rtl=True):
    """پاراگراف پایه: جهت، تراز و فاصلهٔ خطوط ثابت ۲۹."""
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


def _add_footnote_ref_run(par, fn_id):
    """اجرای بالانویس با ارجاع به پاورقی (w:footnoteReference)."""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), EN_FONT)
    rFonts.set(qn('w:hAnsi'), EN_FONT)
    rFonts.set(qn('w:cs'), FA_FONT)
    rPr.append(rFonts)
    va = OxmlElement('w:vertAlign')
    va.set(qn('w:val'), 'superscript')
    rPr.append(va)
    r.append(rPr)
    ref = OxmlElement('w:footnoteReference')
    ref.set(qn('w:id'), str(fn_id))
    r.append(ref)
    par._p.append(r)


def add_rich_text(par, text, fa_pt=FA_SIZE, en_pt=EN_SIZE, bold_all=False):
    """افزودن متن با بولد درون‌خطی (**…**) و توکن‌های پاورقی به پاراگراف."""
    for seg in FN_TOKEN_RE.split(text):
        if seg == '':
            continue
        if seg.isdigit():
            _add_footnote_ref_run(par, int(seg))
            continue
        for sub in re.split(r'(\*\*.+?\*\*)', seg):
            if not sub:
                continue
            if sub.startswith('**') and sub.endswith('**') and len(sub) > 4:
                style_run(par.add_run(sub[2:-2]), fa_pt, en_pt, bold=True)
            else:
                style_run(par.add_run(sub), fa_pt, en_pt, bold=bold_all)


def add_field_paragraph(doc, instr, placeholder, title_align=WD_ALIGN_PARAGRAPH.RIGHT):
    """پاراگراف حاوی فیلد Word (مثل TOC) با متن جای‌دار تا به‌روزرسانی."""
    par = doc.add_paragraph()
    _set_paragraph_bidi(par)
    par.alignment = title_align
    for tag, extra in (('w:fldChar', {'w:fldCharType': 'begin'}),):
        r = OxmlElement('w:r')
        el = OxmlElement(tag)
        for k, v in extra.items():
            el.set(qn(k), v)
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
    run = par.add_run(placeholder)
    style_run(run)
    r = OxmlElement('w:r')
    el = OxmlElement('w:fldChar')
    el.set(qn('w:fldCharType'), 'end')
    r.append(el)
    par._p.append(r)
    return par


def add_page_number_footer(section):
    """پاورقی شمارهٔ صفحه: وسط، فاصلهٔ ۱ سانتی‌متر از لبهٔ پایین."""
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
    rFonts.set(qn('w:ascii'), EN_FONT)
    rFonts.set(qn('w:hAnsi'), EN_FONT)
    rFonts.set(qn('w:cs'), FA_FONT)
    rPr.append(rFonts)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.text = '۱'
    r.append(t)
    fld.append(r)
    par._p.append(fld)


def clear_footer(section):
    """پاک‌سازی پاورقی (بدون شماره)."""
    footer = section.footer
    footer.is_linked_to_previous = False
    footer.paragraphs[0].text = ''


# ------------------------- پاورقی‌های معادل انگلیسی -------------------------

class FootnoteRegistry:
    """ثبت معادل‌های انگلیسی یکتا و ساخت پاورقی برای اولین رخداد بدنه."""

    def __init__(self):
        self.entries = []            # متن پاورقی‌ها به ترتیب ساخت (ایندکس+۱ = شناسه)
        self.keys = {}               # کلید نرمال‌شده -> (شناسه، فصل، شمارهٔ خط)
        self.removed_later = 0       # رخدادهای تکراری پاک‌شده
        self.left_intact = set()     # پرانتزهای دست‌نخورده (خارج از قواعد)

    @staticmethod
    def eligible(content):
        """آیا داخل پرانتز یک معادل انگلیسی مجاز پاورقی است؟"""
        if not re.search(r'[A-Za-z]', content):
            return False
        if re.search(r'[0-9۰-۹٠-٩]', content):     # پرانتزهای عددی/کد ابزار دست نخورند
            return False
        if any(ch in content for ch in '<>=×+%/'):  # عبارات آماری و مسیرها دست نخورند
            return False
        return True

    @staticmethod
    def norm_key(content):
        return re.sub(r'\s+', ' ', content.split(';')[0]).strip().lower()

    def scan(self, chapters):
        """گذر نخست: ثبت اولین رخداد بدنهٔ هر معادل (غیر از تیتر و جدول)."""
        for ch_idx, lines in enumerate(chapters):
            for ln_no, line in enumerate(lines):
                kind = _line_kind(line)
                if kind in ('table', 'fence', 'ref', 'blank'):
                    continue
                for m in ELIG_RE.finditer(line):
                    content = m.group(1)
                    if not self.eligible(content):
                        continue
                    if not _persian_before(line, m.start()):
                        continue
                    key = self.norm_key(content)
                    if key not in self.keys and kind in ('p', 'bullet'):
                        self.entries.append(content.strip())
                        self.keys[key] = (len(self.entries), ch_idx, ln_no)

    def transform_line(self, line, ch_idx, ln_no):
        """گذر دوم: جایگزینی اولین رخداد با توکن پاورقی و حذف تکراری‌ها."""
        kind = _line_kind(line)
        if kind in ('table', 'fence', 'ref', 'blank'):
            return line
        out, last = [], 0
        for m in ELIG_RE.finditer(line):
            content = m.group(1)
            key = self.norm_key(content)
            reg = self.keys.get(key)
            if reg is not None:
                fid, fch, fln = reg
                if (fch, fln) == (ch_idx, ln_no) and kind in ('p', 'bullet'):
                    # اولین رخداد: پرانتز حذف و ارجاع پاورقی جای آن می‌نشیند
                    out.append(line[last:m.start()])
                    out.append(FN_TOKEN.format(fid))
                else:
                    # رخداد تکراری (یا رخداد درون تیتر): فقط حذف پرانتز
                    out.append(line[last:m.start()].rstrip(' '))
                    self.removed_later += 1
                last = m.end()
                continue
            first_word = re.sub(r'[^a-z]', '', content.split()[0].lower()) \
                if content.split() else ''
            if first_word and any(k == first_word or k.startswith(first_word + ' ')
                                  for k in self.keys) \
                    and _persian_or_paren_before(line, m.start()):
                # پرانتز توضیحی دنبالهٔ اصطلاح پاورقی‌شده (مثل «PROCESS Model 4»)
                out.append(line[last:m.start()].rstrip(' '))
                self.removed_later += 1
                last = m.end()
            elif not self.eligible(content):
                pass                      # کد ابزار/عبارت آماری: دست‌نخورده
            else:
                self.left_intact.add(content)   # گزارش موارد دست‌نخورده
        out.append(line[last:])
        return ''.join(out)


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


def attach_footnotes_part(doc, registry):
    """ساخت word/footnotes.xml و پیوند آن به بدنهٔ سند."""
    if not registry.entries:
        return
    body = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
            '<w:footnote w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
            '<w:footnote w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>']
    for i, text in enumerate(registry.entries, start=1):
        body.append(
            f'<w:footnote w:id="{i}"><w:p>'
            '<w:pPr><w:bidi w:val="0"/></w:pPr>'
            '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r><w:rPr><w:rFonts w:ascii="{EN_FONT}" w:hAnsi="{EN_FONT}" w:cs="{EN_FONT}"/>'
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
    """تبدیل بدنهٔ مارک‌داون به فهرست بلوک‌ها؛ فهرست منابع جدا شده است."""
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
    for idx, (kind, payload) in enumerate(blocks):
        nxt = blocks[idx + 1][0] if idx + 1 < len(blocks) else ''
        if kind in ('p', 'h3') and payload.startswith('جدول') and nxt == 'table':
            out.append(('caption', payload))
        else:
            out.append((kind, payload))
    return out


def parse_references(text):
    """استخراج مدخل‌های فهرست منابع یک فصل از متن مارک‌داون."""
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


# ------------------------- رندر اجزای سند -------------------------

def setup_heading_styles(doc):
    """تیترهای واقعی Word (Heading 2/3): راست‌چین، RTL، سیاه، برای فهرست مطالب."""
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
    """رندر بلوک‌های تحلیل‌شده در بدنهٔ سند."""
    for kind, payload in blocks:
        if kind == 'h2':
            par = doc.add_paragraph(style='Heading 2')
            _set_paragraph_bidi(par)          # اطمینان از RTL بودن هر تیتر
            add_rich_text(par, payload, fa_pt=H2_PT, en_pt=H2_PT, bold_all=True)
        elif kind == 'h3':
            par = doc.add_paragraph(style='Heading 3')
            _set_paragraph_bidi(par)
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
        else:
            par = base_paragraph(doc, space_after=0)
            add_rich_text(par, payload)


def render_table(doc, rows):
    """جدول Word با کادر و جهت راست‌به‌چپ."""
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
    """شکل ۲-۱: جدول سه‌ستونهٔ متغیرها + عنوان زیر آن."""
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
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER, space_before=3, space_after=6)
    add_rich_text(par, caption, bold_all=True)


# ------------------------- بخش‌ها و صفحات -------------------------

def setup_section(section, numbering='none', page_start=None, title_pg=False,
                  vcenter=False):
    """تنظیم A4، حاشیه‌ها، جهت بخش و سبک شماره‌گذاری صفحات."""
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.right_margin, section.left_margin = Cm(4), Cm(2.5)
    section.top_margin, section.bottom_margin = Cm(3), Cm(2.5)
    section.footer_distance = Cm(1)
    _set_section_rtl(section)
    # python-docx با add_section ویژگی‌های sectPr قبلی راclone می‌کند؛
    # برای جلوگیری از انباشت، برچسب‌های شماره‌گذاری بخش پیشین پاک می‌شوند.
    for tag in ('w:pgNumType', 'w:titlePg', 'w:vAlign'):
        for el in section._sectPr.findall(qn(tag)):
            section._sectPr.remove(el)
    if vcenter:
        # وسط‌چین عمودی صفحه (جای w:vAlign با رعایت ترتیب طرحواره)
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
        # صفحهٔ نخست فصل: بدون نمایش شماره ولی به حساب می‌آید
        _insert_sectpr_element(section, 'w:titlePg')
        fpf = section.first_page_footer
        fpf.is_linked_to_previous = False
        fpf.paragraphs[0].text = ''


def page_break(doc):
    """درج شکست صفحه."""
    par = doc.add_paragraph()
    run = par.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)


def centered(doc, text, size, bold=False, before=0, after=0):
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER,
                         space_before=before, space_after=after)
    add_rich_text(par, text, fa_pt=size, en_pt=size, bold_all=bold)
    return par


def right(doc, text, size, bold=False, before=0, after=0):
    par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.RIGHT,
                         space_before=before, space_after=after)
    add_rich_text(par, text, fa_pt=size, en_pt=size, bold_all=bold)
    return par


def build_front_matter(doc):
    """صفحات مقدماتی به ترتیب راهنمای دانشگاه."""
    # --- بخش ۱: بسم‌الله، وسط صفحه، بدون کادر و تزیین ---
    setup_section(doc.sections[0], numbering='none', vcenter=True)
    centered(doc, 'بسم الله الرحمن الرحیم', 16, bold=True)

    # --- بخش ۲: سایر مقدماتی ---
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec, numbering='none')

    # صفحهٔ عنوان فارسی (چیدمان فونت‌ها طبق صفحهٔ ۲ راهنما)
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

    # فرم صورت‌جلسهٔ دفاع
    centered(doc, '[فرم صورت‌جلسه دفاع پس از برگزاری جلسه دفاع در این محل قرار می‌گیرد]',
             12, before=120)
    page_break(doc)

    # تعهدنامهٔ اصالت اثر
    centered(doc, '[فرم تعهدنامه اصالت اثر پس از امضا در این محل قرار می‌گیرد]',
             12, before=120)
    page_break(doc)

    # تقدیم
    right(doc, 'تقدیم به:', 14, bold=True, after=12)
    centered(doc, '[متن تقدیم توسط دانشجو تکمیل می‌شود]', 12, before=60)
    page_break(doc)

    # سپاسگزاری
    right(doc, 'سپاسگزاری', 14, bold=True, after=12)
    centered(doc, '[متن سپاسگزاری توسط دانشجو تکمیل می‌شود]', 12, before=60)
    page_break(doc)

    # چکیدهٔ فارسی
    right(doc, 'چکیده:', 14, bold=True, after=6)
    par = base_paragraph(doc, space_after=6)
    add_rich_text(par, '[چکیده پس از تکمیل فصل‌های چهارم و پنجم نوشته می‌شود]')
    par = base_paragraph(doc, space_after=0)
    add_rich_text(par, 'کلمات کلیدی: اضطراب صفتی، طرحواره‌های ناسازگار اولیه، '
                       'رفتارهای مالی پرخطر، معامله‌گران بازارهای مالی، '
                       'مدل‌سازی معادلات ساختاری', fa_pt=12, en_pt=11)


def build_lists_section(doc):
    """فهرست مطالب (فیلد TOC) + فهرست جدول‌ها/شکل‌ها/علائم با اعداد رومی."""
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
               '— [شمارهٔ صفحه پس از صفحه‌بندی نهایی در Word تکمل شود]', 12, after=3)
    page_break(doc)

    centered(doc, 'فهرست علائم', 14, bold=True, after=9)
    par = base_paragraph(doc)
    add_rich_text(par, '[در صورت وجود علائم و اختصارات، در این بخش تکمیل می‌شود]')


# ------------------------- منابع -------------------------

def fa_sort_key(text):
    """کلید مرتب‌سازی الفبایی فارسی با ترتیب صحیح الفبا."""
    order = 'آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی'

    def rank(ch):
        return order.find(ch) if ch in order else 99
    cleaned = re.sub(r'[\u200c\s]+', ' ', text).strip()
    return [rank(ch) for ch in cleaned if not ch.isspace()]


def build_unified_references(doc, ref_sets):
    """ادغام، حذف تکرار، مرتب‌سازی؛ فارسی RTL و لاتین LTR با تورفتگی آویزان."""
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
            near = [k for k in seen if k[:45] == key_low[:45]]
            if near:
                flagged.append((seen[near[0]], entry))
            seen[key_low] = entry
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
        # LTR، چپ‌چین، TNR 11، تورفتگی آویزان APA 7
        par = doc.add_paragraph()
        _set_paragraph_bidi(par, rtl=False)
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt = par.paragraph_format
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        fmt.line_spacing = Pt(LINE_SPACING_PT)
        fmt.left_indent = Cm(1.27)
        fmt.first_line_indent = Cm(-1.27)
        fmt.space_after = Pt(2)
        run = par.add_run(entry)
        run.font.name = EN_FONT
        run.font.size = Pt(11)
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = rPr.makeelement(qn('w:rFonts'), {})
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:cs'), EN_FONT)   # اطمینان از چپ‌به‌راست‌بودن کامل
    return len(fa_list), len(en_list), dup_count, flagged


# ------------------------- ساخت سند -------------------------

def main():
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    raw = [(k, p, p.read_text(encoding='utf-8')) for k, p in CH_SOURCES]

    # --- گذر پاورقی‌ها: ثبت و سپس بازنویسی خطوط بدنه ---
    registry = FootnoteRegistry()
    registry.scan([t.split('## منابع')[0].split('\n') for _, _, t in raw])
    processed = {}
    for kicker, path, text in raw:
        ch_idx = [r[0] for r in raw].index(kicker)
        body, _, refs_part = text.partition('## منابع')
        new_lines = []
        for ln_no, line in enumerate(body.split('\n')):
            new_lines.append(registry.transform_line(line, ch_idx, ln_no))
        processed[kicker] = '\n'.join(new_lines) + ('## منابع' + refs_part)

    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = EN_FONT
    normal.font.size = Pt(EN_SIZE)
    setup_heading_styles(doc)

    # --- صفحات مقدماتی و فهرست‌ها ---
    build_front_matter(doc)
    build_lists_section(doc)

    # --- فصل‌ها: شماره‌گذاری پیوستهٔ فارسی از فصل اول ---
    ref_sets = []
    first_chapter = True
    for kicker, path, text in raw:
        ref_sets.append(parse_references(text))
        blocks = parse_md_body(processed[kicker])
        title = chapter_title_from_h1(text)

        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        setup_section(sec, numbering='hindi',
                      page_start=1 if first_chapter else None,
                      title_pg=True)
        first_chapter = False

        # «فصل …» با سطح_outline صفر تا در فهرست مطالب بیاید
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER,
                             space_before=24, space_after=9)
        pPr = par._p.get_or_add_pPr()
        olvl = OxmlElement('w:outlineLvl')
        olvl.set(qn('w:val'), '0')
        pPr.append(olvl)
        add_rich_text(par, kicker, fa_pt=CHAPTER_KICKER_PT, en_pt=CHAPTER_KICKER_PT,
                      bold_all=True)
        par = base_paragraph(doc, WD_ALIGN_PARAGRAPH.CENTER,
                             space_before=6, space_after=24)
        add_rich_text(par, title, fa_pt=CHAPTER_TITLE_PT, en_pt=CHAPTER_TITLE_PT,
                      bold_all=True)
        render_blocks(doc, blocks)

    # --- منابع واحد ---
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec, numbering='hindi')
    stats = build_unified_references(doc, ref_sets)

    # --- پیوست پاورقی‌ها (پس از اتمام بدنه) ---
    attach_footnotes_part(doc, registry)

    doc.save(OUT_PATH)
    fa_n, en_n, dups, flagged = stats
    print(f'فایل ساخته شد: {OUT_PATH}')
    print(f'منابع: {fa_n} فارسی + {en_n} لاتین | تکرار حذف‌شده: {dups}')
    print(f'پاورقی یکتا: {len(registry.entries)} | رخداد تکراری پاک‌شده: '
          f'{registry.removed_later}')
    print('پرانتزهای دست‌نخورده (گزارش):',
          '; '.join(sorted(registry.left_intact)) or 'هیچ')
    for a, b in flagged:
        print('فلگ شباهت:', a[:60], '<->', b[:60])


if __name__ == '__main__':
    main()
