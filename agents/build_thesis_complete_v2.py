# -*- coding: utf-8 -*-
"""Mission v2: rebuild the complete thesis draft from MARKDOWN sources only.

Pipeline: md --(repo builders)--> fresh intermediate docx in tmpdir
          --(XML assembly here)--> output/final/thesis_complete_draft_v2.docx

Never reads output/final/*.docx. Old Word files are not the base.
"""
import copy
import posixpath
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'agents'))

import build_word
import build_ch4_word
import build_ch5_word

ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
ET.register_namespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('o', 'urn:schemas-microsoft-com:office:office')
ET.register_namespace('v', 'urn:schemas-microsoft-com:vml')
ET.register_namespace('w10', 'urn:schemas-microsoft-com:office:word')
ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
ET.register_namespace('sl', 'http://schemas.openxmlformats.org/schemaLibrary/2006/main')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
RN = 'http://schemas.openxmlformats.org/package/2006/relationships'
OFFDOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CTN = 'http://schemas.openxmlformats.org/package/2006/content-types'

CH13_MD = [ROOT / 'output/drafts/chapter1.md',
           ROOT / 'output/drafts/chapter2_full.md',
           ROOT / 'output/drafts/chapter3_full.md']
CH4_MD = ROOT / 'output/drafts/chapter4_simulated_real_pipeline.md'
CH4_FIG = ROOT / 'output/drafts/chapter4_real_pipeline_fig1.png'
CH5_MD = ROOT / 'output/drafts/chapter5_simulated_real_pipeline.md'
OUT = ROOT / 'output/final/thesis_complete_draft_v2.docx'
REPORT = ROOT / 'data/processed/final_thesis_build_report.md'


def T(p):
    return ''.join(x.text or '' for x in p.findall(f'.//{W}t'))


def run_builders(tmp):
    """Run the repo md->docx builders with ALL outputs redirected to tmp."""
    p13 = Path(tmp) / 'ch13.docx'
    p4 = Path(tmp) / 'ch4.docx'
    p5 = Path(tmp) / 'ch5.docx'
    build_word.OUT_PATH = p13
    build_word.MAP_PATH = Path(tmp) / 'citation_latin_map.md'
    build_word.REGISTRY_PATH = Path(tmp) / 'registry.md'
    build_word.main()
    build_ch4_word.CH4_REGISTRY_OUT = Path(tmp) / 'ch4_registry.md'
    build_ch4_word.build(md_path=CH4_MD, fig_path=CH4_FIG, out_path=p4)
    build_ch5_word.CH5_REGISTRY_OUT = Path(tmp) / 'ch5_registry.md'
    build_ch5_word.build(md_path=CH5_MD, out_path=p5)
    return p13, p4, p5


def body_texts(docx_path, skip_refs=True):
    """Top-level body para texts; optionally cut at the refs heading."""
    root = ET.fromstring(zipfile.ZipFile(docx_path).read('word/document.xml'))
    out = []
    for c in list(root.find(W + 'body')):
        if c.tag != W + 'p':
            continue
        t = re.sub(r'\s+', ' ', T(c)).strip()
        if skip_refs and t == 'منابع':
            break
        if t:
            out.append(t)
    return out


def verify_intermediates(p13, p4, p5, log):
    """Report-only: fresh builds vs committed docx (committed may be stale;
    the mission orders building from the latest md, not matching old docx)."""
    checks = [
        ('ch1-3', str(p13), 'output/final/thesis_ch1_to_ch3.docx',
         'فصل اول'),
        ('ch4', str(p4), 'output/final/chapter4_simulated_real_pipeline.docx',
         None),
        ('ch5', str(p5), 'output/final/chapter5_simulated_real_pipeline.docx',
         None),
    ]
    for name, fresh, committed, frm in checks:
        a = body_texts(fresh)
        b = body_texts(ROOT / committed)
        if frm:
            a = a[a.index(frm):]
            b = b[b.index(frm):]
        if a == b:
            log.append(f'- intermediate {name}: {len(a)} paras, '
                       f'identical-to-committed=True')
            continue
        nd = sum(1 for x, y in zip(a, b) if x != y) + abs(len(a) - len(b))
        log.append(f'- intermediate {name}: fresh={len(a)} committed={len(b)} '
                   f'differing={nd} (committed docx is stale; md wins)')
        for x, y in zip(a, b):
            if x != y:
                log.append(f'  first-diff fresh: {x[:90]}')
                log.append(f'  first-diff committed: {y[:90]}')
                break


def run_el(text, font='B Nazanin', sz='24', bold=False, italic=False,
           vanish=False, rtl=False):
    r = ET.Element(W + 'r')
    rp = ET.SubElement(r, W + 'rPr')
    if vanish:
        ET.SubElement(rp, W + 'vanish')
    rf = ET.SubElement(rp, W + 'rFonts')
    rf.set(W + 'ascii', font)
    rf.set(W + 'hAnsi', font)
    rf.set(W + 'cs', font)
    if bold:
        ET.SubElement(rp, W + 'b')
        ET.SubElement(rp, W + 'bCs')
    if italic:
        ET.SubElement(rp, W + 'i')
        ET.SubElement(rp, W + 'iCs')
    ET.SubElement(rp, W + 'sz').set(W + 'val', sz)
    ET.SubElement(rp, W + 'szCs').set(W + 'val', sz)
    if rtl:
        ET.SubElement(rp, W + 'rtl')
    t = ET.SubElement(r, W + 't')
    t.text = text
    if text != text.strip():
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def para_el(runs, jc=None, rtl=True, exact='580', after=None, outline=None,
            pstyle=None, ind=None):
    # ترتیب CT_PPr در طرحواره: pStyle ... bidi ... spacing ind ... jc ... outlineLvl
    p = ET.Element(W + 'p')
    pr = ET.SubElement(p, W + 'pPr')
    if pstyle:
        ET.SubElement(pr, W + 'pStyle').set(W + 'val', pstyle)
    if rtl:
        ET.SubElement(pr, W + 'bidi')
    if exact or after:
        sp = ET.SubElement(pr, W + 'spacing')
        if exact:
            sp.set(W + 'line', exact)
            sp.set(W + 'lineRule', 'exact')
        if after:
            sp.set(W + 'after', after)
    if ind:
        ie = ET.SubElement(pr, W + 'ind')
        for k, v in ind.items():
            ie.set(W + k, v)
    if jc:
        ET.SubElement(pr, W + 'jc').set(W + 'val', jc)
    if outline is not None:
        ET.SubElement(pr, W + 'outlineLvl').set(W + 'val', str(outline))
    for r in runs:
        p.append(r)
    return p


def text_para(text, **kw):
    return para_el([run_el(text, font=kw.pop('font', 'B Nazanin'),
                           sz=kw.pop('sz', '24'),
                           bold=kw.pop('bold', False),
                           rtl=kw.pop('rtl_run', True))],
                   pstyle=kw.pop('pstyle', None), **kw)


def page_break_para():
    p = ET.Element(W + 'p')
    r = ET.SubElement(p, W + 'r')
    ET.SubElement(r, W + 'br').set(W + 'type', 'page')
    return p


CAP_SPLIT_RE = re.compile(r'^(جدول|شکل) ([۰-۹0-9]+-[۰-۹0-9]+)\.(.*)$',
                            re.DOTALL)


def _run_with(text, rpr):
    r = ET.Element(W + 'r')
    r.append(copy.deepcopy(rpr))
    t = ET.SubElement(r, W + 't')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def seq_caption_surgery(c):
    """Replace the caption number with a LOCKED SEQ field caching the same
    number (visible text byte-identical; TOC \\c "جدول"/"شکل" picks it up)."""
    runs = c.findall(W + 'r')
    t1 = ''.join(x.text or '' for x in runs[0].findall(W + 't'))
    m = CAP_SPLIT_RE.match(t1)
    assert m, f'caption shape: {t1[:60]!r}'
    label, num, rest1 = m.group(1), m.group(2), m.group(3)
    base = runs[0].find(W + 'rPr')
    assert base is not None
    before = ''.join(x.text or ''
                     for r in runs for x in r.findall(W + 't'))
    new = [_run_with(label + ' ', base)]
    rb = ET.Element(W + 'r')
    rb.append(copy.deepcopy(base))
    fb = ET.SubElement(rb, W + 'fldChar')
    fb.set(W + 'fldCharType', 'begin')
    fb.set(W + 'fldLock', 'true')
    new.append(rb)
    ri = ET.Element(W + 'r')
    ri.append(copy.deepcopy(base))
    ET.SubElement(ri, W + 'instrText').text = f'SEQ {label}'
    new.append(ri)
    rs = ET.Element(W + 'r')
    rs.append(copy.deepcopy(base))
    ET.SubElement(rs, W + 'fldChar').set(W + 'fldCharType', 'separate')
    new.append(rs)
    new.append(_run_with(num, base))
    re_ = ET.Element(W + 'r')
    re_.append(copy.deepcopy(base))
    ET.SubElement(re_, W + 'fldChar').set(W + 'fldCharType', 'end')
    new.append(re_)
    new.append(_run_with('.' + rest1, base))
    idx = list(c).index(runs[0])
    for r in reversed(new):
        c.insert(idx + 1, r)
    c.remove(runs[0])
    after = ''.join(x.text or ''
                    for r in c.findall(W + 'r') for x in r.findall(W + 't'))
    assert after == before, (before[:50], after[:50])
    return label


TOC_RUN_SZ = {'TOC1': '24', 'TOC2': '24', 'TOC3': '24',
              'tableoffigures': '24'}


def toc_field_paras(instr, entries):
    out = []
    p = ET.Element(W + 'p')
    ET.SubElement(ET.SubElement(p, W + 'pPr'), W + 'bidi')
    for kt, tx in (('begin', None), (None, instr), ('separate', None)):
        r = ET.SubElement(p, W + 'r')
        rp = ET.SubElement(r, W + 'rPr')
        rf = ET.SubElement(rp, W + 'rFonts')
        rf.set(W + 'ascii', 'B Nazanin')
        rf.set(W + 'hAnsi', 'B Nazanin')
        if kt:
            ET.SubElement(r, W + 'fldChar').set(W + 'fldCharType', kt)
        else:
            ET.SubElement(r, W + 'instrText').text = tx
    out.append(p)
    for t, sid in entries:
        out.append(text_para(t, jc='right', pstyle=sid,
                             sz=TOC_RUN_SZ[sid], rtl_run=True,
                             bold=(sid == 'TOC1')))
    pe = ET.Element(W + 'p')
    ET.SubElement(ET.SubElement(pe, W + 'pPr'), W + 'bidi')
    ET.SubElement(ET.SubElement(pe, W + 'r'), W + 'fldChar').set(
        W + 'fldCharType', 'end')
    out.append(pe)
    return out


def rtl_toc_style(style_id, name, sz, tab_pos, bold=False):
    st = ET.Element(W + 'style')
    st.set(W + 'type', 'paragraph')
    st.set(W + 'styleId', style_id)
    ET.SubElement(st, W + 'name').set(W + 'val', name)
    ET.SubElement(st, W + 'basedOn').set(W + 'val', 'Normal')
    ET.SubElement(st, W + 'next').set(W + 'val', 'Normal')
    ET.SubElement(st, W + 'semiHidden')
    ET.SubElement(st, W + 'unhideWhenUsed')
    pr = ET.SubElement(st, W + 'pPr')
    tabs = ET.SubElement(pr, W + 'tabs')
    tab = ET.SubElement(tabs, W + 'tab')
    tab.set(W + 'val', 'left')
    tab.set(W + 'leader', 'dot')
    tab.set(W + 'pos', str(tab_pos))
    ET.SubElement(pr, W + 'bidi').set(W + 'val', '1')
    ET.SubElement(pr, W + 'jc').set(W + 'val', 'right')
    rp = ET.SubElement(st, W + 'rPr')
    rf = ET.SubElement(rp, W + 'rFonts')
    rf.set(W + 'ascii', 'Times New Roman')
    rf.set(W + 'hAnsi', 'Times New Roman')
    rf.set(W + 'cs', 'B Nazanin')
    if bold:
        ET.SubElement(rp, W + 'b')
        ET.SubElement(rp, W + 'bCs')
    ET.SubElement(rp, W + 'sz').set(W + 'val', sz)
    ET.SubElement(rp, W + 'szCs').set(W + 'val', sz)
    ET.SubElement(rp, W + 'rtl')
    return st


ABSTRACT = ('پژوهش حاضر با هدف تعیین نقش میانجی‌گر طرحواره‌های ناسازگار '
            'اولیه در رابطه بین اضطراب صفتی و رفتارهای مالی پرخطر در '
            'معامله‌گران بازارهای مالی طراحی شد. این پژوهش از نظر هدف '
            'کاربردی و از نظر روش توصیفی-همبستگی از نوع مدل‌سازی معادلات '
            'ساختاری است. جامعهٔ آماری پژوهش را کلیهٔ معامله‌گران حقیقی '
            'فارسی‌زبان بالای ۱۸ سال با حداقل شش ماه سابقهٔ معامله تشکیل '
            'می‌دهند و نمونه‌گیری به روش غیرتصادفی در دسترس به همراه روش '
            'گلوله‌برفی انجام می‌شود؛ حجم نمونهٔ هدف بر اساس قواعد تجربی '
            'مدل‌یابی معادلات ساختاری ۲۰۰ نفر در نظر گرفته شد. ابزارهای '
            'پژوهش شامل فرم اضطراب صفتی پرسشنامهٔ حالت-صفت اسپیلبرگر '
            '(۲۰ گویه)، پرسشنامهٔ طرحواره‌های ناسازگار اولیهٔ یانگ ــ فرم '
            'کوتاه (۷۵ گویه) و مقیاس تحمل ریسک مالی گرابل و لایتون '
            '(۱۳ سؤال) است. برای تحلیل داده‌ها از آمار توصیفی، ضریب '
            'همبستگی پیرسون، مدل‌سازی معادلات ساختاری با شاخص‌های برازش و '
            'تحلیل میانجی‌گری با ماکروی PROCESS استفاده خواهد شد. '
            'یافته‌های تجربی نهایی پس از تحلیل داده‌های واقعی در نسخه '
            'نهایی پایان‌نامه گزارش خواهد شد.')

KEYWORDS = ('اضطراب صفتی، طرحواره‌های ناسازگار اولیه، رفتارهای مالی پرخطر، '
            'معامله‌گران بازارهای مالی، مدل‌سازی معادلات ساختاری')

ABBREVS = [
    ('STAI-Y2', 'فرم اضطراب صفتی پرسشنامهٔ حالت-صفت اسپیلبرگر'),
    ('YSQ-SF', 'پرسشنامهٔ طرحوارهٔ یانگ ــ فرم کوتاه'),
    ('GL-RTS', 'مقیاس تحمل ریسک مالی گرابل و لایتون'),
    ('SEM', 'مدل‌سازی معادلات ساختاری'),
    ('PROCESS', 'ماکروی PROCESS برای تحلیل میانجی‌گری'),
    ('ML', 'برآورد درست‌نمایی بیشینه'),
    ('CFI', 'شاخص برازش تطبیقی'),
    ('TLI', 'شاخص برازش تاکر-لوئیس'),
    ('RMSEA', 'ریشهٔ میانگین مربعات خطای تقریب'),
    ('SRMR', 'ریشهٔ میانگین مربعات باقی‌ماندهٔ استانداردشده'),
    ('VIF', 'عامل تورم واریانس'),
    ('CI', 'فاصلهٔ اطمینان'),
    ('r', 'ضریب همبستگی پیرسون'),
    ('p', 'سطح معناداری'),
    ('β', 'ضریب مسیر استانداردشده'),
    ('χ²', 'آمارهٔ خی‌دو'),
]

TITLE_FA = ('تعیین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه در رابطه بین '
            'اضطراب صفتی و رفتارهای مالی پرخطر در معامله‌گران بازارهای مالی')


def spacer(n=1):
    return [ET.Element(W + 'p') for _ in range(n)]


def abbrev_table():
    tbl = ET.Element(W + 'tbl')
    pr = ET.SubElement(tbl, W + 'tblPr')
    ET.SubElement(pr, W + 'tblStyle').set(W + 'val', 'TableGrid')
    ET.SubElement(pr, W + 'bidiVisual')
    # tblGrid فرزند tbl است (پس از tblPr) نه فرزند tblPr
    grid = ET.SubElement(tbl, W + 'tblGrid')
    for w in ('2200', '6021'):
        ET.SubElement(grid, W + 'gridCol').set(W + 'w', w)
    rows = [('علامت اختصاری', 'معادل فارسی / توضیح', True)] + [
        (s, f, False) for s, f in ABBREVS]
    for sym, fa, head in rows:
        tr = ET.SubElement(tbl, W + 'tr')
        for txt, sym_col in ((sym, True), (fa, False)):
            tc = ET.SubElement(tr, W + 'tc')
            tcpr = ET.SubElement(tc, W + 'tcPr')
            tcw = ET.SubElement(tcpr, W + 'tcW')
            tcw.set(W + 'w', '2200' if sym_col else '6021')
            tcw.set(W + 'type', 'dxa')
            if sym_col and not head:
                tc.append(para_el([run_el(txt, font='Times New Roman',
                                          sz='22')],
                                  jc='center', rtl=False))
            else:
                tc.append(para_el([run_el(txt, bold=head, rtl=True)],
                                  jc='center' if sym_col else 'right'))
    return tbl


def build_front_matter(c_contents, c_tables, c_figs):
    F = []
    F += spacer(4)
    F.append(text_para('بسم الله الرحمن الرحیم', jc='center', sz='28'))
    F.append(page_break_para())
    F += spacer(2)
    F.append(text_para('دانشگاه علم و هنر یزد', jc='center', sz='20',
                       bold=True))
    F += spacer(2)
    for ln in ['تعیین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه',
               'در رابطه بین اضطراب صفتی و رفتارهای مالی پرخطر',
               'در معامله‌گران بازارهای مالی']:
        F.append(text_para(ln, jc='center', sz='48', bold=True))
    F += spacer(2)
    F.append(text_para('پایان‌نامه کارشناسی ارشد روان‌شناسی بالینی',
                       jc='center', sz='28', bold=True))
    F += spacer(3)
    F.append(para_el([run_el('استاد راهنما: ', sz='28', rtl=True),
                      run_el('دکتر سعید وزیری یزدی', sz='28', bold=True,
                             rtl=True)], jc='center'))
    F.append(para_el([run_el('نگارنده: ', sz='28', rtl=True),
                      run_el('مریم قابل', sz='28', bold=True, rtl=True)],
                     jc='center'))
    F += spacer(2)
    F.append(text_para('۱۴۰۵', jc='center', sz='28', bold=True))
    F.append(page_break_para())
    F.append(text_para('تأییدیه هیئت داوران', jc='center', sz='28',
                       bold=True))
    F += spacer(1)
    F.append(text_para('[فرم تأییدیه و صورت‌جلسه دفاع پس از جلسه دفاع در این '
                       'محل قرار می‌گیرد]', jc='center'))
    F.append(page_break_para())
    F.append(text_para('تعهدنامه اصالت اثر', jc='center', sz='28',
                       bold=True))
    F += spacer(1)
    F.append(text_para('[فرم تعهدنامه اصالت اثر پس از تکمیل در این محل قرار '
                       'می‌گیرد]', jc='center'))
    F.append(page_break_para())
    F.append(text_para('تقدیم به:', jc='center', sz='28', bold=True))
    F += spacer(1)
    F.append(text_para('این پژوهش را به خانواده‌ام تقدیم می‌کنم که با حمایت، '
                       'صبوری و دلگرمی خود، مسیر تحصیل و پژوهش را برایم هموار '
                       'ساختند.', jc='center', sz='28'))
    F.append(page_break_para())
    F.append(text_para('سپاسگزاری', jc='center', sz='28', bold=True))
    F += spacer(1)
    F.append(text_para('در آغاز، از استاد گران‌قدر راهنمای این پژوهش، جناب آقای '
                       'دکتر سعید وزیری یزدی، که با راهنمایی‌ها و حمایت‌های '
                       'ارزشمند خود در تمام مراحل انجام این پژوهش یاری‌ام '
                       'کردند، صمیمانه سپاسگزارم. همچنین از تمامی استادان، '
                       'همکاران و شرکت‌کنندگانی که امکان انجام این پژوهش را '
                       'فراهم ساختند، قدردانی می‌کنم.', jc='both', sz='28'))
    F.append(page_break_para())
    F.append(text_para('چکیده:', jc='center', sz='28', bold=True))
    F += spacer(1)
    F.append(text_para(ABSTRACT, jc='both'))
    F.append(para_el([run_el('کلمات کلیدی: ', bold=True, rtl=True),
                      run_el(KEYWORDS, rtl=True)], jc='both'))
    F.append(page_break_para())
    F.append(text_para('فهرست مطالب', jc='center', sz='28', bold=True))
    F += toc_field_paras('TOC \\o "1-3" \\h \\z \\u', c_contents)
    F.append(page_break_para())
    F.append(text_para('فهرست جدول‌ها', jc='center', sz='28', bold=True))
    F += toc_field_paras('TOC \\h \\z \\c "جدول"', c_tables)
    F.append(page_break_para())
    F.append(text_para('فهرست شکل‌ها', jc='center', sz='28', bold=True))
    F += toc_field_paras('TOC \\h \\z \\c "شکل"', c_figs)
    F.append(page_break_para())
    F.append(text_para('فهرست علائم', jc='center', sz='28', bold=True))
    F += spacer(1)
    F.append(abbrev_table())
    return F


CAP_RE = re.compile(r'^(جدول|شکل) [۰-۹0-9]+-[۰-۹0-9]+\.')
SECT_ORDER = ['footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz',
              'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType',
              'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
              'textDirection', 'bidi', 'rtlGutter', 'docGrid',
              'printerSettings']


def sect_props(model, footer_rid, fmt=None, start=None, title_pg=False):
    sp = ET.Element(W + 'sectPr')
    fr = ET.SubElement(sp, W + 'footerReference')
    fr.set(W + 'type', 'default')
    fr.set(R + 'id', footer_rid)
    keep = {}
    for c in list(model):
        tag = c.tag[len(W):]
        if tag in ('footerReference', 'headerReference', 'pgNumType',
                   'titlePg', 'type'):
            continue
        keep.setdefault(tag, []).append(copy.deepcopy(c))
    if fmt:
        pg = ET.Element(W + 'pgNumType')
        pg.set(W + 'fmt', fmt)
        if start:
            pg.set(W + 'start', str(start))
        keep['pgNumType'] = [pg]
    if title_pg:
        keep['titlePg'] = [ET.Element(W + 'titlePg')]
    for tag in SECT_ORDER[1:]:
        for c in keep.get(tag, []):
            sp.append(c)
    return sp


def joint_para(model, footer_rid, **kw):
    p = ET.Element(W + 'p')
    pPr = ET.SubElement(p, W + 'pPr')
    pPr.append(sect_props(model, footer_rid, **kw))
    return p


def footer_xml(numbered, cache, lang_bidi=None):
    ftr = ET.Element(W + 'ftr')
    if not numbered:
        ftr.append(ET.Element(W + 'p'))
        return ftr
    p = ET.SubElement(ftr, W + 'p')
    pr = ET.SubElement(p, W + 'pPr')
    ET.SubElement(pr, W + 'pStyle').set(W + 'val', 'Footer')
    ET.SubElement(pr, W + 'bidi')
    ET.SubElement(pr, W + 'jc').set(W + 'val', 'center')
    fs = ET.SubElement(p, W + 'fldSimple')
    fs.set(W + 'instr', ' PAGE ')
    run = run_el(cache, sz='24', rtl=lang_bidi is not None)
    if lang_bidi:
        lang = ET.SubElement(run.find(W + 'rPr'), W + 'lang')
        lang.set(W + 'bidi', lang_bidi)
    fs.append(run)
    return ftr


TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge',
              'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection',
              'tcFitText', 'vAlign', 'hideMark']

# از XSD انتقالی: rPr پیش از tblPr می‌آید
STYLE_ORDER = ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine',
               'hidden', 'uiPriority', 'semiHidden', 'unhideWhenUsed',
               'qFormat', 'locked', 'personal', 'personalCompose',
               'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr',
               'tcPr', 'tblStylePr']

NUMFMTS = set('''decimal upperRoman lowerRoman upperLetter lowerLetter
    ordinal cardinalText ordinalText hex chicago ideographDigital
    japaneseCounting aiueo iroha decimalFullWidth decimalHalfWidth
    japaneseLegal japaneseDigitalTenThousand decimalEnclosedCircle
    decimalFullWidth2 aiueoFullWidth irohaFullWidth decimalZero bullet
    ganada chosung decimalEnclosedFullstop decimalEnclosedParen
    decimalEnclosedCircleChinese ideographEnclosedCircle
    ideographTraditional ideographZodiac ideographZodiacTraditional
    taiwaneseCounting ideographLegalTraditional taiwaneseCountingThousand
    taiwaneseDigital chineseCounting chineseLegalSimplified
    chineseCountingThousand koreanDigital koreanCounting koreanLegal
    koreanDigital2 vietnameseCounting russianLower russianUpper none
    numberInDash hebrew1 hebrew2 arabicAlpha arabicAbjad hindiVowels
    hindiConsonants hindiNumbers hindiCounting thaiLetters thaiNumbers
    thaiCounting bahtText dollarText custom'''.split())


def _assert_order(el, order, where):
    pos = -1
    seen = set()
    for c in list(el):
        if not c.tag.startswith(W):
            continue
        t = c.tag[len(W):]
        if t not in order:
            continue
        if t not in ('headerReference', 'footerReference', 'tblStylePr'):
            assert t not in seen, f'{where}: duplicate <{t}>'
            seen.add(t)
        p = order.index(t)
        assert p >= pos, f'{where}: <{t}> out of schema order'
        pos = p


def _set_outline_lvl(pPr, val):
    """Set pPr outlineLvl once (builders pre-mark kickers); schema position:
    before divId/cnfStyle/rPr/sectPr."""
    old = pPr.find(W + 'outlineLvl')
    if old is not None:
        assert old.get(W + 'val') == val, (
            old.get(W + 'val'), val)
        return
    el = ET.Element(W + 'outlineLvl')
    el.set(W + 'val', val)
    for c in list(pPr):
        if c.tag.startswith(W) and c.tag[len(W):] in (
                'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'):
            pPr.insert(list(pPr).index(c), el)
            return
    pPr.append(el)


def validate_package(path):
    """Post-build gate: well-formed XML + OPC integrity + ECMA-376 child
    order + numbering-enum values. Raises (fails the build) on any defect."""
    z = zipfile.ZipFile(path)
    parts = {}
    for name in z.namelist():
        if name.endswith('.xml') or name.endswith('.rels'):
            try:
                parts[name] = ET.fromstring(z.read(name))
            except ET.ParseError as e:
                raise AssertionError(f'{name}: not well-formed: {e}')
    # OPC: rel targets exist, no duplicate Ids
    for name, root in parts.items():
        if not name.endswith('.rels'):
            continue
        if name == '_rels/.rels':
            base = ''
        else:
            base = name.split('/_rels/')[0] + '/'
        seen = set()
        for r in root.findall(f'{{{RN}}}Relationship'):
            rid = r.get('Id')
            assert rid not in seen, f'{name}: duplicate {rid}'
            seen.add(rid)
            tgt = r.get('Target', '')
            mode = r.get('TargetMode', 'Internal')
            rtype = r.get('Type', '')
            if name == '_rels/.rels':
                assert rtype.startswith((RN + '/', OFFDOC + '/')), \
                    f'{name}: bad rel Type {rtype}'
            else:
                assert rtype.startswith(OFFDOC + '/') or rtype == (
                    'http://schemas.microsoft.com/office/2007/relationships/'
                    'stylesWithEffects'), f'{name}: bad rel Type {rtype}'
            if mode == 'Internal' and not tgt.startswith('/'):
                norm = posixpath.normpath(base + tgt)
                assert norm in z.namelist(), \
                    f'{name}: missing target {tgt}'
    # OPC: content-type coverage for word parts
    ct = parts['[Content_Types].xml']
    over = {e.get('PartName') for e in ct
            if e.tag == f'{{{CTN}}}Override'}
    dflt = {e.get('Extension') for e in ct
            if e.tag == f'{{{CTN}}}Default'}
    for name in z.namelist():
        if not name.startswith('word/') or name.endswith('.rels'):
            continue
        assert ('/' + name in over
                or name.rsplit('.', 1)[-1] in dflt), f'no CT for {name}'
    # schema child order + enums + field balance
    for name, root in parts.items():
        if not (name.startswith('word/') and name.endswith('.xml')):
            continue
        for el in root.iter():
            if not el.tag.startswith(W):
                continue
            t = el.tag[len(W):]
            if t == 'pPr':
                _assert_order(el, build_word._PPR_ORDER, name)
            elif t == 'rPr':
                _assert_order(el, build_word._RPR_ORDER, name)
            elif t == 'sectPr':
                _assert_order(el, SECT_ORDER, name)
            elif t == 'tblPr':
                _assert_order(el, build_word._TBLPR_ORDER, name)
            elif t == 'tcPr':
                _assert_order(el, TCPR_ORDER, name)
            elif t == 'tbl':
                kids = [c.tag[len(W):] for c in list(el)
                        if c.tag.startswith(W)]
                assert kids[0] == 'tblPr', f'{name}: tbl w/o tblPr first'
                assert 'tblGrid' in kids, f'{name}: tbl w/o tblGrid'
                assert (kids.index('tblGrid') <
                        kids.index('tr')), f'{name}: tblGrid after tr'
            elif t == 'style':
                _assert_order(el, STYLE_ORDER, name)
            elif t in ('pgNumType', 'numFmt'):
                v = el.get(W + ('fmt' if t == 'pgNumType' else 'val'))
                assert v in NUMFMTS, f'{name}: bad num format {v!r}'
            elif t == 'zoom':
                assert el.get(W + 'percent'), f'{name}: zoom w/o percent'
    # body: exactly one direct sectPr, last
    doc = parts['word/document.xml']
    kids = list(doc.find(W + 'body'))
    assert sum(1 for c in kids if c.tag == W + 'sectPr') == 1, \
        'body sectPr count != 1'
    assert kids[-1].tag == W + 'sectPr', 'body must end with sectPr'
    # field balance in document order (fields may span paragraphs)
    for name, root in parts.items():
        if not (name.startswith('word/') and name.endswith('.xml')):
            continue
        depth = 0
        for fc in root.findall(f'.//{W}fldChar'):
            kt = fc.get(W + 'fldCharType')
            if kt == 'begin':
                depth += 1
            elif kt == 'end':
                assert depth > 0, f'{name}: stray fldChar end'
                depth -= 1
            elif kt == 'separate':
                assert depth > 0, f'{name}: stray separate'
        assert depth == 0, f'{name}: unclosed field'
    # image/rel cross-check
    doc = parts['word/document.xml']
    rels = {r.get('Id'): r.get('Target') for r in
            parts['word/_rels/document.xml.rels'].findall(
                f'{{{RN}}}Relationship')}
    for b in doc.findall('.//{http://schemas.openxmlformats.org/drawingml'
                         '/2006/main}blip'):
        emb = b.get(R + 'embed')
        assert emb in rels, f'dangling blip {emb}'
        assert 'word/' + rels[emb] in z.namelist(), f'missing {rels[emb]}'


def extract_refs(blocks):
    refs, cur = [], None
    for c in blocks:
        if c.tag != W + 'p':
            continue
        t = re.sub(r'\s+', ' ', T(c)).strip()
        if not t or t == 'منابع':
            continue
        if t.startswith('الف)'):
            cur = False
            continue
        if t.startswith('ب)'):
            cur = True
            continue
        is_lat = bool(re.match(r'[A-Za-z]', t))
        if cur is not None:
            assert is_lat == cur, f'entry under wrong subhead: {t[:40]}'
        refs.append((is_lat, t))
    return refs


def ref_norm(s):
    return re.sub(r'\s+', ' ',
                  s.replace('‌', ' ').replace(' ', ' ')).strip().lower()


TYPO_MAP = str.maketrans({'ي': 'ی', 'ك': 'ک', 'ۀ': 'ه', 'ة': 'ه'})
PUNCT_RE = re.compile(r'[،؛:.()\[\]"«»\'’‘\-–—/\\*_]+')


def ref_folded(s):
    s = ref_norm(s).translate(TYPO_MAP)
    s = s.replace('‌', '').replace('‍', '')
    return PUNCT_RE.sub('', s)


def typo_score(s):
    return (s.count('ی') + s.count('ک') + s.count('‌')
            - s.count('ي') - s.count('ك'))


def unify_refs(all_entries, log):
    seen, fa, la, dupes, merges, flags = {}, [], [], 0, [], []
    for is_lat, t in all_entries:
        k = ref_norm(t)
        if k in seen:
            dupes += 1
            continue
        seen[k] = (is_lat, t)
    items = list(seen.values())
    by_fold, merged_drop = {}, set()
    for is_lat, t in items:
        f = ref_folded(t)
        if f in by_fold:
            prev_lat, prev_t = by_fold[f]
            assert prev_lat == is_lat
            keep, drop = ((t, prev_t) if typo_score(t) > typo_score(prev_t)
                          else (prev_t, t))
            by_fold[f] = (is_lat, keep)
            merged_drop.add(drop)
            merges.append((keep, drop))
        else:
            by_fold[f] = (is_lat, t)
    for is_lat, t in by_fold.values():
        (la if is_lat else fa).append(t)
    fa.sort(key=lambda e: (build_word.fa_sort_key(e.split('،')[0]),
                           build_word.fa_sort_key(e)))
    la.sort(key=str.lower)
    keys = [ref_norm(t) for _, t in by_fold.values()]
    for i in range(len(keys)):
        for j in range(i + 1, len(keys)):
            if keys[i][:45] == keys[j][:45]:
                flags.append((keys[i][:60], keys[j][:60]))
    log.append(f'- refs exact-dupes removed: {dupes}')
    log.append(f'- refs typo-merges: {len(merges)}')
    for keep, drop in merges:
        log.append(f'  KEEP: {keep[:70]}')
        log.append(f'  DROP: {drop[:70]}')
    log.append(f'- refs near-flags (first-45, kept both): {len(flags)}')
    for a, b in flags[:10]:
        log.append(f'  FLAG: {a} <-> {b}')
    return fa, la


def ref_para_fa(text):
    return para_el([run_el(text, sz='24', rtl=True)], jc='right',
                   after='40')


def ref_para_la(text):
    return para_el([run_el(text, font='Times New Roman', sz='22')],
                   jc='left', after='40',
                   ind={'left': '720', 'hanging': '720'})


def main():
    log = []
    with tempfile.TemporaryDirectory(prefix='v2build_') as tmp:
        p13, p4, p5 = run_builders(tmp)
        verify_intermediates(p13, p4, p5, log)
        zb = zipfile.ZipFile(p13)
        z4 = zipfile.ZipFile(p4)
        z5 = zipfile.ZipFile(p5)
        kids13 = list(ET.fromstring(zb.read('word/document.xml')).find(
            W + 'body'))
        k4 = list(ET.fromstring(z4.read('word/document.xml')).find(
            W + 'body'))
        k5 = list(ET.fromstring(z5.read('word/document.xml')).find(
            W + 'body'))

        def top_paras(kids):
            return [(i, c) for i, c in enumerate(kids) if c.tag == W + 'p']

        tp = top_paras(kids13)
        tmap = {}
        for i, c in tp:
            tmap.setdefault(re.sub(r'\s+', ' ', T(c)).strip(), i)
        i_f1 = tmap['فصل اول']
        i_f2 = tmap['فصل دوم']
        i_f3 = tmap['فصل سوم']
        i_ref = tmap['منابع']
        ch1 = copy.deepcopy(kids13[i_f1:i_f2])
        ch2 = copy.deepcopy(kids13[i_f2:i_f3])
        ch3 = copy.deepcopy(kids13[i_f3:i_ref])
        refs13 = kids13[i_ref:]
        t5 = top_paras(k5)
        i5 = [i for i, c in t5
              if re.sub(r'\s+', ' ', T(c)).strip() == 'منابع'][0]
        ch4 = copy.deepcopy([c for c in k4
                             if c.tag != W + 'sectPr'])
        ch5 = copy.deepcopy(k5[:i5])
        refs5 = k5[i5:]
        assert not [c for c in top_paras(k4)[0:]
                    if re.sub(r'\s+', ' ', T(c[1])).strip() == 'منابع']

        nstrip = 0
        for blk in (ch1, ch2, ch3, ch4, ch5):
            for c in list(blk):
                if c.tag == W + 'sectPr':
                    blk.remove(c)
                    nstrip += 1
                    continue
                if c.tag == W + 'p' and c.find(
                        f'{W}pPr/{W}sectPr') is not None:
                    assert not re.sub(r'\s+', ' ', T(c)).strip()
                    blk.remove(c)
                    nstrip += 1
        for blk, nm in ((ch4, 'ch4'), (ch5, 'ch5')):
            n = sum(1 for c in blk
                    for _ in c.findall(f'.//{W}footnoteReference'))
            assert n == 0, f'{nm} has footnote refs'

        # Bug-2 fix: drop the copied drawing; the figure is re-inserted via
        # python-docx run.add_picture() after the package is written, so the
        # library owns media bytes + rel + drawing XML (standard prefixes).
        ndraw = 0
        for c in ch4:
            for p in ([c] if c.tag == W + 'p' else []):
                for r in p.findall(W + 'r'):
                    for d in r.findall(W + 'drawing'):
                        r.remove(d)
                        ndraw += 1
        assert ndraw == 1, ndraw
        log.append('- ch4 drawing removed for add_picture re-insertion')

        FA_CH = re.compile(r'[\u200c\u0600-\u06FF\uFB50-\uFDFF'
                           r'\uFE70-\uFEFF]')
        pinned = pinned_fa = 0
        for blk in (ch4, ch5):
            for r in [e for c in blk for e in c.findall(f'.//{W}r')]:
                txt = ''.join(x.text or '' for x in r.findall(W + 't'))
                if not txt.strip():
                    continue
                if r.find(W + 'rPr') is None:
                    ET.SubElement(r, W + 'rPr')
                if r.find(f'{W}rPr/{W}rFonts') is None:
                    is_fa = FA_CH.search(txt) is not None
                    rf = ET.Element(W + 'rFonts')
                    for a in ('ascii', 'hAnsi', 'cs'):
                        rf.set(W + a, 'B Nazanin' if is_fa else 'Cambria')
                    r.find(W + 'rPr').insert(0, rf)
                    if is_fa and r.find(f'{W}rPr/{W}rtl') is None:
                        ET.SubElement(r.find(W + 'rPr'), W + 'rtl')
                        pinned_fa += 1
                    pinned += 1
        assert pinned > 300, pinned
        log.append(f'- font-less ch4/ch5 runs pinned: {pinned} '
                   f'(FA->B Nazanin+rtl: {pinned_fa})')

        for blk, frag in ((ch4, 'صرفاً برای نمایش روش تحلیل آماری'),
                          (ch5, 'نمونه آموزشی است که بر پایه')):
            n = sum(1 for c in blk if c.tag == W + 'p' and frag in T(c))
            assert n == 1, (frag[:20], n)

        c_contents, c_tables, c_figs = [], [], []
        n0 = n1 = n2 = 0
        for blk in (ch1, ch2, ch3, ch4, ch5):
            for c in list(blk):
                if c.tag != W + 'p':
                    continue
                t = re.sub(r'\s+', ' ', T(c)).strip()
                if not t:
                    continue
                m = CAP_RE.match(t)
                if m:
                    label = seq_caption_surgery(c)
                    (c_tables if label == 'جدول' else c_figs).append(
                        (t, 'tableoffigures'))
                    continue
                ps = c.find(f'{W}pPr/{W}pStyle')
                sid = ps.get(W + 'val') if ps is not None else ''
                lvl = None
                if (t in ('فصل اول', 'فصل دوم', 'فصل سوم')
                        or t.startswith('فصل چهارم:')
                        or t.startswith('فصل پنجم:')):
                    lvl = ('0', (t, 'TOC1'), 'n0')
                elif sid == 'Heading2':
                    lvl = ('1', (t, 'TOC2'), 'n1')
                elif sid in ('Heading3', 'Heading4'):
                    lvl = ('2', (t, 'TOC3'), 'n2')
                if lvl:
                    _set_outline_lvl(c.find(W + 'pPr'), lvl[0])
                    c_contents.append(lvl[1])
                    if lvl[2] == 'n0':
                        n0 += 1
                    elif lvl[2] == 'n1':
                        n1 += 1
                    else:
                        n2 += 1

        fa, la = unify_refs(extract_refs(refs13) + extract_refs(refs5), log)
        refs = [text_para('منابع و مآخذ', jc='right', bold=True, sz='28',
                          outline=0),
                text_para('الف) منابع فارسی', jc='right', bold=True,
                          sz='28'),
                *[ref_para_fa(t) for t in fa],
                text_para('ب) منابع لاتین', jc='right', bold=True, sz='28'),
                *[ref_para_la(t) for t in la]]
        c_contents.append(('منابع و مآخذ', 'TOC1'))
        n0 += 1

        F = build_front_matter(c_contents, c_tables, c_figs)
        i_lists = next(i for i, c in enumerate(F)
                       if c.tag == W + 'p' and re.sub(
                           r'\s+', ' ', T(c)).strip() == 'فهرست مطالب')
        front7, lists4 = F[:i_lists], F[i_lists:]

        doc = ET.fromstring(zb.read('word/document.xml'))
        body = doc.find(W + 'body')
        model = body.find(W + 'sectPr')
        ids = [r.get('Id') for r in ET.fromstring(
            zb.read('word/_rels/document.xml.rels')).findall(
            f'{{{RN}}}Relationship')]
        n = 1
        while f'rId{n}' in ids:
            n += 1
        rid_empty, rid_roman, rid_body = f'rId{n}', f'rId{n+1}', f'rId{n+2}'

        new = (front7 + [joint_para(model, rid_empty)]
               + lists4 + [joint_para(model, rid_roman, fmt='upperRoman')]
               + ch1 + [joint_para(model, rid_body, fmt='decimal', start=1,
                                   title_pg=True)]
               + ch2 + [joint_para(model, rid_body, fmt='decimal',
                                   title_pg=True)]
               + ch3 + [joint_para(model, rid_body, fmt='decimal',
                                   title_pg=True)]
               + ch4 + [joint_para(model, rid_body, fmt='decimal',
                                   title_pg=True)]
               + ch5 + [page_break_para()] + refs)
        for c in list(body):
            body.remove(c)
        for c in new:
            body.append(c)
        body.append(sect_props(model, rid_body, fmt='decimal',
                               title_pg=True))

        styles = ET.fromstring(zb.read('word/styles.xml'))
        for s in list(styles.findall(W + 'style')):
            if s.get(W + 'styleId') in ('TOC1', 'TOC2', 'TOC3',
                                        'tableoffigures'):
                styles.remove(s)
        pg = model.find(W + 'pgSz')
        pm = model.find(W + 'pgMar')
        tab_pos = (int(pg.get(W + 'w')) - int(pm.get(W + 'left'))
                   - int(pm.get(W + 'right')))
        for sid, nm, sz in [('TOC1', 'toc 1', '24'),
                            ('TOC2', 'toc 2', '24'),
                            ('TOC3', 'toc 3', '24'),
                            ('tableoffigures', 'Table of Figures', '24')]:
            styles.append(rtl_toc_style(sid, nm, sz, tab_pos,
                                        bold=(sid == 'TOC1')))

        rels = ET.fromstring(zb.read('word/_rels/document.xml.rels'))
        for r in list(rels.findall(f'{{{RN}}}Relationship')):
            if re.fullmatch(r'footer\d+\.xml', r.get('Target') or ''):
                rels.remove(r)
        for rid, tgt in ((rid_empty, 'footer1.xml'),
                         (rid_roman, 'footer2.xml'),
                         (rid_body, 'footer3.xml')):
            nr = ET.SubElement(rels, f'{{{RN}}}Relationship')
            nr.set('Id', rid)
            nr.set('Type', f'{OFFDOC}/footer')
            nr.set('Target', tgt)
        ct = ET.fromstring(zb.read('[Content_Types].xml'))
        for e in list(ct):
            if (e.tag == f'{{{CTN}}}Override' and 'footer' in
                    (e.get('PartName') or '')):
                ct.remove(e)
        for k in (1, 2, 3):
            o = ET.SubElement(ct, f'{{{CTN}}}Override')
            o.set('PartName', f'/word/footer{k}.xml')
            o.set('ContentType', 'application/vnd.openxmlformats-officedocument'
                  '.wordprocessingml.footer+xml')
        settings_el = ET.fromstring(zb.read('word/settings.xml'))
        assert settings_el.find(W + 'updateFields') is None
        uf = ET.Element(W + 'updateFields')
        uf.set(W + 'val', '1')
        # CT_Settings: updateFields < hdrShapeDefaults < footnotePr <
        #              endnotePr < compat < docVars < rsids
        placed = False
        for follower in ('hdrShapeDefaults', 'footnotePr', 'endnotePr',
                         'compat', 'docVars', 'rsids'):
            anchor_el = settings_el.find(W + follower)
            if anchor_el is not None:
                settings_el.insert(list(settings_el).index(anchor_el), uf)
                placed = True
                break
        assert placed, 'no updateFields anchor in settings.xml'
        zoom = settings_el.find(W + 'zoom')
        if zoom is not None:
            # CT_Zoom takes w:percent, not w:val="bestFit"
            for k in list(zoom.attrib):
                del zoom.attrib[k]
            zoom.set(W + 'percent', '100')
        settings = ('<?xml version="1.0" encoding="UTF-8" '
                    'standalone="yes"?>\n'
                    + ET.tostring(settings_el, encoding='unicode'))

        with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
            for name in zb.namelist():
                if (name in ('word/document.xml',
                             'word/_rels/document.xml.rels',
                             'word/styles.xml', 'word/settings.xml',
                             '[Content_Types].xml')
                        or re.fullmatch(r'word/footer\d+\.xml', name)):
                    continue
                z.writestr(name, zb.read(name))
            xml_decl = ('<?xml version="1.0" encoding="UTF-8" '
                        'standalone="yes"?>\n')
            z.writestr('word/document.xml',
                       xml_decl + ET.tostring(doc, encoding='unicode'))
            z.writestr('word/_rels/document.xml.rels',
                       xml_decl + ET.tostring(rels, encoding='unicode'))
            z.writestr('word/styles.xml',
                       xml_decl + ET.tostring(styles, encoding='unicode'))
            z.writestr('word/settings.xml', settings)
            z.writestr('[Content_Types].xml',
                       xml_decl + ET.tostring(ct, encoding='unicode'))
            z.writestr('word/footer1.xml',
                       xml_decl + ET.tostring(footer_xml(False, ''),
                                              encoding='unicode'))
            z.writestr('word/footer2.xml',
                       xml_decl + ET.tostring(footer_xml(True, 'i'),
                                              encoding='unicode'))
            z.writestr('word/footer3.xml',
                       xml_decl + ET.tostring(footer_xml(True, '۱', 'fa-IR'),
                                              encoding='unicode'))

    from docx import Document as _DocxDocument
    from docx.shared import Cm
    _d = _DocxDocument(str(OUT))
    _w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    def _is_list_cache(p):
        _ps = p.find(f'{{{_w}}}pPr/{{{_w}}}pStyle')
        return _ps is not None and _ps.get(f'{{{_w}}}val') in (
            'TOC1', 'TOC2', 'TOC3', 'tableoffigures')

    _caps = [p for p in _d.element.body.findall(f'{{{_w}}}p')
             if not _is_list_cache(p) and ''.join(
                 x.text or '' for x in p.findall(f'.//{{{_w}}}t'))
             .startswith('شکل ۴-۱')]
    assert len(_caps) == 1, len(_caps)
    _figp = _caps[0].getprevious()
    assert _figp is not None and _figp.tag == f'{{{_w}}}p'
    assert _figp.find(f'.//{{{_w}}}drawing') is None
    _runs = _figp.findall(f'{{{_w}}}r')
    assert len(_runs) >= 1
    _para = next(p for p in _d.paragraphs if p._p is _figp)
    _para.runs[0].add_picture(str(CH4_FIG), width=Cm(14))
    _d.save(str(OUT))
    validate_package(OUT)
    log.append('- fig 4-1 embedded via python-docx add_picture (14cm)')
    log.append('- package gate: well-formed + OPC + schema orders + enums OK')

    abstract_n = len(ABSTRACT.split())
    report = ['# گزارش ساخت thesis_complete_draft_v2', '',
              'مبنا: فقط فایل‌های متنی main (هیچ فایل Word قدیمی خوانده نشد).',
              'سازنده‌ها: `build_word.py` (فصل ۱-۳) + `build_ch4_word.py` '
              '(فصل ۴ واقع‌نما) + `build_ch5_word.py` (فصل ۵ واقع‌نما).', '']
    report += ['## ترتیب صفحات مقدماتی (۱۱ صفحه)', '',
               '۱. بسم‌الله', '۲. عنوان فارسی', '۳. تأییدیه هیئت داوران '
               '(placeholder)', '۴. تعهدنامه اصالت اثر (placeholder)',
               '۵. تقدیم (متن پیش‌فرض مأموریت)', '۶. سپاسگزاری (متن پیش‌فرض '
               'مأموریت)', '۷. چکیده فارسی + ۵ کلیدواژه', '۸. فهرست مطالب '
               '(TOC واقعی، سه‌سطحی، RTL)', '۹. فهرست جدول‌ها (SEQ با \\c، RTL)',
               '۱۰. فهرست شکل‌ها (SEQ با \\c، RTL)', '۱۱. فهرست علائم (جدول واقعی '
               '۱۶ ردیفه)', '']
    report += ['## آمار', '',
               f'- جدول‌ها: 12 (۱۱ کپشن جدول در فهرست جدول‌ها + جدول گرافیکی شکل ۲-۱ در فهرست شکل‌ها) | شکل‌ها: 1 تصویر (۴-۱)',
               f'- منابع نهایی: {len(fa)} فارسی + {len(la)} لاتین',
               f'- چکیده: {abstract_n} کلمه (حداکثر یک صفحه طبق راهنما)',
               '- تعداد صفحات: ۱۱ صفحه مقدماتی (قطعی) + فصول؛ کل دقیق پس از '
               'باز شدن در Word مشخص می‌شود', '']
    report += ['## placeholderهای باقی‌مانده', '',
               '- فرم تأییدیه و صورت‌جلسه دفاع (پس از جلسه دفاع)',
               '- فرم تعهدنامه اصالت اثر (پس از تکمیل)', '']
    report += ['## وضعیت فصل ۴ و ۵', '',
               'هر دو فصل نسخهٔ آموزشی مبتنی بر داده شبیه‌سازی‌شده‌اند؛ هشدار '
               'دقیقاً یک‌بار در ابتدای هر فصل حفظ شده است.', '']
    report += ['## راستی‌آزمایی میانی‌ها', ''] + log + ['']
    report += ['## یادداشت‌ها', '',
               '- شماره‌گذاری: مقدماتی (۱-۷) بی‌شماره؛ فهرست‌ها (۸-۱۱) رومی؛ '
               'متن فارسی پیوسته از فصل ۱؛ صفحه اول هر فصل بی‌شماره ولی به حساب.',
               '- قالب شماره صفحه متن و پاورقی decimal است (معتبر در طرحواره)؛ '
               'با Numeral=Context در Word فارسی و زبان fa-IR فوتر، ارقام فارسی '
               '۰۱۲۳ نمایش داده می‌شوند. مقدار hindi عضو ST_NumberFormat نیست.',
               '- رفع خطای Unreadable Content: نوع rel فوترها (officeDocument)، '
               'ترتیب pPr (spacing/ind پیش از jc)، ترتیب pPr استایل‌های TOC، '
               'جای tblGrid (فرزند tbl)، جای updateFields (پیش از compat)، '
               'zoom (percent)، ترتیب tblPr جدول‌ها (bidiVisual/tblBorders) و '
               'مقادیر numFmt همگی با XSD انتقالی ECMA-376 اعتبارسنجی شدند؛ '
               'گیت validate_package پس از هر ساخت اجرا می‌شود.',
               '- سایه خاکستری فیلدها تنظیم سطح Word است (View/Options) و در '
               'فایل ذخیره نمی‌شود؛ همه فیلدها نتیجه کش‌شده دارند.',
               '- VIF / نماد CI / نماد β در متن فصل‌ها با همین صورت نیامده‌اند '
               '(مفهوم آن‌ها هست) ولی طبق دستور مأموریت در فهرست علائم‌اند.',
               '- شکل ۲-۱ در منبع فقط کپشن دارد (بدون فایل تصویر)؛ گرافیک آن یک جدول مفهومی است که با همان کپشن در فهرست شکل‌ها پوشش داده می‌شود.',
               '- اصلاح باگ پرانتز: ران‌های فارسی بدون فونت فصل ۴ به B Nazanin با w:rtl ارتقا یافتند (نویسه‌های پرانتز دست‌نخورده: U+0028/U+0029).',
               '- تصویر شکل ۴-۱ با run.add_picture درج شده (کتابخانه مالک بایت‌ها، rel و drawing است).',
               '- کپشن‌ها فیلد SEQ قفل‌شده (fldLock) با همان شماره فارسی دارند؛ فهرست‌ها با \\c ساخته می‌شوند و F9 شماره‌ها را عوض نمی‌کند.', '']
    REPORT.write_text('\n'.join(report), encoding='utf-8')
    print(f'stripped={nstrip} pinned={pinned} outline=({n0},{n1},{n2}) '
          f'tables={len(c_tables)} figs={len(c_figs)}')
    print(f'refs_fa={len(fa)} refs_latin={len(la)} '
          f'abstract_words={abstract_n}')
    print('wrote', OUT)
    print('wrote', REPORT)


if __name__ == '__main__':
    main()
