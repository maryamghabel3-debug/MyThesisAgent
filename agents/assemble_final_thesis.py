"""Mission 21: assemble thesis_final_draft.docx.

New 10-page front matter (basmala, title, approval, commitment,
dedication, acknowledgment, abstract, TOC/lists with real fields)
+ chapters 1-5 (text-untouched) + unified references.

Numbering: front matter unnumbered (empty footer); chapters numbered
from ch1 in Persian digits (titlePg: first page of each chapter
counted but unnumbered).
"""
import copy, re, zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
ET.register_namespace('w', W[1:-1])
ET.register_namespace('r', R[1:-1])

BASE = 'output/final/thesis_ch1_to_ch3.docx'
CH4 = 'output/final/chapter4_simulated_real_pipeline.docx'
CH5 = 'output/final/chapter5_simulated_real_pipeline.docx'
OUT = 'output/final/thesis_final_draft.docx'

WARNING_NEW = ('این فصل نمونه آموزشی مبتنی بر داده شبیه‌سازی‌شده است و صرفاً '
               'برای نمایش روش تحلیل آماری تهیه شده است. متن نهایی فصل پس از '
               'تحلیل داده‌های واقعی پژوهش بازنویسی خواهد شد.')


def T(p):
    return ''.join(x.text or '' for x in p.findall(f'.//{W}t'))


def run_el(text, font='B Nazanin', sz='24', bold=False, italic=False,
           vanish=False):
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
    t = ET.SubElement(r, W + 't')
    t.text = text
    if text != text.strip():
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def para_el(runs, jc=None, rtl=True, exact='580', after=None, outline=None):
    p = ET.Element(W + 'p')
    pr = ET.SubElement(p, W + 'pPr')
    if rtl:
        ET.SubElement(pr, W + 'bidi')
    if jc:
        ET.SubElement(pr, W + 'jc').set(W + 'val', jc)
    if outline is not None:
        ET.SubElement(pr, W + 'outlineLvl').set(W + 'val', str(outline))
    if exact or after:
        sp = ET.SubElement(pr, W + 'spacing')
        if exact:
            sp.set(W + 'line', exact)
            sp.set(W + 'lineRule', 'exact')
        if after:
            sp.set(W + 'after', after)
    for r in runs:
        p.append(r)
    return p


def text_para(text, **kw):
    return para_el([run_el(text, font=kw.pop('font', 'B Nazanin'),
                           sz=kw.pop('sz', '24'),
                           bold=kw.pop('bold', False))], **kw)


def page_break_para():
    p = ET.Element(W + 'p')
    r = ET.SubElement(p, W + 'r')
    ET.SubElement(r, W + 'br').set(W + 'type', 'page')
    return p


def fld_para(kind, instr):
    """Hidden TC marker paragraph."""
    p = ET.Element(W + 'p')
    pr = ET.SubElement(p, W + 'pPr')
    ET.SubElement(pr, W + 'bidi')
    for kt, tx in (('begin', None), (None, instr), ('end', None)):
        r = ET.SubElement(p, W + 'r')
        rp = ET.SubElement(r, W + 'rPr')
        ET.SubElement(rp, W + 'vanish')
        rf = ET.SubElement(rp, W + 'rFonts')
        rf.set(W + 'ascii', 'B Nazanin')
        rf.set(W + 'hAnsi', 'B Nazanin')
        if kt:
            ET.SubElement(r, W + 'fldChar').set(W + 'fldCharType', kt)
        else:
            ET.SubElement(r, W + 'instrText').text = tx
    return p


def toc_field_paras(instr, cache_texts):
    """TOC field: code para + cached title paras + end para."""
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
    for t in cache_texts:
        out.append(text_para(t, jc='both'))
    pe = ET.Element(W + 'p')
    ET.SubElement(ET.SubElement(pe, W + 'pPr'), W + 'bidi')
    re_ = ET.SubElement(pe, W + 'r')
    ET.SubElement(re_, W + 'fldChar').set(W + 'fldCharType', 'end')
    out.append(pe)
    return out


FA_ORDER = 'ابپتثجچحخدرزژسشصضطظعغفقکگلمنوهی'
FA_NORM = str.maketrans({'ك': 'ک', 'ي': 'ی', 'ى': 'ی', 'ة': 'ه',
                         'أ': 'ا', 'إ': 'ا', 'آ': 'ا', 'ؤ': 'و'})
FA_IDX = {c: i for i, c in enumerate(FA_ORDER)}


def norm_text(s):
    return re.sub(r'\s+', ' ', s).strip()


def fa_key(s):
    s = norm_text(s).translate(FA_NORM)
    return tuple(FA_IDX.get(c, 1000 + ord(c)) for c in s)


def extract_refs(children):
    """Merge-compatible: skip subheads, classify by first char, assert fit."""
    refs, cur, seen_head = [], None, False
    for c in children:
        if c.tag != W + 'p':
            continue
        t = norm_text(T(c))
        if not t:
            continue
        if t == 'منابع' and not seen_head:
            seen_head = True
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


def ref_para(text, fa):
    if fa:
        return para_el([run_el(text, font='B Nazanin', sz='24')],
                       jc='right', rtl=True)
    return para_el([run_el(text, font='Times New Roman', sz='22')],
                   jc='left', rtl=True)


def joint(footer_rid, numbered, restart=False):
    p = ET.Element(W + 'p')
    sp = ET.SubElement(ET.SubElement(p, W + 'pPr'), W + 'sectPr')
    fr = ET.SubElement(sp, W + 'footerReference')
    fr.set(W + 'type', 'default')
    fr.set(R + 'id', footer_rid)
    if numbered:
        pg = ET.SubElement(sp, W + 'pgNumType')
        pg.set(W + 'fmt', 'hindi')
        if restart:
            pg.set(W + 'start', '1')
        ET.SubElement(sp, W + 'titlePg')
    return p


def main():
    zb = zipfile.ZipFile(BASE)
    doc = ET.fromstring(zb.read('word/document.xml'))
    body = doc.find(W + 'body')
    kids = list(body)
    sect_final = body.find(W + 'sectPr')

    # ---- footer rIds + sanity ----
    rels = ET.fromstring(zb.read('word/_rels/document.xml.rels'))
    rid_empty = rid_page = None
    for r in rels.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        if r.get('Target') == 'footer1.xml':
            rid_empty = r.get('Id')
        if r.get('Target') == 'footer10.xml':
            rid_page = r.get('Id')
    assert rid_empty and rid_page, 'footer rels'
    assert 'PAGE' not in zb.read('word/footer1.xml').decode('utf8')
    assert 'PAGE' in zb.read('word/footer10.xml').decode('utf8')
    assert sect_final.find(W + 'headerReference') is None

    # ---- slice thesis: ch1/ch2/ch3 body + old refs ----
    tparas = [(i, c) for i, c in enumerate(kids) if c.tag == W + 'p']
    tmap = {T(c).strip(): i for i, c in tparas}
    i_f1, i_f2, i_f3 = tmap['فصل اول'], tmap['فصل دوم'], tmap['فصل سوم']
    i_ref = [i for i, c in tparas if T(c).strip() == 'منابع'][0]
    ch1 = copy.deepcopy(kids[i_f1:i_f2])
    ch2 = copy.deepcopy(kids[i_f2:i_f3])
    ch3 = copy.deepcopy(kids[i_f3:i_ref])
    thesis_old_refs = kids[i_ref:]

    # ---- ch4 full, ch5 minus refs ----
    def load(fn):
        z = zipfile.ZipFile(fn)
        return z, ET.fromstring(z.read('word/document.xml'))
    z4, d4 = load(CH4)
    z5, d5 = load(CH5)
    k4 = list(d4.find(W + 'body'))
    k5 = list(d5.find(W + 'body'))
    ch5_ref_i = [i for i, c in enumerate(k5)
                 if c.tag == W + 'p' and T(c).strip() == 'منابع'][0]
    ch4 = copy.deepcopy(k4)
    ch5 = copy.deepcopy(k5[:ch5_ref_i])
    ch5_old_refs = k5[ch5_ref_i:]

    # ---- strip inline sectPr (must be empty paras) ----
    nstrip = 0
    for blk in (ch1, ch2, ch3, ch4, ch5):
        for c in list(blk):
            if c.tag == W + 'p' and c.find(f'{W}pPr/{W}sectPr') is not None:
                assert T(c).strip() == '', 'non-empty sectPr para'
                blk.remove(c)
                nstrip += 1

    # ---- pin Cambria on font-less moved runs ----
    pinned = 0
    for blk in (ch4, ch5):
        for r in [e for c in blk for e in c.findall(f'.//{W}r')]:
            txt = ''.join(x.text or '' for x in r.findall(W + 't'))
            if not txt.strip():
                continue
            if r.find(W + 'rPr') is None:
                ET.SubElement(r, W + 'rPr')
            if r.find(f'{W}rPr/{W}rFonts') is None:
                rf = ET.Element(W + 'rFonts')
                rf.set(W + 'ascii', 'Cambria')
                rf.set(W + 'hAnsi', 'Cambria')
                rf.set(W + 'cs', 'Cambria')
                r.find(W + 'rPr').insert(0, rf)
                pinned += 1
    assert pinned == 362, pinned

    # ---- ch4/ch5 titles -> 28pt; replace warnings ----
    for blk, head in ((ch4, 'فصل چهارم:'), (ch5, 'فصل پنجم:')):
        hits = [c for c in blk if c.tag == W + 'p'
                and T(c).strip().startswith(head)]
        assert len(hits) == 1, head
        for r in hits[0].findall(W + 'r'):
            sz = r.find(f'{W}rPr/{W}sz')
            if sz is not None:
                sz.set(W + 'val', '56')
            cz = r.find(f'{W}rPr/{W}szCs')
            if cz is not None:
                cz.set(W + 'val', '56')
    nw = 0
    for blk, old in ((ch4, 'صرفاً برای نمایش روش تحلیل آماری'),
                     (ch5, 'نمونه آموزشی است که بر پایه')):
        hits = [c for c in blk if c.tag == W + 'p' and old in T(c)]
        assert len(hits) == 1, old[:20]
        new = para_el([run_el(WARNING_NEW, font='B Nazanin', sz='22',
                              italic=True)], jc='center')
        blk[blk.index(hits[0])] = new
        nw += 1

    # ---- outline levels + TC markers + TOC caches ----
    toc_contents, toc_tables, toc_figs = [], [], []
    n0 = n1 = 0
    for blk in (ch1, ch2, ch3, ch4, ch5):
        for c in list(blk):
            if c.tag != W + 'p':
                continue
            t = re.sub(r'\s+', ' ', T(c)).strip()
            if (t in ('فصل اول', 'فصل دوم', 'فصل سوم')
                    or t.startswith('فصل چهارم:')
                    or t.startswith('فصل پنجم:')):
                ET.SubElement(c.find(W + 'pPr'), W + 'outlineLvl').set(
                    W + 'val', '0')
                toc_contents.append(t)
                n0 += 1
            elif re.match(r'^[۰-۹0-9]+-[۰-۹0-9]+', t):
                ET.SubElement(c.find(W + 'pPr'), W + 'outlineLvl').set(
                    W + 'val', '1')
                toc_contents.append(t)
                n1 += 1
            m = re.match(r'^(جدول|شکل) [۰-۹0-9]+-[۰-۹0-9]+\.', t)
            if m:
                assert '"' not in t, t[:40]
                flag = 'T' if m.group(1) == 'جدول' else 'F'
                blk.insert(blk.index(c) + 1,
                           fld_para(flag, f'TC "{t}" \\f {flag} \\l 1'))
                (toc_tables if flag == 'T' else toc_figs).append(t)

    # ---- unified refs ----
    seen, fa, la = set(), [], []
    for is_lat, t in (extract_refs(thesis_old_refs)
                      + extract_refs(ch5_old_refs)):
        if t in seen:
            continue
        seen.add(t)
        (la if is_lat else fa).append(t)
    fa.sort(key=fa_key)
    la.sort(key=str.lower)
    assert (len(fa), len(la)) == (13, 71), (len(fa), len(la))
    refs = [text_para('منابع', jc='center', bold=True, outline=0),
            text_para('منابع فارسی', bold=True),
            *[ref_para(t, True) for t in fa],
            text_para('منابع لاتین', bold=True),
            *[ref_para(t, False) for t in la]]
    toc_contents.append('منابع')
    n0 += 1

    # ---- front matter ----
    F = []
    F += [ET.Element(W + 'p') for _ in range(4)]
    F.append(text_para('بسم الله الرحمن الرحیم', jc='center', sz='28'))
    F.append(page_break_para())
    for t, s in [('وزارت علوم، تحقیقات و فناوری', '28'),
                 ('دانشگاه علم و هنر یزد', '28'),
                 ('دانشکده انسانی', '28'),
                 ('گروه روان‌شناسی', '28')]:
        F.append(text_para(t, jc='center', sz=s, bold=True))
    F.append(ET.Element(W + 'p'))
    F.append(text_para('پایان‌نامه کارشناسی ارشد', jc='center', sz='32',
                       bold=True))
    F.append(ET.Element(W + 'p'))
    F.append(text_para('عنوان:', jc='center', sz='28', bold=True))
    for t in ['تعیین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه',
              'در رابطه بین اضطراب صفتی و رفتارهای مالی پرخطر',
              'در معامله‌گران بازارهای مالی']:
        F.append(text_para(t, jc='center', sz='32', bold=True))
    F.append(ET.Element(W + 'p'))
    for t in ['نام دانشجو: مریم قابل',
              'استاد راهنما: دکتر سعید وزیری یزدی',
              'دانشکده انسانی', 'گروه روان‌شناسی', '۱۴۰۵']:
        F.append(text_para(t, jc='center', sz='28', bold=True))
    F.append(page_break_para())
    F.append(text_para('تأییدیه پایان‌نامه', jc='center', sz='28', bold=True))
    F.append(ET.Element(W + 'p'))
    F.append(text_para('این پایان‌نامه با عنوان', jc='both'))
    F.append(text_para('«تعیین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه در رابطه '
                       'بین اضطراب صفتی و رفتارهای مالی پرخطر در معامله‌گران '
                       'بازارهای مالی»', jc='center', bold=True))
    F.append(text_para('تهیه‌شده توسط', jc='both'))
    F.append(text_para('مریم قابل', jc='center', bold=True))
    F.append(text_para('به عنوان پایان‌نامه کارشناسی ارشد در رشته روان‌شناسی بالینی',
                       jc='both'))
    F.append(text_para('دانشکده انسانی دانشگاه علم و هنر یزد', jc='both'))
    F.append(text_para('در تاریخ ....... به عنوان پایان‌نامه‌ای کامل و قابل دفاع '
                       'پذیرفته شد.', jc='both'))
    F.append(ET.Element(W + 'p'))
    for t in ['امضاء استاد راهنما: دکتر سعید وزیری یزدی', 'امضاء داور:',
              'امضاء رئیس گروه:', 'تاریخ دفاع:']:
        F.append(text_para(t, jc='both'))
    F.append(page_break_para())
    F.append(text_para('تعهدنامه', jc='center', sz='28', bold=True))
    F.append(ET.Element(W + 'p'))
    F.append(text_para('اینجانب مریم قابل، دانشجوی کارشناسی ارشد روان‌شناسی بالینی '
                       'دانشکده انسانی دانشگاه علم و هنر یزد، تعهد می‌کنم که این '
                       'پایان‌نامه حاصل کار شخصی بنده بوده و تمام منابع مورد استفاده '
                       'در آن به درستی ذکر شده‌اند. در صورت کشف هرگونه تخلف علمی، '
                       'مسئولیت آن بر عهدهٔ اینجانب خواهد بود.', jc='both'))
    F.append(ET.Element(W + 'p'))
    for t in ['نام و نام خانوادگی: مریم قابل', 'امضاء:', 'تاریخ:']:
        F.append(text_para(t, jc='both'))
    F.append(page_break_para())
    F.append(text_para('تقدیم به:', jc='center', sz='28', bold=True))
    F.append(ET.Element(W + 'p'))
    F.append(text_para('[این بخش توسط دانشجو تکمیل می‌شود]', jc='center'))
    F.append(page_break_para())
    F.append(text_para('سپاسگزاری', jc='center', sz='28', bold=True))
    F.append(ET.Element(W + 'p'))
    F.append(text_para('[این بخش توسط دانشجو تکمیل می‌شود]', jc='center'))
    F.append(page_break_para())
    F.append(text_para('چکیده', jc='center', sz='28', bold=True))
    F.append(ET.Element(W + 'p'))
    abstract = ('پژوهش حاضر با هدف تعیین نقش میانجی‌گر طرحواره‌های ناسازگار اولیه '
                'در رابطه بین اضطراب صفتی و رفتارهای مالی پرخطر در معامله‌گران '
                'بازارهای مالی انجام شد. این پژوهش از نظر هدف کاربردی و از نظر روش '
                'توصیفی-همبستگی با رویکرد مدل‌سازی معادلات ساختاری بود. جامعهٔ آماری '
                'پژوهش را کلیهٔ معامله‌گران حقیقی فارسی‌زبان با حداقل ۶ ماه سابقهٔ '
                'معامله تشکیل می‌دادند و نمونه (n=185) با روش نمونه‌گیری در دسترس و '
                'گلوله‌برفی انتخاب شد. ابزارهای پژوهش شامل پرسشنامه اضطراب حالت-صفت '
                'اسپیلبرگر (فرم صفتی)، پرسشنامه طرحواره‌های ناسازگار اولیه یانگ '
                '(فرم کوتاه ۷۵ گویه‌ای) و مقیاس تحمل ریسک مالی گرابل و لایتون بود. '
                'داده‌ها با نرم‌افزارهای SPSS و Amos و ماکروی PROCESS تحلیل شدند. '
                'یافته‌ها نشان داد که اضطراب صفتی با رفتارهای مالی پرخطر (r=0.327)، '
                'طرحواره‌های ناسازگار اولیه (r=0.501) و رفتارهای مالی پرخطر (r=0.321) '
                'رابطه مثبت و معنادار دارد. همچنین طرحواره‌های ناسازگار اولیه نقش '
                'میانجی‌گر جزئی در رابطه بین اضطراب صفتی و رفتارهای مالی پرخطر ایفا '
                'می‌کنند (اثر غیرمستقیم=0.070، CI 95%: 0.011 تا 0.133). نتایج نشان داد '
                'که طرحواره‌های ناسازگار اولیه بخشی از اثر اضطراب صفتی بر رفتارهای مالی '
                'پرخطر را تبیین می‌کنند و مداخله‌های شناختی-رفتاری می‌توانند بر این مسیر '
                'اثر بگذارند.')
    F.append(text_para(abstract, jc='both'))
    F.append(para_el([run_el('کلمات کلیدی: ', bold=True),
                      run_el('اضطراب صفتی؛ طرحواره‌های ناسازگار اولیه؛ رفتارهای مالی '
                             'پرخطر؛ معامله‌گران بازارهای مالی؛ مدل‌سازی معادلات '
                             'ساختاری')], jc='both'))
    F.append(page_break_para())
    F.append(text_para('فهرست مطالب', jc='center', sz='28', bold=True))
    F += toc_field_paras('TOC \\o "1-2" \\h \\z \\u', toc_contents)
    F.append(page_break_para())
    F.append(text_para('فهرست جداول', jc='center', sz='28', bold=True))
    F += toc_field_paras('TOC \\f T \\h \\z', toc_tables)
    F.append(page_break_para())
    F.append(text_para('فهرست شکل‌ها', jc='center', sz='28', bold=True))
    F += toc_field_paras('TOC \\f F \\h \\z', toc_figs)

    # ---- assemble ----
    new = (F + [joint(rid_empty, False)] + ch1 + [joint(rid_page, True, True)]
           + ch2 + [joint(rid_page, True)] + ch3 + [joint(rid_page, True)]
           + ch4 + [joint(rid_page, True)] + ch5 + [page_break_para()] + refs)
    for c in list(body):
        body.remove(c)
    for c in new:
        body.append(c)

    # ---- final sectPr: ch5 props (ordered) ----
    order = ['footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz',
             'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType',
             'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
             'textDirection', 'bidi', 'rtlGutter', 'docGrid',
             'printerSettings']
    keep = {}
    for c in list(sect_final):
        tag = c.tag[len(W):]
        if tag in ('footerReference', 'pgNumType', 'titlePg'):
            sect_final.remove(c)
        else:
            keep.setdefault(tag, []).append(c)
            sect_final.remove(c)
    fr = ET.SubElement(sect_final, W + 'footerReference')
    fr.set(W + 'type', 'default')
    fr.set(R + 'id', rid_page)
    pg = ET.Element(W + 'pgNumType')
    pg.set(W + 'fmt', 'hindi')
    keep.setdefault('pgNumType', []).append(pg)
    keep.setdefault('titlePg', []).append(ET.Element(W + 'titlePg'))
    for tag in order:
        for c in keep.get(tag, []):
            sect_final.append(c)
    for tag, cs in keep.items():
        if tag not in order:
            for c in cs:
                sect_final.append(c)
    body.append(sect_final)

    # ---- image + rels + content types ----
    img = z4.read('word/media/image1.png')
    for r in rels.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        ids = [r.get('Id') for r in rels.findall(
            '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')]
    n = 1
    while f'rId{n}' in ids:
        n += 1
    nr = ET.SubElement(rels, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
    nr.set('Id', f'rId{n}')
    nr.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
    nr.set('Target', 'media/image1.png')
    for c in ch4:
        for e in c.findall(f'.//{W}drawing'):
            for b in e.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
                b.set(R + 'embed', f'rId{n}')
    ct = ET.fromstring(zb.read('[Content_Types].xml'))
    if not [e for e in ct if e.get('Extension') == 'png']:
        d = ET.SubElement(ct, '{http://schemas.openxmlformats.org/package/2006/content-types}Default')
        d.set('Extension', 'png')
        d.set('ContentType', 'image/png')

    # ---- write ----
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for n_ in zb.namelist():
            if n_ in ('word/document.xml', 'word/_rels/document.xml.rels',
                      '[Content_Types].xml'):
                continue
            z.writestr(n_, zb.read(n_))
        z.writestr('word/document.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   + ET.tostring(doc, encoding='unicode'))
        z.writestr('word/_rels/document.xml.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   + ET.tostring(rels, encoding='unicode'))
        z.writestr('[Content_Types].xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   + ET.tostring(ct, encoding='unicode'))
        z.writestr('word/media/image1.png', img)
    assert (nstrip, nw) == (3, 2), (nstrip, nw)
    assert (len(toc_tables), len(toc_figs)) == (11, 2), (len(toc_tables),
                                                         len(toc_figs))
    print(f'stripped_empty_sect={nstrip} pinned={pinned} warnings={nw}')
    print(f'outline_L0={n0} outline_L1={n1} toc_tables={len(toc_tables)} '
          f'toc_figs={len(toc_figs)}')
    print(f'refs_fa={len(fa)} refs_latin={len(la)}')
    print('abstract_words=', len(abstract.split()))


main()
