# -*- coding: utf-8 -*-
"""Merge the thesis Word files into one unified document (NO content changes).

Inputs (committed files on main):
  output/final/thesis_ch1_to_ch3.docx   (base: front matter, ch1-3, footnotes,
                                         page numbers, section setup)
  output/final/chapter4_simulated_real_pipeline.docx  (ch4, 9 tables, 1 figure)
  output/final/chapter5_simulated_real_pipeline.docx  (ch5 + its references)
Output:
  output/final/thesis_complete_draft.docx

Method: XML-level merge. Body elements are MOVED (not re-created), so every
paragraph/table keeps its own formatting. Only structural operations:
  1. cut the per-part reference blocks (thesis ch1-3 refs, ch5 refs),
  2. join bodies with page-break paragraphs (same pattern the builders use),
  3. append ONE unified, deduplicated, alphabetically sorted reference list,
  4. carry over ch4's figure (media file + relationship + reference fix),
  5. pin ch4's style-chain font (Cambria, = its standalone rendering) onto the
     text runs that lack direct fonts, so the Normal-style difference between
     the builders cannot change any rendering.
Styles, footnotes, numbering, headers/footers, section setup and margins of
the base document are left untouched (margins/spacing/page-numbers already
comply with the mission spec).
"""
import re
import zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
for _p, _u in [('w', W), ('r', R),
               ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
               ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
               ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
               ('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')]:
    ET.register_namespace(_p, _u)

BASE = 'output/final/thesis_ch1_to_ch3.docx'
CH4 = 'output/final/chapter4_simulated_real_pipeline.docx'
CH5 = 'output/final/chapter5_simulated_real_pipeline.docx'
OUT = 'output/final/thesis_complete_draft.docx'

NS = {'w': W, 'r': R}
Q = lambda tag: f'{{{W}}}{tag}'          # noqa: E731
QR = lambda tag: f'{{{R}}}{tag}'         # noqa: E731

FA_ORDER = 'ابپتثجچحخدرزژسشصضطظعغفقکگلمنوهی'
FA_NORM = str.maketrans({'ك': 'ک', 'ي': 'ی', 'ى': 'ی', 'ة': 'ه',
                         'أ': 'ا', 'إ': 'ا', 'آ': 'ا', 'ؤ': 'و'})
FA_IDX = {c: i for i, c in enumerate(FA_ORDER)}


def para_text(p):
    return ''.join(t.text or '' for t in p.findall(f'.//{Q("t")}'))


def norm_text(s):
    return re.sub(r'\s+', ' ', s).strip()


def fa_key(s):
    s = norm_text(s).translate(FA_NORM)
    return tuple(FA_IDX.get(c, 1000 + ord(c)) for c in s)


def page_break_para():
    p = ET.Element(Q('p'))
    r = ET.SubElement(p, Q('r'))
    ET.SubElement(r, Q('br'), {'{%s}type' % W: 'page'})
    return p


def split_body(root):
    """Split body children into (content, refs_block). Refs start at the
    paragraph whose full text is exactly 'منابع'."""
    body = root.find(Q('body'))
    kids = list(body)
    assert kids[-1].tag == Q('sectPr'), 'body must end with sectPr'
    ref_at = None
    for i, el in enumerate(kids[:-1]):
        if el.tag == Q('p') and norm_text(para_text(el)) == 'منابع':
            ref_at = i
            break
    if ref_at is None:
        return kids[:-1], []
    nxt = norm_text(para_text(kids[ref_at + 1]))
    assert nxt.startswith('الف)'), f'refs heading not followed by FA subhead: {nxt[:30]}'
    return kids[:ref_at], kids[ref_at:-1]


def extract_refs(refs_block):
    """Return (heading_el, fa_sub_el, latin_sub_el, [(is_latin, norm, el)])."""
    assert refs_block, 'empty refs block'
    head, fa_sub, la_sub, entries = refs_block[0], None, None, []
    assert norm_text(para_text(head)) == 'منابع'
    cur = None
    for el in refs_block[1:]:
        if el.tag != Q('p'):
            entries.append((cur, norm_text(para_text(el)), el))
            continue
        t = norm_text(para_text(el))
        if not t:
            continue
        if t.startswith('الف)'):
            fa_sub, cur = el, False
        elif t.startswith('ب)'):
            la_sub, cur = el, True
        else:
            is_lat = bool(re.match(r'[A-Za-z]', t))
            if cur is not None:
                assert is_lat == cur, f'entry under wrong subhead: {t[:40]}'
            entries.append((is_lat, t, el))
    assert fa_sub is not None and la_sub is not None, 'both subheads required'
    return head, fa_sub, la_sub, entries


def main():
    zb = zipfile.ZipFile(BASE)
    z4 = zipfile.ZipFile(CH4)
    z5 = zipfile.ZipFile(CH5)
    rb = ET.fromstring(zb.read('word/document.xml'))
    r4 = ET.fromstring(z4.read('word/document.xml'))
    r5 = ET.fromstring(z5.read('word/document.xml'))

    base_content, base_refs = split_body(rb)
    ch4_content, ch4_refs = split_body(r4)
    ch5_content, ch5_refs = split_body(r5)
    assert not ch4_refs, 'ch4 unexpectedly has a refs block'
    assert base_refs and ch5_refs, 'refs blocks missing'

    # ---- font pinning for ch4 (see module docstring) ----
    pinned = 0
    for r in [e for p in ch4_content for e in p.findall(f'.//{Q("r")}')]:
        txt = ''.join(t.text or '' for t in r.findall(Q('t')))
        if not txt.strip():
            continue
        rpr = r.find(Q('rPr'))
        if rpr is not None and rpr.find(Q('rFonts')) is not None:
            continue
        if rpr is None:
            rpr = ET.Element(Q('rPr'))
            r.insert(0, rpr)
        else:
            for _ in list(rpr):
                pass
        rf = ET.Element(Q('rFonts'), {'{%s}ascii' % W: 'Cambria',
                                      '{%s}hAnsi' % W: 'Cambria',
                                      '{%s}cs' % W: 'Cambria'})
        rpr.insert(0, rf)
        pinned += 1
    # NOTE: ch4 tables also live outside <w:p>?? no - runs found via .// covers tbl.
    # recount properly over all ch4 content subtrees:
    assert pinned > 0

    # ---- unified references ----
    b_head, b_fa, b_la, b_en = extract_refs(base_refs)
    _, _, _, c_en = extract_refs(ch5_refs)
    seen, fa_list, la_list, dupes = {}, [], [], 0
    for is_lat, t, el in b_en + c_en:
        if t in seen:
            dupes += 1
            continue
        seen[t] = True
        (la_list if is_lat else fa_list).append((t, el))
    fa_list.sort(key=lambda x: fa_key(x[0]))
    la_list.sort(key=lambda x: x[0].lower())

    # ---- image carry-over (ch4 has exactly one) ----
    rels4 = ET.fromstring(z4.read('word/_rels/document.xml.rels'))
    RN = 'http://schemas.openxmlformats.org/package/2006/relationships'
    img_target, img_old_id = None, None
    for rel in rels4.findall(f'{{{RN}}}Relationship'):
        if rel.get('Target', '').startswith('media/'):
            img_target, img_old_id = rel.get('Target'), rel.get('Id')
    assert img_target == 'media/image1.png', img_target
    assert 'word/media/image1.png' not in zb.namelist(), 'media collision'
    rels_b = ET.fromstring(zb.read('word/_rels/document.xml.rels'))
    nums = [int(m.group(1)) for rel in rels_b.findall(f'{{{RN}}}Relationship')
            for m in [re.fullmatch(r'rId(\d+)', rel.get('Id', ''))] if m]
    new_id = f'rId{max(nums) + 1}'
    img_bytes = z4.read('word/' + img_target)
    # repoint the drawing reference inside copied ch4 content
    fixed = 0
    for el in ch4_content:
        for blip in el.findall(f'.//{{http://schemas.openxmlformats.org/drawingml/2006/main}}blip'):
            if blip.get(QR('embed')) == img_old_id:
                blip.set(QR('embed'), new_id)
                fixed += 1
    assert fixed == 1, f'expected 1 drawing repoint, got {fixed}'
    ET.register_namespace('', RN)
    ET.SubElement(rels_b, f'{{{RN}}}Relationship',
                  {'Id': new_id, 'Type': f'{RN}/image',
                   'Target': 'media/image1.png'})
    rels_bytes = ET.tostring(rels_b, encoding='UTF-8', xml_declaration=True)

    # ---- assemble ----
    body = rb.find(Q('body'))
    sect = body.find(Q('sectPr'))
    new_kids = (base_content + [page_break_para()] + ch4_content +
                [page_break_para()] + ch5_content + [page_break_para()] +
                [b_head, b_fa] + [el for _, el in fa_list] +
                [b_la] + [el for _, el in la_list])
    for _ in list(body):
        body.remove(_)
    for el in new_kids:
        body.append(el)
    body.append(sect)

    # ---- write output ----
    ct_root = ET.fromstring(zb.read('[Content_Types].xml'))
    CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
    ET.register_namespace('', CT)
    assert not any(e.get('Extension') == 'png'
                   for e in ct_root.findall(f'{{{CT}}}Default'))
    ET.SubElement(ct_root, f'{{{CT}}}Default',
                  {'Extension': 'png', 'ContentType': 'image/png'})
    ct_bytes = ET.tostring(ct_root, encoding='UTF-8', xml_declaration=True)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zb.namelist():
            if name in ('word/document.xml', 'word/_rels/document.xml.rels',
                        '[Content_Types].xml'):
                continue
            zout.writestr(name, zb.read(name))
        zout.writestr('word/document.xml',
                      ET.tostring(rb, encoding='UTF-8', xml_declaration=True))
        zout.writestr('word/_rels/document.xml.rels', rels_bytes)
        zout.writestr('[Content_Types].xml', ct_bytes)
        zout.writestr('word/media/image1.png', img_bytes)

    print(f'ch4 runs font-pinned: {pinned}')
    print(f'refs: FA={len(fa_list)} Latin={len(la_list)} dupes_removed={dupes}')
    print('wrote', OUT)


main()
