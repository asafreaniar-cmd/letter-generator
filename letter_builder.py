"""
letter_builder.py
-----------------
מנוע יצירת מכתבים על גבי תבנית הבלנק הקיים (הכנסת.docx).
שומר על כל עיצוב המקור: שולי עמוד, כותרת עם לוגו, גופן David 12pt, RTL.
"""

import os
import copy
import re
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from template_data import ensure_template, ensure_signature

def parse_formatting(text):
    if not text:
        return []
    pattern = re.compile(
        r'(?P<bold_under>(\*\*__.*?__\*\*)|(__\*\*.*?\*\*__))|(?P<bold>\*\*.*?\*\*)|(?P<under>__.*?__)'
    )
    runs = []
    last_idx = 0
    for match in pattern.finditer(text):
        start, end = match.span()
        if start > last_idx:
            runs.append({'text': text[last_idx:start], 'bold': False, 'underline': False})
        
        matched_text = match.group()
        if match.group('bold_under'):
            runs.append({'text': matched_text[4:-4], 'bold': True, 'underline': True})
        elif match.group('bold'):
            runs.append({'text': matched_text[2:-2], 'bold': True, 'underline': False})
        elif match.group('under'):
            runs.append({'text': matched_text[2:-2], 'bold': False, 'underline': True})
        last_idx = end
    if last_idx < len(text):
        runs.append({'text': text[last_idx:], 'bold': False, 'underline': False})
    return runs

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'הכנסת.docx')
SIGNATURE_PATH = os.path.join(os.path.dirname(__file__), 'חתימה.png')
ensure_template(TEMPLATE_PATH)
ensure_signature(SIGNATURE_PATH)
FONT_NAME = 'David'
FONT_SIZE_HTP = '24'   # half-points (= 12pt)
NUM_SPACERS = 5        # paragraphs to keep from the original for logo spacing
SIGNATURE_WIDTH = 1374609
SIGNATURE_HEIGHT = 932742
SIGNATURE_OFFSET_X = 695462
SIGNATURE_OFFSET_Y = 121610

# ──────────────────────────────────────────────────────────────────
#  XML helper utilities
# ──────────────────────────────────────────────────────────────────

def _xml_space_preserve(t_el, text):
    """Mark <w:t> element with xml:space='preserve' when needed."""
    if text and (text[0] == ' ' or text[-1] == ' '):
        t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def _make_rPr(bold=False, underline=False, rtl=True, hint_cs=True):
    """Build a <w:rPr> element with David font + optional bold/underline/RTL."""
    rPr = OxmlElement('w:rPr')

    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), FONT_NAME)
    rFonts.set(qn('w:hAnsi'), FONT_NAME)
    rFonts.set(qn('w:cs'), FONT_NAME)
    if hint_cs:
        rFonts.set(qn('w:hint'), 'cs')
    rPr.append(rFonts)

    if bold:
        rPr.append(OxmlElement('w:b'))
        rPr.append(OxmlElement('w:bCs'))

    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), FONT_SIZE_HTP)
    rPr.append(sz)

    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), FONT_SIZE_HTP)
    rPr.append(szCs)

    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

    if rtl:
        rPr.append(OxmlElement('w:rtl'))

    return rPr


def _make_pPr(jc=None, line=240, ind_left=None, ind_right=None, ind_first=None,
              bold=False, underline=False, rtl=True):
    """Build a <w:pPr> element."""
    pPr = OxmlElement('w:pPr')

    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), str(line))
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)

    if ind_left is not None or ind_right is not None or ind_first is not None:
        ind = OxmlElement('w:ind')
        if ind_left is not None:
            ind.set(qn('w:left'), str(ind_left))
        if ind_right is not None:
            ind.set(qn('w:right'), str(ind_right))
        if ind_first is not None:
            ind.set(qn('w:firstLine'), str(ind_first))
        pPr.append(ind)

    if jc:
        jc_el = OxmlElement('w:jc')
        jc_el.set(qn('w:val'), jc)
        pPr.append(jc_el)

    pPr.append(_make_rPr(bold=bold, underline=underline, rtl=rtl, hint_cs=False))
    return pPr


def _make_run(text, bold=False, underline=False, rtl=True, hint_cs=True):
    """Build a <w:r> element with one <w:t>."""
    r = OxmlElement('w:r')
    r.append(_make_rPr(bold=bold, underline=underline, rtl=rtl, hint_cs=hint_cs))
    t = OxmlElement('w:t')
    t.text = text
    _xml_space_preserve(t, text)
    r.append(t)
    return r


def _make_p(text='', jc=None, line=240, bold=False, underline=False, rtl=True,
            ind_left=None, ind_right=None, ind_first=None, extra_runs=None):
    """
    Build a complete <w:p> element.
    `extra_runs` is an optional list of (text, bold, underline, rtl) tuples
    added after the main run (useful for mixed formatting in one paragraph).
    """
    p = OxmlElement('w:p')
    p.append(_make_pPr(jc=jc, line=line, bold=bold, underline=underline, rtl=rtl,
                       ind_left=ind_left, ind_right=ind_right, ind_first=ind_first))
    if text:
        runs = parse_formatting(text)
        for run in runs:
            p.append(_make_run(run['text'], bold=bold or run['bold'], underline=underline or run['underline'], rtl=rtl))
    if extra_runs:
        for run_kwargs in extra_runs:
            p.append(_make_run(**run_kwargs))
    return p


def _make_signature_anchor(doc):
    """Build a floating signature drawing anchored behind the closing paragraph."""
    inline = doc.part.new_pic_inline(
        SIGNATURE_PATH,
        width=SIGNATURE_WIDTH,
        height=SIGNATURE_HEIGHT,
    )

    anchor = OxmlElement('wp:anchor')
    anchor.set('distT', '0')
    anchor.set('distB', '0')
    anchor.set('distL', '114300')
    anchor.set('distR', '114300')
    anchor.set('simplePos', '0')
    anchor.set('relativeHeight', '251659264')
    anchor.set('behindDoc', '1')
    anchor.set('locked', '0')
    anchor.set('layoutInCell', '1')
    anchor.set('allowOverlap', '1')

    simple_pos = OxmlElement('wp:simplePos')
    simple_pos.set('x', '0')
    simple_pos.set('y', '0')
    anchor.append(simple_pos)

    position_h = OxmlElement('wp:positionH')
    position_h.set('relativeFrom', 'column')
    position_h_offset = OxmlElement('wp:posOffset')
    position_h_offset.text = str(SIGNATURE_OFFSET_X)
    position_h.append(position_h_offset)
    anchor.append(position_h)

    position_v = OxmlElement('wp:positionV')
    position_v.set('relativeFrom', 'paragraph')
    position_v_offset = OxmlElement('wp:posOffset')
    position_v_offset.text = str(SIGNATURE_OFFSET_Y)
    position_v.append(position_v_offset)
    anchor.append(position_v)

    anchor.append(copy.deepcopy(inline.extent))

    effect_extent = OxmlElement('wp:effectExtent')
    effect_extent.set('l', '0')
    effect_extent.set('t', '0')
    effect_extent.set('r', '0')
    effect_extent.set('b', '0')
    anchor.append(effect_extent)

    anchor.append(OxmlElement('wp:wrapNone'))
    anchor.append(copy.deepcopy(inline.docPr))
    anchor.append(copy.deepcopy(inline.find(qn('wp:cNvGraphicFramePr'))))
    anchor.append(copy.deepcopy(inline.graphic))
    return anchor


def _append_signature_to_paragraph(doc, paragraph):
    """Append a behind-text signature image to the given paragraph."""
    run = OxmlElement('w:r')
    run_rpr = _make_rPr(rtl=True, hint_cs=False)
    run_rpr.append(OxmlElement('w:noProof'))
    run.append(run_rpr)

    drawing = OxmlElement('w:drawing')
    drawing.append(_make_signature_anchor(doc))
    run.append(drawing)
    paragraph.append(run)


def _make_signature_paragraph(doc):
    """Create a standalone paragraph carrying the floating signature anchor."""
    paragraph = _make_p('', jc=None, line=240, rtl=True)
    _append_signature_to_paragraph(doc, paragraph)
    return paragraph


# ──────────────────────────────────────────────────────────────────
#  Table helper (for multiple recipients)
# ──────────────────────────────────────────────────────────────────

def _make_recipients_table(recipients):
    """
    Build a <w:tbl> element for multiple recipients, laid out in a row,
    right-to-left (first recipient on the right, last on the left).
    Each recipient dict: { 'intro': str, 'name': str, 'title': str }
    """
    from docx.oxml import OxmlElement
    tbl = OxmlElement('w:tbl')

    # Table properties
    tblPr = OxmlElement('w:tblPr')
    tblStyle = OxmlElement('w:tblStyle')
    tblStyle.set(qn('w:val'), 'TableGrid')
    tblPr.append(tblStyle)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '0')
    tblW.set(qn('w:type'), 'auto')
    tblPr.append(tblW)
    # No borders
    tblBorders = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'none')
        tblBorders.append(b)
    tblPr.append(tblBorders)
    # RTL table
    bidiVisual = OxmlElement('w:bidiVisual')
    tblPr.append(bidiVisual)
    tbl.append(tblPr)

    # Grid columns – one per recipient (fixed width of 2160 dxa)
    tblGrid = OxmlElement('w:tblGrid')
    col_width = '2160'
    for _ in recipients:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), col_width)
        tblGrid.append(gridCol)
    tbl.append(tblGrid)

    # Single row of recipients (intro + name + title stacked per cell)
    tr = OxmlElement('w:tr')
    for rec in recipients:
        tc = OxmlElement('w:tc')
        tcPr = OxmlElement('w:tcPr')
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), col_width)
        tcW.set(qn('w:type'), 'dxa')
        tcPr.append(tcW)
        tc.append(tcPr)

        intro = rec.get('intro', 'לכבוד').strip()
        name  = rec.get('name', '').strip()
        title = rec.get('title', '').strip()

        # Recipient intro paragraph
        if name:
            tc.append(_make_p(intro, jc='both', line=240, rtl=True))
            tc.append(_make_p(name, jc='both', line=240, rtl=True))
            if title:
                tc.append(_make_p(title, jc='both', line=240, underline=True, rtl=True))
        else:
            tc.append(_make_p(intro, jc='both', line=240, underline=True, rtl=True))
        
        tr.append(tc)
    tbl.append(tr)
    return tbl


# ──────────────────────────────────────────────────────────────────
#  Table helper (for multiple signers)
# ──────────────────────────────────────────────────────────────────

def _make_signers_table(signers):
    """
    Build a <w:tbl> element for multiple signers, laid out in a row,
    right-to-left (first signer on the right, last on the left).
    Each signer dict: { 'name': str, 'title': str }
    """
    from docx.oxml import OxmlElement
    tbl = OxmlElement('w:tbl')

    # Table properties
    tblPr = OxmlElement('w:tblPr')
    tblStyle = OxmlElement('w:tblStyle')
    tblStyle.set(qn('w:val'), 'TableGrid')
    tblPr.append(tblStyle)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '0')
    tblW.set(qn('w:type'), 'auto')
    tblPr.append(tblW)
    # No borders
    tblBorders = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'none')
        tblBorders.append(b)
    tblPr.append(tblBorders)
    # RTL table
    bidiVisual = OxmlElement('w:bidiVisual')
    tblPr.append(bidiVisual)
    tbl.append(tblPr)

    # Grid columns – one per signer
    tblGrid = OxmlElement('w:tblGrid')
    col_width = str(9000 // len(signers))  # divide ~15cm evenly
    for _ in signers:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), col_width)
        tblGrid.append(gridCol)
    tbl.append(tblGrid)

    # Single row of signers (name + title stacked per cell)
    tr = OxmlElement('w:tr')
    for signer in signers:
        tc = OxmlElement('w:tc')
        tcPr = OxmlElement('w:tcPr')
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), col_width)
        tcW.set(qn('w:type'), 'dxa')
        tcPr.append(tcW)
        tc.append(tcPr)

        # Signer name paragraph
        tc.append(_make_p(signer.get('name', ''), jc='center', line=240, rtl=True))
        # Signer title paragraph
        tc.append(_make_p(signer.get('title', ''), jc='center', line=240, rtl=True))
        tr.append(tc)
    tbl.append(tr)
    return tbl


# ──────────────────────────────────────────────────────────────────
#  Main builder
# ──────────────────────────────────────────────────────────────────

def build_letter(data: dict, output_path: str) -> str:
    """
    Generate a Word letter from `data` and save it to `output_path`.

    Parameters
    ----------
    data : dict
        Keys (all optional except noted):
        - reference        : str  – סימוכין (optional)
        - city             : str  – עיר/מקום (default 'הכנסת, ירושלים')
        - date_hebrew      : str  – תאריך עברי
        - date_gregorian   : str  – תאריך לועזי
        - recipient_intro  : str  – "לכבוד" / "לכל מאן דבעי" etc.
        - recipient_name   : str  – שם הנמען (optional)
        - recipient_title  : str  – תפקיד הנמען (optional)
        - greeting         : str  – פתיח (default 'שלום רב,')
        - subject          : str  – נדון
        - body             : str  – גוף המכתב (שורות חדשות מפרידות פסקאות)
        - closing          : str  – שורת סיום (default 'בכבוד רב,')
        - signers          : list – [{'name': str, 'title': str}, ...]
        - cc               : list – ['...', '...'] – רשימת העתקים
        - note             : str  – הערה פנימית/קישור (optional)

    Returns
    -------
    str : absolute path to the generated .docx file
    """
    doc = Document(TEMPLATE_PATH)
    body = doc.element.body

    # ── 1. Collect first NUM_SPACERS empty paragraphs (logo spacing) ──
    spacers = []
    para_count = 0
    for child in list(body):
        if child.tag == qn('w:p') and para_count < NUM_SPACERS:
            spacers.append(copy.deepcopy(child))
            para_count += 1

    # ── 2. Clear body, keep only sectPr ──
    sectPr = body.find(qn('w:sectPr'))
    sectPr_copy = copy.deepcopy(sectPr) if sectPr is not None else None
    for child in list(body):
        body.remove(child)

    # ── 3. Re-add spacers ──
    for spacer in spacers:
        body.append(spacer)

    def add(p_elem):
        """Insert paragraph/table element before sectPr."""
        if sectPr_copy is not None and sectPr_copy in body:
            idx = list(body).index(sectPr_copy)
            body.insert(idx, p_elem)
        else:
            body.append(p_elem)

    # ── 4. Re-add sectPr so header/footer/margins stay intact ──
    if sectPr_copy is not None:
        body.append(sectPr_copy)

    # ── 5. Build letter content ──

    # Apply vertical body spacing before drawing any content (pushed down from the header logo)
    spacing_val = data.get('body_spacing', 'auto')
    lines_to_add = 0
    if spacing_val == 'auto':
        body_text = data.get('body', '').strip()
        if body_text:
            body_len = len(body_text)
            if body_len < 150:
                lines_to_add = 5
            elif body_len < 300:
                lines_to_add = 4
            elif body_len < 500:
                lines_to_add = 3
            elif body_len < 800:
                lines_to_add = 2
            elif body_len < 1200:
                lines_to_add = 1
    else:
        try:
            lines_to_add = int(spacing_val)
        except ValueError:
            lines_to_add = 0
            
    for _ in range(lines_to_add):
        add(_make_p('', jc='both', line=240, rtl=True))

    # ── 5a. Optional: reference number (סימוכין) ──
    if data.get('reference', '').strip():
        add(_make_p(f"סימוכין: {data['reference'].strip()}",
                    jc='right', line=240, rtl=True))

    # ── 5b. City + Hebrew date ──
    city = data.get('city', 'הכנסת, ירושלים').strip()
    date_heb = data.get('date_hebrew', '').strip()
    city_line = f"{city}, {date_heb}" if date_heb else city
    add(_make_p(city_line, jc='right', line=240, rtl=True))

    # ── 5c. Gregorian date ──
    date_greg = data.get('date_gregorian', '').strip()
    if date_greg:
        add(_make_p(date_greg, jc='right', line=240, rtl=True))

    # ── 5e. Recipient block ──
    recipients = data.get('recipients', [])
    if not recipients:
        recipient_intro = data.get('recipient_intro', 'לכבוד').strip()
        recipient_name  = data.get('recipient_name', '').strip()
        recipient_title = data.get('recipient_title', '').strip()
        recipients = [{'intro': recipient_intro, 'name': recipient_name, 'title': recipient_title}]

    if len(recipients) == 1:
        rec = recipients[0]
        intro = rec.get('intro', 'לכבוד').strip()
        name  = rec.get('name', '').strip()
        title = rec.get('title', '').strip()
        if name:
            add(_make_p(intro, jc='both', line=240, rtl=True))
            add(_make_p(name, jc='both', line=240, rtl=True))
            if title:
                add(_make_p(title, jc='both', line=240, underline=True, rtl=True))
        else:
            add(_make_p(intro, jc='both', line=240, underline=True, rtl=True))
    else:
        add(_make_recipients_table(recipients))

    # ── 5f. Empty line ──
    add(_make_p('', jc='both', line=240, rtl=True))

    # ── 5g. Greeting ──
    greeting = data.get('greeting', 'שלום רב,').strip()
    add(_make_p(greeting, jc='both', line=240, rtl=True))

    # ── 5h. Empty line ──
    add(_make_p('', jc='both', line=240, rtl=True))

    # ── 5i. Subject line: "הנדון:" bold, subject text bold+underline ──
    subject = data.get('subject', '').strip()
    subject_runs = [
        {'text': 'הנדון:', 'bold': True, 'underline': False, 'rtl': True, 'hint_cs': False},
        {'text': ' ', 'bold': True, 'underline': False, 'rtl': True, 'hint_cs': False},
        {'text': subject, 'bold': True, 'underline': True, 'rtl': True, 'hint_cs': True},
    ]
    add(_make_p('', jc='center', line=360, bold=True, rtl=True, extra_runs=subject_runs))

    # ── 5j. Body paragraphs ──
    body_text = data.get('body', '').strip()
    if body_text:
        paragraphs = body_text.split('\n')
        for para_text in paragraphs:
            add(_make_p(para_text.strip(), jc='both', line=360, rtl=True))
    else:
        add(_make_p('', jc='both', line=360, rtl=True))

    # ── 5k. Empty line ──
    add(_make_p('', jc='both', line=240, rtl=True))

    # ── 5l. Signature + closing ──
    # Match the example letter: floating signature anchored in its own paragraph,
    # then "בכבוד רב," and signer lines remain untouched.
    closing = data.get('closing', 'בכבוד רב,').strip()
    if os.path.exists(SIGNATURE_PATH):
        try:
            add(_make_signature_paragraph(doc))
        except Exception:
            pass
    closing_para = _make_p(closing, jc=None, line=240, ind_left=5760, rtl=True)
    add(closing_para)

    # ── 5m. Signers ──
    # ind_right=1410 matches example letter (name positioned near right with inset)
    signers = data.get('signers', [])
    if not signers:
        signers = [{'name': '', 'title': ''}]

    if len(signers) == 1:
        s = signers[0]
        add(_make_p(s.get('name', ''),  jc='right', line=240, ind_right=1410, rtl=True))
        if s.get('title', '').strip():
            add(_make_p(s['title'].strip(), jc='right', line=240, ind_right=1410, rtl=True))
    else:
        # Multiple signers → table
        add(_make_signers_table(signers))

    # ── 5n. Empty lines before CC ──
    add(_make_p('', jc='both', line=240, rtl=True))
    add(_make_p('', jc='both', line=240, rtl=True))

    # ── 5o. CC list (העתק) ──
    cc_list = [c for c in data.get('cc', []) if c.strip()]
    if cc_list:
        first = True
        for cc_item in cc_list:
            if first:
                add(_make_p(f"העתק: {cc_item.strip()}", jc='both', line=240, rtl=True))
                first = False
            else:
                # Each subsequent CC item on its own line, indented to align under first item's text
                add(_make_p(f"            {cc_item.strip()}", jc='both', line=240, rtl=True))

    # ── 5p. Internal note ──
    note = data.get('note', '').strip()
    if note:
        add(_make_p('', jc='both', line=240, rtl=True))
        add(_make_p(f"[{note}]", jc='both', line=240, rtl=True))

    # ── 6. Save ──
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path
