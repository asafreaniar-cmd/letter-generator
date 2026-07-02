"""
pdf_builder.py
--------------
Python-native PDF generator: David 12pt font + Knesset letterhead background.
Matches the Word output exactly — same font, size, bold, underline, RTL layout.
No dependency on Word, LibreOffice or Gotenberg.
"""

import os
import zipfile
import re
from io import BytesIO

from bidi.algorithm import get_display
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas

BASE_DIR = os.path.dirname(__file__)
FONTS_DIR = os.path.join(BASE_DIR, 'fonts')
TEMPLATE_PATH = os.path.join(BASE_DIR, 'הכנסת.docx')
SIGNATURE_PATH = os.path.join(BASE_DIR, 'חתימה.png')

DAVID = 'David'
DAVID_BOLD = 'DavidBold'

# ── Page geometry (A4, matching Word template) ──────────────────────────────
PAGE_W, PAGE_H = A4           # 595.28 × 841.89 pt
MARGIN_RIGHT = 90.0           # twips 1800 → 90pt
MARGIN_LEFT  = 90.0
CONTENT_RIGHT = PAGE_W - MARGIN_RIGHT    # ~505.3 pt from left = right edge of text
CONTENT_LEFT  = MARGIN_LEFT              # ~90 pt
CONTENT_WIDTH = CONTENT_RIGHT - CONTENT_LEFT  # ~415 pt

FONT_SIZE = 12.0

# ── Line spacing (calibrated from Word-generated PDFs printed via macOS) ────
# BODY_START_Y: baseline of the first text line (city/date),
# measured from the BOTTOM of the page in ReportLab coordinates.
# Word places city at 181.2pt from top → 841.9 - 181.2 = 660.7pt from bottom.
BODY_START_Y      = 660.7   # pt from bottom
# LINE_SINGLE: baseline gap between consecutive line=240 paragraphs.
# Word: ~19.7pt (line height ≈11.7pt + template's default spaceAfter ≈8pt).
LINE_SINGLE       = 19.7
# LINE_WITHIN_PARA: baseline gap between lines inside a line=360 body paragraph.
# Word: ~17.8pt (1.5× line height; no spaceAfter between lines within a paragraph).
LINE_WITHIN_PARA  = 17.8
# LINE_BETWEEN_PARA: baseline gap between separate line=360 paragraphs.
# Word: ~25.6pt (17.8 line height + 8pt spaceAfter from template Normal style).
LINE_BETWEEN_PARA = 25.6

# ── Signature geometry ───────────────────────────────────────────────────────
# Dimensions match the EMU values in letter_builder.py → same size as Word renders
SIG_W = 108.2   # pt  = 1374609 EMU / 914400 * 72
SIG_H = 73.4    # pt  = 932742  EMU / 914400 * 72
# Horizontal offset from CONTENT_RIGHT (RTL column, measured from right)
SIG_OFFSET_X = 54.75   # pt  = 695462 EMU / 914400 * 72

# ── Closing and signer positions (matching Word's RTL indent rules) ──────────
# In Word RTL paragraphs, ind_left shifts the physical-RIGHT boundary inward.
# Closing: ind_left=5760 twips=288pt → right edge at CONTENT_RIGHT - 288 = 217.3pt
CLOSING_RIGHT = CONTENT_RIGHT - (5760.0 / 1440.0 * 72)   # ≈ 217.3 pt

# Signer (single): jc='right' RTL = physical-LEFT aligned; ind_right=1410 twips=70.5pt
# shifts the physical-LEFT boundary rightward: CONTENT_LEFT + 70.5 = 160.5pt
SIGNER_LEFT = CONTENT_LEFT + (1410.0 / 1440.0 * 72)   # ≈ 160.5 pt


# ═══════════════════════════════════════════════════════ font registration ═══

_fonts_registered = False

def _register_fonts():
    global _fonts_registered
    if _fonts_registered:
        return
    regular = os.path.join(FONTS_DIR, 'david.ttf')
    bold    = os.path.join(FONTS_DIR, 'davidbd.ttf')
    if not os.path.exists(regular):
        raise FileNotFoundError(f'David font not found at {regular}')
    pdfmetrics.registerFont(TTFont(DAVID, regular))
    pdfmetrics.registerFont(TTFont(DAVID_BOLD, bold if os.path.exists(bold) else regular))
    _fonts_registered = True


# ═══════════════════════════════════════════════════════ image helpers ═══════

def _letterhead_image() -> BytesIO:
    """Extract full-page letterhead JPEG from the DOCX template."""
    with zipfile.ZipFile(TEMPLATE_PATH) as z:
        return BytesIO(z.read('word/media/image2.jpeg'))


def _signature_image():
    if not os.path.exists(SIGNATURE_PATH):
        return None
    with open(SIGNATURE_PATH, 'rb') as f:
        return BytesIO(f.read())


# ═══════════════════════════════════════════════════════ text helpers ════════

def _font_name(bold: bool) -> str:
    return DAVID_BOLD if bold else DAVID


def _rtl(text: str) -> str:
    """Convert Hebrew logical order → visual order for left-to-right rendering."""
    if not text or not text.strip():
        return text
    return get_display(text)


def _str_width(text: str, bold: bool = False, size: float = FONT_SIZE) -> float:
    return pdfmetrics.stringWidth(text, _font_name(bold), size)


def _wrap(text, bold=False, max_w=CONTENT_WIDTH):
    """
    Wrap Hebrew text (logical order) into lines that fit max_w.
    Splits on whitespace; each line is a logical-order substring.
    """
    if not text.strip():
        return ['']
    words = text.split()
    lines: list[str] = []
    cur: list[str] = []
    cur_w = 0.0
    fn = _font_name(bold)
    sp_w = pdfmetrics.stringWidth(' ', fn, FONT_SIZE)

    for word in words:
        w_w = pdfmetrics.stringWidth(_rtl(word), fn, FONT_SIZE)
        needed = w_w + (sp_w + cur_w if cur else 0.0)
        if needed > max_w and cur:
            lines.append(' '.join(cur))
            cur = [word]
            cur_w = w_w
        else:
            cur.append(word)
            cur_w = needed
    if cur:
        lines.append(' '.join(cur))
    return lines


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


def _wrap_formatted(text, max_w=CONTENT_WIDTH, bold_default=False, underline_default=False, size=FONT_SIZE):
    if not text.strip():
        return [[{'text': '', 'bold': bold_default, 'underline': underline_default}]]
    
    # 1. Parse formatting into logical runs
    runs = parse_formatting(text)
    
    # 2. Tokenize runs into words
    words = []
    for run in runs:
        run_text = run['text']
        parts = run_text.split(' ')
        for idx, part in enumerate(parts):
            if part:
                words.append({
                    'text': part,
                    'bold': bold_default or run['bold'],
                    'underline': underline_default or run['underline'],
                    'has_space_after': (idx < len(parts) - 1)
                })
            elif idx < len(parts) - 1:
                if words:
                    words[-1]['has_space_after'] = True
                    
    # 3. Wrap words into lines
    lines = []
    current_line = []
    current_w = 0.0
    
    def word_width(w_dict):
        fn = _font_name(w_dict['bold'])
        return pdfmetrics.stringWidth(_rtl(w_dict['text']), fn, size)
        
    for w in words:
        w_width = word_width(w)
        fn = _font_name(w['bold'])
        space_w = pdfmetrics.stringWidth(' ', fn, size) if w['has_space_after'] else 0.0
        
        needed = w_width + space_w
        if current_line and (current_w + needed > max_w):
            lines.append(current_line)
            current_line = []
            current_w = 0.0
            
        current_line.append(w)
        current_w += needed
        
    if current_line:
        lines.append(current_line)
        
    # 4. Merge words on each line into runs
    merged_lines = []
    for line_words in lines:
        line_runs = []
        for w in line_words:
            text_val = w['text'] + (' ' if w['has_space_after'] else '')
            if line_runs and line_runs[-1]['bold'] == w['bold'] and line_runs[-1]['underline'] == w['underline']:
                line_runs[-1]['text'] += text_val
            else:
                line_runs.append({
                    'text': text_val,
                    'bold': w['bold'],
                    'underline': w['underline']
                })
        merged_lines.append(line_runs)
        
    return merged_lines


def _draw_formatted_line_right(c, line_runs, y, right_x=CONTENT_RIGHT, size=FONT_SIZE):
    total_w = 0.0
    for run in line_runs:
        fn = _font_name(run['bold'])
        total_w += pdfmetrics.stringWidth(_rtl(run['text']), fn, size)
        
    x = right_x - total_w
    for run in reversed(line_runs):
        fn = _font_name(run['bold'])
        visual = _rtl(run['text'])
        w_width = pdfmetrics.stringWidth(visual, fn, size)
        
        c.setFont(fn, size)
        c.drawString(x, y, visual)
        if run['underline']:
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0, 0, 0)
            c.line(x, y - 1.5, x + w_width, y - 1.5)
        x += w_width


def _draw_formatted_line_left(c, line_runs, y, left_x=CONTENT_LEFT, size=FONT_SIZE):
    x = left_x
    for run in reversed(line_runs):
        fn = _font_name(run['bold'])
        visual = _rtl(run['text'])
        w_width = pdfmetrics.stringWidth(visual, fn, size)
        
        c.setFont(fn, size)
        c.drawString(x, y, visual)
        if run['underline']:
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0, 0, 0)
            c.line(x, y - 1.5, x + w_width, y - 1.5)
        x += w_width


def _draw_formatted_line_center(c, line_runs, y, size=FONT_SIZE):
    total_w = 0.0
    for run in line_runs:
        fn = _font_name(run['bold'])
        total_w += pdfmetrics.stringWidth(_rtl(run['text']), fn, size)
        
    x = (CONTENT_LEFT + CONTENT_RIGHT) / 2.0 - total_w / 2.0
    for run in reversed(line_runs):
        fn = _font_name(run['bold'])
        visual = _rtl(run['text'])
        w_width = pdfmetrics.stringWidth(visual, fn, size)
        
        c.setFont(fn, size)
        c.drawString(x, y, visual)
        if run['underline']:
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0, 0, 0)
            c.line(x, y - 1.5, x + w_width, y - 1.5)
        x += w_width


def _draw_formatted_line_justified(c, line_runs, y, size=FONT_SIZE):
    # Split runs back into individual words for justification
    line_words = []
    for run in line_runs:
        parts = run['text'].strip().split(' ')
        parts = [p for p in parts if p]
        for p in parts:
            line_words.append({
                'text': p,
                'bold': run['bold'],
                'underline': run['underline']
            })
            
    if len(line_words) <= 1:
        _draw_formatted_line_right(c, line_runs, y, right_x=CONTENT_RIGHT, size=size)
        return
    
    total_words_w = 0.0
    for w in line_words:
        fn = _font_name(w['bold'])
        total_words_w += pdfmetrics.stringWidth(_rtl(w['text']), fn, size)
        
    gap_w = (CONTENT_WIDTH - total_words_w) / (len(line_words) - 1)
    
    x = CONTENT_LEFT
    for idx, w in enumerate(reversed(line_words)):
        fn = _font_name(w['bold'])
        visual = _rtl(w['text'])
        w_width = pdfmetrics.stringWidth(visual, fn, size)
        
        c.setFont(fn, size)
        c.drawString(x, y, visual)
        if idx < len(line_words) - 1:
            c.drawString(x + w_width, y, " ")
        if w['underline']:
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0, 0, 0)
            c.line(x, y - 1.5, x + w_width, y - 1.5)
        x += w_width + gap_w


# ═══════════════════════════════════════════════════════ draw primitives ═════

def _draw_right(c: rl_canvas.Canvas, text: str, y: float, *,
                bold: bool = False, underline: bool = False,
                size: float = FONT_SIZE, right_x: float = CONTENT_RIGHT):
    if not text.strip():
        return
    lines = _wrap_formatted(text, max_w=CONTENT_WIDTH, bold_default=bold, underline_default=underline, size=size)
    if lines:
        _draw_formatted_line_right(c, lines[0], y, right_x=right_x, size=size)


def _draw_left(c: rl_canvas.Canvas, text: str, y: float, *,
               bold: bool = False, underline: bool = False,
               size: float = FONT_SIZE, left_x: float = CONTENT_LEFT):
    if not text.strip():
        return
    lines = _wrap_formatted(text, max_w=CONTENT_WIDTH, bold_default=bold, underline_default=underline, size=size)
    if lines:
        _draw_formatted_line_left(c, lines[0], y, left_x=left_x, size=size)


def _draw_center(c: rl_canvas.Canvas, text: str, y: float, *,
                 bold: bool = False, underline: bool = False,
                 size: float = FONT_SIZE):
    if not text.strip():
        return
    lines = _wrap_formatted(text, max_w=CONTENT_WIDTH, bold_default=bold, underline_default=underline, size=size)
    if lines:
        _draw_formatted_line_center(c, lines[0], y, size=size)


def _draw_para_right(c: rl_canvas.Canvas, text: str, y: float, *,
                     bold: bool = False, underline: bool = False) -> float:
    """
    Draw a potentially multi-line right-aligned paragraph supporting inline formatting.
    Returns the baseline Y of the LAST line drawn (unchanged if empty).
    """
    if not text.strip():
        return y
    lines = _wrap_formatted(text, bold_default=bold, underline_default=underline)
    for i, line in enumerate(lines):
        _draw_formatted_line_right(c, line, y)
        if i < len(lines) - 1:
            y -= LINE_WITHIN_PARA
    return y


def _draw_para_justified(c: rl_canvas.Canvas, text: str, y: float, *,
                          bold: bool = False) -> float:
    """
    Draw a Hebrew RTL paragraph with full justification supporting inline formatting.
    Returns baseline Y of the last line drawn.
    """
    if not text.strip():
        return y
    lines = _wrap_formatted(text, bold_default=bold)
    for i, line in enumerate(lines):
        is_last = (i == len(lines) - 1)
        if is_last:
            _draw_formatted_line_right(c, line, y)
        else:
            _draw_formatted_line_justified(c, line, y)
            
        if i < len(lines) - 1:
            y -= LINE_WITHIN_PARA
    return y


# Subject line: "הנדון: " bold (no underline) + subject bold+underline, centred
def _draw_subject(c: rl_canvas.Canvas, subject: str, y: float) -> float:
    """
    Renders the subject block centred, bold.
    "הנדון: " has no underline; the subject text is underlined.
    Handles line-wrapping.
    Returns baseline Y of last line drawn.
    """
    prefix = 'הנדון: '
    full   = prefix + subject
    lines  = _wrap(full, bold=True, max_w=CONTENT_WIDTH)

    fn  = _font_name(True)
    sp  = (CONTENT_LEFT + CONTENT_RIGHT) / 2.0

    # Pre-compute width of prefix in visual order
    prefix_visual = _rtl(prefix)
    prefix_w = pdfmetrics.stringWidth(prefix_visual, fn, FONT_SIZE)

    for idx, line in enumerate(lines):
        visual = _rtl(line)
        total_w = pdfmetrics.stringWidth(visual, fn, FONT_SIZE)
        x_start = sp - total_w / 2.0
        x_end   = x_start + total_w

        c.setFont(fn, FONT_SIZE)
        c.drawString(x_start, y, visual)

        # Underline everything, then erase underline under the prefix
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0, 0, 0)
        c.line(x_start, y - 1.5, x_end, y - 1.5)

        # The prefix "הנדון: " is RTL so in visual order it sits at the RIGHT end
        if idx == 0:
            # Erase prefix underline: prefix is at the right end of the visual line
            c.setStrokeColorRGB(1, 1, 1)
            c.line(x_end - prefix_w, y - 1.5, x_end, y - 1.5)
            c.setStrokeColorRGB(0, 0, 0)

        if idx < len(lines) - 1:
            y -= LINE_WITHIN_PARA

    return y


# ═══════════════════════════════════════════════════════ main builder ════════

def build_letter_pdf(data: dict, output_path: str) -> str:
    """
    Generate a PDF letter with David 12pt font matching the Word output.

    Uses the Knesset letterhead (image2.jpeg from הכנסת.docx) as the page
    background, then draws all body text with the David font at correct
    positions so the result is visually identical to the Word export.

    Parameters: same `data` dict as build_letter().
    Returns: absolute path to the generated PDF.
    """
    _register_fonts()
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    c = rl_canvas.Canvas(output_path, pagesize=A4)

    # ── Background: full-page letterhead ────────────────────────────────────
    lh = ImageReader(_letterhead_image())
    c.drawImage(lh, 0, 0, width=PAGE_W, height=PAGE_H, preserveAspectRatio=False)

    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)

    y = BODY_START_Y  # current baseline (moves DOWN = y decreases)

    def step(n: float = LINE_SINGLE):
        nonlocal y
        y -= n

    # ── 5a. Siman (optional) ─────────────────────────────────────────────────
    # jc='right' for RTL = physical LEFT-aligned in Word
    reference = (data.get('reference') or '').strip()
    if reference:
        _draw_left(c, f'סימוכין: {reference}', y)
        step()

    # ── 5b. City + Hebrew date ───────────────────────────────────────────────
    # jc='right' for RTL = physical LEFT-aligned in Word
    city     = (data.get('city') or 'הכנסת, ירושלים').strip()
    date_heb = (data.get('date_hebrew') or '').strip()
    city_line = f'{city}, {date_heb}' if date_heb else city
    _draw_left(c, city_line, y)
    step()

    # ── 5c. Gregorian date ───────────────────────────────────────────────────
    # jc='right' for RTL = physical LEFT-aligned in Word
    date_greg = (data.get('date_gregorian') or '').strip()
    if date_greg:
        _draw_left(c, date_greg, y)
        step()

    # ── 5e. Recipient block ──────────────────────────────────────────────────
    recipients = data.get('recipients', [])
    if not recipients:
        r_intro = (data.get('recipient_intro') or 'לכבוד').strip()
        r_name  = (data.get('recipient_name')  or '').strip()
        r_title = (data.get('recipient_title') or '').strip()
        recipients = [{'intro': r_intro, 'name': r_name, 'title': r_title}]

    if len(recipients) == 1:
        rec = recipients[0]
        intro = rec.get('intro', 'לכבוד').strip()
        name  = rec.get('name', '').strip()
        title = rec.get('title', '').strip()
        if name:
            _draw_right(c, intro, y)
            step()
            _draw_right(c, name, y)
            step()
            if title:
                _draw_right(c, title, y, underline=True)
                step()
        else:
            _draw_right(c, intro, y, underline=True)
            step()
    else:
        # Multiple recipients
        col_pitch = 108.0
        max_lines = 0
        for rec in recipients:
            lines = 2
            if rec.get('title', '').strip():
                lines = 3
            if lines > max_lines:
                max_lines = lines

        for ri, rec in enumerate(recipients):
            col_right = CONTENT_RIGHT - ri * col_pitch
            intro = rec.get('intro', 'לכבוד').strip()
            name  = rec.get('name', '').strip()
            title = rec.get('title', '').strip()

            cur_y = y
            if name:
                _draw_right(c, intro, cur_y, right_x=col_right)
                cur_y -= LINE_SINGLE
                _draw_right(c, name, cur_y, right_x=col_right)
                if title:
                    cur_y -= LINE_SINGLE
                    _draw_right(c, title, cur_y, underline=True, right_x=col_right)
            else:
                _draw_right(c, intro, cur_y, underline=True, right_x=col_right)

        for _ in range(max_lines):
            step()

    # ── 5f. Empty line ───────────────────────────────────────────────────────
    step()

    # ── 5g. Greeting ─────────────────────────────────────────────────────────
    greeting = (data.get('greeting') or 'שלום רב,').strip()
    _draw_right(c, greeting, y)
    step()

    # ── 5h. Empty line ───────────────────────────────────────────────────────
    step()

    # ── 5i. Subject line ─────────────────────────────────────────────────────
    subject = (data.get('subject') or '').strip()
    if subject:
        y = _draw_subject(c, subject, y)
        step(LINE_BETWEEN_PARA)

    # ── 5j. Body paragraphs ──────────────────────────────────────────────────
    body_text = (data.get('body') or '').strip()
    if body_text:
        paras = body_text.split('\n')
        for pi, para in enumerate(paras):
            para = para.strip()
            if not para:
                step(LINE_SINGLE)
                continue
            y = _draw_para_justified(c, para, y)
            # Word adds spaceAfter (≈8pt) after every body paragraph including the last,
            # so the gap is always LINE_BETWEEN_PARA regardless of position.
            step(LINE_BETWEEN_PARA)
    else:
        step(LINE_SINGLE)

    # ── 5k. Extra blank before closing (P17 in DOCX) ─────────────────────────
    step()

    # ── 5l. Signature image (floating, at P18 paragraph position) ────────────
    # OOXML anchor: positionH relativeFrom="column" posOffset=695462 EMU = 54.75pt
    # from the LEFT edge of the column → sig left at CONTENT_LEFT + 54.75pt.
    # positionV relativeFrom="paragraph" posOffset=121610 EMU = 9.57pt below baseline.
    _SIG_OFFSET_Y = 121610.0 / 914400.0 * 72   # ≈ 9.57pt below P18 baseline
    sig_bytes = _signature_image()
    if sig_bytes:
        sig_img = ImageReader(sig_bytes)
        sig_x = CONTENT_LEFT + SIG_OFFSET_X
        # Shift up by 8.39pt (~2.96mm) to match Word rendering exactly
        sig_y = y - _SIG_OFFSET_Y - SIG_H + 8.39
        c.drawImage(sig_img, sig_x, sig_y, width=SIG_W, height=SIG_H, mask='auto')

    # Step from P18 (signature para) to P19 (closing) — this step was missing
    step()

    # ── 5l. Closing text (P19 in DOCX) ───────────────────────────────────────
    # In Word RTL, ind_left=5760 twips shifts the right boundary inward by 288pt
    # so the closing is right-aligned to CLOSING_RIGHT ≈ 217pt (physical left area).
    closing = (data.get('closing') or 'בכבוד רב,').strip()
    _draw_right(c, closing, y, right_x=CLOSING_RIGHT)
    step()

    # ── 5m. Signers ──────────────────────────────────────────────────────────
    # Single signer: jc='right' RTL = physical LEFT-aligned; ind_right=1410 twips
    # shifts the left boundary to SIGNER_LEFT ≈ 160.5pt.
    signers = data.get('signers') or [{'name': '', 'title': ''}]
    if len(signers) == 1:
        s = signers[0]
        name  = (s.get('name')  or '').strip()
        title = (s.get('title') or '').strip()
        if name:
            _draw_left(c, name, y, left_x=SIGNER_LEFT)
            step()
        if title:
            _draw_left(c, title, y, left_x=SIGNER_LEFT)
            step()
    else:
        # Multiple signers: evenly spaced columns centered in content area (table in DOCX)
        col_w = CONTENT_WIDTH / len(signers)
        fn = _font_name(False)
        for si, s in enumerate(signers):
            col_center = CONTENT_LEFT + (si + 0.5) * col_w
            name  = (s.get('name')  or '').strip()
            title = (s.get('title') or '').strip()
            if name:
                visual = _rtl(name)
                w = pdfmetrics.stringWidth(visual, fn, FONT_SIZE)
                c.setFont(fn, FONT_SIZE)
                c.drawString(col_center - w / 2, y, visual)
            if title:
                visual = _rtl(title)
                w = pdfmetrics.stringWidth(visual, fn, FONT_SIZE)
                c.setFont(fn, FONT_SIZE)
                c.drawString(col_center - w / 2, y - LINE_SINGLE, visual)
        step()
        step()

    # ── 5n–5o. CC list ───────────────────────────────────────────────────────
    step()
    cc_list = [item for item in (data.get('cc') or []) if item.strip()]
    if cc_list:
        step()
        # Word has exactly one space after colon
        prefix = 'העתק: '
        w_prefix = _str_width(prefix, bold=False, size=FONT_SIZE)
        for ci, item in enumerate(cc_list):
            if ci == 0:
                # Draw the prefix "העתק: " right-aligned to CONTENT_RIGHT
                _draw_right(c, prefix, y)
            # Draw the item text right-aligned to CONTENT_RIGHT - w_prefix
            _draw_right(c, item.strip(), y, right_x=CONTENT_RIGHT - w_prefix)
            step()

    # ── 5p. Internal note ────────────────────────────────────────────────────
    note = (data.get('note') or '').strip()
    if note:
        step()
        _draw_right(c, f'[{note}]', y)

    c.save()
    return output_path
