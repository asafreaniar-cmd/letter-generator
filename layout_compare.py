"""
layout_compare.py
-----------------
השוואת PDF חזותית עמוד-מול-עמוד כדי לזהות סטיות עימוד.

ההשוואה מתבצעת על ייצוג מרונדר של כל עמוד ולכן רגישה ל:
- שבירות שורה ופסקה
- ריווח ומיקום טקסט
- מיקום חתימה/לוגו/תאריך
- הבדלים בגודל עמוד או במספר עמודים
"""

from __future__ import annotations

from typing import Tuple

from PIL import Image, ImageChops, ImageStat

try:
    import fitz
except ImportError:  # pragma: no cover - handled at runtime
    fitz = None


class LayoutComparisonError(RuntimeError):
    pass


def _render_page(pdf_document, page_index: int, dpi: int) -> Image.Image:
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    page = pdf_document.load_page(page_index)
    pixmap = page.get_pixmap(matrix=matrix, alpha=False)
    return Image.frombytes('RGB', [pixmap.width, pixmap.height], pixmap.samples)


def _pad_to_same_size(left: Image.Image, right: Image.Image) -> Tuple[Image.Image, Image.Image]:
    width = max(left.width, right.width)
    height = max(left.height, right.height)

    if left.size != (width, height):
        padded_left = Image.new('RGB', (width, height), 'white')
        padded_left.paste(left, (0, 0))
        left = padded_left

    if right.size != (width, height):
        padded_right = Image.new('RGB', (width, height), 'white')
        padded_right.paste(right, (0, 0))
        right = padded_right

    return left, right


def compare_pdfs(candidate_path: str, reference_path: str, *, dpi: int = 144, tolerance: float = 0.001) -> dict:
    if fitz is None:
        raise LayoutComparisonError('PyMuPDF אינו מותקן ולכן אי אפשר לבצע השוואת עימוד אוטומטית.')

    candidate = fitz.open(candidate_path)
    reference = fitz.open(reference_path)
    try:
        pages = []
        identical = candidate.page_count == reference.page_count

        for page_index in range(max(candidate.page_count, reference.page_count)):
            if page_index >= candidate.page_count or page_index >= reference.page_count:
                identical = False
                pages.append({
                    'page': page_index + 1,
                    'identical': False,
                    'reason': 'מספר העמודים שונה בין הקבצים.',
                })
                continue

            candidate_image = _render_page(candidate, page_index, dpi)
            reference_image = _render_page(reference, page_index, dpi)
            candidate_image, reference_image = _pad_to_same_size(candidate_image, reference_image)

            diff = ImageChops.difference(candidate_image, reference_image)
            bbox = diff.getbbox()
            grayscale = diff.convert('L')
            stat = ImageStat.Stat(grayscale)
            mean_delta = (stat.mean[0] / 255.0) if stat.mean else 0.0
            rms_delta = (stat.rms[0] / 255.0) if stat.rms else 0.0
            non_zero_histogram = grayscale.point(lambda value: 255 if value else 0).histogram()
            changed_pixels = non_zero_histogram[255] if len(non_zero_histogram) > 255 else 0
            total_pixels = grayscale.width * grayscale.height
            changed_ratio = (changed_pixels / total_pixels) if total_pixels else 0.0
            page_identical = bbox is None or (mean_delta <= tolerance and changed_ratio <= tolerance)

            if not page_identical:
                identical = False

            pages.append({
                'page': page_index + 1,
                'identical': page_identical,
                'candidate_size': list(candidate_image.size),
                'reference_size': list(reference_image.size),
                'mean_delta': round(mean_delta, 6),
                'rms_delta': round(rms_delta, 6),
                'changed_pixels_ratio': round(changed_ratio, 6),
            })

        max_mean_delta = max((page.get('mean_delta', 0.0) for page in pages), default=0.0)
        max_changed_ratio = max((page.get('changed_pixels_ratio', 0.0) for page in pages), default=0.0)

        return {
            'identical': identical,
            'within_tolerance': identical,
            'tolerance': tolerance,
            'dpi': dpi,
            'page_count': {
                'candidate': candidate.page_count,
                'reference': reference.page_count,
                'match': candidate.page_count == reference.page_count,
            },
            'max_mean_delta': round(max_mean_delta, 6),
            'max_changed_pixels_ratio': round(max_changed_ratio, 6),
            'pages': pages,
        }
    finally:
        candidate.close()
        reference.close()
