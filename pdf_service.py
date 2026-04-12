"""
pdf_service.py
--------------
שכבת המרת PDF צד-שרת עם מנועי המרה מתחלפים.

ה-DOCX הוא מקור האמת. ה-PDF תמיד נוצר מתוך ה-DOCX שנבנה על השרת,
ולא מ-HTML/CSS או canvas.

שני פרופילים מרכזיים:
- portable: המרה עצמאית לגמרי בשרת באמצעות Gotenberg / LibreOffice
- exact: המרה דרך Word מקומי או שירות Word מרוחק, כדי לשמר עימוד כמו Word
"""

from __future__ import annotations

import base64
import json
import os
import platform
import shutil
import subprocess
import tempfile
import threading
import time
from dataclasses import dataclass
from typing import Optional

import requests

from layout_compare import LayoutComparisonError, compare_pdfs

try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:  # pragma: no cover - handled at runtime
    docx2pdf_convert = None


DOCX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
PDF_MIMETYPE = 'application/pdf'
SUPPORTED_ENGINES = ('local_word', 'remote_word', 'gotenberg', 'libreoffice')
SUPPORTED_PROFILES = ('auto', 'portable', 'exact')
_LOCAL_WORD_LOCK = threading.Lock()


class PDFServiceError(RuntimeError):
    def __init__(self, message: str, status_code: int = 500):
        super().__init__(message)
        self.status_code = status_code


@dataclass
class PDFConversionResult:
    engine: str
    requested_profile: str
    comparison: Optional[dict] = None


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {'1', 'true', 'yes', 'on'}


def _normalize_profile(profile: Optional[str]) -> str:
    value = (profile or '').strip().lower()
    if not value or value == 'default':
        return _default_profile()
    if value not in SUPPORTED_PROFILES:
        raise PDFServiceError(
            f'פרופיל PDF לא נתמך: {value}. השתמש ב-auto, portable או exact.',
            status_code=400,
        )
    return value


def _normalize_engine(engine: Optional[str]) -> str:
    value = (engine or '').strip().lower()
    if not value:
        return 'auto'
    if value not in SUPPORTED_ENGINES and value != 'auto':
        raise PDFServiceError(
            f'מנוע PDF לא נתמך: {value}.',
            status_code=400,
        )
    return value


def _remote_word_available() -> bool:
    return bool(os.environ.get('REMOTE_WORD_URL'))


def _windows_word_installed() -> bool:
    try:
        import winreg  # type: ignore
    except ImportError:
        winreg = None  # type: ignore

    if winreg is not None:
        registry_paths = [
            r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
            r'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
        ]
        for subkey in registry_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey) as key:  # type: ignore[attr-defined]
                    value, _ = winreg.QueryValueEx(key, '')  # type: ignore[attr-defined]
                    if value and os.path.exists(value):
                        return True
            except OSError:
                continue

    common_paths = [
        os.path.expandvars(r'%ProgramFiles%\Microsoft Office\root\Office16\WINWORD.EXE'),
        os.path.expandvars(r'%ProgramFiles(x86)%\Microsoft Office\root\Office16\WINWORD.EXE'),
        os.path.expandvars(r'%ProgramFiles%\Microsoft Office\Office16\WINWORD.EXE'),
        os.path.expandvars(r'%ProgramFiles(x86)%\Microsoft Office\Office16\WINWORD.EXE'),
    ]
    if any(path and os.path.exists(path) for path in common_paths):
        return True

    return shutil.which('WINWORD.EXE') is not None


def _local_word_available() -> bool:
    if docx2pdf_convert is None:
        return False

    system = platform.system()
    if system == 'Darwin':
        return os.path.exists('/Applications/Microsoft Word.app')
    if system == 'Windows':
        return _windows_word_installed()
    return False


def _gotenberg_available() -> bool:
    return bool(os.environ.get('GOTENBERG_URL'))


def _libreoffice_binary() -> Optional[str]:
    configured = os.environ.get('LIBREOFFICE_BIN', 'soffice')
    return shutil.which(configured)


def _libreoffice_available() -> bool:
    return _libreoffice_binary() is not None


def _default_profile() -> str:
    configured = (os.environ.get('PDF_PROFILE_DEFAULT') or '').strip().lower()
    if configured:
        if configured not in SUPPORTED_PROFILES:
            raise PDFServiceError(
                f'PDF_PROFILE_DEFAULT לא נתמך: {configured}.',
                status_code=500,
            )
        return configured
    if _local_word_available() or _remote_word_available():
        return 'exact'
    if _gotenberg_available() or _libreoffice_available():
        return 'portable'
    return 'auto'


def get_pdf_capabilities() -> dict:
    engines = {
        'local_word': _local_word_available(),
        'remote_word': _remote_word_available(),
        'gotenberg': _gotenberg_available(),
        'libreoffice': _libreoffice_available(),
    }
    portable_available = engines['gotenberg'] or engines['libreoffice']
    return {
        'configured_engine': _normalize_engine(os.environ.get('PDF_ENGINE', 'auto')),
        'default_profile': _default_profile(),
        'engines': engines,
        'profiles': {
            'auto': any(engines.values()),
            'portable': portable_available,
            'exact': engines['local_word'] or engines['remote_word'],
        },
        'comparison': {
            'enabled_by_default': _env_bool('PDF_COMPARE_ENABLED', False),
            'reference_engine': _normalize_engine(
                os.environ.get('PDF_COMPARE_REFERENCE_ENGINE', '')
            ),
        },
    }


def _engine_order(requested_profile: str, explicit_engine: Optional[str] = None) -> list[str]:
    if explicit_engine:
        return [_normalize_engine(explicit_engine)]

    configured_engine = _normalize_engine(os.environ.get('PDF_ENGINE', 'auto'))
    portable_engines = ['gotenberg', 'libreoffice']
    exact_engines = ['local_word', 'remote_word']
    all_engines = ['local_word', 'remote_word', 'gotenberg', 'libreoffice']

    if requested_profile == 'exact':
        if configured_engine in exact_engines:
            return [configured_engine] + [engine for engine in exact_engines if engine != configured_engine]
        return exact_engines

    if requested_profile == 'portable':
        if configured_engine in portable_engines:
            return [configured_engine] + [engine for engine in portable_engines if engine != configured_engine]
        return portable_engines

    if configured_engine in all_engines:
        return [configured_engine] + [engine for engine in all_engines if engine != configured_engine]

    if _local_word_available() or _remote_word_available():
        return all_engines
    return portable_engines


def _local_word_workdir() -> str:
    """Return a fixed working directory for local Word conversions.

    Using the SAME file paths every time prevents macOS from popping up
    a new file-access permission dialog for each unique DOCX filename.
    Once the user approves access to this directory, all subsequent
    conversions reuse the same approved paths silently.
    """
    configured = os.environ.get('LOCAL_WORD_WORKDIR', '').strip()
    if configured:
        workdir = configured
    else:
        workdir = os.path.join(os.path.dirname(__file__), 'instance', 'word-work')
    os.makedirs(workdir, exist_ok=True)
    return workdir


_WORD_WORK_INPUT = 'input.docx'
_WORD_WORK_OUTPUT = 'output.pdf'


def _convert_with_local_word(docx_path: str, pdf_path: str):
    if not _local_word_available():
        raise PDFServiceError('Microsoft Word או docx2pdf אינם זמינים על השרת.', status_code=503)

    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    attempts = int(os.environ.get('LOCAL_WORD_RETRIES', '3'))
    retry_delay = float(os.environ.get('LOCAL_WORD_RETRY_DELAY', '2.0'))
    last_error = None

    workdir = _local_word_workdir()
    fixed_docx = os.path.join(workdir, _WORD_WORK_INPUT)
    fixed_pdf = os.path.join(workdir, _WORD_WORK_OUTPUT)

    with _LOCAL_WORD_LOCK:
        for attempt in range(1, attempts + 1):
            try:
                # Copy source DOCX to fixed path so Word always sees the
                # same filename and macOS doesn't ask for permission again.
                shutil.copyfile(docx_path, fixed_docx)

                if os.path.exists(fixed_pdf):
                    os.remove(fixed_pdf)

                docx2pdf_convert(fixed_docx, fixed_pdf)

                if not os.path.exists(fixed_pdf):
                    raise PDFServiceError('Word לא יצר קובץ PDF.', status_code=502)

                with open(fixed_pdf, 'rb') as handle:
                    if not handle.read(5).startswith(b'%PDF-'):
                        raise PDFServiceError('Word החזיר קובץ שאינו PDF תקין.', status_code=502)

                # Copy the result to the requested destination.
                shutil.copyfile(fixed_pdf, pdf_path)
                return
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                if attempt >= attempts:
                    break
                time.sleep(retry_delay)

    if isinstance(last_error, PDFServiceError):
        raise last_error
    raise PDFServiceError(
        f'Word נכשל בהמרת PDF: {last_error}',
        status_code=502,
    ) from last_error


def _write_pdf_bytes(pdf_bytes: bytes, output_path: str):
    if not pdf_bytes.startswith(b'%PDF-'):
        raise PDFServiceError('שירות ההמרה החזיר קובץ שאינו PDF תקין.')
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'wb') as handle:
        handle.write(pdf_bytes)


def _convert_with_gotenberg(docx_path: str, pdf_path: str):
    base_url = os.environ.get('GOTENBERG_URL')
    if not base_url:
        raise PDFServiceError('GOTENBERG_URL לא הוגדר.', status_code=503)

    url = f"{base_url.rstrip('/')}/forms/libreoffice/convert"
    timeout = int(os.environ.get('GOTENBERG_TIMEOUT', '180'))
    with open(docx_path, 'rb') as handle:
        response = requests.post(
            url,
            files={'files': (os.path.basename(docx_path), handle, DOCX_MIMETYPE)},
            data={
                'losslessImageCompression': 'true',
                'quality': '100',
            },
            timeout=timeout,
        )

    if response.status_code >= 400:
        raise PDFServiceError(
            f'Gotenberg נכשל ({response.status_code}): {response.text[:300]}',
            status_code=502,
        )

    _write_pdf_bytes(response.content, pdf_path)


def _convert_with_libreoffice(docx_path: str, pdf_path: str):
    soffice = _libreoffice_binary()
    if not soffice:
        raise PDFServiceError('LibreOffice headless אינו זמין על השרת.', status_code=503)

    timeout = int(os.environ.get('LIBREOFFICE_TIMEOUT', '180'))
    with tempfile.TemporaryDirectory(prefix='letter-pdf-') as tmpdir:
        command = [
            soffice,
            '--headless',
            '--convert-to',
            'pdf:writer_pdf_Export',
            '--outdir',
            tmpdir,
            docx_path,
        ]
        process = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
        if process.returncode != 0:
            detail = (process.stderr or process.stdout or '').strip()
            raise PDFServiceError(
                f'LibreOffice נכשל בהמרת PDF: {detail[:400]}',
                status_code=502,
            )

        produced = os.path.join(
            tmpdir,
            f"{os.path.splitext(os.path.basename(docx_path))[0]}.pdf",
        )
        if not os.path.exists(produced):
            detail = (process.stderr or process.stdout or '').strip()
            raise PDFServiceError(
                f'LibreOffice לא יצר PDF. {detail[:400]}',
                status_code=502,
            )
        shutil.copyfile(produced, pdf_path)


def _remote_word_headers() -> dict:
    headers = {}
    api_key = os.environ.get('REMOTE_WORD_API_KEY', '').strip()
    if not api_key:
        return headers

    header_name = os.environ.get('REMOTE_WORD_API_KEY_HEADER', 'Authorization').strip() or 'Authorization'
    if header_name.lower() == 'authorization':
        headers[header_name] = f'Bearer {api_key}'
    else:
        headers[header_name] = api_key
    return headers


def _remote_word_form_fields(filename: str) -> dict:
    fields = {}
    filename_field = os.environ.get('REMOTE_WORD_FILENAME_FIELD', '').strip()
    if filename_field:
        fields[filename_field] = filename

    format_field = os.environ.get('REMOTE_WORD_FORMAT_FIELD', '').strip()
    if format_field:
        fields[format_field] = 'pdf'

    extra_json = os.environ.get('REMOTE_WORD_EXTRA_FORM_JSON', '').strip()
    if extra_json:
        try:
            payload = json.loads(extra_json)
        except json.JSONDecodeError as exc:
            raise PDFServiceError(
                f'REMOTE_WORD_EXTRA_FORM_JSON אינו JSON תקין: {exc}',
                status_code=500,
            ) from exc
        if not isinstance(payload, dict):
            raise PDFServiceError('REMOTE_WORD_EXTRA_FORM_JSON חייב להיות אובייקט JSON.', status_code=500)
        for key, value in payload.items():
            fields[str(key)] = '' if value is None else str(value)
    return fields


def _download_remote_pdf(url: str, destination_path: str, headers: dict, timeout: int):
    response = requests.get(url, headers=headers, timeout=timeout)
    if response.status_code >= 400:
        raise PDFServiceError(
            f'שירות Word מרוחק החזיר קישור הורדה לא תקין ({response.status_code}).',
            status_code=502,
        )
    _write_pdf_bytes(response.content, destination_path)


def _handle_remote_word_response(response: requests.Response, destination_path: str, headers: dict, timeout: int):
    content_type = (response.headers.get('Content-Type') or '').lower()
    body = response.content

    if PDF_MIMETYPE in content_type or body.startswith(b'%PDF-'):
        _write_pdf_bytes(body, destination_path)
        return

    try:
        payload = response.json()
    except ValueError as exc:
        snippet = body[:160].decode('utf-8', errors='replace')
        raise PDFServiceError(
            f'שירות Word מרוחק החזיר תשובה לא מזוהה: {snippet}',
            status_code=502,
        ) from exc

    direct_base64 = payload.get('pdf_base64') or payload.get('base64')
    if direct_base64:
        try:
            pdf_bytes = base64.b64decode(direct_base64)
        except Exception as exc:  # noqa: BLE001
            raise PDFServiceError('לא ניתן לפענח את ה-PDF שהוחזר משירות Word מרוחק.', status_code=502) from exc
        _write_pdf_bytes(pdf_bytes, destination_path)
        return

    download_url = payload.get('download_url') or payload.get('file_url') or payload.get('url')
    if download_url:
        _download_remote_pdf(download_url, destination_path, headers, timeout)
        return

    raise PDFServiceError(
        'שירות Word מרוחק לא החזיר PDF, pdf_base64 או download_url.',
        status_code=502,
    )


def _convert_with_remote_word(docx_path: str, pdf_path: str):
    url = os.environ.get('REMOTE_WORD_URL', '').strip()
    if not url:
        raise PDFServiceError(
            'פרופיל exact דורש REMOTE_WORD_URL לשירות Word מרוחק.',
            status_code=503,
        )

    timeout = int(os.environ.get('REMOTE_WORD_TIMEOUT', '240'))
    headers = _remote_word_headers()
    form_fields = _remote_word_form_fields(os.path.basename(docx_path))
    file_field = os.environ.get('REMOTE_WORD_FILE_FIELD', 'file').strip() or 'file'

    with open(docx_path, 'rb') as handle:
        response = requests.post(
            url,
            headers=headers,
            data=form_fields or None,
            files={file_field: (os.path.basename(docx_path), handle, DOCX_MIMETYPE)},
            timeout=timeout,
        )

    if response.status_code >= 400:
        raise PDFServiceError(
            f'שירות Word מרוחק נכשל ({response.status_code}): {response.text[:300]}',
            status_code=502,
        )

    _handle_remote_word_response(response, pdf_path, headers, timeout)


def _convert_with_engine(engine: str, docx_path: str, pdf_path: str):
    if engine == 'gotenberg':
        _convert_with_gotenberg(docx_path, pdf_path)
        return
    if engine == 'libreoffice':
        _convert_with_libreoffice(docx_path, pdf_path)
        return
    if engine == 'local_word':
        _convert_with_local_word(docx_path, pdf_path)
        return
    if engine == 'remote_word':
        _convert_with_remote_word(docx_path, pdf_path)
        return
    raise PDFServiceError(f'מנוע PDF לא נתמך: {engine}', status_code=500)


def _engine_is_available(engine: str) -> bool:
    if engine == 'remote_word':
        return _remote_word_available()
    if engine == 'local_word':
        return _local_word_available()
    if engine == 'gotenberg':
        return _gotenberg_available()
    if engine == 'libreoffice':
        return _libreoffice_available()
    return False


def _reference_engine_for_comparison(primary_engine: str, requested_profile: str) -> Optional[str]:
    configured = _normalize_engine(os.environ.get('PDF_COMPARE_REFERENCE_ENGINE', ''))
    if configured != 'auto':
        return configured or None

    if primary_engine != 'local_word' and _local_word_available():
        return 'local_word'
    if primary_engine != 'remote_word' and _remote_word_available():
        return 'remote_word'
    if requested_profile == 'portable' and _gotenberg_available():
        return 'gotenberg'
    return None


def _run_comparison(docx_path: str, pdf_path: str, primary_engine: str, requested_profile: str) -> dict:
    reference_engine = _reference_engine_for_comparison(primary_engine, requested_profile)
    if not reference_engine:
        return {
            'status': 'skipped',
            'reason': 'לא הוגדר מנוע ייחוס להשוואת עימוד.',
        }

    if reference_engine == primary_engine:
        return {
            'status': 'skipped',
            'reason': 'מנוע הייחוס זהה למנוע ההפקה ולכן אין השוואה מועילה.',
            'reference_engine': reference_engine,
        }

    if not _engine_is_available(reference_engine):
        return {
            'status': 'skipped',
            'reason': 'מנוע הייחוס אינו זמין כעת.',
            'reference_engine': reference_engine,
        }

    dpi = int(os.environ.get('PDF_COMPARE_DPI', '144'))
    tolerance = float(os.environ.get('PDF_COMPARE_TOLERANCE', '0.001'))

    try:
        with tempfile.TemporaryDirectory(prefix='letter-pdf-compare-') as tmpdir:
            reference_path = os.path.join(tmpdir, 'reference.pdf')
            _convert_with_engine(reference_engine, docx_path, reference_path)
            comparison = compare_pdfs(
                candidate_path=pdf_path,
                reference_path=reference_path,
                dpi=dpi,
                tolerance=tolerance,
            )
    except (PDFServiceError, LayoutComparisonError) as exc:
        return {
            'status': 'error',
            'reason': str(exc),
            'reference_engine': reference_engine,
        }

    comparison['status'] = 'completed'
    comparison['reference_engine'] = reference_engine
    comparison['primary_engine'] = primary_engine
    comparison['source_of_truth'] = 'DOCX rendered through the reference engine'
    return comparison


def convert_docx_to_pdf(
    docx_path: str,
    pdf_path: str,
    *,
    profile: Optional[str] = None,
    compare_layout: Optional[bool] = None,
    engine: Optional[str] = None,
) -> PDFConversionResult:
    requested_profile = _normalize_profile(profile)
    explicit_engine = _normalize_engine(engine)
    if explicit_engine == 'auto':
        explicit_engine = None

    if requested_profile == 'exact' and not (_local_word_available() or _remote_word_available()) and not explicit_engine:
        raise PDFServiceError(
            'נדרש מנוע Word כדי להפיק PDF זהה ל-Word. הפעל Microsoft Word על השרת או הגדר REMOTE_WORD_URL.',
            status_code=503,
        )

    engine_order = _engine_order(requested_profile, explicit_engine=explicit_engine)
    errors = []
    chosen_engine = None

    for candidate_engine in engine_order:
        if not _engine_is_available(candidate_engine):
            errors.append(f'{candidate_engine}: לא זמין')
            continue
        try:
            _convert_with_engine(candidate_engine, docx_path, pdf_path)
            chosen_engine = candidate_engine
            break
        except PDFServiceError as exc:
            errors.append(f'{candidate_engine}: {exc}')
            if requested_profile == 'exact' or explicit_engine:
                break

    if not chosen_engine:
        message = ' ; '.join(errors) if errors else 'לא נמצא מנוע PDF זמין.'
        raise PDFServiceError(message, status_code=503)

    should_compare = compare_layout if compare_layout is not None else _env_bool('PDF_COMPARE_ENABLED', False)
    comparison = None
    if should_compare:
        comparison = _run_comparison(docx_path, pdf_path, chosen_engine, requested_profile)

    return PDFConversionResult(
        engine=chosen_engine,
        requested_profile=requested_profile,
        comparison=comparison,
    )
