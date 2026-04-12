"""
remote_word_service.py
----------------------
שירות HTTP פשוט ל-Windows שממיר DOCX ל-PDF באמצעות Microsoft Word.

הוא מיועד לפריסה על שרת Windows/VPS נפרד, כדי שהאפליקציה הראשית
תוכל לבקש ממנו PDF זהה ל-Word גם כשה-Mac המקומי כבוי.

Endpoints:
- GET  /health
- POST /convert   (multipart/form-data, field name defaults to "file")
"""

from __future__ import annotations

import os
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file

try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:  # pragma: no cover - handled at runtime
    docx2pdf_convert = None


BASE_DIR = Path(__file__).resolve().parent
SERVICE_ROOT = Path(os.environ.get('REMOTE_WORD_STORAGE_ROOT', BASE_DIR / 'remote-word-data')).resolve()
INBOX_DIR = SERVICE_ROOT / 'inbox'
OUTBOX_DIR = SERVICE_ROOT / 'outbox'
FILE_FIELD = os.environ.get('REMOTE_WORD_SERVICE_FILE_FIELD', 'file').strip() or 'file'
API_KEY = os.environ.get('REMOTE_WORD_SERVICE_API_KEY', '').strip()
HOST = os.environ.get('REMOTE_WORD_HOST', '0.0.0.0')
PORT = int(os.environ.get('REMOTE_WORD_PORT', '8090'))
KEEP_FILES = os.environ.get('REMOTE_WORD_KEEP_FILES', '0').strip().lower() in {'1', 'true', 'yes', 'on'}

INBOX_DIR.mkdir(parents=True, exist_ok=True)
OUTBOX_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)


def _require_api_key():
    if not API_KEY:
        return None

    auth_header = request.headers.get('Authorization', '').strip()
    if auth_header == f'Bearer {API_KEY}':
        return None

    custom_header = request.headers.get('X-API-Key', '').strip()
    if custom_header == API_KEY:
        return None

    return jsonify({'error': 'Unauthorized'}), 401


def _safe_stem(filename: str) -> str:
    stem = Path(filename).stem
    safe = ''.join(ch if ch.isalnum() or ch in '-_ ' else '_' for ch in stem).strip('_ ')
    return safe or 'document'


@app.route('/health')
def health():
    return jsonify({
        'status': 'ok',
        'service': 'remote_word',
        'word_available': docx2pdf_convert is not None,
        'docx2pdf_available': docx2pdf_convert is not None,
        'file_field': FILE_FIELD,
        'storage_root': str(SERVICE_ROOT),
    })


@app.route('/convert', methods=['POST'])
def convert():
    auth_error = _require_api_key()
    if auth_error is not None:
        return auth_error

    if docx2pdf_convert is None:
        return jsonify({'error': 'docx2pdf is not installed on this server'}), 500

    uploaded = request.files.get(FILE_FIELD)
    if uploaded is None:
        return jsonify({'error': f'missing multipart file field "{FILE_FIELD}"'}), 400

    original_name = uploaded.filename or 'document.docx'
    safe_stem = _safe_stem(original_name)
    request_id = uuid.uuid4().hex[:12]
    docx_name = f'{safe_stem}_{request_id}.docx'
    pdf_name = f'{safe_stem}_{request_id}.pdf'
    docx_path = INBOX_DIR / docx_name
    pdf_path = OUTBOX_DIR / pdf_name

    uploaded.save(docx_path)

    try:
        docx2pdf_convert(str(docx_path), str(pdf_path))
    except Exception as exc:  # noqa: BLE001
        return jsonify({'error': f'Word conversion failed: {exc}'}), 502

    if not pdf_path.exists():
        return jsonify({'error': 'Word did not produce a PDF file'}), 502

    response = send_file(
        pdf_path,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=pdf_name,
        conditional=False,
    )

    if not KEEP_FILES:
        @response.call_on_close
        def _cleanup():  # pragma: no cover - cleanup hook
            try:
                docx_path.unlink(missing_ok=True)
            except Exception:
                pass
            try:
                pdf_path.unlink(missing_ok=True)
            except Exception:
                pass

    return response


if __name__ == '__main__':
    print(f'remote_word_service listening on http://{HOST}:{PORT}')
    app.run(host=HOST, port=PORT, debug=False)
