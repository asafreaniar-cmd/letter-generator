"""
app.py
------
Backend ל-web app של מחולל המכתבים.

המערכת בנויה לפריסה צד-שרת:
- יצירת DOCX בצד השרת
- המרת PDF בצד השרת מתוך ה-DOCX, דרך Gotenberg / LibreOffice / remote Word
- אחסון טיוטות ומסמכים תחת storage ייעודי של השרת
"""

import os
import uuid
from datetime import datetime
from urllib.parse import quote

from flask import Flask, abort, jsonify, request, send_file, send_from_directory

from letter_builder import build_letter
from pdf_service import PDFServiceError, convert_docx_to_pdf, get_pdf_capabilities
from storage import AppStorage


BASE_DIR = os.path.dirname(__file__)
STATIC_DIR = os.path.join(BASE_DIR, 'static')
STORAGE = AppStorage(BASE_DIR)

app = Flask(
    __name__,
    static_folder=STATIC_DIR,
    static_url_path='',
    root_path=BASE_DIR,
    instance_path=os.path.join(BASE_DIR, 'instance'),
)


def _safe_subject(subject: str) -> str:
    subject = (subject or '').strip()[:40] or 'מכתב'
    safe = ''.join(char if char.isalnum() or char in '-_ ' else '_' for char in subject).strip('_')
    return safe or 'מכתב'


def _draft_or_404(draft_id: str) -> dict:
    draft = STORAGE.load_draft(draft_id)
    if draft is None:
        abort(404, description='טיוטה לא נמצאה')
    return draft


def _bool_from_request(value, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {'1', 'true', 'yes', 'on'}


@app.route('/')
def index():
    return send_from_directory(STATIC_DIR, 'index.html')


@app.route('/manifest.webmanifest')
def manifest():
    return send_from_directory(STATIC_DIR, 'manifest.webmanifest')


@app.route('/sw.js')
def service_worker():
    return send_from_directory(STATIC_DIR, 'sw.js')


@app.route('/api/health')
def health():
    capabilities = get_pdf_capabilities()
    return jsonify({
        'status': 'ok',
        'storage_root': STORAGE.root,
        **capabilities,
    })


@app.route('/api/generate', methods=['POST'])
def generate_word():
    data = request.get_json(force=True)
    if not data:
        return jsonify({'error': 'נתונים חסרים'}), 400

    safe_subject = _safe_subject(data.get('subject', ''))
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'{safe_subject}_{timestamp}.docx'
    output_path = STORAGE.document_path(filename)

    try:
        build_letter(data, output_path)
    except Exception as exc:
        app.logger.exception('שגיאה ביצירת ה-Word')
        return jsonify({'error': str(exc)}), 500

    return jsonify({
        'filename': filename,
        'url': f'/api/download/{quote(filename, safe="")}',
    })


@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    data = request.get_json(force=True)
    if not data:
        return jsonify({'error': 'נתונים חסרים'}), 400

    requested_profile = data.get('pdf_profile') or data.get('pdfProfile')
    compare_layout = (
        _bool_from_request(data.get('compare_layout'))
        if 'compare_layout' in data
        else None
    )
    explicit_engine = data.get('pdf_engine')
    safe_subject = _safe_subject(data.get('subject', ''))
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    docx_filename = f'{safe_subject}_{timestamp}.docx'
    pdf_filename = f'{safe_subject}_{timestamp}.pdf'
    docx_path = STORAGE.document_path(docx_filename)
    pdf_path = STORAGE.document_path(pdf_filename)

    try:
        build_letter(data, docx_path)
        result = convert_docx_to_pdf(
            docx_path,
            pdf_path,
            profile=requested_profile,
            compare_layout=compare_layout,
            engine=explicit_engine,
        )
    except PDFServiceError as exc:
        app.logger.exception('שגיאה בהמרת PDF')
        return jsonify({'error': str(exc)}), exc.status_code
    except Exception as exc:
        app.logger.exception('שגיאה ביצירת ה-PDF')
        return jsonify({'error': str(exc)}), 500

    return jsonify({
        'filename': pdf_filename,
        'url': f'/api/download/{quote(pdf_filename, safe="")}',
        'preview_url': f'/api/download/{quote(pdf_filename, safe="")}?download=0',
        'docx_url': f'/api/download/{quote(docx_filename, safe="")}',
        'engine': result.engine,
        'requested_profile': result.requested_profile,
        'comparison': result.comparison,
    })


@app.route('/api/download/<filename>')
def download(filename):
    if '..' in filename or '/' in filename or '\\' in filename:
        abort(400)

    file_path = STORAGE.document_path(filename)
    if not os.path.exists(file_path):
        abort(404)

    is_download = request.args.get('download', '1') != '0'
    mimetype = (
        'application/pdf'
        if filename.endswith('.pdf')
        else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    return send_file(file_path, as_attachment=is_download, download_name=filename, mimetype=mimetype)


@app.route('/api/drafts', methods=['GET'])
def list_drafts():
    return jsonify(STORAGE.list_drafts())


@app.route('/api/drafts', methods=['POST'])
def create_draft():
    data = request.get_json(force=True)
    if not data:
        return jsonify({'error': 'נתונים חסרים'}), 400

    draft_id = str(uuid.uuid4())
    now = datetime.now().isoformat()
    data['_id'] = draft_id
    data.setdefault('_name', data.get('subject', 'טיוטה') or 'טיוטה')
    data['_created_at'] = now
    data['_updated_at'] = now
    STORAGE.save_draft(draft_id, data)
    return jsonify({'id': draft_id, 'name': data['_name']}), 201


@app.route('/api/drafts/<draft_id>', methods=['GET'])
def get_draft(draft_id):
    return jsonify(_draft_or_404(draft_id))


@app.route('/api/drafts/<draft_id>', methods=['PUT'])
def update_draft(draft_id):
    data = request.get_json(force=True)
    if not data:
        return jsonify({'error': 'נתונים חסרים'}), 400

    existing = _draft_or_404(draft_id)
    data['_id'] = draft_id
    data['_created_at'] = existing.get('_created_at', datetime.now().isoformat())
    data['_updated_at'] = datetime.now().isoformat()
    data.setdefault('_name', data.get('subject', existing.get('_name', 'טיוטה')) or 'טיוטה')
    STORAGE.save_draft(draft_id, data)
    return jsonify({'id': draft_id, 'name': data['_name']})


@app.route('/api/drafts/<draft_id>', methods=['DELETE'])
def delete_draft(draft_id):
    if not STORAGE.delete_draft(draft_id):
        abort(404)
    return jsonify({'deleted': draft_id})


if __name__ == '__main__':
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', 5000))
    print(f'🏛️  מחולל מכתבים – פועל על http://{host}:{port}')
    app.run(debug=True, host=host, port=port)
