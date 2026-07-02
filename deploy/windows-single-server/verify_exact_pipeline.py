from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from urllib.parse import urljoin

import requests
from docx2pdf import convert as docx2pdf_convert


REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from layout_compare import compare_pdfs  # noqa: E402


def download_file(url: str, destination: Path):
    with requests.get(url, stream=True, timeout=180) as response:
        response.raise_for_status()
        with destination.open('wb') as handle:
            for chunk in response.iter_content(chunk_size=1024 * 256):
                if chunk:
                    handle.write(chunk)


def main() -> int:
    parser = argparse.ArgumentParser(description='Verify that the generated PDF is identical to an independent Word PDF.')
    parser.add_argument('--base-url', default='http://127.0.0.1:8080/', help='Base URL of the deployed app.')
    parser.add_argument('--dpi', type=int, default=144)
    parser.add_argument('--tolerance', type=float, default=0.0)
    args = parser.parse_args()

    base_url = args.base_url.rstrip('/') + '/'
    payload = {
        'subject': 'בדיקת exact production',
        'recipient_intro': 'לכבוד',
        'recipient_name': 'Dr. Example כהן',
        'recipient_title': 'CEO, Health Corp',
        'greeting': 'שלום רב,',
        'body': (
            'פסקה ראשונה לבדיקת עימוד מקצה לקצה.\n\n'
            'שורה מעורבת Hebrew + English ABC 123 לבדיקה.\n\n'
            'The PDF must be identical to Word.'
        ),
        'closing': 'בכבוד רב,',
        'signers': [
            {
                'name': 'חה"כ יונתן משריקי',
                'title': 'יו"ר ועדת הבריאות',
            }
        ],
        'pdf_profile': 'exact',
    }

    response = requests.post(urljoin(base_url, 'api/generate-pdf'), json=payload, timeout=300)
    response.raise_for_status()
    result = response.json()

    if result.get('engine') not in {'local_word', 'remote_word'}:
        print(f"Unexpected engine: {result.get('engine')}", file=sys.stderr)
        return 2

    work_root = REPO_ROOT / 'instance' / 'storage' / 'verify'
    work_root.mkdir(parents=True, exist_ok=True)
    run_dir = work_root / 'latest'
    if run_dir.exists():
        shutil.rmtree(run_dir)
    run_dir.mkdir(parents=True, exist_ok=True)

    docx_path = run_dir / 'generated.docx'
    api_pdf_path = run_dir / 'generated.pdf'
    reference_pdf_path = run_dir / 'reference.pdf'

    download_file(urljoin(base_url, result['docx_url'].lstrip('/')), docx_path)
    download_file(urljoin(base_url, result['url'].lstrip('/')), api_pdf_path)

    docx2pdf_convert(str(docx_path), str(reference_pdf_path))
    if not reference_pdf_path.exists():
        print('Word reference PDF was not created.', file=sys.stderr)
        return 3

    comparison = compare_pdfs(
        str(api_pdf_path),
        str(reference_pdf_path),
        dpi=args.dpi,
        tolerance=args.tolerance,
    )

    print(comparison)
    return 0 if comparison.get('identical') else 1


if __name__ == '__main__':
    raise SystemExit(main())
