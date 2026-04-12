"""
storage.py
----------
אחסון צד-שרת למסמכים וטיוטות.

ברירת המחדל שומרת תחת instance/storage כדי שהקוד יהיה מוכן
לפריסה על שרת או על volume ייעודי, ולא תלוי בתיקיות פיתוח מקומיות.
"""

import json
import os
from typing import Optional


class AppStorage:
    def __init__(self, base_dir: str, storage_root: Optional[str] = None):
        root = storage_root or os.environ.get('STORAGE_ROOT')
        if root:
            self.root = os.path.abspath(root)
        else:
            self.root = os.path.join(base_dir, 'instance', 'storage')

        self.drafts_dir = os.path.join(self.root, 'drafts')
        self.documents_dir = os.path.join(self.root, 'documents')

        os.makedirs(self.drafts_dir, exist_ok=True)
        os.makedirs(self.documents_dir, exist_ok=True)

    def draft_path(self, draft_id: str) -> str:
        return os.path.join(self.drafts_dir, f'{draft_id}.json')

    def document_path(self, filename: str) -> str:
        return os.path.join(self.documents_dir, filename)

    def save_draft(self, draft_id: str, data: dict):
        with open(self.draft_path(draft_id), 'w', encoding='utf-8') as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)

    def load_draft(self, draft_id: str) -> Optional[dict]:
        path = self.draft_path(draft_id)
        if not os.path.exists(path):
            return None
        with open(path, encoding='utf-8') as handle:
            return json.load(handle)

    def delete_draft(self, draft_id: str) -> bool:
        path = self.draft_path(draft_id)
        if not os.path.exists(path):
            return False
        os.remove(path)
        return True

    def list_drafts(self) -> list[dict]:
        drafts = []
        for filename in sorted(os.listdir(self.drafts_dir)):
            if not filename.endswith('.json'):
                continue
            draft_id = filename[:-5]
            try:
                with open(os.path.join(self.drafts_dir, filename), encoding='utf-8') as handle:
                    draft = json.load(handle)
            except Exception:
                continue
            drafts.append({
                'id': draft_id,
                'name': draft.get('_name', 'ללא שם'),
                'subject': draft.get('subject', ''),
                'updated_at': draft.get('_updated_at', ''),
            })
        drafts.sort(key=lambda item: item.get('updated_at', ''), reverse=True)
        return drafts
