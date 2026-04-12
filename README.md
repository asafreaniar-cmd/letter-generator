---
title: Letter Generator
emoji: 📝
colorFrom: blue
colorTo: blue
sdk: docker
app_port: 8080
pinned: false
---

# מחולל מכתבים: Mobile-First Web App

מערכת ליצירת מכתבים רשמיים מהטלפון או מהדפדפן, עם:

- יצירת `DOCX` בצד השרת
- יצירת `PDF` מתוך ה-`DOCX` בצד השרת
- שמירת טיוטות ומסמכים תחת storage של השרת
- `PWA` שניתן להוסיף למסך הבית באייפון

## מה השתנה

המערכת כבר לא בנויה ככלי מקומי שתלוי ב-`Microsoft Word` על ה-Mac.

במקום זה:

- ה-`DOCX` נוצר בשרת באמצעות `python-docx`
- ה-`PDF` נוצר בשרת מתוך ה-`DOCX`, עם שני מסלולים:
  - `portable`: `Gotenberg` או `LibreOffice headless`
  - `exact`: `Microsoft Word` על השרת או שירות Word מרוחק שהוא מנוע העימוד המחייב
- הקבצים נשמרים תחת `STORAGE_ROOT`
- הממשק כולל `manifest` + `service worker` כדי להתנהג כמו אפליקציה

## פיתוח מקומי

```bash
pip3 install -r requirements.txt
PDF_ENGINE=gotenberg \
PDF_PROFILE_DEFAULT=portable \
GOTENBERG_URL=https://demo.gotenberg.dev \
HOST=0.0.0.0 \
PORT=8080 \
python3 app.py
```

הערה:

- `https://demo.gotenberg.dev` מתאים לבדיקה בלבד, לא לייצור
- בייצור יש להרים שירות Gotenberg משלך
- פרופיל `exact` יעבוד אם `Microsoft Word` מותקן על השרת הזה, או אם מוגדר `REMOTE_WORD_URL`

## פריסה מומלצת

### עם Compose

```bash
docker compose up --build
```

הקבצים יישמרו תחת:

```text
./server-data
```

ברירת המחדל של ה-compose:

- `app` על פורט `8080`
- `gotenberg` על רשת פנימית
- `STORAGE_ROOT=/data`
- `portable` הוא פרופיל ברירת המחדל

## Exact PDF

כדי לקבל `PDF` זהה ל-Word, צריך להשתמש במנוע Word אמיתי:

- `Microsoft Word` על השרת המקומי באמצעות `docx2pdf`
- או שירות Word מרוחק שהאפליקציה תשלח אליו `DOCX` ותקבל ממנו `PDF`

החוזה הנתמך כעת:

- `POST REMOTE_WORD_URL`
- `multipart/form-data`
- שדה קובץ ברירת מחדל: `file`
- אפשר לצרף `Authorization: Bearer <token>` או כותרת אחרת באמצעות `REMOTE_WORD_API_KEY_HEADER`
- התשובה יכולה להיות:
  - `application/pdf`
  - JSON עם `download_url`
  - JSON עם `pdf_base64`

זה מאפשר לחבר:

- שירות Windows/Word ארגוני משלך
- שירות conversion מסחרי/פנימי שמרנדר `DOCX -> PDF` עם מנוע Word
- backend נפרד שמבצע conversion בסביבת Office ומחזיר `PDF`

יש גם שירות מוכן לפריסה ב-Windows בתוך הפרויקט:

- [remote_word_service.py](/Users/user/Desktop/כנסת/cloud/letter-generator/remote_word_service.py)
- [remote_word_service_requirements.txt](/Users/user/Desktop/כנסת/cloud/letter-generator/remote_word_service_requirements.txt)
- [DEPLOY_EXACT.md](/Users/user/Desktop/כנסת/cloud/letter-generator/DEPLOY_EXACT.md)

## בדיקת סטיות עימוד

אפשר להפעיל בדיקת עימוד אוטומטית:

- `PDF_COMPARE_ENABLED=1`
- `PDF_COMPARE_REFERENCE_ENGINE=remote_word`

במצב כזה, אחרי יצירת ה-PDF המערכת תפיק PDF ייחוס מאותו `DOCX` דרך מנוע הייחוס ותבצע השוואה חזותית עמוד-מול-עמוד. התוצאה מוחזרת ב-JSON של `/api/generate-pdf`.

## משתני סביבה

| משתנה | תיאור |
|---|---|
| `HOST` | כתובת bind של השרת |
| `PORT` | פורט השרת |
| `STORAGE_ROOT` | נתיב שמירת טיוטות ומסמכים בשרת |
| `PDF_ENGINE` | `local_word`, `remote_word`, `gotenberg`, `libreoffice`, או `auto` |
| `PDF_PROFILE_DEFAULT` | `exact`, `portable`, או `auto` |
| `GOTENBERG_URL` | כתובת שירות Gotenberg |
| `REMOTE_WORD_URL` | כתובת שירות Word מרוחק |
| `REMOTE_WORD_API_KEY` | טוקן לשירות Word מרוחק |
| `PDF_COMPARE_ENABLED` | `1` כדי לבצע השוואת עימוד |
| `PDF_COMPARE_REFERENCE_ENGINE` | מנוע הייחוס להשוואה, בדרך כלל `remote_word` |

## API

| Method | Path | תיאור |
|---|---|---|
| `GET` | `/` | האפליקציה |
| `GET` | `/api/health` | סטטוס backend |
| `POST` | `/api/generate` | יצירת `DOCX` |
| `POST` | `/api/generate-pdf` | יצירת `PDF` |
| `GET` | `/api/download/<filename>` | הורדה או preview inline |
| `GET` | `/api/drafts` | רשימת טיוטות |
| `POST` | `/api/drafts` | שמירת טיוטה |
| `GET` | `/api/drafts/<id>` | טעינת טיוטה |
| `PUT` | `/api/drafts/<id>` | עדכון טיוטה |
| `DELETE` | `/api/drafts/<id>` | מחיקת טיוטה |

## קבצים חשובים

| קובץ | תפקיד |
|---|---|
| `app.py` | Flask API |
| `letter_builder.py` | יצירת `DOCX` מהתבנית |
| `pdf_service.py` | המרת `DOCX -> PDF` בשרת לפי פרופיל/מנוע |
| `layout_compare.py` | השוואת עימוד חזותית בין PDFים |
| `remote_word_service.py` | שירות Word מרוחק ל-Windows |
| `storage.py` | אחסון טיוטות ומסמכים בשרת |
| `static/index.html` | ממשק mobile-first |
| `static/manifest.webmanifest` | PWA manifest |
| `static/sw.js` | service worker |
| `compose.yaml` | פריסה מומלצת עם Gotenberg |
| `ARCHITECTURE.md` | שתי הארכיטקטורות האפשריות |

## החלטה מקצועית

אם המטרה היא עצמאות מלאה מה-Mac, פרופיל `portable` הוא המסלול הנכון.

אם המטרה העליונה היא fidelity מקסימלי לעימוד של Word, צריך להגדיר שירות `remote_word` ולבחור פרופיל `exact`. פירוט מלא נמצא ב-[ARCHITECTURE.md](/Users/user/Desktop/כנסת/cloud/letter-generator/ARCHITECTURE.md).
