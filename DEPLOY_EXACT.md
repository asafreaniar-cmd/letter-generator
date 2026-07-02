# פריסה ללא תלות ב-Mac

כדי שייצוא `Word` ו-`PDF` יעבוד גם כשהמחשב האישי כבוי, צריך לפרוס את המערכת על תשתית שתמיד דולקת.

אם חשוב לשמור על עימוד זהה ל-Word, זו הארכיטקטורה הנכונה:

## מסלול מומלץ: שרת Windows יחיד

זה הפתרון הפשוט ביותר בפועל:

- שרת Windows אחד, קבוע ודולק תמיד
- עליו רצה האפליקציה הראשית
- עליו מותקן Microsoft Word
- האפליקציה משתמשת ב-`local_word` דרך `docx2pdf`
- אין Mac בשרשרת

קבצי הפריסה:

- [deploy/windows-single-server/bootstrap.ps1](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/bootstrap.ps1)
- [deploy/windows-single-server/install.ps1](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/install.ps1)
- [deploy/windows-single-server/install_caddy.ps1](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/install_caddy.ps1)
- [deploy/windows-single-server/README.md](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/README.md)
- [deploy/windows-single-server/Caddyfile.example](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/Caddyfile.example)
- [deploy/windows-single-server/verify_exact.ps1](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/verify_exact.ps1)
- [deploy/windows-single-server/verify_exact_pipeline.py](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/verify_exact_pipeline.py)
- [windows_requirements.txt](/Users/user/Desktop/כנסת/cloud/letter-generator/windows_requirements.txt)

אם אתה רוצה הקמה מהירה ויציבה, זה המסלול שכדאי לבחור.

## 1. האפליקציה הראשית

הקבצים הבאים כבר מוכנים לפריסה:

- [Dockerfile](/Users/user/Desktop/כנסת/cloud/letter-generator/Dockerfile)
- [compose.yaml](/Users/user/Desktop/כנסת/cloud/letter-generator/compose.yaml)
- [app.py](/Users/user/Desktop/כנסת/cloud/letter-generator/app.py)

את החלק הזה אפשר להרים על Linux VPS או container platform.

## 2. חלופה: שירות Word מרוחק

קובץ השירות:

- [remote_word_service.py](/Users/user/Desktop/כנסת/cloud/letter-generator/remote_word_service.py)

זה שירות Flask ל-Windows שמקבל `DOCX` ומחזיר `PDF` דרך `Microsoft Word`.

התקנה על שרת Windows:

```powershell
py -m venv .venv
.venv\Scripts\activate
pip install -r remote_word_service_requirements.txt
set REMOTE_WORD_SERVICE_API_KEY=change-me
set REMOTE_WORD_HOST=0.0.0.0
set REMOTE_WORD_PORT=8090
python remote_word_service.py
```

## 3. חיבור האפליקציה הראשית לשירות Word

בשרת של האפליקציה הראשית:

```bash
PDF_ENGINE=remote_word
PDF_PROFILE_DEFAULT=exact
REMOTE_WORD_URL=http://WINDOWS_SERVER_IP:8090/convert
REMOTE_WORD_API_KEY=change-me
REMOTE_WORD_API_KEY_HEADER=Authorization
```

## 4. תוצאה

במבנה הזה:

- הטלפון מדבר רק עם השרת
- השרת יוצר `DOCX`
- השרת שולח את ה-`DOCX` לשירות Word מרוחק
- שירות Word מחזיר `PDF` זהה ל-Word
- אין תלות ב-Mac האישי

## מה לא מספיק

לא מספיק רק לפתוח את האפליקציה על ה-Mac עם `localhost` או IP פנימי.

במצב כזה:

- אם ה-Mac כבוי, אין backend
- אם ה-IP משתנה, הקישור נשבר
- אם Word רץ על ה-Mac, הייצוא נשאר תלוי במחשב האישי

## מסקנה

כדי שזה יעבוד גם כשהמחשב כבוי וגם בלי לשנות את צורת המכתב, חייבים:

1. שרת קבוע לאפליקציה הראשית
2. מנוע Word על שרת קבוע, מקומי לאותו שרת או מרוחק ממנו
