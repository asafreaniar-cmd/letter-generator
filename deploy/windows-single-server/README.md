# שרת Windows יחיד: הפתרון הפשוט ביותר

אם המטרה היא:

- שהאפליקציה תעבוד גם כשה-Mac כבוי
- שייצוא `PDF` יהיה זהה ל-Word
- והקמה תהיה כמה שיותר פשוטה

זה הנתיב המומלץ:

## מה רץ על אותו שרת

- האפליקציה הראשית (`Flask`)
- מנוע `Word` דרך `docx2pdf`
- אחסון קבצים
- reverse proxy עם `Caddy` או `IIS`

## דרישות

1. Windows Server 2022 או Windows 11 Pro קבוע
2. Microsoft Word מותקן ומופעל
3. Python 3.11+
4. הדפסת הגופן `David` בשרת
5. דומיין קבוע, למשל `letters.example.com`

אם אתה מתקין Office LTSC 2024 דרך Office Deployment Tool, יש גם קבצים מוכנים:

- [office-ltsc-2024-config.xml.example](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/office-ltsc-2024-config.xml.example)
- [install_office_ltsc_2024.ps1](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/install_office_ltsc_2024.ps1)

## התקנה

העתק את כל הפרויקט לשרת, למשל ל:

```text
C:\letter-generator
```

הרץ PowerShell כמנהל והפעל:

```powershell
cd C:\letter-generator
powershell -ExecutionPolicy Bypass -File .\deploy\windows-single-server\install.ps1
```

זה יבצע:

- יצירת `venv`
- התקנת dependencies
- הגדרת `PDF_ENGINE=local_word`
- הגדרת `PDF_PROFILE_DEFAULT=exact`
- יצירת Scheduled Task להפעלה בכל login
- הפעלת השרת עכשיו

אם כבר יש לך דומיין מוכן:

```powershell
cd C:\letter-generator
powershell -ExecutionPolicy Bypass -File .\deploy\windows-single-server\bootstrap.ps1 -Domain letters.example.com
```

זה יבצע גם את התקנת `Caddy` ויכין `HTTPS`.

## HTTPS ודומיין

התקן `Caddy` על אותו שרת והשתמש בקובץ:

- [Caddyfile.example](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/Caddyfile.example)
- [install_caddy.ps1](/Users/user/Desktop/כנסת/cloud/letter-generator/deploy/windows-single-server/install_caddy.ps1)

דוגמה:

```text
letters.example.com {
    encode gzip
    reverse_proxy 127.0.0.1:8080
}
```

## הערה חשובה על Word automation

כדי ש-Word automation יעבוד יציב אחרי reboot בלי התערבות ידנית:

- עדיף להקדיש משתמש Windows ייעודי לשרת
- להגדיר auto-login לאותו משתמש
- להריץ את האפליקציה תחת אותו משתמש

זה פשוט יותר ויציב יותר מאשר Office automation מתוך Windows Service קלאסי.

## בדיקת זהות מלאה בין Word ל-PDF

לאחר ההתקנה, הרץ:

```powershell
cd C:\letter-generator
powershell -ExecutionPolicy Bypass -File .\deploy\windows-single-server\verify_exact.ps1
```

זה יבצע:

- יצירת מסמך מבחן דרך ה-API
- הורדת `DOCX` ו-`PDF`
- המרה עצמאית נוספת של אותו `DOCX` דרך Word
- השוואה חזותית בין שני ה-PDF

אם היציאה היא `0`, ה-PDF של האפליקציה זהה חזותית ל-PDF הייחוס של Word.
