# Israeli Tech Junior Job Tracker

מנוע סריקה וסינון אוטונומי המנטר מעל **325 חברות הייטק וסטארטאפים בישראל**, ומאתר משרות פיתוח תוכנה, AI/ML, דאטה ו-Junior MLOps המיועדות ל-**0–2 שנות ניסיון** (כולל משרות סטודנט, בוגרים טריים וג'וניור — עם אפשרות התאמה אישית מלאה דרך `config.json`).

הפרויקט רץ באופן עצמאי בענן באמצעות **GitHub Actions** (4 פעמים ביום), מסנן רעשים (פוסל משרות מחוץ לישראל, דפי קטגוריות/חיפוש, תפקידי מכירות/תפעול ו-QA ידני), מציג את כל המשרות ב**לוח אינטרנטי אינטראקטיבי** (`index.html` ב-GitHub Pages) ומייצר גם קובץ אקסל יומי להורדה.

---

## הפעלה מהירה למי שעושה Fork

אם עשית **Fork** למאגר הזה וברצונך שהסריקה והאתר יעבדו אוטומטית בחשבון ה-GitHub שלך, בצע את 4 השלבים הבאים:

### שלב 1: התאמת דרישות החיפוש שלך (`config.json`)
אין צורך לגעת בקוד ה-Python. ניתן לערוך ישירות באתר של GitHub את הקובץ `config.json` (לחיצה על הקובץ -> אייקון העיפרון):
* `min_yoe` / `max_yoe`: טווח שנות הניסיון הרצוי (ברירת מחדל: `0` עד `2`).
* `location`: מיקום המשרה (ברירת מחדל: `"Israel"`).
* `include_student_and_graduate`: האם לכלול משרות סטודנט ובוגרים טריים (`true` / `false`).
* `target_roles`: רשימת תפקידי הפיתוח/טכנולוגיה שהסורק יאשר.
* `excluded_roles`: רשימת תחומים שייפסלו מראש (כגון `"Manual QA"`, `"DevOps"`).

### שלב 2: הגדרת הרשאות כתיבה לבוט (חובה)
כדי שהבוט יוכל לעדכן אוטומטית את מאגר המשרות שנצפו (`seen_jobs.json`) לאחר כל ריצה:
1. היכנס במאגר שלך ל-**Settings**.
2. בתפריט השמאלי, בחר **Actions** -> **General**.
3. גלול לסעיף **Workflow permissions**.
4. סמן **Read and write permissions** ולחץ **Save**.

### שלב 3: הגדרת מפתח Gemini API (חינמי)
הסורק נעזר במודל Gemini (`gemini-flash-lite-latest`) כדי לקרוא את מלוא דרישות המשרה ולוודא התאמה לקריטריונים:
1. צור מפתח API חינמי ב-[Google AI Studio](https://aistudio.google.com/).
2. במאגר שלך ב-GitHub, היכנס ל-**Settings** -> **Secrets and variables** -> **Actions**.
3. לחץ על **New repository secret**.
4. תחת **Name** הזן: `GEMINI_API_KEY`.
5. תחת **Secret** הדבק את המפתח ולחץ **Add secret**.

### שלב 4: הפעלת ה-Actions ולוח המשרות (GitHub Pages)
1. היכנס ללשונית **Actions** -> אם מופיע כפתור ירוק **I understand my workflows, go ahead and enable them**, לחץ עליו. להרצת בדיקה ראשונה: **Daily Israeli Tech Job Crawler** -> **Run workflow**.
2. להפעלת האתר: היכנס ל-**Settings** -> **Pages**, בחר תחת **Branch** את **`main`** ותיקיית **`/ (root)`**, ולחץ **Save**.
3. בתוך כדקה לוח המשרות יעלה לאוויר בכתובת:  
   `https://<your-username>.github.io/tech-job-tracker-Israel/`

---

## שימוש בלוח המשרות, ברפרלים פרטיים ובדוחות

* **לוח המשרות באתר (`index.html`):** טוען אוטומטית את `seen_jobs.json`, מאפשר חיפוש וסינון לפי תחום או סוג רפרל, ומאפשר לסמן משרות כ**"הגשתי"**, לשמור להמשך או להסיר משרות לא מתאימות (הסימונים נשמרים מקומית בדפדפן ב-`localStorage`).
* **טעינת רפרלים אישיים (פרטי בדפדפן בלבד):** ניתן ללחוץ באתר על כפתור **"רפרלים אישיים"** ולטעון פעם אחת קובץ JSON מקומי מהמחשב בפורמט:
  ```json
  [
    {
      "name": "Company Name",
      "referral_contact": "שם איש הקשר",
      "referral_type": "איש קשר אישי"
    }
  ]
  ```
  המיפוי נשמר אך ורק ב-`localStorage` בדפדפן שלך ואינו עולה ל-GitHub, כך שתוכל לראות את שמות אנשי הקשר שלך על כל משרה מבלי לחשוף מידע פרטי במאגר ציבורי.
* **קובץ אקסל יומי:** בסיום כל ריצה ב-**Actions**, ניתן להיכנס לריצה האחרונה ולהוריד מתחתית העמוד (תחת **Artifacts**) את הקובץ `daily-jobs-excel`.

---

## הרצה מקומית (על המחשב שלך)

1. שכפול המאגר והתקנת התלויות:
```bash
git clone https://github.com/<your-username>/tech-job-tracker-Israel.git
cd tech-job-tracker-Israel
pip install -r requirements.txt
```

2. יצירת קובץ `.env` על בסיס `.env.example` עם מפתח ה-Gemini שלך:
```env
GEMINI_API_KEY="your-free-gemini-api-key"
```

3. התאמת `config.json` (במידת הצורך) והפעלת הסורק:
```bash
python job_crawler.py
```

---

## מבנה הקבצים המרכזיים
* `config.json` — קובץ הגדרות הדרישות (שנות ניסיון, תפקידים, מיקום).
* `job_crawler.py` — מנוע הסריקה המקבילי והסינון מול Gemini AI.
* `index.html` — ממשק הווב האינטראקטיבי למעקב אחר משרות וטעינת רפרלים מקומית.
* `seen_jobs.json` — מאגר המשרות שאותרו (מתעדכן אוטומטית על ידי הבוט).
* `companies_registry_public.json` — רשימת החברות והקישורים לעמודי הקריירה.
* `.env.example` — קובץ תבנית למשתני הסביבה להרצה מקומית.
* `.github/workflows/job_crawler.yml` — הגדרת התזמון האוטומטי ב-GitHub Actions (4 פעמים ביום).