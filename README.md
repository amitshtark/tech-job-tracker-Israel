# 🚀 Israeli Tech Junior Job Tracker

מנוע סריקה וסינון אוטונומי המנטר מעל **330 חברות הייטק וסטארטאפים בישראל**, ומאתר משרות פיתוח תוכנה, AI/ML ודאטה המיועדות ל-**0–2 שנות ניסיון** (כולל משרות סטודנט, בוגרים טריים וג'וניור).

הפרויקט רץ באופן עצמאי בענן באמצעות **GitHub Actions** (4 פעמים ביום), מסנן רעשים (פוסל משרות מחוץ לישראל, תפקידי מכירות/תפעול ו-QA ידני), ומייצר קובץ אקסל יומי להורדה עם שמות אנשי הקשר להגשת רפרל.

---

## ⚙️ מדריך הפעלה מהירה למי שעושה Fork

אם עשית **Fork** למאגר הזה וברצונך שהסריקה תרוץ אוטומטית בחשבון ה-GitHub שלך, עקוב אחר 3 הצעדים הבאים:

### שלב 1: הגדרת הרשאות כתיבה לבוט (חובה!)
כדי שהבוט יוכל לעדכן את מאגר המשרות שנצפו (`seen_jobs.json`):
1. היכנס במאגר שלך ל-**Settings** (הגדרות).
2. בתפריט השמאלי, גלול אל **Actions** -> **General**.
3. גלול למטה עד לסעיף **Workflow permissions**.
4. סמן את: **Read and write permissions** ולחץ **Save**.

### שלב 2: הגדרת מפתח Gemini API (חינמי)
הבוט נעזר ב-AI (בבאלק מהיר וחסכוני) כדי לוודא שדרישות הניסיון והמיקום מתאימים:
1. השג מפתח חינמי בלחיצת כפתור מ-[Google AI Studio](https://aistudio.google.com/).
2. במאגר שלך ב-GitHub, היכנס ל-**Settings** -> **Secrets and variables** -> **Actions**.
3. לחץ על **New repository secret**.
4. תחת **Name** רשום: `GEMINI_API_KEY`.
5. תחת **Secret** הדבק את המפתח שקיבלת ולחץ **Add secret**.

### שלב 3: הפעלת ה-Actions
1. היכנס ללשונית **Actions** בראש העמוד.
2. אם מופיע כפתור ירוק בשם **I understand my workflows, go ahead and enable them**, לחץ עליו.
3. לבדיקה ראשונה: לחץ בתפריט השמאלי על **Daily Israeli Tech Job Crawler** -> לחץ על **Run workflow** מצד ימין.

---

## 📊 איפה מוצאים את המשרות החדשות?
בסיום כל ריצה (אוטומטית או ידנית):
1. נכנסים ללשונית **Actions**.
2. לוחצים על הריצה האחרונה (המסומנת ב-V ירוק).
3. בתחתית העמוד, תחת הסעיף **Artifacts**, מחכה קובץ בשם `daily-jobs-excel` להורדה בלחיצה אחת.

---

## 💻 הרצה מקומית (על המחשב שלך)

1. שכפול המאגר והתקנת תלויות:
```bash
git clone [https://github.com/](https://github.com/)<your-username>/tech-job-tracker-Israel.git
cd tech-job-tracker-Israel
pip install -r requirements.txt