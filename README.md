# מוקד שפ״ע אריאל — תמונת מצב

אתר יחיד, self-contained: מפה כקנבס מלא־מסך עם שכבות מידע צפות, 5 מדדי־על,
סינון צולב, ומצב סיפור בן 7 תצוגות לישיבת בוקר.

## הרצה

הקובץ `index.html` עצמאי לחלוטין (הספריות והנתונים מוטמעים). אפשר פשוט לפתוח אותו:

```bash
open index.html
```

או דרך שרת מקומי (מומלץ):

```bash
npm run serve      # → http://localhost:8777
```

הבקשות החיצוניות היחידות בזמן ריצה: אריחי מפת OpenStreetMap (עם נפילה אוטומטית
לקו גבול השיפוט אם אין רשת), וגופני Google (עם fallback לגופן מערכת).

## בנייה מחדש מהנתונים הגולמיים

```bash
npm install                    # xlsx (ל-build.js בלבד)
npm run build                  # build.js -> dist/site_data.json ; assemble.js -> index.html
```

`build.py` הוא מימוש־מראה ב-pandas (ראה DECISIONS.md §2):

```bash
python3 -m pip install pandas openpyxl && python3 build.py
```

## מבנה

| קובץ | תפקיד |
|---|---|
| `index.html` | **התוצר** — האתר, קובץ יחיד עובד (~940KB) |
| `index.template.html` | תבנית: כל ה־CSS/JS של האפליקציה + placeholders להטמעה |
| `build.js` | קורא XLSX + JSON גולמי → `dist/site_data.json` + מריץ 18 בדיקות קבלה |
| `build.py` | מימוש־מראה של `build.js` ב-pandas |
| `assemble.js` | מטמיע `vendor/*` + `dist/site_data.json` → `index.html` |
| `serve.js` | שרת סטטי קטן לבדיקה מקומית |
| `data/street_geo_final.json` | 107 רחובות ממופים (OSM) |
| `data/ariel_boundary.geojson` | גבול שיפוט (OSM relation 10011903) |
| `vendor/` | Leaflet 1.9.4, markercluster 1.5.3, leaflet.heat 0.2.0, Chart.js 4.4.4 |
| `ACCEPTANCE.md` | 18 בדיקות קבלה, מספר מול מספר — **18/18 עוברות** |
| `DECISIONS.md` | כל החלטה שנאלצתי לקבל ולא נכתבה באפיון |

## קלט

מונח ב־`~/Downloads/`: `tickets (99).xlsx` (12,309), `subjects (21).xlsx` (165),
`רחובות אריאל.xlsx` (132). לשינוי הנתיב — ראש `build.js` / `build.py`.

---

גיאומטריה: © OpenStreetMap contributors, ODbL. relation 10011903 (admin_level=8).
