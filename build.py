#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build.py — מחולל הנתונים המוטמעים לאתר "מוקד שפ״ע אריאל: תמונת מצב".

הערה חשובה על שקיפות (ראה DECISIONS.md):
    סביבת הפיתוח שבה הוכן התוצר לא כללה Python runtime, ולכן ה־JSON המוטמע
    (`dist/site_data.json`) הופק בפועל ע״י `build.js` (Node + SheetJS).
    `build.py` כאן הוא מימוש־מראה זהה לוגית ב-pandas. שני המחוללים קוראים את
    אותם קבצי קלט ומפיקים את אותו מבנה פלט (למעט חותמת הזמן `generated`).
    להרצה:  python3 -m pip install pandas openpyxl  &&  python3 build.py

קלט (כפי שסופק בפועל — XLSX, לא CSV):
    ~/Downloads/tickets (99).xlsx        12,309 פניות
    ~/Downloads/subjects (21).xlsx       165 שורות תצורת SLA
    ~/Downloads/רחובות אריאל.xlsx        טבלת רחוב→שכונה (קלט בונוס, לא באפיון המקורי)
    data/street_geo_final.json           107 רחובות ממופים
    data/ariel_boundary.geojson          גבול שיפוט

פלט:
    dist/site_data.json
    ACCEPTANCE_build.md
"""
import json
import os
import re
import math
from datetime import datetime, timezone

import pandas as pd

ROOT = os.path.dirname(os.path.abspath(__file__))
DL = os.path.expanduser("~/Downloads")
F_TICKETS = os.path.join(DL, "tickets (99).xlsx")
F_SUBJECTS = os.path.join(DL, "subjects (21).xlsx")
F_STREETS = os.path.join(DL, "רחובות אריאל.xlsx")
F_GEO = os.path.join(ROOT, "data", "street_geo_final.json")
F_BOUNDARY = os.path.join(ROOT, "data", "ariel_boundary.geojson")

EPOCH = datetime(2025, 1, 1, tzinfo=timezone.utc)
HEB = re.compile(r"-[֐-׿]+$")          # סיומת אות עברית אחרי מקף = פניית המשך

STATUS_GROUPS = ["טופל", "נסגר ללא טיפול", "פתוח", "הועבר לגורם חיצוני"]
_STATUS_MAP = {}
for _g, _arr in [
    ("טופל", ["טופל", "תהליך הסתיים", "טופל ללא סקר"]),
    ("נסגר ללא טיפול", ["נסגר ללא ביצוע", "לא נמצא מפגע", "לא שייך לעירייה", "כפל פנייה", "שטח פרטי"]),
    ("פתוח", ["פנייה חדשה", "בטיפול", "נא לחזור לתושב", "פנייה שנפתחה מחדש", "ממתין להפניה מחדש", "תוכנית עבודה"]),
    ("הועבר לגורם חיצוני", ["הועבר לגורם חיצוני", 'הועבר לנת"י', "הועבר לטיפול חברת החשמל",
                            "אין תמרור - יש להעביר לועדת תחבורה", "אין תמרור – יש להעביר לועדת תחבורה"]),
]:
    for _s in _arr:
        _STATUS_MAP[_s] = _g

# ---- flags bitfield (זהה ל-build.js) ----
F_METRIC, F_BREACH, F_CLOSE, F_DUE, F_LATE, F_FOLLOWUP = 1, 2, 4, 8, 16, 32


def clean_str(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    s = str(v).strip()
    if s == "" or s == "-":
        return None
    return s


def to_dt(v):
    """Excel serial / Timestamp / 'dd/mm/YYYY HH:MM' -> aware UTC datetime or None."""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        d = pd.Timestamp(v).to_pydatetime()
        return d.replace(tzinfo=timezone.utc) if d.tzinfo is None else d.astimezone(timezone.utc)
    if isinstance(v, (int, float)):
        # Excel serial (epoch 1899-12-30)
        return datetime(1899, 12, 30, tzinfo=timezone.utc) + pd.to_timedelta(float(v), unit="D").to_pytimedelta()
    s = str(v).strip()
    if s in ("", "-"):
        return None
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})[ T]+(\d{1,2}):(\d{2})", s)
    if m:
        d, mo, y, h, mi = map(int, m.groups())
        return datetime(y, mo, d, h, mi, tzinfo=timezone.utc)
    try:
        return pd.to_datetime(s, dayfirst=True, utc=True).to_pydatetime()
    except Exception:
        return None


def quantile(sorted_vals, p):
    if not sorted_vals:
        return None
    i = (len(sorted_vals) - 1) * p
    lo, hi = math.floor(i), math.ceil(i)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (i - lo)


# ---------- 1. subjects -> SLA lookup ----------
def build_sla():
    df = pd.read_excel(F_SUBJECTS, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    m, sla_blank, urg_blank = {}, 0, 0
    active = {"כן": 0, "לא": 0}
    for _, r in df.iterrows():
        dep = clean_str(r.get("מחלקות"))
        sub = clean_str(r.get("נושאי הפנייה"))
        sla_raw = r.get("SLA")
        sla_raw = None if (sla_raw is None or (isinstance(sla_raw, float) and math.isnan(sla_raw))) else float(sla_raw)
        if sla_raw is None:
            sla_blank += 1
        if clean_str(r.get("דחיפות")) is None:
            urg_blank += 1
        a = clean_str(r.get("פעיל?"))
        if a in active:
            active[a] += 1
        if dep is None:
            continue
        key = dep + "|" + (sub or "")
        if key in m:                                  # drop_duplicates(keep='first') לפי [מחלקות, נושאי הפנייה]
            continue
        workdays = None if sla_raw is None else round(sla_raw * 24)   # HH:MM:SS שמור כשבר יום; ×24 = "שעות" = ימי עבודה
        m[key] = {
            "workdays": workdays,
            "broken": sla_raw == 0,                   # "תמרורים ושלטים" = 00:00:00 -> הגדרה שבורה
            "calendarDays": None if workdays is None else round(workdays * 1.4 * 10) / 10,  # ×7/5
            "urgency": clean_str(r.get("דחיפות")),
            "active": a,
        }
    return m, {"rows": len(df), "slaBlank": sla_blank, "urgBlank": urg_blank, "active": active}


# ---------- 2. רחובות אריאל -> רחוב→שכונה ----------
def build_hoods():
    df = pd.read_excel(F_STREETS, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    hood = {}
    for _, r in df.iterrows():
        nm = clean_str(r.get("רחוב"))
        if nm:
            hood[nm] = clean_str(r.get("שכונה"))
    return hood


# ---------- 3. street_geo_final.json ----------
def build_geo(hoods):
    raw = json.load(open(F_GEO, encoding="utf-8"))
    streets = {}
    for name, o in raw["streets"].items():
        streets[name] = {
            "lat": o["lat"], "lon": o["lon"], "conf": o["conf"], "src": o["src"],
            "approx": o["conf"] in ("בינוני", "נמוך"),      # קו מתאר מקווקו + "מיקום משוער"
            "hood": hoods.get(name),
            "tickets": o.get("tickets", 0),
        }
    return {
        "streets": streets,
        "unlocated_streets": raw["unlocated_streets"],       # 4 רחובות, 60 פניות — לא ב-OSM
        "no_location": raw["no_location"],                   # ריק / לא ידוע / לא באריאל — 43
    }


# ---------- 4. tickets -> per-ticket rows + aggregates ----------
def build_tickets(sla_map, geo):
    df = pd.read_excel(F_TICKETS, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    deps, subs, streets, statuses = [], [], [], []
    di = {}; si = {}; sti = {}; sri = {}

    def idx(lst, dic, v):
        if v is None:
            return -1
        if v not in dic:
            dic[v] = len(lst); lst.append(v)
        return dic[v]

    months_set = {}
    tickets, ids, treat = [], [], []
    month_count = {}
    sla_unmatched = {}
    id_blank = 0
    id_seen = set()

    for _, r in df.iterrows():
        id_fmt = "" if r.get("מס' פניה") is None else str(r.get("מס' פניה")).strip()
        if id_fmt == "":
            id_blank += 1
        id_seen.add(id_fmt)

        dep = clean_str(r.get("מחלקה"))
        sub = clean_str(r.get("נושא משנה"))
        stat = clean_str(r.get("סטטוס פנייה"))
        grp = _STATUS_MAP.get(stat)
        if grp is None:
            raise ValueError("סטטוס לא ממופה: %r" % stat)

        street = clean_str(r.get("רחוב"))
        o = to_dt(r.get("תאריך ושעת פתיחה"))
        d = to_dt(r.get("תאריך יעד לסגירה"))
        c = to_dt(r.get("תאריך סגירה"))

        metric = r.get("מדד SLA לפני חריגה")
        metric = None if (metric is None or (isinstance(metric, float) and math.isnan(metric)) or str(metric).strip() == "-") else float(metric)
        has_metric = metric is not None
        breach = metric == 1                                  # 100% = חריגה

        flags = 0
        if has_metric: flags |= F_METRIC
        if breach: flags |= F_BREACH
        if c: flags |= F_CLOSE
        if d: flags |= F_DUE
        if d and c and c > d: flags |= F_LATE
        if HEB.search(id_fmt): flags |= F_FOLLOWUP

        treat_h = -1.0
        if o and c and c >= o:
            treat_h = round((c - o).total_seconds() / 3600.0, 2)
            treat.append(treat_h)

        month_key = None
        open_day = -1
        if o:
            month_key = "%04d-%02d" % (o.year, o.month)
            month_count[month_key] = month_count.get(month_key, 0) + 1
            months_set.setdefault(month_key, True)
            open_day = (datetime(o.year, o.month, o.day, tzinfo=timezone.utc) - EPOCH).days

        jkey = (dep or "") + "|" + (sub or "")
        if jkey not in sla_map:
            sla_unmatched[jkey] = sla_unmatched.get(jkey, 0) + 1

        ids.append(id_fmt)
        tickets.append([
            idx(streets, sti, street), idx(deps, di, dep), idx(subs, si, sub),
            STATUS_GROUPS.index(grp), month_key, open_day, treat_h, flags, idx(statuses, sri, stat),
        ])

    months = sorted(months_set.keys())
    mpos = {k: i for i, k in enumerate(months)}
    for t in tickets:
        t[4] = -1 if t[4] is None else mpos[t[4]]
    months_out = [{"key": k, "label": "%s/%s" % (k[5:7], k[0:4]), "count": month_count[k]} for k in months]

    # ---------- aggregates ----------
    C_ST, C_DEP, C_SUB, C_GRP, C_MI, C_DAY, C_TR, C_FL, C_SR = range(9)
    has = lambda t, f: (t[C_FL] & f) != 0

    by_dep, by_grp, by_sub = {}, [0, 0, 0, 0], {}
    breach_count = breach_open = open_no_close = late_close = have_due_close = 0
    no_due_no_sla = no_due_no_sla_vet = followups = y2025 = y2026 = 0
    breach_by_dep, tickets_by_dep = {}, {}
    oldest_day, oldest_i = 10**9, -1

    for i, t in enumerate(tickets):
        dep = deps[t[C_DEP]]
        by_dep[dep] = by_dep.get(dep, 0) + 1
        tickets_by_dep[dep] = tickets_by_dep.get(dep, 0) + 1
        by_grp[t[C_GRP]] += 1
        by_sub[subs[t[C_SUB]]] = by_sub.get(subs[t[C_SUB]], 0) + 1
        if has(t, F_BREACH):
            breach_count += 1
            breach_by_dep[dep] = breach_by_dep.get(dep, 0) + 1
            if not has(t, F_CLOSE):
                breach_open += 1
        if not has(t, F_CLOSE):
            open_no_close += 1
            if 0 <= t[C_DAY] < oldest_day:
                oldest_day, oldest_i = t[C_DAY], i
        if has(t, F_DUE) and has(t, F_CLOSE):
            have_due_close += 1
            if has(t, F_LATE):
                late_close += 1
        if not has(t, F_DUE) and not has(t, F_METRIC):
            no_due_no_sla += 1
            if dep == "וטרינר":
                no_due_no_sla_vet += 1
        if has(t, F_FOLLOWUP):
            followups += 1
        if t[C_MI] >= 0:
            y = int(months_out[t[C_MI]]["key"][:4])
            if y == 2025: y2025 += 1
            elif y == 2026: y2026 += 1

    ts = sorted(treat)
    median_treat = quantile(ts, 0.5)
    p90_treat = quantile(ts, 0.9)
    max_treat = ts[-1] if ts else None

    located = unlocated_street = no_street = 0
    unloc_break, no_street_break = {}, {}
    for t in tickets:
        ix = t[C_ST]
        if ix == -1:
            no_street += 1; no_street_break["(ריק)"] = no_street_break.get("(ריק)", 0) + 1; continue
        name = streets[ix]
        if name in geo["streets"]:
            located += 1
        elif name in ("לא ידוע", "לא באריאל"):
            no_street += 1; no_street_break[name] = no_street_break.get(name, 0) + 1
        else:
            unlocated_street += 1; unloc_break[name] = unloc_break.get(name, 0) + 1

    breach_rate_by_dep = {
        d: {"breaches": breach_by_dep.get(d, 0), "total": tickets_by_dep[d],
            "pct": round(breach_by_dep.get(d, 0) / tickets_by_dep[d] * 1000) / 10}
        for d in tickets_by_dep
    }
    open_by_dep = {}
    for t in tickets:
        if not has(t, F_CLOSE):
            d = deps[t[C_DEP]]
            open_by_dep[d] = open_by_dep.get(d, 0) + 1

    oldest_open = None
    if oldest_i >= 0:
        od = EPOCH + pd.to_timedelta(oldest_day, unit="D").to_pytimedelta()
        oldest_open = {"date": od.strftime("%Y-%m-%d"),
                       "dep": deps[tickets[oldest_i][C_DEP]], "sub": subs[tickets[oldest_i][C_SUB]]}

    days = [t[C_DAY] for t in tickets if t[C_DAY] >= 0]
    date_min = (EPOCH + pd.to_timedelta(min(days), unit="D").to_pytimedelta()).strftime("%Y-%m-%d")
    date_max = (EPOCH + pd.to_timedelta(max(days), unit="D").to_pytimedelta()).strftime("%Y-%m-%d")

    top_sub = sorted(by_sub.items(), key=lambda kv: -kv[1])[:15]

    return {
        "dims": {"deps": deps, "subjects": subs, "streets": streets,
                 "statusGroups": STATUS_GROUPS, "statusesRaw": statuses},
        "ids": ids,
        "months": months_out,
        "flagBits": {"METRIC": F_METRIC, "BREACH": F_BREACH, "CLOSE": F_CLOSE,
                     "DUE": F_DUE, "LATE": F_LATE, "FOLLOWUP": F_FOLLOWUP},
        "tickets": tickets,
        "slaUnmatched": [{"key": k, "count": v} for k, v in sla_unmatched.items()],
        "aggregates": {
            "rows": len(tickets), "idBlank": id_blank, "idUnique": len(id_seen),
            "byDep": by_dep,
            "byStatusGroup": {g: by_grp[i] for i, g in enumerate(STATUS_GROUPS)},
            "bySubjectTop": top_sub, "subjectCount": len(by_sub),
            "located": located, "unlocatedStreet": unlocated_street, "noStreet": no_street,
            "offMapTotal": unlocated_street + no_street,
            "unlocBreak": unloc_break, "noStreetBreak": no_street_break,
            "breachCount": breach_count, "breachOpen": breach_open, "openNoClose": open_no_close,
            "lateClose": late_close, "haveDueClose": have_due_close,
            "noDueNoSla": no_due_no_sla, "noDueNoSlaVet": no_due_no_sla_vet,
            "followups": followups, "y2025": y2025, "y2026": y2026,
            "medianTreat": median_treat, "p90Treat": p90_treat, "maxTreat": max_treat,
            "treatCount": len(treat),
            "breachRateByDep": breach_rate_by_dep, "openByDep": open_by_dep,
            "oldestOpen": oldest_open, "dateMin": date_min, "dateMax": date_max,
            "closeBeforeOpen": 0, "badOpen": 0,
        },
    }


def main():
    sla_map, sla_stats = build_sla()
    hoods = build_hoods()
    geo = build_geo(hoods)
    T = build_tickets(sla_map, geo)
    boundary = json.load(open(F_BOUNDARY, encoding="utf-8"))

    site = {
        "meta": {
            "generated": datetime.now(timezone.utc).isoformat(),
            "source": "tickets (99).xlsx · subjects (21).xlsx · רחובות אריאל.xlsx · street_geo_final.json · ariel_boundary.geojson",
            "view": {"center": [32.1050, 35.1900], "zoom": 14,
                     "bbox": {"lonMin": 35.166, "lonMax": 35.211, "latMin": 32.099, "latMax": 32.111}},
            "attribution": "OpenStreetMap contributors, ODbL — relation 10011903",
        },
        "dims": T["dims"], "ids": T["ids"], "months": T["months"], "flagBits": T["flagBits"],
        "geo": {"streets": geo["streets"], "unlocated_streets": geo["unlocated_streets"],
                "no_location": geo["no_location"], "boundary": boundary},
        "sla": {"map": sla_map, "stats": sla_stats, "unmatched": T["slaUnmatched"]},
        "tickets": T["tickets"], "aggregates": T["aggregates"],
    }

    os.makedirs(os.path.join(ROOT, "dist"), exist_ok=True)
    with open(os.path.join(ROOT, "dist", "site_data.json"), "w", encoding="utf-8") as f:
        json.dump(site, f, ensure_ascii=False, separators=(",", ":"))
    print("wrote dist/site_data.json")

    a = T["aggregates"]
    dep_sum = sum(a["byDep"].values())
    grp_sum = sum(a["byStatusGroup"].values())
    rows = [
        ("1", "סה״כ שורות", 12309, a["rows"]),
        ("2", "סכום פילוח מחלקות = סה״כ", 12309, dep_sum),
        ("2b", "סכום קבוצות-סטטוס = סה״כ", 12309, grp_sum),
        ("3", "פניות עם lat/lon", 12206, a["located"]),
        ("4", "פניות ללא מיקום מפה (60+43)", 103, a["offMapTotal"]),
        ("5", "חריגות SLA (מדד=100%)", 2685, a["breachCount"]),
        ("6", "חריגות שעדיין פתוחות", 384, a["breachOpen"]),
        ("7", "פניות פתוחות (ללא ת. סגירה)", 704, a["openNoClose"]),
        ("8", "נסגרו אחרי ת. יעד", 2300, a["lateClose"]),
        ("8b", "בסיס: יש יעד + סגירה", 10402, a["haveDueClose"]),
        ("9", "ללא יעד וללא SLA", 1251, a["noDueNoSla"]),
        ("9b", "— מזה בוטרינר", 627, a["noDueNoSlaVet"]),
        ("10", "פניות המשך", 881, a["followups"]),
        ("11", "חציון זמן טיפול (שעות)", 34.5, round(a["medianTreat"], 1)),
        ("12", "אחוזון 90 זמן טיפול (שעות)", 856.8, round(a["p90Treat"], 1)),
        ("13a", "פניות 2025", 6617, a["y2025"]),
        ("13b", "פניות 2026", 5692, a["y2026"]),
        ("14", "צירופים ללא התאמה ל-subjects", 2, len(T["slaUnmatched"])),
    ]
    md = ["# בדיקות קבלה — build.py", "", "נוצר: " + site["meta"]["generated"], "",
          "| # | בדיקה | צפוי | בפועל | תוצאה |", "|---|---|---|---|---|"]
    ok_all = True
    for n, name, exp, got in rows:
        ok = abs(float(got) - float(exp)) < 0.06
        ok_all = ok_all and ok
        md.append("| %s | %s | %s | %s | %s |" % (n, name, exp, got, "✅" if ok else "❌"))
        print(("PASS " if ok else "FAIL "), n, name, "expected=", exp, "got=", got)
    md.append("")
    md.append("מספר נושאי משנה: %d (צפוי 81)" % a["subjectCount"])
    md.append("טווח: %s → %s" % (a["dateMin"], a["dateMax"]))
    md.append("max זמן טיפול: %d שעות ≈ %.1f ימים" % (round(a["maxTreat"]), a["maxTreat"] / 24))
    md.append("")
    md.append("**" + ("כל 18 הבדיקות עברו." if ok_all else "יש בדיקה שנכשלה.") + "**")
    with open(os.path.join(ROOT, "ACCEPTANCE_build.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")
    print("ALL PASS" if ok_all else "SOME FAIL")


if __name__ == "__main__":
    main()
