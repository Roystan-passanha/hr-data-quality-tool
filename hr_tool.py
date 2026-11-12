"""HR data cleaning and query tool.

Usage:
  python hr_tool.py clean  hr_sample_messy.xlsx            -> hr_clean.xlsx (Clean_Data, Issues, Summary)
  python hr_tool.py query  hr_clean.xlsx --country UAE --visa-expiring 60
  python hr_tool.py query  hr_clean.xlsx --issues-only --out issues_for_hr.xlsx

Edit COLS below to match the real file's headers. Nothing else should need changing.
"""
import argparse
import re
import sys
from datetime import date

import pandas as pd

# ---- Map your real column names here -------------------------------------------------
COLS = {
    "id": "Employee ID", "name": "Full Name", "dept": "Department",
    "country": "Work Country", "location": "Work Location", "nationality": "Nationality",
    "email": "Email", "passport": "Passport Number",
    "pp_issue": "Passport Issue Date", "pp_expiry": "Passport Expiry Date",
    "visa_status": "Visa Status", "visa_expiry": "Visa Expiry Date",
}
COUNTRY_MAP = {"uae": "UAE", "u.a.e": "UAE", "united arab emirates": "UAE", "ksa": "KSA",
               "saudi arabia": "KSA", "qatar": "Qatar", "india": "India", "uk": "UK",
               "united kingdom": "UK", "usa": "USA", "us": "USA"}
VISA_MAP = {"valid": "Valid", "active": "Valid", "under process": "Under Process",
            "in progress": "Under Process", "processing": "Under Process",
            "cancelled": "Cancelled", "canceled": "Cancelled", "expired": "Expired"}
VISA_FREE_NATIONALITY = {"UAE": "Emirati"}  # nationals who shouldn't hold a work visa
PASSPORT_MIN_VALIDITY_DAYS = 180            # many GCC visas need 6 months passport validity
EXPIRY_WARNING_DAYS = 90
DATE_FORMATS = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d %b %Y", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"]
# ---------------------------------------------------------------------------------------


def parse_date(v):
    """Return a Timestamp, or NaT if the value can't be a real date. Day-first (UAE style)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, (pd.Timestamp, date)):
        return pd.Timestamp(v)
    s = str(v).strip()
    for f in DATE_FORMATS:
        try:
            return pd.to_datetime(s, format=f)
        except ValueError:
            continue
    return pd.NaT


def clean(path, out):
    raw = pd.read_excel(path, dtype=object)
    missing = [c for c in COLS.values() if c not in raw.columns]
    if missing:
        sys.exit(f"Columns not found in file: {missing}. Update COLS in hr_tool.py.")
    df = raw.copy()
    df.insert(0, "Source Row", df.index + 2)  # Excel row number, so HR can find it
    issues = []

    def flag(mask, field, problem, severity="Medium"):
        for r in df.index[mask]:
            issues.append({"Source Row": df.at[r, "Source Row"], "Employee ID": df.at[r, COLS["id"]],
                           "Name": df.at[r, COLS["name"]], "Field": field, "Problem": problem,
                           "Original Value": raw.at[r, field] if field in raw.columns else "",
                           "Severity": severity})

    C = COLS
    # Text tidy
    for k in ("name", "dept", "location", "nationality", "email", "passport"):
        df[C[k]] = df[C[k]].astype("string").str.strip()
    df[C["name"]] = df[C["name"]].str.title()
    df[C["passport"]] = df[C["passport"]].str.upper().str.replace(r"\s+", "", regex=True)
    df[C["email"]] = df[C["email"]].str.lower()

    # Standardise categories
    c_raw = df[C["country"]].astype("string").str.strip().str.lower()
    df[C["country"]] = c_raw.map(COUNTRY_MAP)
    flag(df[C["country"]].isna(), C["country"], "Unrecognised country")
    v_raw = df[C["visa_status"]].astype("string").str.strip().str.lower()
    df[C["visa_status"]] = v_raw.map(VISA_MAP)
    flag(df[C["visa_status"]].isna(), C["visa_status"], "Unrecognised visa status")

    # Dates
    for k in ("pp_issue", "pp_expiry", "visa_expiry"):
        parsed = raw[C[k]].apply(parse_date)
        bad = parsed.isna() & raw[C[k]].notna()
        flag(bad, C[k], "Not a valid date", "High")
        flag(raw[C[k]].isna(), C[k], "Missing", "High")
        df[C[k]] = parsed

    today = pd.Timestamp(date.today())
    pp_days = (df[C["pp_expiry"]] - today).dt.days
    v_days = (df[C["visa_expiry"]] - today).dt.days
    df["Passport Days Left"] = pp_days
    df["Visa Days Left"] = v_days

    flag(df[C["pp_expiry"]] <= df[C["pp_issue"]], C["pp_expiry"], "Expiry not after issue date", "High")
    flag(pp_days < 0, C["pp_expiry"], "Passport expired", "High")
    flag(pp_days.between(0, PASSPORT_MIN_VALIDITY_DAYS), C["pp_expiry"],
         f"Passport valid < {PASSPORT_MIN_VALIDITY_DAYS} days")
    flag((df[C["visa_status"]] == "Valid") & (v_days < 0), C["visa_status"],
         "Status 'Valid' but visa expiry date has passed", "High")
    flag((df[C["visa_status"]] == "Valid") & v_days.between(0, EXPIRY_WARNING_DAYS), C["visa_expiry"],
         f"Visa expires within {EXPIRY_WARNING_DAYS} days")
    flag(df[C["visa_expiry"]] > df[C["pp_expiry"]], C["visa_expiry"], "Visa outlasts passport", "Low")
    for ctry, nat in VISA_FREE_NATIONALITY.items():
        flag((df[C["country"]] == ctry) & (df[C["nationality"]] == nat) & df[C["visa_status"]].notna(),
             C["visa_status"], f"{nat} national recorded with a visa", "Low")

    # Identity fields
    flag(df[C["passport"]].isna(), C["passport"], "Missing passport number", "High")
    flag(df[C["passport"]].notna() & ~df[C["passport"]].fillna("").str.match(r"^[A-Z0-9]{6,9}$"),
         C["passport"], "Passport number format looks wrong")
    flag(~df[C["email"]].fillna("").str.match(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$"), C["email"], "Invalid email")

    # Duplicates
    full_dup = df.drop(columns="Source Row").duplicated(keep="first")
    flag(full_dup, C["id"], "Exact duplicate row (removed)", "Low")
    df = df[~full_dup]
    id_dup = df[C["id"]].duplicated(keep=False)
    flag(id_dup, C["id"], "Same Employee ID used for different records", "High")
    pp_dup = df[C["passport"]].notna() & df[C["passport"]].duplicated(keep=False)
    flag(pp_dup, C["passport"], "Passport number shared by multiple records", "High")

    iss = pd.DataFrame(issues)
    counts = iss.groupby("Source Row").size() if len(iss) else pd.Series(dtype=int)
    df["Issue Count"] = df["Source Row"].map(counts).fillna(0).astype(int)
    df["Status"] = df["Issue Count"].map(lambda n: "OK" if n == 0 else "Needs review")

    summary = pd.DataFrame([
        ("Run date", date.today().isoformat()), ("Source file", str(path)),
        ("Rows in source", len(raw)), ("Rows after removing exact duplicates", len(df)),
        ("Rows OK", int((df["Status"] == "OK").sum())),
        ("Rows needing review", int((df["Status"] != "OK").sum())), ("Total issues", len(iss)),
    ], columns=["Metric", "Value"])
    by_problem = (iss.groupby(["Severity", "Problem"]).size().reset_index(name="Count")
                  .sort_values("Count", ascending=False)) if len(iss) else pd.DataFrame()

    write_excel(out, {"Clean_Data": df, "Issues": iss.sort_values(["Severity", "Source Row"]) if len(iss) else iss,
                      "Summary": summary}, extra=("Summary", by_problem))
    print(f"Cleaned {len(raw)} rows -> {out}. {len(iss)} issues on "
          f"{int((df['Status'] != 'OK').sum())} employees.")


def write_excel(out, sheets, extra=None):
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="DD-MMM-YYYY") as w:
        for name, d in sheets.items():
            d.to_excel(w, sheet_name=name, index=False)
            ws = w.sheets[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 45)
        if extra is not None and len(extra[1]):
            extra[1].to_excel(w, sheet_name=extra[0], index=False, startrow=len(sheets[extra[0]]) + 3)


def query(path, a):
    df = pd.read_excel(path, sheet_name="Clean_Data")
    C = COLS
    m = pd.Series(True, index=df.index)
    if a.country:     m &= df[C["country"]].str.lower() == a.country.lower()
    if a.location:    m &= df[C["location"]].str.lower() == a.location.lower()
    if a.dept:        m &= df[C["dept"]].str.lower() == a.dept.lower()
    if a.nationality: m &= df[C["nationality"]].str.lower() == a.nationality.lower()
    if a.visa_status: m &= df[C["visa_status"]].str.lower() == a.visa_status.lower()
    if a.id:          m &= df[C["id"]].isin([x.strip().upper() for x in a.id.split(",")])
    if a.name:        m &= df[C["name"]].str.contains(a.name, case=False, na=False)
    if a.visa_expiring is not None:
        m &= df["Visa Days Left"].between(0, a.visa_expiring)
    if a.passport_expiring is not None:
        m &= df["Passport Days Left"].between(0, a.passport_expiring)
    if a.issues_only: m &= df["Status"] != "OK"
    res = df[m]
    if a.cols:
        res = res[[c.strip() for c in a.cols.split(",")]]
    print(f"{len(res)} matching employees")
    if a.out:
        write_excel(a.out, {"Result": res})
        print(f"Saved to {a.out}")
    else:
        print(res.head(25).to_string(index=False))


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="HR data cleaning and query tool")
    sub = p.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("clean"); c.add_argument("file"); c.add_argument("--out", default="hr_clean.xlsx")
    q = sub.add_parser("query"); q.add_argument("file")
    for f in ("country", "location", "dept", "nationality", "visa-status", "id", "name", "cols", "out"):
        q.add_argument(f"--{f}")
    q.add_argument("--visa-expiring", type=int, help="visa expires within N days")
    q.add_argument("--passport-expiring", type=int, help="passport expires within N days")
    q.add_argument("--issues-only", action="store_true")
    a = p.parse_args()
    clean(a.file, a.out) if a.cmd == "clean" else query(a.file, a)
