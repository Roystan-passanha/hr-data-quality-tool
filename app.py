"""HR Data Assistant: local web app for hr_tool.py. Start with:  streamlit run app.py"""
import io
import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import hr_tool
from hr_tool import COLS as C

st.set_page_config(page_title="HR Data Assistant", layout="wide")


class _NoAudit:
    """Placeholder for the audit log used in the full version."""
    @staticmethod
    def log(*args, **kwargs):
        pass


auth = _NoAudit()
st.title("HR Data Assistant")
st.caption("Runs locally. Employee data never leaves this computer. Try it with hr_sample_messy.xlsx from make_sample.py.")


def to_excel(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "out.xlsx"
        hr_tool.write_excel(p, sheets)
        buf.write(p.read_bytes())
    return buf.getvalue()


# ---------- Step 1: upload and clean ----------
st.header("1. Upload and clean")
up = st.file_uploader("Employee Excel file (.xlsx)", type=["xlsx"])
if up and st.button("Clean data", type="primary"):
    with tempfile.TemporaryDirectory() as d:
        src, out = Path(d) / "in.xlsx", Path(d) / "clean.xlsx"
        src.write_bytes(up.getvalue())
        try:
            hr_tool.clean(src, out)
        except SystemExit as e:
            st.error(str(e))
            st.stop()
        st.session_state.clean = pd.read_excel(out, sheet_name="Clean_Data")
        st.session_state.issues = pd.read_excel(out, sheet_name="Issues")
        st.session_state.clean_file = out.read_bytes()
        st.session_state.source_name = up.name
        _c, _i = st.session_state.clean, st.session_state.issues
        auth.log("Uploaded and cleaned file", f"file={up.name} rows={len(_c)} issues={len(_i)} "
                 f"needs_review={int((_c['Status'] != 'OK').sum())}")

if "clean" in st.session_state:
    df, iss = st.session_state.clean, st.session_state.issues
    a, b, c, d = st.columns(4)
    a.metric("Employees", len(df))
    b.metric("OK", int((df["Status"] == "OK").sum()))
    c.metric("Need review", int((df["Status"] != "OK").sum()))
    d.metric("High severity issues", int((iss["Severity"] == "High").sum()) if len(iss) else 0)

    with st.expander("Issues found", expanded=False):
        if len(iss):
            st.dataframe(iss.groupby(["Severity", "Problem"]).size().reset_index(name="Count")
                         .sort_values("Count", ascending=False), hide_index=True, use_container_width=True)
            sev = st.multiselect("Severity", ["High", "Medium", "Low"], default=["High"])
            st.dataframe(iss[iss["Severity"].isin(sev)], hide_index=True, use_container_width=True)
    st.download_button("Download cleaned file + issues report", st.session_state.clean_file,
                       file_name="hr_clean.xlsx", on_click=auth.log,
                       args=("Downloaded cleaned file", f"source={st.session_state.source_name}"))

    # ---------- Step 2: requests ----------
    st.header("2. Request data")
    st.write("Type a request in plain words, or use the filters below.")
    st.caption("Examples: *UAE visa expiring in 60 days* · *passport expiring 180 days Finance* · "
               "*Riyadh Indian* · *expired visas KSA* · *EMP00012, EMP00450* · *issues in Dubai*")
    text = st.text_input("Request", placeholder="e.g. UAE staff whose visa expires in the next 60 days")

    def opts(col):
        return sorted(df[col].dropna().astype(str).unique())

    with st.expander("Filters", expanded=not text):
        f1, f2, f3 = st.columns(3)
        countries = f1.multiselect("Country", opts(C["country"]))
        locations = f2.multiselect("Location", opts(C["location"]))
        depts = f3.multiselect("Department", opts(C["dept"]))
        f4, f5, f6 = st.columns(3)
        nats = f4.multiselect("Nationality", opts(C["nationality"]))
        vstat = f5.multiselect("Visa status", opts(C["visa_status"]))
        only_issues = f6.checkbox("Only records with issues")
        f7, f8 = st.columns(2)
        v_days = f7.number_input("Visa expires within N days (0 = ignore)", 0, 3650, 0)
        p_days = f8.number_input("Passport expires within N days (0 = ignore)", 0, 3650, 0)

    m = pd.Series(True, index=df.index)
    understood = []

    # --- plain-text parsing (simple rules, no AI) ---
    if text:
        t = text.lower()

        def match_values(col, label):
            hits = [v for v in opts(col) if re.search(rf"\b{re.escape(v.lower())}\b", t)]
            if hits:
                understood.append(f"{label}: {', '.join(hits)}")
                return df[col].astype(str).isin(hits)
            return None

        aliases = {"uae": "UAE", "emirates": "UAE", "dubai": None, "saudi": "KSA", "ksa": "KSA"}
        for k, v in aliases.items():
            if v and re.search(rf"\b{k}\b", t) and v in opts(C["country"]) and v.lower() not in t:
                t += f" {v.lower()}"
        for col, label in [(C["country"], "Country"), (C["location"], "Location"),
                           (C["dept"], "Department"), (C["nationality"], "Nationality")]:
            r = match_values(col, label)
            if r is not None:
                m &= r
        ids = re.findall(r"[a-z]{2,4}\d{3,}", t)
        if ids:
            ids = [i.upper() for i in ids]
            m &= df[C["id"]].isin(ids)
            understood.append(f"Employee IDs: {', '.join(ids)}")
        num = re.search(r"(\d+)\s*(day|days|d\b|month|months)", t)
        n = None
        if num:
            n = int(num.group(1)) * (30 if num.group(2).startswith("month") else 1)
        if "visa" in t and n:
            m &= df["Visa Days Left"].between(0, n); understood.append(f"Visa expiring within {n} days")
        elif "passport" in t and n:
            m &= df["Passport Days Left"].between(0, n); understood.append(f"Passport expiring within {n} days")
        elif n:
            m &= df["Visa Days Left"].between(0, n) | df["Passport Days Left"].between(0, n)
            understood.append(f"Visa or passport expiring within {n} days")
        if "expired" in t:
            if "passport" in t:
                m &= df["Passport Days Left"] < 0; understood.append("Passport already expired")
            else:
                m &= df["Visa Days Left"] < 0; understood.append("Visa already expired")
        if "missing passport" in t:
            m &= df[C["passport"]].isna(); understood.append("Missing passport number")
        if re.search(r"\b(issue|issues|problem|problems|error|errors)\b", t):
            m &= df["Status"] != "OK"; understood.append("Only records with issues")
        for s in opts(C["visa_status"]):
            if re.search(rf"\b{re.escape(s.lower())}\b", t) and s.lower() != "expired":
                m &= df[C["visa_status"]] == s; understood.append(f"Visa status: {s}")

    # --- filter widgets ---
    for vals, col, label in [(countries, C["country"], "Country"), (locations, C["location"], "Location"),
                             (depts, C["dept"], "Department"), (nats, C["nationality"], "Nationality"),
                             (vstat, C["visa_status"], "Visa status")]:
        if vals:
            m &= df[col].astype(str).isin(vals); understood.append(f"{label}: {', '.join(vals)}")
    if only_issues:
        m &= df["Status"] != "OK"; understood.append("Only records with issues")
    if v_days:
        m &= df["Visa Days Left"].between(0, v_days); understood.append(f"Visa expiring within {v_days} days")
    if p_days:
        m &= df["Passport Days Left"].between(0, p_days); understood.append(f"Passport expiring within {p_days} days")

    if text and not understood:
        st.warning("I couldn't match anything in that request. Try a country, city, department, "
                   "nationality, employee ID, or 'visa/passport expiring in N days'.")
    elif understood:
        st.info("Understood as: " + " | ".join(understood))
        res = df[m]
        default_cols = [C["id"], C["name"], C["dept"], C["country"], C["location"], C["nationality"],
                        C["passport"], C["pp_expiry"], "Passport Days Left", C["visa_status"],
                        C["visa_expiry"], "Visa Days Left", "Status"]
        cols = st.multiselect("Columns in output", list(df.columns), default=default_cols)
        st.success(f"{len(res)} matching employees")
        q_desc = f"request='{text}' | " + " | ".join(understood) + f" | results={len(res)}"
        if st.session_state.get("last_q") != q_desc:
            auth.log("Ran request", q_desc)
            st.session_state.last_q = q_desc
        st.dataframe(res[cols], hide_index=True, use_container_width=True)
        name = re.sub(r"[^a-z0-9]+", "_", (text or "filtered_request").lower()).strip("_")[:50] or "request"
        st.download_button("Download as Excel", to_excel({"Result": res[cols]}),
                           file_name=f"{name}.xlsx", type="primary", on_click=auth.log,
                           args=("Downloaded request result", f"file={name}.xlsx rows={len(res)} | " + " | ".join(understood)))
else:
    st.info("Upload an Excel file and click **Clean data** to start.")
