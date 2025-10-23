# app.py
from datetime import datetime, date
from zoneinfo import ZoneInfo
import re
import os
import pandas as pd
import streamlit as st
from openpyxl.utils.datetime import from_excel

st.set_page_config(page_title="Power BI Enrollment (Cleaned)", layout="centered")
st.title("Power BI Enrollment (Cleaned)")
st.caption("Upload two Enrollment exports (.xlsx). This outputs a plain, unstyled file ready for Power BI (no totals, no helper columns).")

PB_FOLDER = r"C:\Users\Daniella.Thompson\OneDrive - hchsp.org\Power Bi Data"
os.makedirs(PB_FOLDER, exist_ok=True)

def coerce_to_dt(v):
    if pd.isna(v):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return from_excel(v)
        except Exception:
            return None
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                continue
    return None

def most_recent(series):
    dates, texts = [], []
    for v in series.dropna().unique():
        dt = coerce_to_dt(v)
        if dt:
            dates.append(dt)
        else:
            sv = str(v).strip()
            if sv:
                texts.append(sv)
    if dates:
        return max(dates)
    return texts[0] if texts else None

def normalize(s: str) -> str:
    s = s.lower()
    s = re.sub(r"[\s\-\–\—_:()]+", " ", s)
    return s.strip()

def find_cols(cols, keywords):
    out = []
    for c in cols:
        if not isinstance(c, str):
            continue
        n = normalize(c)
        if any(k in n for k in keywords):
            out.append(c)
    return out

def collapse_row_values(row, col_names):
    vals = []
    for c in col_names:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
            vals.append(row[c])
    if not vals:
        return None
    dts = [coerce_to_dt(v) for v in vals]
    dts = [d for d in dts if d]
    if dts:
        return max(dts)
    return str(vals[0]).strip()

def read_and_standardize(file):
    tmp = pd.read_excel(file, header=None, nrows=30)
    header_row = None
    for r in range(tmp.shape[0]):
        row_vals = tmp.iloc[r].astype(str).fillna("")
        if any("ST: Participant PID" in v for v in row_vals):
            header_row = r
            break
    if header_row is None:
        st.error("Couldn't find 'ST: Participant PID' in one of the files.")
        return None

    file.seek(0)
    df = pd.read_excel(file, header=header_row)
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]
    if "Participant PID" not in df.columns:
        st.error("A file is missing 'Participant PID'.")
        return None
    df = df.dropna(subset=["Participant PID"])
    return df

def clean_and_collapse(df):
    df = df.groupby("Participant PID", as_index=False).agg(most_recent)
    all_cols = list(df.columns)
    immun_cols = find_cols(all_cols, ["immun"])
    tb_cols = find_cols(all_cols, ["tb", "tuberc", "ppd"])
    lead_cols = find_cols(all_cols, ["lead", "pb"])
    scn_en_cols = find_cols(all_cols, ["special care needs english"])
    scn_es_cols = find_cols(all_cols, ["special care needs spanish"])
    scn_cols = scn_en_cols + scn_es_cols

    if immun_cols:
        df["Immunizations"] = df.apply(lambda r: collapse_row_values(r, immun_cols), axis=1)
        df.drop(columns=[c for c in immun_cols if c in df.columns], inplace=True)
    if tb_cols:
        df["TB Test"] = df.apply(lambda r: collapse_row_values(r, tb_cols), axis=1)
        df.drop(columns=[c for c in tb_cols if c in df.columns], inplace=True)
    if lead_cols:
        df["Lead Test"] = df.apply(lambda r: collapse_row_values(r, lead_cols), axis=1)
        df.drop(columns=[c for c in lead_cols if c in df.columns], inplace=True)
    if scn_cols:
        df["Child's Special Care Needs"] = df.apply(lambda r: collapse_row_values(r, scn_cols), axis=1)
        df.drop(columns=[c for c in scn_cols if c in df.columns], inplace=True)

    def to_iso_or_val(v):
        dt = coerce_to_dt(v)
        return dt.date().isoformat() if dt else (None if pd.isna(v) else v)

    for col in ["Immunizations", "TB Test", "Lead Test", "Child's Special Care Needs"]:
        if col in df.columns:
            df[col] = df[col].apply(to_iso_or_val)

    leading = [c for c in ["Participant PID", "Participant Name", "Center", "Campus", "School"] if c in df.columns]
    metrics = [c for c in ["Immunizations", "TB Test", "Lead Test", "Child's Special Care Needs"] if c in df.columns]
    others = [c for c in df.columns if c not in set(leading + metrics)]
    df = df[leading + metrics + others]
    return df

uploaded_files = st.file_uploader("Upload two Enrollment exports (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if st.button("Process & Save to OneDrive"):
    if not uploaded_files or len(uploaded_files) != 2:
        st.error("Please upload exactly two Excel files.")
        st.stop()

    dfs = []
    for f in uploaded_files:
        df_part = read_and_standardize(f)
        if df_part is None:
            st.stop()
        dfs.append(df_part)

    combined = pd.concat(dfs, ignore_index=True)
    cleaned = clean_and_collapse(combined)

    csv_name = "PowerBI_Enrollment_Clean.csv"
    xlsx_name = "PowerBI_Enrollment_Clean.xlsx"
    central_now = datetime.now(ZoneInfo("America/Chicago"))
    stamp = central_now.strftime("%Y%m%d_%H%M%S")
    csv_ts = f"PowerBI_Enrollment_Clean_{stamp}.csv"
    xlsx_ts = f"PowerBI_Enrollment_Clean_{stamp}.xlsx"

    cleaned.to_csv(os.path.join(PB_FOLDER, csv_name), index=False)
    cleaned.to_excel(os.path.join(PB_FOLDER, xlsx_name), index=False)
    cleaned.to_csv(os.path.join(PB_FOLDER, csv_ts), index=False)
    cleaned.to_excel(os.path.join(PB_FOLDER, xlsx_ts), index=False)

    st.success("✅ Clean files saved to your OneDrive Power BI folder.")
    st.code(PB_FOLDER)
    st.download_button("⬇️ Download CSV (recommended)", data=cleaned.to_csv(index=False).encode("utf-8"), file_name=csv_name, mime="text/csv")
