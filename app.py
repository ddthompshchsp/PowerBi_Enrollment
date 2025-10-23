from datetime import datetime, date
from zoneinfo import ZoneInfo
import re
import os
import pandas as pd
import streamlit as st
from openpyxl.utils.datetime import from_excel  # for Excel serials

st.set_page_config(page_title="Power BI Enrollment – No Styles/No Totals", layout="centered")
st.title("Power BI Enrollment – No Styles/No Totals")
st.caption("Uploads the same Enrollment export as your original app and outputs the **exact same data table** but with **no styling and no totals**, saved directly to OneDrive.")

# === OneDrive path for Power BI ===
PB_FOLDER = r"C:\Users\Daniella.Thompson\OneDrive - hchsp.org\Power Bi Data"
os.makedirs(PB_FOLDER, exist_ok=True)

# -------- Utilities from your original logic --------
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
    for v in pd.unique(series.dropna()):
        dt = coerce_to_dt(v)
        if dt:
            dates.append(dt)
        else:
            s = str(v).strip()
            if s:
                texts.append(s)
    if dates:
        return max(dates)
    return texts[0] if texts else None

def normalize(s: str) -> str:
    s = s.lower() if isinstance(s, str) else str(s).lower()
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

# -------- Core processing (mirrors your original, minus styling/totals) --------
def process_enrollment(file):
    # Detect header row by searching for 'ST: Participant PID' anywhere in first 30 rows
    tmp = pd.read_excel(file, header=None, nrows=30)
    header_row = None
    for r in range(tmp.shape[0]):
        row_vals = tmp.iloc[r].astype(str).fillna("")
        if any("ST: Participant PID" in v for v in row_vals):
            header_row = r
            break
    if header_row is None:
        st.error("Couldn't find 'ST: Participant PID' in the file.")
        return None

    file.seek(0)
    df = pd.read_excel(file, header=header_row)
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]

    if "Participant PID" not in df.columns:
        st.error("The file is missing 'Participant PID'.")
        return None

    # Deduplicate PIDs using most_recent per column
    df = (
        df.dropna(subset=["Participant PID"])
          .groupby("Participant PID", as_index=False)
          .agg(most_recent)
    )

    all_cols = list(df.columns)
    immun_cols = find_cols(all_cols, ["immun"])
    tb_cols    = find_cols(all_cols, ["tb", "tuberc", "ppd"])
    lead_cols  = find_cols(all_cols, ["lead", "pb"])
    scn_en_cols = find_cols(all_cols, ["special care needs english"])
    scn_es_cols = find_cols(all_cols, ["special care needs spanish"])
    scn_cols = scn_en_cols + scn_es_cols

    # Collapse multi-columns to one, like the original
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

    # Apply the same date threshold logic as your original code (but in-dataframe)
    general_cutoff = datetime(2025, 5, 11)
    field_cutoff   = datetime(2025, 8, 1)

    # Build a set for special fields that use field_cutoff
    special_date_cols = set([c for c in ["Immunizations", "TB Test", "Lead Test", "Child's Special Care Needs"] if c in df.columns])

    def transform_cell(val, colname):
        # Missing -> "X" (as original)
        if val in (None, "", "nan", "NaT"):
            return "X"
        dt = coerce_to_dt(val)
        if dt:
            # special columns use field_cutoff, others use general_cutoff
            cutoff = field_cutoff if colname in special_date_cols else general_cutoff
            if dt < cutoff:
                return "X"
            # format date plain (Power BI-friendly)
            return dt.date().isoformat()
        # Non-date, keep as-is
        return val

    # Transform all cells with the original rule
    for col in df.columns:
        df[col] = df[col].apply(lambda v, c=col: transform_cell(v, c))

    # Order columns: like original you often want identifiers first
    leading = [c for c in ["Participant PID", "Participant Name", "Center", "Campus", "School", "Area"] if c in df.columns]
    metrics = [c for c in ["Immunizations", "TB Test", "Lead Test", "Child's Special Care Needs"] if c in df.columns]
    others  = [c for c in df.columns if c not in set(leading + metrics)]
    df = df[leading + metrics + others]

    # Explicitly ensure no total-like rows (safety; groupby already removes them)
    def row_has_total_text(row):
        joined = " | ".join([str(x) for x in row.values]).lower()
        return "total" in joined and "participant pid" not in joined
    df = df[~df.apply(row_has_total_text, axis=1)]

    return df

# -------- UI --------
uploaded = st.file_uploader("Upload Enrollment.xlsx (same export you used before)", type=["xlsx"])

if st.button("Process & Save to OneDrive"):
    if not uploaded:
        st.error("Please upload your Enrollment.xlsx export.")
        st.stop()

    cleaned_df = process_enrollment(uploaded)
    if cleaned_df is None:
        st.stop()

    # Save plain outputs (no styles, no totals)
    base_csv = "PowerBI_Enrollment_Main_NoStyles.csv"
    base_xlsx = "PowerBI_Enrollment_Main_NoStyles.xlsx"
    cleaned_df.to_csv(os.path.join(PB_FOLDER, base_csv), index=False)
    cleaned_df.to_excel(os.path.join(PB_FOLDER, base_xlsx), index=False)

    st.success("✅ Saved plain, no-styles/no-totals files to OneDrive.")
    st.code(os.path.join(PB_FOLDER, base_csv))

    # Also offer downloads now
    st.download_button("⬇️ Download CSV", data=cleaned_df.to_csv(index=False).encode("utf-8"),
                       file_name=base_csv, mime="text/csv")
    # For Excel download without writing engine to buffer, write to temp then read back
    tmp_xlsx = cleaned_df.to_excel(index=False, engine="openpyxl")
