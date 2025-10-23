from datetime import datetime, date
from zoneinfo import ZoneInfo
import re
import os
import pandas as pd
import streamlit as st
from openpyxl.utils.datetime import from_excel  # used by coerce_to_dt

st.set_page_config(page_title="Power BI Enrollment (Cleaned)", layout="centered")
st.title("Power BI Enrollment (Cleaned)")
st.caption(
    "Upload two Excel files: (1) a participant-level export (with Participant PID) and "
    "(2) a Funded vs Enrolled export WITHOUT PIDs. "
    "This app will produce plain CSV/XLSX files in your OneDrive folder ready for Power BI."
)

# === Your OneDrive path for Power BI ===
PB_FOLDER = r"C:\Users\Daniella.Thompson\OneDrive - hchsp.org\Power Bi Data"
os.makedirs(PB_FOLDER, exist_ok=True)

# ---------- Utilities ----------
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
    s = str(s)
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

# ---------- Readers ----------
def detect_header_row(file, landmark_terms=("ST: Participant PID", "Participant PID")):
    """Find a header row by looking for a landmark string in first 60 rows."""
    tmp = pd.read_excel(file, header=None, nrows=60)
    header_row = None
    for r in range(tmp.shape[0]):
        row_vals = tmp.iloc[r].astype(str).fillna("")
        row_text = " | ".join(row_vals)
        if any(term in row_text for term in landmark_terms):
            header_row = r
            break
    file.seek(0)
    return header_row

def read_participant_export(file):
    """Read a participant-level export (expects PID somewhere in columns)."""
    header_row = detect_header_row(file)
    if header_row is None:
        return None  # not participant style
    df = pd.read_excel(file, header=header_row)
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]
    # robust PID presence (allow variations like 'Participant PID' or 'PID')
    pid_col = None
    for c in df.columns:
        if isinstance(c, str) and "pid" in c.lower():
            pid_col = c
            break
    if pid_col is None:
        return None
    df = df.dropna(subset=[pid_col])
    # Standardize PID column name
    if pid_col != "Participant PID":
        df = df.rename(columns={pid_col: "Participant PID"})
    return df

def read_funded_export(file):
    """Read a funded/enrolled summary WITHOUT PID. Very flexible header handling."""
    # Try first row as header; if messy, try to auto-promote first non-empty row
    try:
        df0 = pd.read_excel(file, header=0)
    except Exception:
        file.seek(0)
        df0 = pd.read_excel(file, header=None)
        # promote first row with enough non-nulls as header
        header_idx = df0.apply(lambda r: r.notna().sum(), axis=1).idxmax()
        headers = df0.iloc[header_idx].astype(str).tolist()
        df0 = df0.iloc[header_idx+1:].reset_index(drop=True)
        df0.columns = headers

    # Drop totally empty columns
    df0 = df0.dropna(axis=1, how="all")
    # Remove empty rows
    df0 = df0.dropna(how="all")

    # Fuzzy match key columns
    cols = list(df0.columns)
    name_cols = find_cols(cols, ["center", "campus", "school", "site", "location"])
    funded_cols = find_cols(cols, ["funded", "slots", "capacity"])
    enrolled_cols = find_cols(cols, ["enrolled", "actual", "current", "served"])
    area_cols = find_cols(cols, ["area", "region"])

    # choose the first reasonable match for each
    name_col = name_cols[0] if name_cols else None
    funded_col = funded_cols[0] if funded_cols else None
    enrolled_col = enrolled_cols[0] if enrolled_cols else None
    area_col = area_cols[0] if area_cols else None

    # if we can't find funded or enrolled, this is likely not the funded sheet
    if not (funded_col and enrolled_col and name_col):
        return None

    # Clean totals/headers footers: remove rows that look like totals
    def is_totalish(x):
        s = str(x).strip().lower()
        return any(t in s for t in ["total", "grand total", "filtered total"])

    df0 = df0[~df0[name_col].astype(str).str.lower().str.contains("total", na=False)]

    out = pd.DataFrame({
        "Area": df0[area_col] if area_col else None,
        "Center/Campus": df0[name_col],
        "Funded": pd.to_numeric(df0[funded_col], errors="coerce"),
        "Enrolled": pd.to_numeric(df0[enrolled_col], errors="coerce"),
    })
    # Trim whitespace
    out["Center/Campus"] = out["Center/Campus"].astype(str).str.strip()
    if "Area" in out.columns:
        out["Area"] = out["Area"].astype(str).str.strip()
    # Drop rows without a campus name
    out = out.dropna(subset=["Center/Campus"])
    # Remove any rows where both Funded and Enrolled are NaN
    out = out.dropna(subset=["Funded", "Enrolled"], how="all")
    return out

# ---------- Cleaning ----------
def clean_and_collapse_participants(df):
    """Collapse related columns into single fields; leave lean participant table."""
    df = df.groupby("Participant PID", as_index=False).agg(most_recent)
    all_cols = list(df.columns)

    immun_cols = find_cols(all_cols, ["immun"])
    tb_cols    = find_cols(all_cols, ["tb", "tuberc", "ppd"])
    lead_cols  = find_cols(all_cols, ["lead", "pb"])
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

    # ISO date conversion for these metrics
    def to_iso_or_val(v):
        dt = coerce_to_dt(v)
        return dt.date().isoformat() if dt else (None if pd.isna(v) else v)

    for col in ["Immunizations", "TB Test", "Lead Test", "Child's Special Care Needs"]:
        if col in df.columns:
            df[col] = df[col].apply(to_iso_or_val)

    leading = [c for c in ["Participant PID", "Participant Name", "Center", "Campus", "School", "Area"] if c in df.columns]
    metrics = [c for c in ["Immunizations", "TB Test", "Lead Test", "Child's Special Care Needs"] if c in df.columns]
    others  = [c for c in df.columns if c not in set(leading + metrics)]
    df = df[leading + metrics + others]
    return df

def participants_to_center_counts(df):
    """Aggregate participant table to counts by center/campus for use next to Funded."""
    # choose a name column
    name_col = None
    for c in ["Center", "Campus", "School", "Center/Campus"]:
        if c in df.columns:
            name_col = c
            break
    if name_col is None:
        return None
    grp = df.groupby(name_col, dropna=False, as_index=False).agg(TotalStudents=("Participant PID", "nunique"))
    grp = grp.rename(columns={name_col: "Center/Campus"})
    return grp

# ---------- UI ----------
uploaded_files = st.file_uploader(
    "Upload: one participant export AND one funded/enrolled export (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("Process & Save to OneDrive"):
    if not uploaded_files or len(uploaded_files) < 1:
        st.error("Please upload at least one Excel file.")
        st.stop()

    participant_frames = []
    funded_frames = []

    # Classify each upload
    for f in uploaded_files:
        # Try participant first
        df_part = read_participant_export(f)
        if df_part is not None:
            participant_frames.append(df_part)
            continue
        # If not participant, try funded
        f.seek(0)
        df_fund = read_funded_export(f)
        if df_fund is not None:
            funded_frames.append(df_fund)
            continue
        # If neither, warn but continue
        st.warning(f"File '{f.name}' did not match participant or funded format. Skipped.")

    # --- Participant output ---
    if participant_frames:
        combined_part = pd.concat(participant_frames, ignore_index=True)
        cleaned_part = clean_and_collapse_participants(combined_part)

        part_csv = "PowerBI_Enrollment_Participants.csv"
        part_xlsx = "PowerBI_Enrollment_Participants.xlsx"
        cleaned_part.to_csv(os.path.join(PB_FOLDER, part_csv), index=False)
        cleaned_part.to_excel(os.path.join(PB_FOLDER, part_xlsx), index=False)

        st.success("✅ Participant table saved.")
        st.code(os.path.join(PB_FOLDER, part_csv))
        st.download_button("⬇️ Download Participants CSV", data=cleaned_part.to_csv(index=False).encode("utf-8"),
                           file_name=part_csv, mime="text/csv")
    else:
        st.info("No participant-style file detected (no PID found).")

    # --- Funded output ---
    if funded_frames:
        combined_fund = pd.concat(funded_frames, ignore_index=True)
        # ensure numeric
        for c in ["Funded", "Enrolled"]:
            if c in combined_fund.columns:
                combined_fund[c] = pd.to_numeric(combined_fund[c], errors="coerce")
        fund_csv = "PowerBI_Enrollment_Funded.csv"
        fund_xlsx = "PowerBI_Enrollment_Funded.xlsx"
        combined_fund.to_csv(os.path.join(PB_FOLDER, fund_csv), index=False)
        combined_fund.to_excel(os.path.join(PB_FOLDER, fund_xlsx), index=False)
        st.success("✅ Funded vs Enrolled table saved.")
        st.code(os.path.join(PB_FOLDER, fund_csv))
        st.download_button("⬇️ Download Funded CSV", data=combined_fund.to_csv(index=False).encode("utf-8"),
                           file_name=fund_csv, mime="text/csv")
    else:
        st.info("No funded/enrolled-style file detected.")

    # --- Optional: produce a join-ready counts file from participants ---
    if participant_frames:
        center_counts = participants_to_center_counts(cleaned_part)
        if center_counts is not None:
            counts_csv = "PowerBI_Enrollment_StudentCounts.csv"
            center_counts.to_csv(os.path.join(PB_FOLDER, counts_csv), index=False)
            st.success("✅ Student counts by Center/Campus saved (join to Funded in Power BI).")
            st.code(os.path.join(PB_FOLDER, counts_csv))
            st.download_button("⬇️ Download Student Counts CSV", data=center_counts.to_csv(index=False).encode("utf-8"),
                               file_name=counts_csv, mime="text/csv")
