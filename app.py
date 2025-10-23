from pathlib import Path

app_code = r"""# app.py
from datetime import datetime, date
from zoneinfo import ZoneInfo
import os, re
import pandas as pd
import streamlit as st
from openpyxl.utils.datetime import from_excel  # Excel serials -> datetime

st.set_page_config(page_title="Power BI Enrollment — One Clean File (Unstyled)", layout="centered")
st.title("Power BI Enrollment — One Clean File (Unstyled)")
st.caption("Upload the same two **original** files you use today: "
           "(1) Participant export (has PID) and (2) Funded vs Enrolled (no PID). "
           "You'll get **one Excel file** with the same data and sheets as your formatted version, "
           "but **no styles** and **no totals**, saved to OneDrive.")

# -------- OneDrive target --------
PB_FOLDER = r"C:\Users\Daniella.Thompson\OneDrive - hchsp.org\Power Bi Data"
os.makedirs(PB_FOLDER, exist_ok=True)

# -------- Utilities (matches your original logic) --------
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

# -------- Participant processing (same rules; no styles/totals) --------
GENERAL_CUTOFF = datetime(2025, 5, 11)
FIELD_CUTOFF   = datetime(2025, 8, 1)

METRIC_COLS = ["Immunizations", "TB Test", "Lead Test", "Child's Special Care Needs"]

def detect_header_row(file, landmark_terms=("ST: Participant PID","Participant PID","PID")):
    tmp = pd.read_excel(file, header=None, nrows=120)
    header_row = None
    for r in range(tmp.shape[0]):
        row_vals = tmp.iloc[r].astype(str).fillna("")
        row_text = " | ".join(row_vals)
        if any(term.lower() in row_text.lower() for term in landmark_terms):
            header_row = r
            break
    file.seek(0)
    return header_row

def read_participant(file):
    hdr = detect_header_row(file)
    if hdr is None:
        return None
    df = pd.read_excel(file, header=hdr)
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]
    pid_col = None
    for c in df.columns:
        if isinstance(c, str) and "pid" in c.lower():
            pid_col = c
            break
    if pid_col is None:
        return None
    if pid_col != "Participant PID":
        df = df.rename(columns={pid_col: "Participant PID"})
    df = df.dropna(subset=["Participant PID"])
    return df

def clean_participants_like_original(df):
    # collapse to most recent per PID
    df = df.groupby("Participant PID", as_index=False).agg(most_recent)
    all_cols = list(df.columns)

    immun_cols = find_cols(all_cols, ["immun"])
    tb_cols    = find_cols(all_cols, ["tb","tuberc","ppd"])
    lead_cols  = find_cols(all_cols, ["lead","pb"])
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

    special_date_cols = set([c for c in METRIC_COLS if c in df.columns])

    def transform_cell(val, colname):
        if val in (None, "", "nan", "NaT"):
            return "X"
        dt = coerce_to_dt(val)
        if dt:
            cutoff = FIELD_CUTOFF if colname in special_date_cols else GENERAL_CUTOFF
            if dt < cutoff:
                return "X"
            return dt.date().isoformat()
        return val

    for col in df.columns:
        df[col] = df[col].apply(lambda v, c=col: transform_cell(v, c))

    leading = [c for c in ["Participant PID","Participant Name","Center","Campus","School","Area"] if c in df.columns]
    metrics = [c for c in METRIC_COLS if c in df.columns]
    others  = [c for c in df.columns if c not in set(leading + metrics)]
    df = df[leading + metrics + others]

    # remove any residual total-like rows (shouldn't exist after PID groupby)
    def row_has_total_text(row):
        joined = " | ".join([str(x) for x in row.values]).lower()
        return " total" in joined or joined.startswith("total")
    df = df[~df.apply(row_has_total_text, axis=1)]
    return df

def center_summary(df):
    metrics = [c for c in METRIC_COLS if c in df.columns]
    name_col = None
    for c in ["Center","Campus","School","Center/Campus"]:
        if c in df.columns:
            name_col = c
            break
    if name_col is None:
        return None
    def row_complete(row):
        for c in metrics:
            v = row.get(c, None)
            if v is None or str(v).strip().upper() == "X" or str(v).strip() == "":
                return False
        return True
    base = df[[name_col,"Participant PID"] + metrics].copy()
    base["CompleteFlag"] = base.apply(row_complete, axis=1)
    grp = base.groupby(name_col, dropna=False).agg(
        Completed_Students=("CompleteFlag","sum"),
        Total_Students=("Participant PID","nunique")
    ).reset_index()
    grp["Completion_Rate"] = grp["Completed_Students"] / grp["Total_Students"]
    grp = grp.rename(columns={name_col:"Center/Campus"})
    return grp

def scn_summary(df):
    if "Child's Special Care Needs" not in df.columns:
        return None
    name_col = None
    for c in ["Center","Campus","School","Center/Campus"]:
        if c in df.columns:
            name_col = c
            break
    if name_col is None:
        return None
    scn_ok = df["Child's Special Care Needs"].astype(str).str.upper() != "X"
    agg = df.groupby(name_col, dropna=False).agg(
        Completed_SCN=(scn_ok,"sum"),
        Total_Students=("Participant PID","nunique")
    ).reset_index()
    agg["Remaining"] = agg["Total_Students"] - agg["Completed_SCN"]
    agg["Completion_Rate"] = (agg["Completed_SCN"] / agg["Total_Students"]).fillna(0)
    agg = agg.rename(columns={name_col:"Center/Campus"})
    return agg

# -------- Funded vs Enrolled parsing (saved as a sheet; no totals; unstyled) --------
FUND_KEYS = ["funded","fund","slot","capacity"]
ENR_KEYS  = ["enrolled","enrol","actual","current","served"]
NAME_KEYS = ["center","campus","school","site","location","center name","campus name"]
AREA_KEYS = ["area","region"]

def find_cols_soft(cols, keys):
    out = []
    for c in cols:
        n = normalize(c)
        if any(k in n for k in keys):
            out.append(c)
    return out

def read_funded(file):
    df_raw = pd.read_excel(file, header=None, dtype=object)
    # find probable header row
    hdr = None
    limit = min(60, len(df_raw))
    for r in range(limit):
        row = df_raw.iloc[r].astype(str).fillna("")
        row_norm = [normalize(x) for x in row]
        has_name = any(any(k in cell for k in NAME_KEYS) for cell in row_norm)
        has_fund = any(any(k in cell for k in FUND_KEYS) for cell in row_norm)
        has_enr  = any(any(k in cell for k in ENR_KEYS) for cell in row_norm)
        if has_name and (has_fund or has_enr):
            hdr = r
            break
    if hdr is None:
        non_empty_rows = df_raw.apply(lambda r: r.notna().sum(), axis=1)
        hdr = int(non_empty_rows.idxmax()) if non_empty_rows.max() > 0 else 0
    headers = df_raw.iloc[hdr].astype(str).tolist()
    df0 = df_raw.iloc[hdr+1:].reset_index(drop=True)
    df0.columns = headers
    df0 = df0.dropna(axis=1, how="all").dropna(how="all")
    cols = list(df0.columns)
    name_cols = find_cols_soft(cols, NAME_KEYS)
    funded_cols = find_cols_soft(cols, FUND_KEYS)
    enrolled_cols = find_cols_soft(cols, ENR_KEYS)
    area_cols = find_cols_soft(cols, AREA_KEYS)
    if not name_cols:
        return None
    name_col = name_cols[0]
    funded_col = funded_cols[0] if funded_cols else None
    enrolled_col = enrolled_cols[0] if enrolled_cols else None
    area_col = area_cols[0] if area_cols else None
    # remove totals
    df0[name_col] = df0[name_col].astype(str)
    df0 = df0[~df0[name_col].str.lower().str.contains("total", na=False)]
    out = pd.DataFrame()
    out["Center/Campus"] = df0[name_col].astype(str).str.strip()
    if area_col: out["Area"] = df0[area_col].astype(str).str.strip()
    if funded_col: out["Funded"] = pd.to_numeric(df0[funded_col], errors="coerce")
    if enrolled_col: out["Enrolled"] = pd.to_numeric(df0[enrolled_col], errors="coerce")
    out = out[out["Center/Campus"].str.len() > 0]
    return out

# -------- UI --------
st.subheader("Uploads")
part_file = st.file_uploader("Participant export (.xlsx) — has PID", type=["xlsx"], key="part")
fund_file = st.file_uploader("Funded vs Enrolled (.xlsx) — no PID", type=["xlsx"], key="fund")

if st.button("Create One Clean Excel (save to OneDrive)"):
    if not part_file or not fund_file:
        st.error("Please upload BOTH the participant export and the funded vs enrolled file.")
        st.stop()

    # participants
    df_part_raw = read_participant(part_file)
    if df_part_raw is None:
        st.error("Could not detect a Participant PID column in the participant file.")
        st.stop()
    df_part = clean_participants_like_original(df_part_raw)

    # funded
    df_fund = read_funded(fund_file)
    if df_fund is None:
        st.error("Could not detect Funded/Enrolled and Center/Campus headers in the funded file.")
        st.stop()

    # summaries like your original (no styles, no totals)
    df_center = center_summary(df_part)
    df_scn    = scn_summary(df_part)

    # one Excel file, multiple sheets, unstyled
    out_name = "PowerBI_Enrollment_Unstyled.xlsx"
    out_path = os.path.join(PB_FOLDER, out_name)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_part.to_excel(writer, index=False, sheet_name="Main")
        if df_center is not None:
            df_center.to_excel(writer, index=False, sheet_name="Center Summary")
        if df_scn is not None:
            df_scn.to_excel(writer, index=False, sheet_name="Child's Special Care Needs Summary")
        # include funded sheet for completeness
        df_fund.to_excel(writer, index=False, sheet_name="Funded (raw)")

    st.success("✅ One clean Excel (no styling, no totals) saved to OneDrive.")
    st.code(out_path)
    # also allow direct download
    with open(out_path, "rb") as f:
        st.download_button("⬇️ Download PowerBI_Enrollment_Unstyled.xlsx", data=f.read(), file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
"""

reqs = """streamlit>=1.37
pandas>=2.2
openpyxl>=3.1
pillow>=10.3
tzdata>=2024.1
"""

Path("/mnt/data/app.py").write_text(app_code)
Path("/mnt/data/requirements.txt").write_text(reqs)

"/mnt/data app.py and requirements.txt written (single Excel output, unstyled, no totals)"
