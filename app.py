import re
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import os

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Enrollment (Unstyled • No Totals)", layout="centered")
st.title("HCHSP Enrollment (Unstyled • No Totals)")
st.caption("Upload the VF Average Funded Enrollment report and the 25–26 Applied/Accepted report. "
           "This outputs a single Excel file with the same data and columns as your original app, "
           "but WITHOUT totals rows and WITHOUT styling, saved straight to OneDrive.")

# ----------------------------
# Inputs
# ----------------------------
vf_file = st.file_uploader("Upload *VF_Average_Funded_Enrollment_Level.xlsx*", type=["xlsx"], key="vf")
aa_file = st.file_uploader("Upload *25-26 Applied/Accepted.xlsx*", type=["xlsx"], key="aa")
process  = st.button("Process & Save to OneDrive")

# ----------------------------
# OneDrive path (hard-coded)
# ----------------------------
PB_FOLDER = r"C:\Users\Daniella.Thompson\OneDrive - hchsp.org\Power Bi Data"
os.makedirs(PB_FOLDER, exist_ok=True)

# ----------------------------
# HARD-CODED LICENSED CAPACITY
# ----------------------------
LIC_CAP = {
    "alvarez": 138, "camarena": 192, "chapa": 154, "edinburg": 232,
    "edinburg north": 147, "escandon": 131, "farias": 153, "guerra": 144,
    "guzman": 343, "longoria": 125, "mercedes": 213, "mission": 165,
    "monte alto": 100, "palacios": 135, "salinas": 90, "sam fordyce": 121,
    "sam houston": 134, "san carlos": 105, "san juan": 182, "seguin": 150,
    "singleterry": 130, "thigpen": 136, "wilson": 119
}

DASH_CLASS = r"[-‐-‒–—]"  # accept ASCII hyphen + Unicode dashes

# ----------------------------
# Normalization / matching
# ----------------------------
def _norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()

def _canonicalize_center(s: str) -> str:
    if s is None:
        return ""
    txt = str(s)
    txt = re.sub(rf"^\s*HCHSP\s*{DASH_CLASS}{{1,}}\s*", "", txt, flags=re.I)  # strip "HCHSP — "
    txt = re.sub(r"\([^)]*\)", " ", txt)                                     # remove "(...)"
    txt = txt.lower()
    txt = re.sub(r"[^a-z0-9\s]", " ", txt)
    filler = {"head", "start", "headstart", "hs", "ehs", "center", "campus", "elementary", "school", "program"}
    tokens = [t for t in txt.split() if t and t not in filler]
    return " ".join(tokens).strip()

_CANON_TO_OFFICIAL = {_canonicalize_center(k): k for k in LIC_CAP}

def lic_cap_for(center_name: str):
    if not isinstance(center_name, str):
        return None
    canon = _canonicalize_center(center_name)
    if canon in _CANON_TO_OFFICIAL:
        return LIC_CAP[_CANON_TO_OFFICIAL[canon]]
    best_key, best_len = None, 0
    for canon_k, off in _CANON_TO_OFFICIAL.items():
        if canon_k in canon or canon in canon_k:
            if len(canon_k) > best_len:
                best_key, best_len = off, len(canon_k)
    return LIC_CAP.get(best_key) if best_key else None

# ----------------------------
# Helpers (parsing)
# ----------------------------
def _first_nonempty_strings(row, max_cols=8):
    vals = []
    for j in range(min(max_cols, row.shape[0])):
        v = row.iloc[j]
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s:
            vals.append(s)
    return vals

def _row_has_totals(cells_lower: list[str]) -> bool:
    joined = " | ".join(cells_lower)
    return (
        "class totals" in joined
        or "totals for class" in joined
        or re.search(r"\bclass\s*total", joined) is not None
    )

def _last_two_numbers(row):
    nums = []
    for v in row:
        x = pd.to_numeric(v, errors="coerce")
        if pd.notna(x):
            nums.append(float(x))
    if len(nums) >= 2:
        return nums[-2], nums[-1]
    return None, None

# ----------------------------
# Parsers
# ----------------------------
def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
    # Output columns: Center | Class | Funded | Enrolled | PctRatio
    # - Accept any dash between 'HCHSP' and center
    # - Capture FULL class name (incl. parentheses)
    # - Totals row: Enrolled=col 3, Funded=col 4, Percent ratio=col 6 (fallback to last-two-numbers)
    records = []
    current_center = None
    current_class  = None

    re_center = re.compile(rf"^\s*HCHSP\s*{DASH_CLASS}{{1,}}\s*(.+)$", re.I)
    re_class  = re.compile(r"^\s*Class\s+(?!Totals:)(.+?)\s*$", re.I)  # avoid "Class Totals:"

    for i in range(len(vf_df_raw)):
        row = vf_df_raw.iloc[i, :]
        cells = _first_nonempty_strings(row, max_cols=8)
        if not cells:
            continue

        first = cells[0]
        lower_cells = [c.lower() for c in cells]

        m_center = re_center.match(first)
        if m_center:
            current_center = _norm_ws(m_center.group(1))
            continue

        if _row_has_totals(lower_cells) and current_center and current_class:
            enrolled = pd.to_numeric(row.iloc[3], errors="coerce")
            funded   = pd.to_numeric(row.iloc[4], errors="coerce")
            pct_ratio= pd.to_numeric(row.iloc[6], errors="coerce")  # ratio (1.00, 1.12, ...)
            if pd.isna(enrolled) or pd.isna(funded):
                e, f = _last_two_numbers(row)
                enrolled = e if e is not None else enrolled
                funded   = f if f is not None else funded
            records.append({
                "Center": current_center,
                "Class": f"Class {current_class}",
                "Funded": 0.0 if pd.isna(funded) else float(funded),
                "Enrolled": 0.0 if pd.isna(enrolled) else float(enrolled),
                "PctRatio": None if pd.isna(pct_ratio) else float(pct_ratio),
            })
            continue

        m_class = re_class.match(first)
        if m_class:
            current_class = m_class.group(1).strip()
            continue

    tidy = pd.DataFrame(records)
    if tidy.empty:
        raise ValueError("Could not parse VF report (check class/center markers and column indices).")
    tidy["Center"] = tidy["Center"].map(_norm_ws)
    return tidy


def parse_applied_accepted(aa_df_raw: pd.DataFrame) -> pd.DataFrame:
    header_row_idx = aa_df_raw.index[aa_df_raw.iloc[:, 0].astype(str).str.startswith("ST: Participant PID", na=False)]
    if len(header_row_idx) == 0:
        raise ValueError("Could not find header row in Applied/Accepted report.")
    header_row_idx = int(header_row_idx[0])
    headers = aa_df_raw.iloc[header_row_idx].tolist()
    body = pd.DataFrame(aa_df_raw.iloc[header_row_idx + 1:].values, columns=headers)

    center_col = "ST: Center Name"
    status_col = "ST: Status"
    date_col = "ST: Status End Date"

    is_blank_date = body[date_col].isna() | body[date_col].astype(str).str.strip().eq("")
    body = body[is_blank_date].copy()
    body[center_col] = (
        body[center_col]
        .astype(str)
        .str.replace(rf"^\s*HCHSP\s*{DASH_CLASS}{{1,}}\s*", "", regex=True)
        .map(_norm_ws)
    )

    counts = body.groupby(center_col)[status_col].value_counts().unstack(fill_value=0)
    for c in ["Accepted", "Applied"]:
        if c not in counts.columns:
            counts[c] = 0
    return counts[["Accepted", "Applied"]].astype(int).reset_index().rename(columns={center_col: "Center"})

# ----------------------------
# Builder (same as original, but totals will be dropped after)
# ----------------------------
def build_output_table(vf_tidy: pd.DataFrame, counts: pd.DataFrame) -> pd.DataFrame:
    if "PctRatio" in vf_tidy.columns and vf_tidy["PctRatio"].notna().any():
        vf_tidy["PctInt"] = pd.array((vf_tidy["PctRatio"] * 100).round(0), dtype="Int64")
    else:
        pct = (vf_tidy["Enrolled"] * 100).div(pd.Series(vf_tidy["Funded"]).replace(0, np.nan))
        vf_tidy["PctInt"] = pd.array(pct.round(0), dtype="Int64")

    merged = vf_tidy.merge(counts, on="Center", how="left").fillna({"Accepted": 0, "Applied": 0})
    applied_by_center  = merged.groupby("Center")["Applied"].max()
    accepted_by_center = merged.groupby("Center")["Accepted"].max()

    rows = []
    for center, group in merged.groupby("Center", sort=True):
        funded_sum   = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        pct_total    = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA
        accepted_val = int(accepted_by_center.get(center, 0))
        applied_val  = int(applied_by_center.get(center, 0))

        # Center Total row (will be dropped later)
        rows.append({
            "Center": f"{center} Total",
            "Room#/Age/Lang": "",
            "Lic Cap.": lic_cap_for(center),
            "Funded": funded_sum, "Enrolled": enrolled_sum,
            "Applied": applied_val, "Accepted": accepted_val,
            "Lacking/Overage": funded_sum - enrolled_sum, "Waitlist": accepted_val if enrolled_sum >= funded_sum else "",
            "% Enrolled of Funded": pct_total,
            "Comments": ""
        })

        # Class rows
        for _, r in group.iterrows():
            rows.append({
                "Center": r["Center"],
                "Room#/Age/Lang": r["Class"],
                "Lic Cap.": "",
                "Funded": int(r["Funded"]), "Enrolled": int(r["Enrolled"]),
                "Applied": "", "Accepted": "", "Lacking/Overage": "", "Waitlist": "",
                "% Enrolled of Funded": int(r["PctInt"]) if pd.notna(r["PctInt"]) else pd.NA,
                "Comments": ""
            })

    final = pd.DataFrame(rows)

    # Agency total (will be dropped later)
    agency_funded   = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Funded"].sum())
    agency_enrolled = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Enrolled"].sum())
    counts_totals   = counts[["Applied","Accepted"]].sum()
    agency_applied  = int(counts_totals["Applied"])
    agency_accepted = int(counts_totals["Accepted"])
    agency_pct      = int(round(agency_enrolled / agency_funded * 100, 0)) if agency_funded > 0 else pd.NA
    agency_lacking  = agency_funded - agency_enrolled

    final = pd.concat([final, pd.DataFrame([{
        "Center": "Agency Total",
        "Room#/Age/Lang": "",
        "Lic Cap.": "",
        "Funded": agency_funded, "Enrolled": agency_enrolled,
        "Applied": agency_applied, "Accepted": agency_accepted,
        "Lacking/Overage": agency_lacking, "Waitlist": "",
        "% Enrolled of Funded": agency_pct,
        "Comments": ""
    }])], ignore_index=True)

    final = final[[
        "Center","Room#/Age/Lang","Lic Cap.","Funded","Enrolled",
        "Applied","Accepted","Lacking/Overage","Waitlist","% Enrolled of Funded","Comments"
    ]]
    return final

# ----------------------------
# Main
# ----------------------------
if process and vf_file and aa_file:
    try:
        vf_raw = pd.read_excel(vf_file, sheet_name=0, header=None)
        aa_raw = pd.read_excel(aa_file, sheet_name=0, header=None)

        vf_tidy = parse_vf(vf_raw)
        aa_counts = parse_applied_accepted(aa_raw)
        final_df_with_totals = build_output_table(vf_tidy, aa_counts)

        # Drop ALL totals rows (center totals + agency total)
        mask_totals = final_df_with_totals["Center"].astype(str).str.endswith(" Total", na=False) | \
                      final_df_with_totals["Center"].astype(str).eq("Agency Total")
        final_df = final_df_with_totals[~mask_totals].reset_index(drop=True)

        # Save ONE unstyled Excel file to OneDrive
        out_name = "HCHSP_Enrollment_Unstyled_NoTotals.xlsx"
        out_path = os.path.join(PB_FOLDER, out_name)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Head Start Enrollment")

        st.success("✅ Saved unstyled Excel (no totals) to OneDrive.")
        st.code(out_path)
        st.dataframe(final_df, use_container_width=True)

        # Also return a direct download
        with open(out_path, "rb") as f:
            st.download_button("⬇️ Download HCHSP_Enrollment_Unstyled_NoTotals.xlsx",
                               data=f.read(),
                               file_name=out_name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Processing error: {e}")
