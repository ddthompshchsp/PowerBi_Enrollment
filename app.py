
import io
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Enrollment", layout="centered")
st.title("HCHSP Enrollment (Unstyled • No Totals)")
st.caption("Upload **VF_Average_Funded_Enrollment_Level.xlsx** and **25–26 Applied/Accepted.xlsx**. "
           "This produces ONE Excel for download with the same data and columns as your formatted version, "
           "but **no totals rows**, and it **excludes 'Lic Cap.' and 'Comments'**.")

# ===== Helpers =====
DASH_CLASS = r"[-‐-‒–—]"

def _norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()

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
    return ("class totals" in joined) or ("totals for class" in joined) or (re.search(r"\bclass\s*total", joined) is not None)

def _last_two_numbers(row):
    nums = []
    for v in row:
        x = pd.to_numeric(v, errors="coerce")
        if pd.notna(x):
            nums.append(float(x))
    if len(nums) >= 2:
        return nums[-2], nums[-1]
    return None, None

# ===== Parse VF (funded/enrolled by class) =====
def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
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

# ===== Parse Applied/Accepted =====
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

# ===== Build output table (drop Lic Cap./Comments; will drop totals later) =====
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

        # Center Total row (we will drop after build)
        rows.append({
            "Center": f"{center} Total",
            "Room#/Age/Lang": "",
            "Funded": funded_sum, "Enrolled": enrolled_sum,
            "Applied": applied_val, "Accepted": accepted_val,
            "Lacking/Overage": funded_sum - enrolled_sum,
            "Waitlist": accepted_val if enrolled_sum >= funded_sum else "",
            "% Enrolled of Funded": pct_total
        })

        # Class rows
        for _, r in group.iterrows():
            rows.append({
                "Center": r["Center"],
                "Room#/Age/Lang": r["Class"],
                "Funded": int(r["Funded"]), "Enrolled": int(r["Enrolled"]),
                "Applied": "", "Accepted": "", "Lacking/Overage": "", "Waitlist": "",
                "% Enrolled of Funded": int(r["PctInt"]) if pd.notna(r["PctInt"]) else pd.NA
            })

    final = pd.DataFrame(rows)

    # Agency Total row (will be dropped later)
    agency_funded   = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Funded"].sum())
    agency_enrolled = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Enrolled"].sum())
    counts_totals   = counts[["Applied","Accepted"]].sum()
    agency_applied  = int(counts_totals["Applied"])
    agency_accepted = int(counts_totals["Accepted"])
    agency_pct      = int(round(agency_enrolled / agency_funded * 100, 0)) if agency_funded > 0 else pd.NA

    final = pd.concat([final, pd.DataFrame([{
        "Center": "Agency Total",
        "Room#/Age/Lang": "",
        "Funded": agency_funded, "Enrolled": agency_enrolled,
        "Applied": agency_applied, "Accepted": agency_accepted,
        "Lacking/Overage": agency_funded - agency_enrolled, "Waitlist": "",
        "% Enrolled of Funded": agency_pct
    }])], ignore_index=True)

    final = final[[
        "Center","Room#/Age/Lang","Funded","Enrolled","Applied","Accepted",
        "Lacking/Overage","Waitlist","% Enrolled of Funded"
    ]]
    return final

# ===== UI =====
vf_up = st.file_uploader("Upload VF_Average_Funded_Enrollment_Level.xlsx", type=["xlsx"], key="vf")
aa_up = st.file_uploader("Upload 25-26 Applied/Accepted.xlsx", type=["xlsx"], key="aa")
if st.button("Process & Download"):
    if not vf_up or not aa_up:
        st.error("Please upload both files.")
        st.stop()
    try:
        vf_raw = pd.read_excel(vf_up, sheet_name=0, header=None)
        aa_raw = pd.read_excel(aa_up, sheet_name=0, header=None)

        vf_tidy = parse_vf(vf_raw)
        aa_counts = parse_applied_accepted(aa_raw)
        df_full = build_output_table(vf_tidy, aa_counts)

        # Drop ALL totals rows (center totals + agency total)
        mask_totals = df_full["Center"].astype(str).str.endswith(" Total", na=False) | \
                      df_full["Center"].astype(str).eq("Agency Total")
        df = df_full[~mask_totals].reset_index(drop=True)

        # Preview
        st.dataframe(df, use_container_width=True)

        # Build Excel in memory for download
        out_name = "HCHSP_Enrollment_Unstyled_NoTotals.xlsx"
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Head Start Enrollment")
        excel_buf.seek(0)
        st.download_button(
            "⬇️ Download HCHSP_Enrollment_Unstyled_NoTotals.xlsx",
            data=excel_buf.getvalue(),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Processing error: {e}")
