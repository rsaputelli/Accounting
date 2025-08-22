
import io
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="WAPA SFA Builder (Streamlit)", layout="wide")

# -----------------------------
# Utilities
# -----------------------------

MONTHS = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

ACCOUNT_RE = re.compile(r"^\d{4,6}(\s*[·\-\–]\s*)?.+")

def first_nonblank_label(row: pd.Series) -> str:
    for v in row:
        if isinstance(v, str) and v.strip():
            return v.strip()
    return ""

def autodetect_header(df: pd.DataFrame, fallback_header_row: int = 0) -> pd.DataFrame:
    """
    Find the first row that contains a cell named 'Account' (case-insensitive) and
    treat that row as the header row. If none found, fallback to provided index.
    """
    # Look for 'Account' anywhere in the first ~10 rows
    search_rows = min(len(df), 15)
    for r in range(search_rows):
        row_vals = [str(v).strip().lower() for v in df.iloc[r].tolist()]
        if any(v == "account" for v in row_vals):
            new_df = df.iloc[r + 1 : ].copy()
            new_df.columns = df.iloc[r].tolist()
            return new_df.reset_index(drop=True)
    # Fallback
    new_df = df.iloc[fallback_header_row + 1 : ].copy()
    new_df.columns = df.iloc[fallback_header_row].tolist()
    return new_df.reset_index(drop=True)

def load_excel_dynamic(file_bytes) -> pd.DataFrame:
    raw = pd.read_excel(file_bytes, header=None, engine="openpyxl")
    # Try to detect header row
    df = autodetect_header(raw)
    # Normalize "Account" column if missing by extracting first nonblank label
    if "Account" not in df.columns:
        df["Account"] = df.apply(first_nonblank_label, axis=1)
    return df

def detect_month_columns(monthly_df: pd.DataFrame) -> Tuple[str, str, str]:
    """
    From a monthly BvA df, find the most recent pair of '<Month> Actual' and '<Month> Budget' headers.
    Return (month_name, actual_col, budget_col).
    """
    headers = [str(h) for h in monthly_df.columns]
    month_actuals = []
    month_budgets = []
    for h in headers:
        for m in MONTHS:
            if m.lower() in h.lower():
                if "actual" in h.lower():
                    month_actuals.append((h, m, headers.index(h)))
                if "budget" in h.lower():
                    month_budgets.append((h, m, headers.index(h)))
    # Choose the rightmost month that has both actual+budget
    best = None
    best_idx = -1
    for m in MONTHS:
        a = [idx for (h, mm, idx) in month_actuals if mm == m]
        b = [idx for (h, mm, idx) in month_budgets if mm == m]
        if a and b:
            idx = max(max(a), max(b))
            if idx > best_idx:
                # get column names
                a_name = [h for (h, mm, _) in month_actuals if mm == m and _ == max(a)][0]
                b_name = [h for (h, mm, _) in month_budgets if mm == m and _ == max(b)][0]
                best = (m, a_name, b_name)
                best_idx = idx
    if not best:
        # Provide helpful diagnostics
        raise ValueError(
            "Could not detect '<Month> Actual' and '<Month> Budget' in the Monthly BvA file.\n"
            f"Headers seen:\n{headers}"
        )
    return best

def is_account_label(txt) -> bool:
    if not isinstance(txt, str):
        return False
    return ACCOUNT_RE.match(txt.strip()) is not None

def to_num(x):
    try:
        return float(x)
    except Exception:
        return 0.0

def safe_pct(num, den):
    try:
        denf = float(den)
        if abs(denf) < 1e-12:
            return None
        return float(num) / denf
    except Exception:
        return None

def choose_col(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    for a in aliases:
        if a in df.columns:
            return a
    # case-insensitive fallback
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for a in aliases:
        al = a.lower()
        if al in lower_map:
            return lower_map[al]
    return None

def build_map(df: pd.DataFrame, key_col: str, val_col: str) -> Dict[str, float]:
    series = pd.to_numeric(df[val_col], errors="coerce").fillna(0.0)
    keys = df[key_col].fillna("")
    return dict(zip(keys, series))

def ordered_accounts_from(df: pd.DataFrame, key_col: str = "Account") -> List[str]:
    seen = set()
    order = []
    for _, row in df.iterrows():
        key = row.get(key_col, "")
        if isinstance(key, str) and key and key not in seen and is_account_label(key):
            seen.add(key)
            order.append(key)
    return order

def format_excel_bytes(df: pd.DataFrame,
                       currency_cols: List[str],
                       percent_cols: List[str],
                       sheet_name: str = "Report",
                       use_delta_symbol: bool = True) -> bytes:
    """Return an .xlsx file as bytes with number formats applied."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        wb = writer.book
        ws = writer.sheets[sheet_name]
        # Formats
        currency_fmt = wb.add_format({"num_format": '_($* #,##0.00_);[Red]_($* (#,##0.00)'})
        percent_fmt  = wb.add_format({"num_format": '0.00%'})
        header_fmt   = wb.add_format({"bold": True, "align": "center"})
        # Apply header format
        ws.set_row(0, None, header_fmt)
        # Apply column formats
        headers = list(df.columns)
        for idx, name in enumerate(headers):
            col_letter = idx
            if name in currency_cols:
                ws.set_column(col_letter, col_letter, 14, currency_fmt)
            elif name in percent_cols:
                ws.set_column(col_letter, col_letter, 10, percent_fmt)
            else:
                ws.set_column(col_letter, col_letter, 22)
        # Optional: replace Δ with 'Delta' in header if needed
        if not use_delta_symbol:
            for c, name in enumerate(headers):
                if "Δ" in name:
                    ws.write(0, c, name.replace("Δ", "Delta"), header_fmt)
    out.seek(0)
    return out.getvalue()

# -----------------------------
# UI
# -----------------------------

st.title("WAPA SFA Consolidated Report (Streamlit)")
st.caption("Upload the four inputs. The app will produce a consolidated Excel with month-named columns, YTD and Annual variances.")

with st.sidebar:
    st.header("Settings")
    fy_start = st.number_input("Fiscal year start month", min_value=1, max_value=12, value=1, help="WAPA uses calendar FY (January = 1). Used for labels only in this app.")
    use_delta_symbol = st.checkbox("Use Δ symbol in headers", value=True)
    show_preview = st.checkbox("Show preview table", value=True)

col1, col2 = st.columns(2)
with col1:
    sfa_bva = st.file_uploader("SFA BvA.xlsx (YTD & Annual)", type=["xlsx"])
    sfa_comp = st.file_uploader("SFA Comp.xlsx (Prior Year)", type=["xlsx"])
with col2:
    sfa_bva_month = st.file_uploader("SFA BvA Month.xlsx (Monthly Actual/Budget)", type=["xlsx"])
    budget_overview = st.file_uploader("Budget Overview.xlsx (optional for ordering)", type=["xlsx"])

build = st.button("Build Consolidated Report")

if build:
    if not (sfa_bva and sfa_bva_month):
        st.error("Please upload at least SFA BvA.xlsx and SFA BvA Month.xlsx.")
        st.stop()

    # Load files with robust header detection
    try:
        df_bva = load_excel_dynamic(sfa_bva)
        df_bva_month = load_excel_dynamic(sfa_bva_month)
        df_comp = load_excel_dynamic(sfa_comp) if sfa_comp else pd.DataFrame()
        df_bo = load_excel_dynamic(budget_overview) if budget_overview else pd.DataFrame()
    except Exception as e:
        st.exception(e)
        st.stop()

    # Detect month columns in monthly file
    try:
        month_name, cm_actual_col, cm_budget_col = detect_month_columns(df_bva_month)
    except Exception as e:
        st.error("Monthly BvA: " + str(e))
        st.stop()

    # Ensure Account column present
    for dfn in (df_bva, df_bva_month, df_comp, df_bo):
        if not dfn.empty and "Account" not in dfn.columns:
            dfn["Account"] = dfn.apply(first_nonblank_label, axis=1)

    # Column selections / aliases
    col_actuals = choose_col(df_bva, ["Actuals", "YTD Actuals", "Current YTD", "Total Actuals"])
    col_btd     = choose_col(df_bva, ["Budget to Date", "YTD Budget", "Budget YTD"])
    col_annual  = choose_col(df_bva, ["Annual Budget", "Total Budget", "Full Budget"])
    col_prevyr  = choose_col(df_comp, ["Previous Year", "Prior Year", "PY Actuals"]) if not df_comp.empty else None

    # Build maps
    maps = {}
    maps["cm_actual"] = build_map(df_bva_month, "Account", cm_actual_col)
    maps["cm_budget"] = build_map(df_bva_month, "Account", cm_budget_col)
    if col_actuals: maps["ytd_actual"] = build_map(df_bva, "Account", col_actuals)
    if col_btd:     maps["ytd_budget"] = build_map(df_bva, "Account", col_btd)
    if col_annual:  maps["annual_budget"] = build_map(df_bva, "Account", col_annual)
    if col_prevyr:  maps["prev_year"] = build_map(df_comp, "Account", col_prevyr)

    # Build account order (prefer Budget Overview if it includes Account; fallback to BVA order)
    account_order = []
    if not df_bo.empty and "Account" in df_bo.columns:
        account_order = ordered_accounts_from(df_bo, "Account")
    if not account_order:
        account_order = ordered_accounts_from(df_bva, "Account")
    # Comprehensive account set
    acct_keys = set()
    for m in maps.values():
        acct_keys.update([k for k in m.keys() if isinstance(k, str) and k.strip()])
    # Maintain order
    ordered_keys = [a for a in account_order if a in acct_keys] + [a for a in acct_keys if a not in account_order]

    # Assemble output
    cm_actual_name = f"{month_name} Actual"
    cm_budget_name = f"{month_name} Budget"
    cm_delta_name  = f"{month_name} Δ" if use_delta_symbol else f"{month_name} Delta"
    cm_pct_name    = f"{month_name} %"

    rows = []
    for acc in ordered_keys:
        if not is_account_label(acc):
            continue  # skip headers/totals from raw files
        row = {"Account": acc}
        ca = maps["cm_actual"].get(acc, 0.0)
        cb = maps["cm_budget"].get(acc, 0.0)
        row[cm_actual_name] = ca
        row[cm_budget_name] = cb
        row[cm_delta_name]  = ca - cb
        row[cm_pct_name]    = safe_pct(row[cm_delta_name], cb)

        ya = maps.get("ytd_actual", {}).get(acc, 0.0)
        yb = maps.get("ytd_budget", {}).get(acc, 0.0)
        ab = maps.get("annual_budget", {}).get(acc, 0.0)
        py = maps.get("prev_year", {}).get(acc, None)

        row["Actuals"]         = ya
        row["Budget to Date"]  = yb
        row["YTD Δ" if use_delta_symbol else "YTD Delta"] = ya - yb if (ya is not None and yb is not None) else None
        row["YTD %"]           = safe_pct(ya - yb, yb)

        row["Annual Budget"]   = ab
        row["Annual Δ" if use_delta_symbol else "Annual Delta"] = ya - ab if (ya is not None and ab is not None) else None
        row["Annual %"]        = safe_pct(ya - ab, ab)

        if col_prevyr:
            row["Previous Year"] = py

        rows.append(row)

    out_df = pd.DataFrame(rows)

    # Order columns
    desired = [
        "Account",
        cm_actual_name, cm_budget_name, cm_delta_name, cm_pct_name,
        "Actuals", "Budget to Date",
        "YTD Δ" if use_delta_symbol else "YTD Delta", "YTD %",
        "Annual Budget",
        "Annual Δ" if use_delta_symbol else "Annual Delta", "Annual %"
    ]
    if col_prevyr:
        desired += ["Previous Year"]
    # Append any other columns that slipped in
    remaining = [c for c in out_df.columns if c not in desired]
    out_df = out_df[[c for c in desired if c in out_df.columns] + remaining]

    # Replace very small numbers with blanks for neatness
    def blank_small(x):
        try:
            return "" if abs(float(x)) < 1e-9 else x
        except:
            return x

    for c in out_df.columns:
        out_df[c] = out_df[c].apply(blank_small)

    # Preview
    if show_preview:
        st.subheader("Preview")
        st.dataframe(out_df, use_container_width=True)

    # Build downloadable Excel
    currency_cols = [
        cm_actual_name, cm_budget_name, cm_delta_name,
        "YTD Δ" if use_delta_symbol else "YTD Delta",
        "Annual Budget",
        "Annual Δ" if use_delta_symbol else "Annual Delta",
        "Previous Year" if col_prevyr else ""
    ]
    percent_cols = [cm_pct_name, "YTD %", "Annual %"]

    xlsx_bytes = format_excel_bytes(out_df, currency_cols, percent_cols, sheet_name="Report", use_delta_symbol=use_delta_symbol)

    st.download_button(
        "⬇️ Download Consolidated Excel",
        data=xlsx_bytes,
        file_name=f"WAPA_SFA_Consolidated_{month_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Diagnostics
    with st.expander("Diagnostics"):
        st.write("Detected columns:")
        st.json({
            "Monthly": {"month": month_name, "actual_col": cm_actual_col, "budget_col": cm_budget_col},
            "BvA": {"Actuals": col_actuals, "Budget to Date": col_btd, "Annual Budget": col_annual},
            "Comp": {"Previous Year": col_prevyr},
            "FY start": fy_start
        })
        st.write("Row counts:")
        st.json({
            "SFA BvA": len(df_bva),
            "SFA BvA Month": len(df_bva_month),
            "SFA Comp": len(df_comp) if not df_comp.empty else 0,
            "Budget Overview (optional)": len(df_bo) if not df_bo.empty else 0,
            "Output": len(out_df),
        })
