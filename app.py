import io
import os
import re
from datetime import date

import pandas as pd
import streamlit as st

try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False


# ==========================================
# 0. GENERIC COLUMN-MATCHING HELPERS
# ==========================================
def _find_first(columns, keywords, exclude_keywords=None):
    """Returns the first column whose lowercased name contains ANY of `keywords`
    and NONE of `exclude_keywords`. Fixes an operator-precedence bug in the
    original code where `"Main" in c or "Group" in c and "Sub" not in c` could
    accidentally match the wrong column."""
    exclude_keywords = exclude_keywords or []
    for c in columns:
        cl = c.lower()
        if any(k.lower() in cl for k in keywords) and not any(e.lower() in cl for e in exclude_keywords):
            return c
    return None


def _match_col(columns, candidates):
    """Exact (case-insensitive) match first, then substring match, across a list
    of acceptable header names. Used for the Day Book, which previously required
    exact Tally header text and broke on the slightest export variation."""
    for cand in candidates:
        for c in columns:
            if cand.lower() == c.lower():
                return c
    for cand in candidates:
        for c in columns:
            if cand.lower() in c.lower():
                return c
    return None


def _read_excel_fast(file_bytes):
    """python-calamine (Rust-based) reads large .xlsx files dramatically faster
    than the default openpyxl engine. Falls back automatically if it isn't
    installed or the file trips it up."""
    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine="calamine")
    except Exception:
        return pd.read_excel(io.BytesIO(file_bytes))


def get_file_bytes(uploaded_file, fallback_filename):
    """Prefers an uploaded file; falls back to a same-named .xlsx sitting next
    to app.py, preserving the original workflow for anyone who doesn't want to
    use the uploader."""
    if uploaded_file is not None:
        return uploaded_file.getvalue(), uploaded_file.name
    if os.path.exists(fallback_filename):
        with open(fallback_filename, "rb") as f:
            return f.read(), fallback_filename
    return None, None


# ==========================================
# 1. PARSING UNIFIED TRIAL BALANCE & MASTER GROUPS
# ==========================================
@st.cache_data(show_spinner=False)
def load_ledger_master(file_bytes, source_name):
    try:
        df = _read_excel_fast(file_bytes)
    except Exception as e:
        return pd.DataFrame(), f"Error reading Ledger Master ({source_name}): {e}"

    df.columns = df.columns.astype(str).str.strip()

    name_col = _find_first(df.columns, ["name", "particulars", "ledger"])
    sub_col = _find_first(df.columns, ["sub"])
    main_col = _find_first(df.columns, ["main"]) or _find_first(df.columns, ["group"], exclude_keywords=["sub"])

    if not name_col:
        return pd.DataFrame(), f"Ledger Master ({source_name}) needs a column indicating Ledger Name. Found columns: {list(df.columns)}"

    rename_dict = {name_col: "Party Name"}
    if sub_col:
        rename_dict[sub_col] = "Sub-Group"
    if main_col:
        rename_dict[main_col] = "Main Group"

    df = df.rename(columns=rename_dict)
    df["Party Name"] = df["Party Name"].astype(str).str.strip()

    # Guard against duplicate ledger entries silently picking an arbitrary row
    dupes = df["Party Name"][df["Party Name"].duplicated()].unique().tolist()

    return df, (f"⚠️ Duplicate ledger entries found in Ledger Master, first match used: {', '.join(dupes[:10])}" if dupes else None)


def _clean_amount(val_str):
    if pd.isna(val_str):
        return 0.0
    val_str = str(val_str).strip().replace(",", "")
    if not val_str or val_str == "nan":
        return 0.0
    match = re.search(r"(-?[\d\.]+)", val_str)
    return float(match.group(1)) if match else 0.0


def _clean_amount_series(s):
    """Vectorized version of _clean_amount for whole columns — this is the fast
    path used on the (usually huge) Day Book. Falls back to the slower regex
    parser only for the handful of rows that don't parse cleanly, instead of
    running it on every single row."""
    cleaned = s.astype(str).str.strip().str.replace(",", "", regex=False)
    numeric = pd.to_numeric(cleaned, errors="coerce")
    bad_mask = numeric.isna() & ~cleaned.isin(["", "nan"])
    if bad_mask.any():
        numeric.loc[bad_mask] = cleaned.loc[bad_mask].apply(_clean_amount)
    return numeric.fillna(0.0)


@st.cache_data(show_spinner=False)
def parse_trial_balance(file_bytes, source_name):
    try:
        df = _read_excel_fast(file_bytes)
    except Exception as e:
        return {}, f"Error reading Trial Balance ({source_name}): {e}"

    df.columns = df.columns.astype(str).str.strip()

    name_col = _find_first(df.columns, ["name", "particulars", "ledger"])
    # "Opening" takes priority; a generic "Balance" match is excluded if it's
    # actually the Closing column (original bug: op_col could silently grab
    # "Closing Balance" if it happened to appear before "Opening Balance").
    op_col = _find_first(df.columns, ["opening"]) or _find_first(df.columns, ["balance"], exclude_keywords=["closing"])
    cl_col = _find_first(df.columns, ["closing"])

    if not name_col or not op_col:
        return {}, f"Could not find required columns in Trial Balance ({source_name}). Found columns: {list(df.columns)}"

    parsed_data = {}
    skipped_dupes = []
    for _, row in df.iterrows():
        name = str(row[name_col]).strip()
        if not name or name == "nan" or "Total" in name:
            continue
        if name in parsed_data:
            skipped_dupes.append(name)
        parsed_data[name] = {
            "Opening": _clean_amount(row[op_col]),
            "File_Closing": _clean_amount(row[cl_col]) if cl_col else 0.0,
        }

    warning = f"⚠️ Duplicate rows in Trial Balance, last one used: {', '.join(sorted(set(skipped_dupes))[:10])}" if skipped_dupes else None
    return parsed_data, warning


# ==========================================
# 2. PARSING TALLY'S DAY BOOK WITH DATE FILTERING
# ==========================================
@st.cache_data(show_spinner=False)
def parse_day_book(file_bytes, source_name, start_date=None, end_date=None):
    try:
        df = _read_excel_fast(file_bytes)
    except Exception as e:
        return pd.DataFrame(), f"Error reading Day Book ({source_name}): {e}"

    df.columns = df.columns.astype(str).str.strip()

    date_col = _match_col(df.columns, ["Date"])
    vch_type_col = _match_col(df.columns, ["Vch Type", "Voucher Type"])
    vch_no_col = _match_col(df.columns, ["Vch No.", "Vch No", "Voucher No.", "Voucher No"])
    particulars_col = _match_col(df.columns, ["Particulars", "Party Name", "Ledger Name", "Ledger"])
    dr_col = _match_col(df.columns, ["Debit Amount", "Debit"])
    cr_col = _match_col(df.columns, ["Credit Amount", "Credit"])

    missing = [label for label, col in [("Date", date_col), ("Particulars", particulars_col)] if not col]
    if not dr_col and not cr_col:
        missing.append("Debit/Credit Amount")
    if missing:
        return pd.DataFrame(), f"Day Book ({source_name}) is missing required column(s): {', '.join(missing)}. Found columns: {list(df.columns)}"

    # Vectorized replacement for the old row-by-row Python loop. Tally's Day Book
    # only prints Date / Vch Type / Vch No. on the FIRST row of each voucher and
    # leaves it blank on continuation lines — that's a forward-fill, and pandas
    # can do the whole file's worth in one vectorized pass instead of a Python
    # for-loop over every row.
    df["_Date"] = pd.to_datetime(df[date_col], errors="coerce").ffill()

    if vch_type_col:
        vch_type_series = df[vch_type_col].astype(str).str.strip()
        header_mask = vch_type_series.replace({"nan": ""}) != ""
        vch_type_series = vch_type_series.where(header_mask, None).ffill().fillna("Adjustment")
    else:
        header_mask = pd.Series(False, index=df.index)
        vch_type_series = pd.Series("Adjustment", index=df.index)

    if vch_no_col:
        vch_no_series = df[vch_no_col].astype(str).str.strip().replace({"nan": ""})
        vch_no_series = vch_no_series.where(header_mask, None).ffill().fillna("")
    else:
        vch_no_series = pd.Series("", index=df.index)

    particulars_series = df[particulars_col].astype(str).str.strip()

    dr_amt = _clean_amount_series(df[dr_col]) if dr_col else pd.Series(0.0, index=df.index)
    cr_amt = _clean_amount_series(df[cr_col]) if cr_col else pd.Series(0.0, index=df.index)

    result = pd.DataFrame({
        "Date": df["_Date"],
        "Party Name": particulars_series,
        "Voucher Type": vch_type_series,
        "_VchNo": vch_no_series,
        "Debit Amount": dr_amt,
        "Credit Amount": cr_amt,
    })

    mask = result["Date"].notna()
    mask &= ~particulars_series.isin(["", "nan"])
    mask &= (dr_amt > 0) | (cr_amt > 0)
    if start_date:
        mask &= result["Date"].dt.date >= start_date
    if end_date:
        mask &= result["Date"].dt.date <= end_date

    result = result.loc[mask].copy()
    result["Ref Number"] = (result["Voucher Type"] + "-" + result["_VchNo"]).where(result["_VchNo"] != "", "JV")
    result = result.drop(columns=["_VchNo"]).reset_index(drop=True)

    return result, None


# ==========================================
# 3. CHRONOLOGICAL FIFO TRACKING ENGINE & AGING
# ==========================================
def run_ledger_engine(party_name, op_bal, party_txs, engine_mode, base_date, end_date):
    """Single-pass FIFO settlement engine. Replaces the old pair of functions
    (calculate_aging_split_rows + get_aging_data), which each re-ran the exact
    same due/clearance matching independently — doubling the work for every
    ledger. `party_txs` must already be this party's day-book rows, sorted by
    Date (the caller pre-groups the day book once instead of filtering it
    fresh for every ledger).

    Returns: (invoice_split_df, aging_buckets, outstanding_items)
    """
    dues = []
    clearances = []

    if engine_mode == "debit_dominant":
        if op_bal > 0:
            dues.append({"Ref": "Opening Balance (Dr)", "Date": base_date, "Amount": op_bal, "Remaining": op_bal, "Settlements": []})
        elif op_bal < 0:
            clearances.append({"Date": base_date, "Amount": abs(op_bal), "Type": "Opening Advance (Cr)"})
    else:  # credit_dominant
        if op_bal < 0:
            dues.append({"Ref": "Opening Balance (Cr)", "Date": base_date, "Amount": abs(op_bal), "Remaining": abs(op_bal), "Settlements": []})
        elif op_bal > 0:
            clearances.append({"Date": base_date, "Amount": op_bal, "Type": "Opening Advance (Dr)"})

    if party_txs is not None and not party_txs.empty:
        # Iterate as plain (unnamed) tuples in a fixed column order — faster than
        # iterrows() and avoids itertuples()'s fragile auto-renaming of columns
        # with spaces (e.g. "Debit Amount") into positional "_N" attributes.
        cols = party_txs[["Date", "Voucher Type", "Ref Number", "Debit Amount", "Credit Amount"]]
        for tx_date, tx_vtype, tx_ref, tx_dr, tx_cr in cols.itertuples(index=False, name=None):
            if engine_mode == "debit_dominant":
                if tx_dr > 0:
                    dues.append({"Ref": tx_ref, "Date": tx_date, "Amount": tx_dr, "Remaining": tx_dr, "Settlements": []})
                if tx_cr > 0:
                    clearances.append({"Date": tx_date, "Amount": tx_cr, "Type": tx_vtype})
            else:  # credit_dominant
                if tx_cr > 0:
                    dues.append({"Ref": tx_ref, "Date": tx_date, "Amount": tx_cr, "Remaining": tx_cr, "Settlements": []})
                if tx_dr > 0:
                    clearances.append({"Date": tx_date, "Amount": tx_dr, "Type": tx_vtype})

    for cl in clearances:
        cl_remaining = cl["Amount"]
        cl_date = cl["Date"]
        cl_type = cl["Type"]

        for due in dues:
            if cl_remaining <= 0:
                break
            if due["Remaining"] > 0:
                allocated = min(cl_remaining, due["Remaining"])
                cl_remaining -= allocated
                due["Remaining"] -= allocated
                is_final = due["Remaining"] == 0
                due["Settlements"].append({"Date": cl_date, "Amount": allocated, "Type": cl_type, "IsFinalSettlement": is_final})
        cl["Remaining"] = cl_remaining

    # --- Invoice-wise settlement breakdown (party detail / audit view) ---
    split_results = []
    for due in dues:
        if not due["Settlements"]:
            split_results.append({
                "Invoice/Ref Number": due["Ref"],
                "Invoice Date": due["Date"].strftime("%Y-%m-%d") if pd.notna(due["Date"]) else "-",
                "Invoice Value": due["Amount"],
                "Settlement Date": "-",
                "Settled Amount": 0.0,
                "Settled Via": "-",
                "Remaining Outstanding": due["Remaining"],
                "Days to Settle": "N/A",
                "Highlight": False
            })
        else:
            for stl in due["Settlements"]:
                days_taken = max(0, (stl["Date"] - due["Date"]).days)
                split_results.append({
                    "Invoice/Ref Number": due["Ref"],
                    "Invoice Date": due["Date"].strftime("%Y-%m-%d") if pd.notna(due["Date"]) else "-",
                    "Invoice Value": due["Amount"],
                    "Settlement Date": stl["Date"].strftime("%Y-%m-%d"),
                    "Settled Amount": stl["Amount"],
                    "Settled Via": stl["Type"],
                    "Remaining Outstanding": due["Remaining"],
                    "Days to Settle": days_taken,
                    "Highlight": stl["IsFinalSettlement"]
                })
    split_df = pd.DataFrame(split_results)

    # --- Aging buckets (same due/clearance state, no re-simulation) ---
    buckets = {"0-90 Days": 0.0, "91-120 Days": 0.0, "121-180 Days": 0.0, "> 180 Days": 0.0}
    end_dt_pd = pd.to_datetime(end_date)
    outstanding_items = []

    total_due_remaining = sum(d["Remaining"] for d in dues)
    total_cl_remaining = sum(c["Remaining"] for c in clearances)

    if total_due_remaining > 0:
        for due in dues:
            if due["Remaining"] > 0:
                age = max(0, (end_dt_pd - due["Date"]).days)
                if age <= 90: buckets["0-90 Days"] += due["Remaining"]
                elif age <= 120: buckets["91-120 Days"] += due["Remaining"]
                elif age <= 180: buckets["121-180 Days"] += due["Remaining"]
                else: buckets["> 180 Days"] += due["Remaining"]
                outstanding_items.append({
                    "Reference": due["Ref"], "Date": due["Date"].strftime("%Y-%m-%d"),
                    "Age (Days)": age, "Outstanding Amount": due["Remaining"]
                })
    elif total_cl_remaining > 0:
        for cl in clearances:
            if cl["Remaining"] > 0:
                age = max(0, (end_dt_pd - cl["Date"]).days)
                if age <= 90: buckets["0-90 Days"] += cl["Remaining"]
                elif age <= 120: buckets["91-120 Days"] += cl["Remaining"]
                elif age <= 180: buckets["121-180 Days"] += cl["Remaining"]
                else: buckets["> 180 Days"] += cl["Remaining"]
                outstanding_items.append({
                    "Reference": cl.get("Ref", cl["Type"]), "Date": cl["Date"].strftime("%Y-%m-%d"),
                    "Age (Days)": age, "Outstanding Amount": cl["Remaining"]
                })

    return split_df, buckets, outstanding_items


CASH_BANK_TERMS = ["cash", "bank"]


@st.cache_data(show_spinner="⚙️ Running the reconciliation & aging engine...")
def compute_ledger_engine(scoped_parties, trial_balance_dict, group_lookup, day_book_df, base_date, end_date, skip_cash_bank_fifo):
    """Runs the FIFO engine for every scoped ledger exactly once and caches the
    result. Because this is cached on its actual inputs (not on unrelated UI
    state like the search box or sort dropdown), typing in the search field or
    changing the sort order no longer re-runs the whole engine — Streamlit
    just returns the cached DataFrame instantly.

    `group_lookup` is a dict {party_name: (main_group, sub_group)}.
    """
    # Group the Day Book by party ONCE instead of re-filtering the full frame
    # for every single ledger (that was O(parties x transactions) before).
    if not day_book_df.empty:
        day_book_groups = {name: grp.sort_values("Date") for name, grp in day_book_df.groupby("Party Name", sort=False)}
    else:
        day_book_groups = {}
    empty_txs = day_book_df.iloc[0:0]

    summary_records = []
    aging_records = []
    buyer_cycles = []
    supplier_cycles = []
    supplier_raw_points = []
    buyer_raw_points = []

    for party in scoped_parties:
        party_info = trial_balance_dict[party]
        p_op = party_info["Opening"]

        current_main_group, current_sub_group = group_lookup.get(party, ("Unmapped Main Group", "Unmapped Sub-Group"))

        is_creditor = any(term in current_sub_group.lower() or term in current_main_group.lower() for term in ["creditor", "payable", "liability", "liabilities"])
        engine_strategy = "credit_dominant" if is_creditor else "debit_dominant"

        is_cash_bank = any(term in current_sub_group.lower() or term in current_main_group.lower() for term in CASH_BANK_TERMS)

        party_txs = day_book_groups.get(party, empty_txs)
        daybook_debit = party_txs["Debit Amount"].sum() if not party_txs.empty else 0.0
        daybook_credit = party_txs["Credit Amount"].sum() if not party_txs.empty else 0.0

        run_fifo = not (skip_cash_bank_fifo and is_cash_bank)

        if run_fifo:
            p_report, buckets, out_items = run_ledger_engine(party, p_op, party_txs, engine_strategy, base_date, end_date)
        else:
            # Cash/Bank ledgers: usually the highest transaction volume in a
            # Day Book but invoice-level "aging" is meaningless for them —
            # skip the FIFO simulation entirely and just use the totals.
            p_report, buckets, out_items = pd.DataFrame(), None, None

        if not p_report.empty:
            settled_rows = p_report[
                (p_report["Days to Settle"] != "N/A") &
                (~p_report["Invoice/Ref Number"].str.contains("Opening Balance", case=False, na=False))
            ]
            avg_days = settled_rows["Days to Settle"].mean() if not settled_rows.empty else 0.0

            if not settled_rows.empty:
                inv_dates = pd.to_datetime(settled_rows["Invoice Date"], errors="coerce")
                pts = pd.DataFrame({
                    "Date_Key": inv_dates,
                    "Month_Sort": inv_dates.dt.strftime("%Y-%m"),
                    "Days": settled_rows["Days to Settle"].astype(float).values
                }).dropna(subset=["Date_Key"])
                target_list = supplier_raw_points if engine_strategy == "credit_dominant" else buyer_raw_points
                target_list.extend(pts.to_dict("records"))
        else:
            avg_days = 0.0

        if engine_strategy == "debit_dominant" and avg_days > 0:
            buyer_cycles.append(avg_days)
        elif engine_strategy == "credit_dominant" and avg_days > 0:
            supplier_cycles.append(avg_days)

        calculated_closing = p_op + daybook_debit - daybook_credit
        variance = calculated_closing - party_info["File_Closing"]

        # --- Aging Analysis Classification Engine (skipped for Cash/Bank) ---
        cat = None
        if run_fifo:
            if engine_strategy == "debit_dominant":
                if calculated_closing > 0.01:
                    cat = "Debtors with Debit balance"
                elif calculated_closing < -0.01:
                    cat = "Debtors with Credit balance"
            else:
                if calculated_closing < -0.01:
                    cat = "Creditors with Credit balance"
                elif calculated_closing > 0.01:
                    cat = "Creditors with Debit balance"

        if cat and buckets is not None:
            aging_records.append({
                "Category Metric": cat,
                "Party Name": party,
                "Closing Balance": abs(calculated_closing),
                "0-90 Days": buckets["0-90 Days"],
                "91-120 Days": buckets["91-120 Days"],
                "121-180 Days": buckets["121-180 Days"],
                "> 180 Days": buckets["> 180 Days"],
                "Items": out_items
            })

        summary_records.append({
            "Party Name": party,
            "Main Group": current_main_group,
            "Sub-Group": current_sub_group,
            "Settlement": round(avg_days, 1),
            "Opening Bal": p_op,
            "Debits": daybook_debit,
            "Credits": daybook_credit,
            "Closing (Calc)": calculated_closing,
            "Closing (Excel)": party_info["File_Closing"],
            "Variance": round(variance, 2),
            "Engine Strategy": engine_strategy,
            "Volume": daybook_debit + daybook_credit
        })

    master_summary_df = pd.DataFrame(summary_records)
    return master_summary_df, aging_records, buyer_cycles, supplier_cycles, supplier_raw_points, buyer_raw_points


# ==========================================
# 4. EXPORT HELPERS (EXCEL & PDF)
# ==========================================
def build_excel_report(summary_df=None, aging_df=None, extra_sheets=None):
    """Builds a multi-sheet Excel workbook from whichever pieces are actually
    passed in. `summary_df`/`aging_df` are the full-scope dashboard/aging
    tables — only pass these when the user explicitly asked for the full
    report. `extra_sheets` is a list of (sheet_name, dataframe) tuples for
    single-ledger exports, so a "download this party's statement" button
    never has to drag the whole ledger dashboard along with it."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if summary_df is not None and not summary_df.empty:
            summary_df.drop(columns=["Engine Strategy", "Volume"], errors="ignore").to_excel(
                writer, index=False, sheet_name="Ledger Summary"
            )
        if aging_df is not None and not aging_df.empty:
            aging_df.drop(columns=["Items"], errors="ignore").to_excel(
                writer, index=False, sheet_name="Aging Summary"
            )
        for sheet_name, df in (extra_sheets or []):
            if df is None or df.empty:
                continue
            safe_name = re.sub(r"[\[\]\:\*\?/\\]", "_", sheet_name)[:31]
            df.drop(columns=["Highlight"], errors="ignore").to_excel(
                writer, index=False, sheet_name=safe_name
            )

        # Auto-size columns for readability
        for sheet in writer.sheets.values():
            for col_cells in sheet.columns:
                length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
                sheet.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 40)

    buffer.seek(0)
    return buffer


def _plain_currency(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return str(val)
    return f"Rs. {v:,.2f}"


def build_summary_pdf(filtered_df, totals, start_dt, end_dt):
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Ledger Reconciliation Summary", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 7, f"Period: {start_dt} to {end_dt}", ln=True)
    pdf.ln(2)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Totals", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, f"Opening Balance: {_plain_currency(totals['op'])}", ln=True)
    pdf.cell(0, 6, f"Total Debits: {_plain_currency(totals['dr'])}", ln=True)
    pdf.cell(0, 6, f"Total Credits: {_plain_currency(totals['cr'])}", ln=True)
    pdf.cell(0, 6, f"Closing (Calculated): {_plain_currency(totals['cl'])}", ln=True)
    pdf.cell(0, 6, f"Net Variance: {_plain_currency(totals['var'])}", ln=True)
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Ledger-wise Detail", ln=True)
    headers = ["Party", "Main Group", "Sub-Group", "Opening", "Debits", "Credits", "Closing", "Variance"]
    widths = [55, 35, 35, 28, 28, 28, 28, 24]

    def draw_header():
        pdf.set_font("Helvetica", "B", 8)
        for h, w in zip(headers, widths):
            pdf.cell(w, 7, h, border=1)
        pdf.ln()
        pdf.set_font("Helvetica", "", 7)

    draw_header()
    for _, row in filtered_df.iterrows():
        if pdf.get_y() > 190:
            pdf.add_page()
            draw_header()
        values = [
            str(row["Party Name"])[:32],
            str(row["Main Group"])[:20],
            str(row["Sub-Group"])[:20],
            f"{row['Opening Bal']:,.0f}",
            f"{row['Debits']:,.0f}",
            f"{row['Credits']:,.0f}",
            f"{row['Closing (Calc)']:,.0f}",
            f"{row['Variance']:,.0f}",
        ]
        for v, w in zip(values, widths):
            pdf.cell(w, 6, v, border=1)
        pdf.ln()

    return bytes(pdf.output())


# ==========================================
# 5. STREAMLIT FRONTEND & INTERACTIVE UI
# ==========================================
st.set_page_config(page_title="Tracker", layout="wide")
st.title("📊 Tracker")

st.markdown(
    """
    <style>
    ::-webkit-scrollbar { width: 12px !important; height: 12px !important; }
    ::-webkit-scrollbar-track { background: #f1f1f1 !important; }
    ::-webkit-scrollbar-thumb { background: #c1c1c1 !important; border-radius: 4px !important; }
    [data-testid="stDataFrame"] canvas { cursor: pointer; }
    .block-container { padding-top: 1.5rem !important; padding-bottom: 1.5rem !important; }
    div[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: bold; }
    .total-box { background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 10px; border-radius: 5px; margin-top: 10px; text-align: center; font-size: 14px; }
    </style>
    """,
    unsafe_allow_html=True
)


def highlight_final_clearance(row):
    if row.get("Highlight"):
        return ["background-color: #d4edda; color: #155724; font-weight: bold;"] * len(row)
    return [""] * len(row)


def format_indian_currency(val):
    try:
        val_float = float(val)
    except (TypeError, ValueError):
        return str(val)

    is_negative = val_float < 0
    val_float = abs(val_float)

    s = f"{val_float:,.2f}"
    parts = s.split('.')
    main_part = parts[0].replace(",", "")

    if len(main_part) <= 3:
        formatted = "₹" + s
    else:
        last_three = main_part[-3:]
        remaining = main_part[:-3]
        remaining_grouped = re.sub(r'(\d)(?=(\d\d)+(?!\d))', r'\1,', remaining)
        formatted = f"₹{remaining_grouped},{last_three}.{parts[1]}"

    return f"-{formatted}" if is_negative else formatted


# ----------------------------------------
# SIDEBAR: DATA SOURCES & SETTINGS
# ----------------------------------------
with st.sidebar:
    st.header("📁 Data Sources")
    st.caption(
        "Upload your files here, or leave a box empty to auto-load "
        "**Trial_Balance.xlsx / Day_book.xlsx / Ledger_Master.xlsx** "
        "if they're sitting next to app.py (old workflow still works)."
    )
    tb_upload = st.file_uploader("Trial Balance", type=["xlsx", "xls"], key="tb_upload")
    db_upload = st.file_uploader("Day Book", type=["xlsx", "xls"], key="db_upload")
    lm_upload = st.file_uploader("Ledger Master", type=["xlsx", "xls"], key="lm_upload")

    st.markdown("---")
    st.header("⚙️ Settings")
    start_dt = st.date_input("Start Date:", value=date(2025, 4, 1))
    end_dt = st.date_input("End Date:", value=date(2026, 3, 31))
    opening_bal_date = st.date_input(
        "Opening Balances As Of:",
        value=start_dt,
        help="The date your Trial Balance's Opening column is struck as of. "
             "Used as the anchor date for aging opening balances."
    )

    st.markdown("---")
    st.header("⚡ Performance")
    skip_cash_bank_fifo = st.checkbox(
        "Skip invoice-level aging for Cash/Bank ledgers",
        value=True,
        help="Cash and Bank ledgers are usually your highest-volume ledgers in the Day Book, "
             "but invoice-by-invoice FIFO aging isn't meaningful for them anyway. Skipping the "
             "detailed engine for these can significantly speed up large files. Their Opening/"
             "Debits/Credits/Closing totals are still calculated normally."
    )

days_delta = (end_dt - start_dt).days
base_date = pd.to_datetime(opening_bal_date)

# ----------------------------------------
# RESOLVE FILE SOURCES
# ----------------------------------------
tb_bytes, tb_source = get_file_bytes(tb_upload, "Trial_Balance.xlsx")
db_bytes, db_source = get_file_bytes(db_upload, "Day_book.xlsx")
lm_bytes, lm_source = get_file_bytes(lm_upload, "Ledger_Master.xlsx")

missing_sources = []
if tb_bytes is None:
    missing_sources.append("Trial Balance")
if db_bytes is None:
    missing_sources.append("Day Book")
if lm_bytes is None:
    missing_sources.append("Ledger Master")

if missing_sources:
    st.info(
        f"👋 Waiting on: **{', '.join(missing_sources)}**. "
        f"Upload them in the sidebar, or place `Trial_Balance.xlsx`, `Day_book.xlsx`, "
        f"and `Ledger_Master.xlsx` next to `app.py` and refresh."
    )
    st.stop()

trial_balance_dict, tb_warning = parse_trial_balance(tb_bytes, tb_source)
master_groups_df, lm_warning = load_ledger_master(lm_bytes, lm_source)
day_book_df, db_warning = parse_day_book(db_bytes, db_source, start_date=start_dt, end_date=end_dt)

for warning in [tb_warning, lm_warning, db_warning]:
    if warning:
        (st.error if "Error" in warning else st.warning)(warning)

if not trial_balance_dict:
    st.error("No usable rows found in the Trial Balance. Please check the file and try again.")
    st.stop()

st.caption(f"Loaded — Trial Balance: `{tb_source}` · Day Book: `{db_source}` · Ledger Master: `{lm_source}`")

# ----------------------------------------
# BUILD WORKING REGISTRY (LEDGER -> GROUP MAPPING)
# ----------------------------------------
all_trial_parties = list(trial_balance_dict.keys())
base_records = []
for party in all_trial_parties:
    matching_meta = master_groups_df[master_groups_df["Party Name"] == party] if not master_groups_df.empty else pd.DataFrame()

    sub_grp = matching_meta["Sub-Group"].values[0] if not matching_meta.empty and "Sub-Group" in master_groups_df.columns else "Unmapped Sub-Group"
    main_grp = matching_meta["Main Group"].values[0] if not matching_meta.empty and "Main Group" in master_groups_df.columns else "Unmapped Main Group"

    base_records.append({
        "Party Name": party,
        "Sub-Group": str(sub_grp).strip(),
        "Main Group": str(main_grp).strip()
    })

working_registry_df = pd.DataFrame(base_records)
unmapped_count = (working_registry_df["Main Group"] == "Unmapped Main Group").sum()
if unmapped_count:
    st.warning(f"⚠️ {unmapped_count} ledger(s) from the Trial Balance were not found in the Ledger Master and are grouped as 'Unmapped'.")

st.markdown("### 🗂️ Filter Hierarchy Configuration")
filter_col1, filter_col2, filter_col3 = st.columns([2, 2, 1])

with filter_col1:
    unique_main_groups = sorted(working_registry_df["Main Group"].unique())
    selected_main_groups = st.multiselect("Main Groups:", unique_main_groups, default=[])

with filter_col2:
    if selected_main_groups:
        filtered_sub_df = working_registry_df[working_registry_df["Main Group"].isin(selected_main_groups)]
    else:
        filtered_sub_df = working_registry_df

    unique_sub_groups = sorted(filtered_sub_df["Sub-Group"].unique())
    selected_sub_groups = st.multiselect("Sub-Groups:", unique_sub_groups, default=[])

with filter_col3:
    st.write("")
    st.write("")
    only_variance = st.checkbox("Only show variances", value=False, help="Show only ledgers where Closing (Calculated) doesn't match Closing (Excel).")

final_scoped_df = working_registry_df.copy()
if selected_main_groups:
    final_scoped_df = final_scoped_df[final_scoped_df["Main Group"].isin(selected_main_groups)]
if selected_sub_groups:
    final_scoped_df = final_scoped_df[final_scoped_df["Sub-Group"].isin(selected_sub_groups)]

scoped_parties = sorted(final_scoped_df["Party Name"].unique())

if not scoped_parties:
    st.warning("⚠️ No ledgers match the selected Main Group / Sub-Group filter combination.")
else:
    group_lookup = {
        row["Party Name"]: (row["Main Group"], row["Sub-Group"])
        for _, row in final_scoped_df.drop_duplicates("Party Name").iterrows()
    }

    master_summary_df, aging_records, buyer_cycles, supplier_cycles, supplier_raw_points, buyer_raw_points = compute_ledger_engine(
        tuple(scoped_parties), trial_balance_dict, group_lookup, day_book_df, base_date, end_dt, skip_cash_bank_fifo
    )

    def generate_timeline_chart_df(raw_datapoints, is_daily_mode):
        if not raw_datapoints:
            if is_daily_mode:
                idx = pd.date_range(start=start_dt, end=end_dt)
                return pd.DataFrame(index=idx, columns=["Days"]).fillna(0)
            else:
                idx = pd.date_range(start=start_dt, end=end_dt, freq='ME').strftime('%Y-%m')
                df_empty = pd.DataFrame(index=idx, columns=["Days"]).fillna(0)
                return df_empty

        tdf = pd.DataFrame(raw_datapoints)

        if is_daily_mode:
            grouped = tdf.groupby("Date_Key")["Days"].mean().reset_index()
            idx = pd.date_range(start=start_dt, end=end_dt)
            grouped = grouped.set_index("Date_Key").reindex(idx).fillna(0)
            return grouped[["Days"]]
        else:
            grouped = tdf.groupby("Month_Sort")["Days"].mean().reset_index()
            idx = pd.date_range(start=start_dt, end=end_dt, freq='ME').strftime('%Y-%m')
            grouped = grouped.set_index("Month_Sort").reindex(idx).fillna(0)
            return grouped[["Days"]]

    is_daily = (days_delta <= 31)

    # ----------------------------------------
    # TABBED INTERFACE SYSTEM
    # ----------------------------------------
    tab1, tab2, tab3 = st.tabs(["📊 Dashboard & Reconciliation", "⏳ Aging Analysis", "📅 Ledger Trends"])

    # =====================================================================
    # TAB 1: DASHBOARD & RECONCILIATION
    # =====================================================================
    with tab1:
        st.subheader("🏢 Company-Level Performance Metrics")
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            company_collection_days = sum(buyer_cycles) / len(buyer_cycles) if buyer_cycles else 0.0
            st.metric("Avg Collection Days (Buyers/Debtors)", f"{company_collection_days:.1f} Days")
        with m2:
            company_repayment_days = sum(supplier_cycles) / len(supplier_cycles) if supplier_cycles else 0.0
            st.metric("Avg Repayment Days (Suppliers/Creditors)", f"{company_repayment_days:.1f} Days")
        with m3:
            st.metric("Ledgers in Scope", f"{len(master_summary_df)}")
        with m4:
            mismatch_count = int((master_summary_df["Variance"].abs() > 0.05).sum())
            st.metric("Ledgers with Variance", f"{mismatch_count}")

        st.markdown("---")
        st.subheader("🏆 Commercial Volume & Aging Leaders")
        top_col1, top_col2 = st.columns(2)

        with top_col1:
            st.markdown("#### 🟥 Suppliers Timeline Summary & Top 10 Leaders")
            supp_chart_df = generate_timeline_chart_df(supplier_raw_points, is_daily)
            st.bar_chart(supp_chart_df, use_container_width=True, color="#d9534f")

            suppliers_df = master_summary_df[master_summary_df["Engine Strategy"] == "credit_dominant"].nlargest(10, "Volume")
            if not suppliers_df.empty:
                supp_display = suppliers_df[["Party Name", "Debits", "Credits", "Settlement"]].copy()
                supp_display["Debits"] = supp_display["Debits"].apply(format_indian_currency)
                supp_display["Credits"] = supp_display["Credits"].apply(format_indian_currency)
                supp_display["Settlement"] = supp_display["Settlement"].apply(lambda x: f"{x} Days" if x > 0 else "N/A")
                supp_display = supp_display.rename(columns={
                    "Party Name": "Supplier Name", "Debits": "Total Debits", "Credits": "Total Credits", "Settlement": "Avg Repayment Days"
                })
                st.dataframe(supp_display, use_container_width=True, hide_index=True)
            else:
                st.caption("No supplier match metrics available.")

        with top_col2:
            st.markdown("#### 🟦 Buyers Timeline Summary & Top 10 Leaders")
            buyer_chart_df = generate_timeline_chart_df(buyer_raw_points, is_daily)
            st.bar_chart(buyer_chart_df, use_container_width=True, color="#0275d8")

            buyers_df = master_summary_df[master_summary_df["Engine Strategy"] == "debit_dominant"].nlargest(10, "Volume")
            if not buyers_df.empty:
                buyer_display = buyers_df[["Party Name", "Debits", "Credits", "Settlement"]].copy()
                buyer_display["Debits"] = buyer_display["Debits"].apply(format_indian_currency)
                buyer_display["Credits"] = buyer_display["Credits"].apply(format_indian_currency)
                buyer_display["Settlement"] = buyer_display["Settlement"].apply(lambda x: f"{x} Days" if x > 0 else "N/A")
                buyer_display = buyer_display.rename(columns={
                    "Party Name": "Buyer Name", "Debits": "Total Debits", "Credits": "Total Credits", "Settlement": "Avg Collection Days"
                })
                st.dataframe(buyer_display, use_container_width=True, hide_index=True)
            else:
                st.caption("No buyer match metrics available.")

        st.markdown("---")
        st.subheader("📋 Scoped Ledger Overview & Reconciliation Dashboard")

        ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([2, 2, 1])
        with ctrl_col1:
            search_query = st.text_input("🔍 Search Within Filtered Ledgers:", "").strip()
        with ctrl_col2:
            sort_by = st.selectbox("↕️ Sort Table By:", [
                "Party Name", "Main Group", "Sub-Group", "Settlement",
                "Opening Bal", "Closing (Calc)", "Variance"
            ])
        with ctrl_col3:
            sort_order = st.radio("Order:", ["Ascending", "Descending"], horizontal=True)

        filtered_df = master_summary_df.copy()
        if search_query:
            filtered_df = filtered_df[filtered_df["Party Name"].str.contains(search_query, case=False, na=False)]
        if only_variance:
            filtered_df = filtered_df[filtered_df["Variance"].abs() > 0.05]

        ascending_bool = True if sort_order == "Ascending" else False
        filtered_df = filtered_df.sort_values(by=sort_by, ascending=ascending_bool).reset_index(drop=True)

        total_op = filtered_df["Opening Bal"].sum()
        total_dr = filtered_df["Debits"].sum()
        total_cr = filtered_df["Credits"].sum()
        total_cl_calc = filtered_df["Closing (Calc)"].sum()
        total_variance = filtered_df["Variance"].sum()

        display_summary_df = filtered_df.copy()
        display_summary_df["Variance_Raw"] = display_summary_df["Variance"]
        display_summary_df["Opening Bal"] = display_summary_df["Opening Bal"].apply(format_indian_currency)
        display_summary_df["Debits"] = display_summary_df["Debits"].apply(format_indian_currency)
        display_summary_df["Credits"] = display_summary_df["Credits"].apply(format_indian_currency)
        display_summary_df["Closing (Calc)"] = display_summary_df["Closing (Calc)"].apply(format_indian_currency)
        display_summary_df["Closing (Excel)"] = display_summary_df["Closing (Excel)"].apply(format_indian_currency)
        display_summary_df["Variance"] = display_summary_df["Variance"].apply(format_indian_currency)
        display_summary_df["Settlement"] = display_summary_df["Settlement"].apply(lambda x: f"{x} Days" if x > 0 else "N/A")

        display_summary_df["Variance"] = display_summary_df.apply(
            lambda r: f"⚠️ {r['Variance']}" if abs(r["Variance_Raw"]) > 0.05 else r["Variance"], axis=1
        )
        cols_to_show = [c for c in display_summary_df.columns if c not in ["Engine Strategy", "Volume", "Variance_Raw"]]

        st.markdown("_Click anywhere on a ledger row below to view its granular item-wise chronological breakdown. A ⚠️ next to Variance flags a mismatch between calculated and file closing balances._")

        st.dataframe(
            display_summary_df[cols_to_show],
            use_container_width=True,
            selection_mode="single-row",
            on_select="rerun",
            key="master_dashboard",
            hide_index=True,
            column_config={
                "Party Name": st.column_config.TextColumn("Party Name", width="medium"),
                "Main Group": st.column_config.TextColumn("Main Group", width="small"),
                "Sub-Group": st.column_config.TextColumn("Sub-Group", width="small"),
                "Settlement": st.column_config.TextColumn("Avg Days", width="small"),
                "Opening Bal": st.column_config.TextColumn("Opening Balance", width="small"),
                "Debits": st.column_config.TextColumn("Debits", width="small"),
                "Credits": st.column_config.TextColumn("Credits", width="small"),
                "Closing (Calc)": st.column_config.TextColumn("Closing (Calc)", width="small"),
                "Closing (Excel)": st.column_config.TextColumn("Closing (Excel)", width="small"),
                "Variance": st.column_config.TextColumn("Variance", width="small"),
            }
        )

        t_col1, t_col2, t_col3, t_col4, t_col5 = st.columns(5)
        with t_col1:
            st.markdown(f"<div class='total-box'><b>Total Opening Balance</b><br>{format_indian_currency(total_op)}</div>", unsafe_allow_html=True)
        with t_col2:
            st.markdown(f"<div class='total-box'><b>Total Debits (Daybook)</b><br>{format_indian_currency(total_dr)}</div>", unsafe_allow_html=True)
        with t_col3:
            st.markdown(f"<div class='total-box'><b>Total Credits (Daybook)</b><br>{format_indian_currency(total_cr)}</div>", unsafe_allow_html=True)
        with t_col4:
            st.markdown(f"<div class='total-box'><b>Total Closing (Calculated)</b><br>{format_indian_currency(total_cl_calc)}</div>", unsafe_allow_html=True)
        with t_col5:
            st.markdown(f"<div class='total-box'><b>Net Scope Variance</b><br>{format_indian_currency(total_variance)}</div>", unsafe_allow_html=True)

        # --------------- EXPORTS ---------------
        st.markdown("#### ⬇️ Export This View")
        exp_col1, exp_col2 = st.columns(2)
        with exp_col1:
            excel_buffer = build_excel_report(
                summary_df=filtered_df,
                aging_df=pd.DataFrame(aging_records) if aging_records else None
            )
            st.download_button(
                "📥 Download Excel (Summary + Aging)",
                data=excel_buffer,
                file_name=f"Ledger_Report_{start_dt}_to_{end_dt}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with exp_col2:
            if FPDF_AVAILABLE:
                pdf_bytes = build_summary_pdf(
                    filtered_df,
                    {"op": total_op, "dr": total_dr, "cr": total_cr, "cl": total_cl_calc, "var": total_variance},
                    start_dt, end_dt
                )
                st.download_button(
                    "📄 Download PDF Summary",
                    data=pdf_bytes,
                    file_name=f"Ledger_Summary_{start_dt}_to_{end_dt}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                st.caption("PDF export needs `fpdf2` — add it to requirements.txt and reinstall.")

        selected_party = None
        sel_state = st.session_state.get("master_dashboard")
        if sel_state and sel_state.get("selection", {}).get("rows"):
            selected_row_idx = sel_state["selection"]["rows"][0]
            if selected_row_idx < len(filtered_df):
                selected_party = filtered_df.iloc[selected_row_idx]["Party Name"]

        if selected_party:
            st.markdown("---")
            st.subheader(f"🔍 Detailed Transaction History: {selected_party}")

            party_info = filtered_df[filtered_df["Party Name"] == selected_party].iloc[0]
            op_val = party_info["Opening Bal"]
            engine_strat = party_info["Engine Strategy"]

            audit_col1, audit_col2, audit_col3 = st.columns(3)
            with audit_col1:
                st.metric("Opening Balance (From File)", format_indian_currency(op_val))
            with audit_col2:
                st.metric("Closing Balance (From File)", format_indian_currency(party_info["Closing (Excel)"]))
            with audit_col3:
                var_val = party_info["Variance"]
                status_label = "✅ Balanced" if abs(var_val) < 0.05 else "⚠️ Variance Found"
                st.metric(f"Audit Status: {status_label}", format_indian_currency(var_val), delta="Zero variance target")

            selected_party_txs = day_book_df[day_book_df["Party Name"] == selected_party].sort_values("Date") if not day_book_df.empty else day_book_df
            report_df, _, _ = run_ledger_engine(selected_party, op_val, selected_party_txs, engine_strat, base_date, end_dt)

            if not report_df.empty:
                formatted_df = report_df.copy()
                formatted_df["Invoice Value"] = formatted_df["Invoice Value"].apply(format_indian_currency)
                formatted_df["Settled Amount"] = formatted_df["Settled Amount"].apply(format_indian_currency)
                formatted_df["Remaining Outstanding"] = formatted_df["Remaining Outstanding"].apply(format_indian_currency)

                display_df = formatted_df.style.apply(highlight_final_clearance, axis=1)
                st.dataframe(display_df, use_container_width=True, column_config={"Highlight": None}, hide_index=True)

                st.download_button(
                    f"📥 Download {selected_party}'s Ledger Statement (Excel)",
                    data=build_excel_report(extra_sheets=[(f"{selected_party} Statement", report_df)]),
                    file_name=f"{selected_party}_Statement.xlsx".replace("/", "-"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="party_excel_dl",
                )

                settled_rows = report_df[
                    (report_df["Days to Settle"] != "N/A") &
                    (~report_df["Invoice/Ref Number"].str.contains("Opening Balance", case=False, na=False))
                ]
                if not settled_rows.empty:
                    avg_days = settled_rows["Days to Settle"].mean()
                    st.metric(label="Average Days to Settle Transactions (Excluding Opening Balance)", value=f"{avg_days:.1f} Days")
                else:
                    st.metric(label="Average Days to Settle Transactions", value="N/A")
            else:
                st.write("No transaction activity history for this ledger.")
        else:
            st.info("💡 Tip: Click on any row in the table above to view its granular invoice-by-invoice aging breakdown.")

    # =====================================================================
    # TAB 2: AGING ANALYSIS DASHBOARD
    # =====================================================================
    with tab2:
        st.subheader("⏳ Category-Level Aging Summary")
        st.markdown("_Select a Category Metric row below to view its Ledger breakdown._")

        if aging_records:
            aging_df = pd.DataFrame(aging_records)

            cat_df = aging_df.groupby("Category Metric")[["Closing Balance", "0-90 Days", "91-120 Days", "121-180 Days", "> 180 Days"]].sum().reset_index()

            all_cats = pd.DataFrame({"Category Metric": [
                "Debtors with Debit balance",
                "Debtors with Credit balance",
                "Creditors with Credit balance",
                "Creditors with Debit balance"
            ]})

            cat_df = pd.merge(all_cats, cat_df, on="Category Metric", how="left").fillna(0)

            formatted_cat_df = cat_df.copy()
            for col in ["Closing Balance", "0-90 Days", "91-120 Days", "121-180 Days", "> 180 Days"]:
                formatted_cat_df[col] = formatted_cat_df[col].apply(format_indian_currency)

            st.dataframe(
                formatted_cat_df,
                use_container_width=True,
                hide_index=True,
                selection_mode="single-row",
                on_select="rerun",
                key="aging_category_sel"
            )

            aging_excel_buffer = build_excel_report(summary_df=master_summary_df, aging_df=aging_df)
            st.download_button(
                "📥 Download Full Aging Analysis (Excel)",
                data=aging_excel_buffer,
                file_name=f"Aging_Analysis_{start_dt}_to_{end_dt}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="aging_excel_dl",
            )

            sel_cat_state = st.session_state.get("aging_category_sel", {})
            if sel_cat_state and sel_cat_state.get("selection", {}).get("rows"):
                idx = sel_cat_state["selection"]["rows"][0]
                if idx < len(cat_df):
                    sel_cat_name = cat_df.iloc[idx]["Category Metric"]

                    st.markdown("---")
                    st.subheader(f"Ledger Level Breakdown: {sel_cat_name}")
                    st.markdown("_Select a Ledger row below to view its itemized breakdown._")

                    ledger_df = aging_df[aging_df["Category Metric"] == sel_cat_name].copy()
                    if not ledger_df.empty:
                        disp_ledger = ledger_df.drop(columns=["Category Metric", "Items"]).reset_index(drop=True)
                        fmt_ledger = disp_ledger.copy()
                        for col in ["Closing Balance", "0-90 Days", "91-120 Days", "121-180 Days", "> 180 Days"]:
                            fmt_ledger[col] = fmt_ledger[col].apply(format_indian_currency)

                        st.dataframe(
                            fmt_ledger,
                            use_container_width=True,
                            hide_index=True,
                            selection_mode="single-row",
                            on_select="rerun",
                            key="aging_ledger_sel"
                        )

                        sel_led_state = st.session_state.get("aging_ledger_sel", {})
                        if sel_led_state and sel_led_state.get("selection", {}).get("rows"):
                            led_idx = sel_led_state["selection"]["rows"][0]
                            if led_idx < len(disp_ledger):
                                sel_party = disp_ledger.iloc[led_idx]["Party Name"]

                                st.markdown("---")
                                st.subheader(f"🧾 Outstanding Items: {sel_party}")
                                items = ledger_df[ledger_df["Party Name"] == sel_party].iloc[0]["Items"]
                                if items:
                                    items_df = pd.DataFrame(items)
                                    items_df["Outstanding Amount"] = items_df["Outstanding Amount"].apply(format_indian_currency)
                                    st.dataframe(items_df, use_container_width=True, hide_index=True)
                                else:
                                    st.info("No outstanding items found for this ledger.")
                    else:
                        st.info(f"No ledgers currently match the '{sel_cat_name}' status.")
        else:
            st.write("No aging data available based on current filters.")

    # =====================================================================
    # TAB 3: LEDGER TRENDS (MONTH-WISE COLLECTION / REPAYMENT DAYS)
    # =====================================================================
    with tab3:
        st.subheader("📅 Month-wise Collection / Repayment Trend")
        st.markdown("_Pick a ledger to see how its average settlement days moved month to month over the selected period._")

        trend_party_list = sorted(master_summary_df["Party Name"].unique())
        if not trend_party_list:
            st.info("No ledgers in the current scope.")
        else:
            trend_party = st.selectbox("Select Ledger:", trend_party_list, key="trend_party_select")

            trend_info = master_summary_df[master_summary_df["Party Name"] == trend_party].iloc[0]
            trend_op = trend_info["Opening Bal"]
            trend_strategy = trend_info["Engine Strategy"]
            metric_label = "Collection Days" if trend_strategy == "debit_dominant" else "Repayment Days"
            chart_color = "#0275d8" if trend_strategy == "debit_dominant" else "#d9534f"

            trend_party_txs = (
                day_book_df[day_book_df["Party Name"] == trend_party].sort_values("Date")
                if not day_book_df.empty else day_book_df
            )
            trend_report_df, _, _ = run_ledger_engine(trend_party, trend_op, trend_party_txs, trend_strategy, base_date, end_dt)

            settled = pd.DataFrame()
            if not trend_report_df.empty:
                settled = trend_report_df[
                    (trend_report_df["Days to Settle"] != "N/A") &
                    (~trend_report_df["Invoice/Ref Number"].str.contains("Opening Balance", case=False, na=False))
                ].copy()

            if not trend_report_df.empty:
                st.download_button(
                    f"📥 Download {trend_party}'s Ledger Statement (Excel)",
                    data=build_excel_report(extra_sheets=[(f"{trend_party} Statement", trend_report_df)]),
                    file_name=f"{trend_party}_Statement.xlsx".replace("/", "-"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="trend_statement_excel_dl",
                )

            if settled.empty:
                st.info(f"No settled transactions found for **{trend_party}** in the selected period — nothing to trend yet.")
            else:
                settled["Invoice Date"] = pd.to_datetime(settled["Invoice Date"])
                settled["Month"] = settled["Invoice Date"].dt.strftime("%Y-%m")
                settled["Days to Settle"] = settled["Days to Settle"].astype(float)

                monthly = settled.groupby("Month").agg(
                    Avg_Days=("Days to Settle", "mean"),
                    Transactions=("Days to Settle", "count"),
                    Total_Settled=("Settled Amount", "sum")
                ).reset_index()

                # Fill in months with no settlement activity so the chart doesn't silently skip gaps
                full_months = pd.date_range(start=start_dt, end=end_dt, freq="MS").strftime("%Y-%m")
                monthly = monthly.set_index("Month").reindex(full_months).reset_index().rename(columns={"index": "Month"})
                monthly["Avg_Days"] = monthly["Avg_Days"].fillna(0.0)
                monthly["Transactions"] = monthly["Transactions"].fillna(0).astype(int)
                monthly["Total_Settled"] = monthly["Total_Settled"].fillna(0.0)

                m_col1, m_col2 = st.columns(2)
                with m_col1:
                    overall_avg = settled["Days to Settle"].mean()
                    st.metric(f"Overall Avg {metric_label}", f"{overall_avg:.1f} Days")
                with m_col2:
                    st.metric("Total Settled Transactions", f"{int(settled.shape[0])}")

                st.markdown(f"#### {metric_label} by Month — {trend_party}")
                chart_df = monthly.set_index("Month")[["Avg_Days"]].rename(columns={"Avg_Days": metric_label})
                st.bar_chart(chart_df, use_container_width=True, color=chart_color)

                disp_monthly = monthly.copy()
                disp_monthly["Avg_Days"] = disp_monthly["Avg_Days"].apply(lambda x: f"{x:.1f} Days" if x > 0 else "-")
                disp_monthly["Total_Settled"] = disp_monthly["Total_Settled"].apply(format_indian_currency)
                disp_monthly = disp_monthly.rename(columns={
                    "Avg_Days": f"Avg {metric_label}",
                    "Transactions": "Settled Txns",
                    "Total_Settled": "Total Settled Value"
                })
                st.dataframe(disp_monthly, use_container_width=True, hide_index=True)

                trend_export_df = monthly.rename(columns={"Avg_Days": f"Avg {metric_label}", "Total_Settled": "Total Settled Value"})

                st.download_button(
                    f"📥 Download {trend_party}'s Monthly Trend (Excel)",
                    data=build_excel_report(extra_sheets=[(f"{trend_party} Monthly Trend", trend_export_df)]),
                    file_name=f"{trend_party}_Monthly_Trend.xlsx".replace("/", "-"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="trend_excel_dl",
                )
