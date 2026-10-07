"""Local unified search for bank-statement CSV, XLS and XLSX files.

Run with: streamlit run app.py
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
from dataclasses import dataclass
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Any, Iterable

import pandas as pd

import db

try:
    import streamlit as st
except ModuleNotFoundError:  # Enables parser verification before the UI dependency is installed.
    st = None


HEADER_ALIASES = {
    "date": ("transaction date", "txn date", "value date", "date"),
    "reference": (
        "utr", "utr no",
        "cheque. no./ref. no.", "cheque no.", "chq./ref.no.", "ref no",
        "reference no", "reference", "transaction id", "tran. id", "tran id",
    ),
    "description": ("transaction remarks", "description", "narration", "particulars", "remarks"),
    "debit": ("withdrawal amt", "withdrawal", "debit"),
    "credit": ("deposit amt", "deposit", "credit"),
    "balance": ("closing balance", "balance"),
}

ASSIGNMENTS_PATH = Path.home() / ".statementhub" / "party_assignments.json"


@dataclass
class ImportedStatement:
    filename: str
    bank: str
    transactions: pd.DataFrame
    original_columns: list[str]
    skipped_rows: int


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def header_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def party_match_key(description: str) -> str:
    text = header_key(description)
    text = re.sub(r"\b(?:upi|imps|neft|rtgs|inb|mmt|funds?|transfer|payment|paytm|phonepe|gpay)\b", " ", text)
    text = re.sub(r"\b(?:cnrbh\d{8,}|in\d{12,}|\d{12,})\b", " ", text)
    text = re.sub(r"\b\d{1,2}\s+\d{1,2}\s+\d{2,4}\b", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or header_key(description)


def transaction_key(row: dict[str, Any]) -> str:
    parts = [
        clean_text(row.get("Source file")),
        clean_text(row.get("Bank")),
        clean_text(row.get("Date")),
        clean_text(row.get("Amount")),
        clean_text(row.get("Type")),
        clean_text(row.get("UTR / Reference")),
        clean_text(row.get("Description")),
    ]
    return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()


def find_header_row(raw: pd.DataFrame) -> int:
    """Find a row containing a recognizable transaction-table header."""
    best_row, best_score = -1, -1
    for idx, row in raw.iterrows():
        cells = [header_key(v) for v in row.tolist()]
        score = 0
        for aliases in HEADER_ALIASES.values():
            if any(any(alias in cell for alias in aliases) for cell in cells):
                score += 1
        if score > best_score:
            best_row, best_score = idx, score
    if best_score < 3:
        raise ValueError("Could not find a transaction header row. The statement needs date, description, and amount/reference columns.")
    return int(best_row)


def choose_column(columns: Iterable[str], field: str) -> str | None:
    normalized = {column: header_key(column) for column in columns}
    for alias in HEADER_ALIASES[field]:
        for column, key in normalized.items():
            if alias in key:
                return column
    return None


def parse_amount(value: Any) -> float | None:
    text = clean_text(value).replace(",", "")
    if not text or text.lower() in {"nan", "none", "-", "--"}:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()₹Rs. ")
    try:
        amount = float(text)
        return -amount if negative else amount
    except ValueError:
        return None


def infer_bank(filename: str, raw: pd.DataFrame, header_row: int) -> str:
    # Only inspect the statement cover area. Transaction narration often mentions
    # counterparties at other banks and must not decide the source bank.
    sample = " ".join(clean_text(v) for v in raw.iloc[:header_row].astype(object).to_numpy().flatten()).lower()
    name = filename.lower()
    if "hdfc" in name or "hdfc" in sample:
        return "HDFC Bank"
    if "icici" in name or "icici" in sample or "icic" in sample:
        return "ICICI Bank"
    if "canara" in name or "canara" in sample or "cnrb" in sample:
        return "Canara Bank"
    if "axis" in name or "axis" in sample:
        return "Axis Bank"
    return "Unknown bank"


def read_statement(source: Any, filename: str) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()
    if suffix == ".csv":
        # Indian bank exports are commonly UTF-8, UTF-8 with BOM, or Windows-1252.
        data = source.getvalue() if hasattr(source, "getvalue") else Path(source).read_bytes()
        for encoding in ("utf-8-sig", "utf-8", "cp1252"):
            try:
                return pd.read_csv(io.BytesIO(data), header=None, dtype=object, encoding=encoding, engine="python")
            except UnicodeDecodeError:
                continue
        raise ValueError("Could not decode the CSV file.")
    if suffix not in {".xls", ".xlsx"}:
        raise ValueError("Only .csv, .xls, and .xlsx statements are supported.")
    engine = "xlrd" if suffix == ".xls" else "openpyxl"
    data = source.getvalue() if hasattr(source, "getvalue") else Path(source).read_bytes()
    return pd.read_excel(io.BytesIO(data), header=None, dtype=object, engine=engine)


def extract_identifier(row: pd.Series, reference_col: str | None, description_col: str | None) -> str:
    direct = clean_text(row.get(reference_col, "")) if reference_col else ""
    # Prefer a populated cheque/reference field. ICICI's internal Tran. Id (S123...) is
    # less useful than the UPI/NEFT identifier embedded in its remarks.
    if direct and direct.lower() not in {"0", "000000000000", "nan"} and not re.fullmatch(r"[SM]\d+", direct, re.I):
        return direct
    description = clean_text(row.get(description_col, "")) if description_col else ""
    # UTRs and UPI references are normally an all-numeric 12+ digit segment, or a
    # bank prefix such as IN followed by digits. Keep the complete matching segment.
    identifiers = re.findall(
        r"(?<![A-Za-z0-9])(?:CNRBH\d{8,}|IN\d{12,}|\d{12,})(?![A-Za-z0-9])",
        description,
        flags=re.I,
    )
    if identifiers:
        return identifiers[0]
    return direct if direct.lower() not in {"0", "000000000000", "nan"} else ""


def normalize_statement(source: Any, filename: str) -> ImportedStatement:
    raw = read_statement(source, filename)
    header_row = find_header_row(raw)
    columns = [clean_text(v) or f"Column {i + 1}" for i, v in enumerate(raw.iloc[header_row].tolist())]
    data = raw.iloc[header_row + 1:].copy()
    data.columns = columns
    data = data.dropna(axis=1, how="all")
    data = data.loc[:, ~data.columns.duplicated()].copy()

    date_col = choose_column(data.columns, "date")
    reference_col = choose_column(data.columns, "reference")
    description_col = choose_column(data.columns, "description")
    debit_col = choose_column(data.columns, "debit")
    credit_col = choose_column(data.columns, "credit")
    balance_col = choose_column(data.columns, "balance")
    if not date_col or not description_col or not (debit_col or credit_col):
        raise ValueError("The detected table is missing a date, description, or debit/credit column.")

    records: list[dict[str, Any]] = []
    for _, row in data.iterrows():
        date_value = pd.to_datetime(row.get(date_col), errors="coerce", dayfirst=True)
        debit = parse_amount(row.get(debit_col)) if debit_col else None
        credit = parse_amount(row.get(credit_col)) if credit_col else None
        description = clean_text(row.get(description_col))
        if pd.isna(date_value) or (debit is None and credit is None) or not description:
            continue
        amount = credit if credit is not None else debit
        transaction_type = "Credit" if credit is not None else "Debit"
        original = {column: clean_text(row.get(column)) for column in data.columns if clean_text(row.get(column))}
        records.append({
            "Date": date_value.strftime("%d-%m-%Y"),
            "Amount": amount,
            "Type": transaction_type,
            "UTR / Reference": extract_identifier(row, reference_col, description_col),
            "Description": description,
            "Party match key": party_match_key(description),
            "Balance": parse_amount(row.get(balance_col)) if balance_col else None,
            "Original columns": original,
        })

    bank = infer_bank(filename, raw, header_row)
    normalized = pd.DataFrame(records)
    if normalized.empty:
        raise ValueError("No transactions were found below the detected header row.")
    normalized.insert(0, "Bank", bank)
    normalized.insert(0, "Source file", filename)
    normalized["fingerprint"] = normalized.apply(lambda row: db.compute_transaction_fingerprint(row.to_dict()), axis=1)
    normalized["Transaction key"] = normalized["fingerprint"]
    normalized["Duplicate UTR"] = normalized["UTR / Reference"].ne("") & normalized["UTR / Reference"].duplicated(keep=False)
    return ImportedStatement(filename, bank, normalized, list(data.columns), len(data) - len(normalized))


def import_paths(paths: list[str]) -> pd.DataFrame:
    """Command-line verification helper; it also proves the supplied three formats import."""
    all_rows = []
    for path in paths:
        result = normalize_statement(path, Path(path).name)
        all_rows.append(result.transactions)
        print(f"{result.filename}: {len(result.transactions)} transactions, {result.skipped_rows} non-transaction rows")
    combined = pd.concat(all_rows, ignore_index=True)
    nonempty = combined[combined["UTR / Reference"] != ""]
    duplicate_count = nonempty["UTR / Reference"].duplicated(keep=False).sum()
    print(f"Total: {len(combined)} transactions; duplicate identifier rows: {duplicate_count}")
    return combined


def load_uploaded(files: list[Any]) -> tuple[pd.DataFrame, list[ImportedStatement], list[str]]:
    imported, errors = [], []
    for uploaded in files:
        try:
            imported.append(normalize_statement(uploaded, uploaded.name))
        except Exception as exc:  # Show a file-specific, actionable error in the UI.
            errors.append(f"{uploaded.name}: {exc}")
    if not imported:
        return pd.DataFrame(), [], errors
    combined = pd.concat([item.transactions for item in imported], ignore_index=True)
    ids = combined["UTR / Reference"]
    combined["Duplicate UTR"] = ids.ne("") & ids.duplicated(keep=False)
    return combined, imported, errors


def empty_assignments() -> dict[str, dict[str, str]]:
    return {"party_mappings": {}, "transaction_overrides": {}}


def load_assignments() -> dict[str, dict[str, str]]:
    if not ASSIGNMENTS_PATH.exists():
        return empty_assignments()
    try:
        data = json.loads(ASSIGNMENTS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return empty_assignments()
    return {
        "party_mappings": dict(data.get("party_mappings", {})),
        "transaction_overrides": dict(data.get("transaction_overrides", {})),
    }


def save_assignments(assignments: dict[str, dict[str, str]]) -> None:
    ASSIGNMENTS_PATH.parent.mkdir(parents=True, exist_ok=True)
    ASSIGNMENTS_PATH.write_text(json.dumps(assignments, indent=2, ensure_ascii=False), encoding="utf-8")


def apply_assignments(transactions: pd.DataFrame, assignments: dict[str, dict[str, str]]) -> pd.DataFrame:
    transactions = transactions.copy()
    mapped = transactions["Party match key"].map(assignments["party_mappings"]).fillna("")
    overrides = transactions["Transaction key"].map(assignments["transaction_overrides"]).fillna("")
    party_names = overrides.where(overrides.ne(""), mapped)
    transactions["Party Name"] = party_names
    transactions["Assigned Business / Party Name"] = party_names
    transactions["Name Source"] = ""
    transactions.loc[mapped.ne(""), "Name Source"] = "Sender/receiver mapping"
    transactions.loc[overrides.ne(""), "Name Source"] = "Transaction override"
    return transactions


def result_label(row: pd.Series) -> str:
    assigned = clean_text(row.get("Party Name")) or clean_text(row.get("Assigned Business / Party Name")) or "No assigned name"
    reference = clean_text(row.get("UTR / Reference")) or "No reference"
    amount = row.get("Amount")
    amount_text = f"{amount:,.2f}" if isinstance(amount, (int, float)) else clean_text(amount)
    return f"{row['Date']} | {row['Type']} | Rs {amount_text} | {assigned} | {reference}"


def export_excel(transactions: pd.DataFrame) -> bytes:
    export = transactions.copy()
    export = export.drop(
        columns=[
            "Original columns", "Original statement columns", "Party match key",
            "Transaction key", "Assigned Business / Party Name", "Name Source", "fingerprint",
        ],
        errors="ignore",
    )
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, index=False, sheet_name="Transactions")
    return output.getvalue()


def search_transactions(transactions: pd.DataFrame, query: str, mode: str) -> pd.DataFrame:
    """Search exactly one selected category: UTR/reference, date, or amount."""
    query = clean_text(query)
    if not query:
        return transactions.copy()

    if mode == "date":
        is_date = bool(re.fullmatch(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", query)) or bool(
            re.fullmatch(r"\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}", query)
        )
        if is_date:
            parsed_date = pd.to_datetime(query, dayfirst=True, errors="coerce")
            if not pd.isna(parsed_date):
                return transactions[transactions["Date"].eq(parsed_date.strftime("%d-%m-%Y"))].copy()
        return transactions[transactions["Date"].astype(str).str.contains(re.escape(query), case=False, na=False)].copy()

    if mode == "amount":
        amount = parse_amount(query)
        if amount is not None:
            return transactions[(transactions["Amount"] - amount).abs().le(0.005)].copy()
        return transactions[transactions["Amount"].astype(str).str.contains(re.escape(query), na=False)].copy()

    utr_mask = transactions["UTR / Reference"].fillna("").str.contains(re.escape(query), case=False, regex=True)
    return transactions[utr_mask].copy()


def render_backup_restore_section(combined: pd.DataFrame, assignments: dict[str, Any], sb_client: Any) -> None:
    with st.expander("💾 Backup & Restore Center"):
        st.markdown("**Create an offline backup or restore your database & party mappings.**")
        b_col1, b_col2 = st.columns(2)
        with b_col1:
            st.write("**Download Complete Backup**")
            st.caption(f"• **Transactions:** {len(combined):,} records\n• **Party Mappings:** {len(assignments.get('party_mappings', {}))} rules")
            if not combined.empty:
                txns_list = json.loads(combined.to_json(orient="records", date_format="iso"))
                backup_data = {
                    "version": "1.0",
                    "exported_at": datetime.now().isoformat(),
                    "total_transactions": len(combined),
                    "total_party_mappings": len(assignments.get("party_mappings", {})),
                    "party_mappings": assignments.get("party_mappings", {}),
                    "transaction_overrides": assignments.get("transaction_overrides", {}),
                    "transactions": txns_list,
                }
                backup_bytes = json.dumps(backup_data, indent=2, ensure_ascii=False).encode("utf-8")
                timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    "⬇️ Download System Backup (.json)",
                    data=backup_bytes,
                    file_name=f"statementhub_backup_{timestamp_str}.json",
                    mime="application/json",
                    help="Downloads all transactions and party mappings to a secure JSON file",
                )
            else:
                st.caption("No transactions available to backup yet.")

        with b_col2:
            st.write("**Restore Database from Backup**")
            restore_file = st.file_uploader(
                "Upload Backup File (.json)",
                type=["json"],
                key="restore_backup_file",
                help="Select a previously exported StatementHub backup JSON file to restore.",
            )
            if restore_file is not None:
                try:
                    loaded_backup = json.loads(restore_file.getvalue().decode("utf-8"))
                    restored_txns = loaded_backup.get("transactions", [])
                    restored_mappings = loaded_backup.get("party_mappings", {})
                    restored_overrides = loaded_backup.get("transaction_overrides", {})

                    st.info(f"Found **{len(restored_txns):,} transactions** and **{len(restored_mappings)} party mappings** in backup.")

                    if st.button("🚀 Restore Now", type="primary"):
                        assignments["party_mappings"].update(restored_mappings)
                        assignments["transaction_overrides"].update(restored_overrides)
                        save_assignments(assignments)

                        if sb_client:
                            for k, v in restored_mappings.items():
                                db.save_party_mapping(sb_client, k, v)
                            if restored_txns:
                                r_df = pd.DataFrame(restored_txns)
                                db.sync_transactions_with_dedup(sb_client, r_df)
                            st.session_state["cached_transactions"] = db.fetch_all_transactions(sb_client)
                        else:
                            if restored_txns:
                                st.session_state["cached_transactions"] = pd.DataFrame(restored_txns)

                        st.toast("Database and party mappings restored successfully!", icon="✅")
                        st.rerun()
                except Exception as exc:
                    st.error(f"Invalid backup file format: {exc}")


def app() -> None:
    if st is None:
        raise RuntimeError("The browser interface needs Streamlit. Run: pip install -r requirements.txt")
    st.set_page_config(page_title="Unified Bank Search", page_icon="🔎", layout="wide")
    sb_client = db.get_supabase_client()

    col_title, col_status = st.columns([4, 1], vertical_alignment="center")
    with col_title:
        st.title("Unified Bank Search")
    with col_status:
        if sb_client:
            if hasattr(st, "badge"):
                st.badge("🟢 Supabase Connected", color="green")
            else:
                st.caption("🟢 **Supabase Connected**")
        else:
            if hasattr(st, "badge"):
                st.badge("💾 Local Mode", color="gray")
            else:
                st.caption("💾 **Local Mode**")

    if not sb_client:
        with st.expander("☁️ Connect Supabase (Cloud Database)", expanded=False):
            st.markdown(
                "**Enable Cloud Persistence & Daily Deduplication:**\n\n"
                "1. Create a free project at [supabase.com](https://supabase.com)\n"
                "2. Run the SQL below in Supabase's **SQL Editor**\n"
                "3. Add your project URL & Key into `.streamlit/secrets.toml`"
            )
            st.code(db.SUPABASE_SCHEMA_SQL, language="sql")

    assignments = load_assignments()
    if sb_client:
        db_mappings = db.fetch_party_mappings(sb_client)
        if db_mappings:
            assignments["party_mappings"].update(db_mappings)

    files = st.file_uploader(
        "Add or replace statements", type=["csv", "xls", "xlsx"], accept_multiple_files=True,
        help="Select statements to upload. With Supabase, existing transactions are automatically skipped and only new ones are inserted.",
    )

    current_file_signature = [(f.name, getattr(f, "size", 0)) for f in files] if files else []

    # Process new uploads ONLY when the file set changes
    if files and st.session_state.get("last_uploaded_signature") != current_file_signature:
        uploaded_df, imports, errors = load_uploaded(files)
        for error in errors:
            st.error(error)
        if not uploaded_df.empty:
            if sb_client:
                new_count, skipped_count = db.sync_transactions_with_dedup(sb_client, uploaded_df)
                st.session_state["upload_msg"] = (
                    f"Processed {len(uploaded_df):,} transactions from {len(imports)} statement(s): "
                    f"**{new_count:,} new inserted into database**, **{skipped_count:,} duplicates skipped**."
                )
                st.session_state["cached_transactions"] = db.fetch_all_transactions(sb_client)
            else:
                st.session_state["upload_msg"] = f"Loaded {len(uploaded_df):,} transactions from {len(imports)} statement(s)."
                st.session_state["cached_transactions"] = uploaded_df
        st.session_state["last_uploaded_signature"] = current_file_signature

    if not files and not sb_client:
        st.session_state.pop("cached_transactions", None)
        st.session_state.pop("upload_msg", None)
        st.session_state.pop("last_uploaded_signature", None)

    if sb_client and ("cached_transactions" not in st.session_state or st.session_state["cached_transactions"].empty):
        st.session_state["cached_transactions"] = db.fetch_all_transactions(sb_client)

    combined = st.session_state.get("cached_transactions", pd.DataFrame())

    if "upload_msg" in st.session_state and files:
        st.success(st.session_state["upload_msg"])

    if combined.empty:
        if sb_client:
            st.info("No statements in database yet. Upload statements above to get started.")
        else:
            st.info("Select your bank statements to begin. The app recognizes HDFC-style, ICICI-style, and CSV statement layouts, plus similar variants.")
        render_backup_restore_section(combined, assignments, sb_client)
        return

    combined = apply_assignments(combined, assignments)

    # Filters: Bank Account & Time Period
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        all_banks = ["All Banks"] + sorted([b for b in combined["Bank"].dropna().unique().tolist() if b])
        selected_bank = st.selectbox("Bank Account", options=all_banks, index=0)
    with filter_col2:
        period_options = [
            "All Time",
            "This Month",
            "Last 30 Days",
            "Current Financial Year (Apr-Mar)",
            "Custom Date Range",
        ]
        selected_period = st.selectbox("Period", options=period_options, index=0)

    custom_start_date = None
    custom_end_date = None
    if selected_period == "Custom Date Range":
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            custom_start_date = st.date_input("From Date", value=None)
        with d_col2:
            custom_end_date = st.date_input("To Date", value=None)

    filtered_df = combined.copy()
    if selected_bank != "All Banks":
        filtered_df = filtered_df[filtered_df["Bank"] == selected_bank].copy()

    if selected_period != "All Time" and not filtered_df.empty:
        parsed_dates = pd.to_datetime(filtered_df["Date"], format="%d-%m-%Y", errors="coerce")
        valid_dates = parsed_dates.dropna()
        ref_date = valid_dates.max() if not valid_dates.empty else pd.Timestamp(datetime.now())

        if selected_period == "This Month":
            mask = (parsed_dates.dt.year == ref_date.year) & (parsed_dates.dt.month == ref_date.month)
            filtered_df = filtered_df[mask].copy()
        elif selected_period == "Last 30 Days":
            cutoff = ref_date - pd.Timedelta(days=30)
            mask = (parsed_dates >= cutoff) & (parsed_dates <= ref_date)
            filtered_df = filtered_df[mask].copy()
        elif selected_period == "Current Financial Year (Apr-Mar)":
            fy_start_year = ref_date.year if ref_date.month >= 4 else ref_date.year - 1
            fy_start = pd.Timestamp(year=fy_start_year, month=4, day=1)
            fy_end = pd.Timestamp(year=fy_start_year + 1, month=3, day=31, hour=23, minute=59, second=59)
            mask = (parsed_dates >= fy_start) & (parsed_dates <= fy_end)
            filtered_df = filtered_df[mask].copy()
        elif selected_period == "Custom Date Range":
            mask = pd.Series(True, index=filtered_df.index)
            if custom_start_date:
                mask &= (parsed_dates >= pd.Timestamp(custom_start_date))
            if custom_end_date:
                mask &= (parsed_dates <= pd.Timestamp(custom_end_date))
            filtered_df = filtered_df[mask].copy()

    search_modes = {"UTR / Reference": "utr", "Date": "date", "Amount": "amount"}
    selected_mode = st.radio("Search by", list(search_modes.keys()), horizontal=True)
    mode = search_modes[selected_mode]

    if mode == "utr":
        all_ids = sorted(filtered_df.loc[filtered_df["UTR / Reference"] != "", "UTR / Reference"].drop_duplicates().tolist())
        effective_query = st.selectbox(
            "Search UTR / Reference",
            options=all_ids,
            index=None,
            placeholder="Start typing a UTR (for example: 3964)",
            accept_new_options=True,
        ) or ""
    elif mode == "date":
        effective_query = st.text_input(
            "Search Date",
            placeholder="Enter date (e.g. 07-10-2026, 10-2026, or 2026)",
        ).strip()
    else:  # amount
        effective_query = st.text_input(
            "Search Amount",
            placeholder="Enter amount (e.g. 25000, 25,000, or ₹25,000)",
        ).strip()

    results = search_transactions(filtered_df, effective_query, mode)

    result_ids = results[results["UTR / Reference"] != ""]["UTR / Reference"]
    duplicated = result_ids[result_ids.duplicated(keep=False)].unique().tolist()
    if duplicated:
        st.warning(f"Duplicate identifier detected in these results: {', '.join(duplicated[:8])}" + (" …" if len(duplicated) > 8 else ""))
    col_res_header, col_res_opt = st.columns([2, 1])
    with col_res_header:
        st.subheader(f"Results ({len(results):,})")
    with col_res_opt:
        apply_reusable = st.toggle(
            "Apply to all matching transactions",
            value=True,
            help="When turned on, saving a party name will automatically apply to all past and future transactions with the same sender/receiver. When turned off, it applies only to that specific transaction.",
        )

    st.caption("💡 Enter or edit party names directly in the **Party Name** column below.")

    results = results.reset_index(drop=True)

    if "editor_version" not in st.session_state:
        st.session_state.editor_version = 0

    editor_key = f"transactions_editor_{st.session_state.editor_version}"

    display_columns = [
        "Bank", "Date", "Description", "Party Name", "Amount", "Type",
        "UTR / Reference", "Balance", "Duplicate UTR",
    ]
    edited_results = st.data_editor(
        results[display_columns],
        disabled=[col for col in display_columns if col != "Party Name"],
        use_container_width=True,
        hide_index=True,
        key=editor_key,
        column_config={
            "Amount": st.column_config.NumberColumn(format="₹ %0.2f"),
            "Balance": st.column_config.NumberColumn(format="₹ %0.2f"),
            "Party Name": st.column_config.TextColumn(
                "Party Name",
                help="Click to add or edit party name directly for this transaction",
            ),
        },
    )

    editor_state = st.session_state.get(editor_key, {})
    edited_rows = editor_state.get("edited_rows", {})

    changes_to_apply: dict[int, str] = {}
    if edited_rows:
        for idx_raw, col_edits in edited_rows.items():
            if "Party Name" in col_edits:
                changes_to_apply[int(idx_raw)] = clean_text(col_edits["Party Name"])
    else:
        diff = (
            edited_results["Party Name"].fillna("").astype(str).str.strip()
            != results["Party Name"].fillna("").astype(str).str.strip()
        )
        if diff.any():
            for idx in results.index[diff]:
                changes_to_apply[int(idx)] = clean_text(edited_results.loc[idx, "Party Name"])

    if changes_to_apply:
        updated_count = 0
        for idx, new_val in changes_to_apply.items():
            if idx in results.index:
                row = results.loc[idx]
                match_key = row["Party match key"]
                txn_key = row["Transaction key"]
                if new_val:
                    if apply_reusable:
                        assignments["party_mappings"][match_key] = new_val
                        assignments["transaction_overrides"].pop(txn_key, None)
                        if sb_client:
                            db.save_party_mapping(sb_client, match_key, new_val)
                    else:
                        assignments["transaction_overrides"][txn_key] = new_val
                        if sb_client:
                            db.save_transaction_override(sb_client, txn_key, new_val)
                else:
                    assignments["party_mappings"].pop(match_key, None)
                    assignments["transaction_overrides"].pop(txn_key, None)
                    if sb_client:
                        db.delete_party_mapping(sb_client, match_key)
                updated_count += 1

        save_assignments(assignments)
        if "cached_transactions" in st.session_state:
            st.session_state["cached_transactions"] = apply_assignments(st.session_state["cached_transactions"], assignments)
        st.session_state.editor_version += 1
        st.toast(f"Saved party name for {updated_count} transaction{'s' if updated_count > 1 else ''}.", icon="✅")
        st.rerun()

    with st.expander("Manage saved party mappings"):
        has_mappings = bool(assignments["party_mappings"] or assignments["transaction_overrides"])
        if not has_mappings:
            st.info("No party names have been assigned yet.")
        else:
            if assignments["party_mappings"]:
                st.write("**Sender / Receiver Mappings (Reusable)**")
                mappings_df = pd.DataFrame([
                    {"Narration Key": k, "Assigned Party Name": v}
                    for k, v in assignments["party_mappings"].items()
                ])
                st.dataframe(mappings_df, use_container_width=True, hide_index=True)
            if assignments["transaction_overrides"]:
                st.write(f"**Single-transaction overrides:** {len(assignments['transaction_overrides'])} saved.")
            if st.button("Clear all saved party mappings"):
                save_assignments(empty_assignments())
                if sb_client:
                    db.clear_all_party_mappings(sb_client)
                st.toast("Cleared all party mappings.", icon="🗑️")
                st.rerun()

    export = export_excel(results)
    st.download_button(
        "Download displayed results as Excel",
        export,
        "unified-search-results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    render_backup_restore_section(combined, assignments, sb_client)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--verify", nargs="+")
    args, _ = parser.parse_known_args()
    if args.verify:
        import_paths(args.verify)
    else:
        app()
