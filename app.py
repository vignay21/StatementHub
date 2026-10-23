"""Local unified search for bank-statement CSV, XLS and XLSX files.

Run with: streamlit run app.py
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Any, Iterable

import pandas as pd

import db

try:
    import streamlit as st
    import streamlit.components.v1 as components
    import altair as alt
except ModuleNotFoundError:  # Enables parser verification before the UI dependency is installed.
    st = None
    components = None
    alt = None


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
PARTY_STORE_PATH = Path.home() / ".statementhub" / "stored_parties.json"


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
    text = re.sub(r"\s+", " ", str(value)).strip()
    if text.startswith('="') and text.endswith('"'):
        text = text[2:-1].strip()
    elif text.startswith("="):
        text = text[1:].strip().strip('"\'')
    return text


def format_inr(val: Any, prefix: str = "₹ ") -> str:
    """Format a number into Indian currency format (e.g. ₹ 62,41,515.00)."""
    if val is None or pd.isna(val):
        return f"{prefix}0.00"
    try:
        num = float(val)
    except (ValueError, TypeError):
        return f"{prefix}0.00"

    is_neg = num < 0
    num = abs(num)
    s = f"{num:.2f}"
    int_part, dec_part = s.split(".")

    if len(int_part) <= 3:
        formatted_int = int_part
    else:
        last_three = int_part[-3:]
        remaining = int_part[:-3]
        formatted_remaining = re.sub(r"(\d)(?=(\d{2})+(?!\d))", r"\1,", remaining)
        formatted_int = f"{formatted_remaining},{last_three}"

    sign = "-" if is_neg else ""
    return f"{sign}{prefix}{formatted_int}.{dec_part}"


def detect_internal_transfers(data: pd.DataFrame) -> pd.Series:
    """Automatically detect internal self money transfers between accounts.

    Criteria:
    1. Party Name set to 'SELF TRANSFER', 'SELF', 'CONTRA', or 'OWN ACCOUNT'.
    2. Same UTR appearing as Debit in Bank A and Credit in Bank B.
    """
    if data.empty:
        return pd.Series(False, index=data.index)

    # 1. Party Name condition
    party_str = data["Party Name"].fillna("").astype(str).str.strip().str.lower()
    party_mask = party_str.str.contains(r"\bself\b|\bcontra\b|own account|self transfer", regex=True)

    # 2. Cross-bank matching UTR with both Credit & Debit
    utr_mask = pd.Series(False, index=data.index)
    utr_series = data["UTR / Reference"].fillna("").astype(str).str.strip()
    valid_utr_df = data[utr_series != ""].copy()

    if not valid_utr_df.empty:
        contra_utrs: set[str] = set()
        for utr, group in valid_utr_df.groupby("UTR / Reference"):
            if len(group) > 1:
                banks = group["Bank"].dropna().unique()
                types = group["Type"].dropna().unique()
                if len(banks) > 1 and "Credit" in types and "Debit" in types:
                    contra_utrs.add(utr)

        if contra_utrs:
            utr_mask = data["UTR / Reference"].isin(contra_utrs)

    # 3. Description keyword condition (e.g. 'SELF TRANSFER', 'TRANSFER TO SELF', 'OWN A/C')
    desc_str = data["Description"].fillna("").astype(str).str.strip().str.lower()
    desc_mask = desc_str.str.contains(
        r"\bself\s*transfer\b|\btransfer\s+to\s+self\b|\bown\s+acc(?:ount)?\b|\bto\s+own\s+a/?c\b",
        regex=True,
    )

    return party_mask | utr_mask | desc_mask


def detect_fd_investments(data: pd.DataFrame) -> pd.Series:
    """Automatically detect Fixed Deposits and Capital Investments.

    Criteria:
    1. Party Name contains FD, FIXED DEPOSIT, TERM DEPOSIT, INVESTMENT, or DIGITALFD.
    2. Description/Narration contains DIGITALFD, FIXED DEPOSIT, TERM DEPOSIT, FD THROUGH, FD RENEW, etc.
    """
    if data.empty:
        return pd.Series(False, index=data.index)

    party_str = data["Party Name"].fillna("").astype(str).str.strip().str.lower()
    party_mask = party_str.str.contains(r"\bfd\b|fixed deposit|term deposit|investment|digitalfd", regex=True)

    desc_str = data["Description"].fillna("").astype(str).str.strip().str.lower()
    desc_mask = desc_str.str.contains(
        r"digitalfd|fixed deposit|term deposit|\bfd through\b|\bfd renew\b|\bfd clos\b|\btd renew\b|\btd clos\b",
        regex=True,
    )

    return party_mask | desc_mask


def header_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def party_match_key(description: str) -> str:
    raw = clean_text(description).strip()
    if not raw:
        return ""

    # 1. Slash format: UPI/..., CLG/..., INET-IMPS-..., MOB-IMPS-...
    if "/" in raw:
        parts = [p.strip() for p in raw.split("/") if p.strip()]
        if parts:
            p0_upper = parts[0].upper()
            if "UPI" in p0_upper:
                # Format A: UPI/CR(or DR)/12-digit RRN/PARTY/BANK/VPA/Remark//hash/date
                if len(parts) >= 4 and parts[1].upper() in ("CR", "DR"):
                    name = parts[3]
                    if name.lower() == "bank acco":
                        acc = re.sub(r"[^a-z0-9]+", " ", parts[5].lower()).strip() if len(parts) > 5 else ""
                        rem = re.sub(r"[^a-z0-9]+", " ", parts[6].lower()).strip() if len(parts) > 6 else ""
                        rem = re.sub(r"\b(?:payment|advance|dr|cr)\b", "", rem).strip()
                        key = f"bank acco {acc} {rem}".strip()
                        return re.sub(r"\s+", " ", key)
                    return header_key(name)
                # Format B: UPI/PARTY/VPA/Payment fr/Bank/RRN/hash
                elif len(parts) >= 2:
                    first = parts[1]
                    if first.isdigit():
                        vpa_parts = [p for p in parts if "@" in p]
                        if vpa_parts:
                            vpa_prefix = vpa_parts[0].split("@")[0]
                            return header_key(vpa_prefix)
                    else:
                        return header_key(first)

            # Cheque clearing: CLG/PARTY/CHEQUE_NO/BANK/...
            elif p0_upper == "CLG" and len(parts) >= 2:
                if parts[1].upper() != "CHEQUE":
                    return header_key(parts[1])

            # IMPS: INET-IMPS-CR/PARTY/...
            elif "IMPS" in p0_upper and len(parts) >= 2:
                for part in parts[1:]:
                    if not part.isdigit() and not re.match(r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", part):
                        return header_key(part)

    # 2. Hyphen format: UPI-..., IMPS-..., RTGS-..., NEFT-...
    if "-" in raw:
        parts = [p.strip() for p in raw.split("-") if p.strip()]
        if len(parts) >= 2:
            p0_upper = parts[0].upper()
            if "UPI" in p0_upper or "IMPS" in p0_upper or "RTGS" in p0_upper or "NEFT" in p0_upper:
                for candidate in parts[1:]:
                    if (
                        not candidate.isdigit()
                        and "@" not in candidate
                        and not re.match(r"^[A-Z]{4}0[A-Z0-9]{6}$", candidate, re.IGNORECASE)
                        and not re.match(r"^[A-Za-z0-9]{12,24}$", candidate)
                    ):
                        return header_key(candidate)

    # 3. NEFT pattern: MB NEFT Dr <ref> <PARTY> <IFSC> <acc> <remark>
    neft_match = re.search(r"\b(?:mb\s+neft|neft)\s+(?:dr|cr)?\s*(?:cnrbh\w+|\w+\d+)?\s+([a-z\s]+?)\s+[a-z]{4}\d{7}", raw, re.IGNORECASE)
    if neft_match:
        party = neft_match.group(1).strip()
        if party:
            return header_key(party)

    # 4. Handle pre-cleaned or space-delimited UPI text:
    text = header_key(raw)
    masked_split = re.split(r"\b(?:xx\w*|\*\*\w*)\b", text)
    if len(masked_split) > 1 and masked_split[0].strip():
        prefix = masked_split[0].strip()
        prefix = re.sub(r"\b(?:upi|imps|neft|rtgs|cr|dr)\b", " ", prefix).strip()
        if len(prefix) >= 3:
            return prefix

    # 5. Generic sanitization: strip transaction hashes, dates, RRNs, keywords
    text = re.sub(r"\b(?:upi|imps|neft|rtgs|inb|mmt|funds?|transfer|payment|paytm|phonepe|gpay|dr|cr)\b", " ", text)
    text = re.sub(r"\b(?:cnrbh\d{8,}|in\d{12,}|\d{12,})\b", " ", text)
    text = re.sub(r"\b\d{1,2}\s+\d{1,2}\s+\d{2,4}\b", " ", text)
    text = re.sub(r"\b(?:axl|ybl|ptm|paz|ibl|ici|hdf|smy|axb|uti|citi|sbin|cnrb)[a-z0-9]{16,}\b", " ", text)
    text = re.sub(r"\b[a-z0-9]{22,}\b", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or header_key(raw)


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


def get_gemini_api_key() -> str:
    """Retrieve Gemini API key from session state, secrets.toml, or environment variable."""
    if st and "gemini_api_key" in st.session_state and st.session_state["gemini_api_key"]:
        return str(st.session_state["gemini_api_key"]).strip()
    try:
        if st and hasattr(st, "secrets"):
            if "gemini" in st.secrets and "api_key" in st.secrets["gemini"]:
                return str(st.secrets["gemini"]["api_key"]).strip()
            if "GEMINI_API_KEY" in st.secrets:
                return str(st.secrets["GEMINI_API_KEY"]).strip()
            if "gemini_api_key" in st.secrets:
                return str(st.secrets["gemini_api_key"]).strip()
    except Exception:
        pass
    sec_file = Path(".streamlit/secrets.toml")
    if sec_file.exists():
        try:
            try:
                import tomllib
                with open(sec_file, "rb") as f:
                    data = tomllib.load(f)
            except ImportError:
                import toml
                data = toml.load(sec_file)
            if "gemini" in data and "api_key" in data["gemini"]:
                return str(data["gemini"]["api_key"]).strip()
            if "GEMINI_API_KEY" in data:
                return str(data["GEMINI_API_KEY"]).strip()
        except Exception:
            pass
    return os.environ.get("GEMINI_API_KEY", "").strip()


def save_gemini_api_key(api_key: str) -> bool:
    """Persist Gemini API key into .streamlit/secrets.toml and session state."""
    clean_k = api_key.strip()
    if not clean_k:
        return False
    if st:
        st.session_state["gemini_api_key"] = clean_k
    secrets_path = Path(".streamlit/secrets.toml")
    try:
        data = {}
        if secrets_path.exists():
            try:
                import toml
                data = toml.load(secrets_path)
            except Exception:
                pass
        if "gemini" not in data:
            data["gemini"] = {}
        data["gemini"]["api_key"] = clean_k
        import toml
        secrets_path.parent.mkdir(parents=True, exist_ok=True)
        with open(secrets_path, "w", encoding="utf-8") as f:
            toml.dump(data, f)
        return True
    except Exception:
        return True


def extract_transactions_from_pdf_with_gemini(source: Any, filename: str) -> pd.DataFrame:
    """Extract bank transactions from a PDF file using Gemini multimodal AI."""
    api_key = get_gemini_api_key()
    if not api_key:
        raise ValueError(
            f"Cannot process PDF '{filename}' because Gemini API key is missing. "
            "Please enter your Gemini API key in the 'Gemini AI Settings' expander above or add it to .streamlit/secrets.toml."
        )

    pdf_bytes = source.getvalue() if hasattr(source, "getvalue") else Path(source).read_bytes()
    if not pdf_bytes:
        raise ValueError(f"The PDF file '{filename}' is empty.")

    try:
        from google import genai
        from google.genai import types
    except ImportError:
        raise RuntimeError("Google GenAI library is not installed. Run: pip install google-genai")

    client = genai.Client(api_key=api_key)

    prompt = (
        "You are an expert banking statement extraction engine. "
        "Extract all individual bank transactions from this PDF statement into a clean tabular JSON format. "
        "Each transaction must have:\n"
        "- 'Date': Transaction date in DD-MM-YYYY format.\n"
        "- 'Description': Full transaction narration or description.\n"
        "- 'Debit': Withdrawal/debit amount as a number (or null/0 if it is a credit).\n"
        "- 'Credit': Deposit/credit amount as a number (or null/0 if it is a debit).\n"
        "- 'Reference': Cheque number, UTR, UPI transaction ID, or IMPS reference if available, else empty string.\n"
        "- 'Balance': Account balance after transaction as a number, or null if not available.\n\n"
        "Return ONLY a JSON array of transaction objects."
    )

    models_to_try = ["gemini-3.8-flash", "gemini-3.5-flash", "gemini-flash-latest", "gemini-2.5-flash"]
    last_err = None
    for model_name in models_to_try:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    prompt,
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )
            # Record successful Gemini API request & token count
            try:
                tokens_used = getattr(getattr(response, "usage_metadata", None), "total_token_count", 0) or 0
                sb_for_log = db.get_supabase_client()
                db.record_gemini_api_call(sb_for_log, tokens=tokens_used)
            except Exception:
                pass

            raw_text = response.text.strip()
            if raw_text.startswith("```"):
                raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                raw_text = re.sub(r"\s*```$", "", raw_text)
            records = json.loads(raw_text)
            if isinstance(records, dict):
                for v in records.values():
                    if isinstance(v, list):
                        records = v
                        break
            if isinstance(records, list) and records:
                df = pd.DataFrame(records)
                cols = ["Date", "Description", "Debit", "Credit", "Reference", "Balance"]
                for c in cols:
                    if c not in df.columns:
                        df[c] = ""
                df["Debit"] = df["Debit"].apply(lambda v: "" if v in (0, 0.0, "0", "0.0", None) else v)
                df["Credit"] = df["Credit"].apply(lambda v: "" if v in (0, 0.0, "0", "0.0", None) else v)
                header_row = pd.DataFrame([[c for c in cols]], columns=cols)
                return pd.concat([header_row, df[cols]], ignore_index=True)
            else:
                raise ValueError("No transaction records were identified in the PDF statement.")
        except Exception as e:
            last_err = e
            continue

    raise ValueError(f"Failed to extract transactions from PDF using Gemini: {last_err}")


def read_statement(source: Any, filename: str) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()
    if suffix == ".pdf":
        return extract_transactions_from_pdf_with_gemini(source, filename)
    if suffix == ".csv":
        # Indian bank exports are commonly UTF-8, UTF-8 with BOM, Windows-1252, or Latin-1.
        data = source.getvalue() if hasattr(source, "getvalue") else Path(source).read_bytes()
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = data.decode(encoding)
                reader = csv.reader(io.StringIO(text))
                rows = list(reader)
                if not rows:
                    return pd.DataFrame()
                max_cols = max(len(r) for r in rows)
                padded = []
                for r in rows:
                    cleaned_row = [clean_text(cell) for cell in r]
                    cleaned_row.extend([""] * (max_cols - len(cleaned_row)))
                    padded.append(cleaned_row)
                return pd.DataFrame(padded, dtype=object)
            except (UnicodeDecodeError, csv.Error):
                continue
            except Exception:
                continue
        raise ValueError("Could not decode the CSV file.")
    if suffix not in {".xls", ".xlsx"}:
        raise ValueError("Only .csv, .xls, .xlsx, and .pdf statements are supported.")
    data = source.getvalue() if hasattr(source, "getvalue") else Path(source).read_bytes()
    primary_engine = "openpyxl" if suffix == ".xlsx" else "xlrd"
    fallback_engine = "xlrd" if primary_engine == "openpyxl" else "openpyxl"
    try:
        return pd.read_excel(io.BytesIO(data), header=None, dtype=object, engine=primary_engine)
    except Exception:
        try:
            return pd.read_excel(io.BytesIO(data), header=None, dtype=object, engine=fallback_engine)
        except Exception:
            raise


def extract_identifier(row: pd.Series, reference_col: str | None, description_col: str | None) -> str:
    description = clean_text(row.get(description_col, "")) if description_col else ""

    # 1. Cheque Clearing transactions: extract the actual 6-digit cheque number
    if "CLG" in description.upper():
        clg_parts = [p.strip() for p in description.split("/") if p.strip()]
        if len(clg_parts) >= 3 and clg_parts[0].upper() == "CLG":
            if clg_parts[2].isdigit():
                return clg_parts[2]
            elif clg_parts[1].isdigit():
                return clg_parts[1]
        chq_match = re.search(r"\bCLG/[^/]+/(\d{6,8})\b", description, re.I)
        if chq_match:
            return chq_match.group(1)

    # 2. General cheque pattern (e.g. CHEQUE NO. 123456 or CHQ/123456)
    chq_gen = re.search(r"\b(?:chq|cheque|cheq|chk)\b[^\w\n]*(\d{6,8})\b", description, re.I)
    if chq_gen:
        return chq_gen.group(1)

    # 3. Dedicated NEFT / RTGS patterns:
    # e.g. NEFT Cr-0811OP6182246121-DBSS0IN0811-NAGARJUNA MEDICAL AGENCIES-
    #      NEFT Cr-1012i30089986601-IBKL0NEFT01-...
    #      NEFT Cr-SIBLN26263187292-...
    #      RTGS CR-SBIN0021370-P SUDHAKAR-K JHANSI-SBINR52026100144594496
    neft_match = re.search(r"\b(?:NEFT|RTGS)\s+(?:Cr|Dr)[-\s/]+([A-Za-z0-9]{10,24})\b", description, re.I)
    if neft_match:
        candidate = neft_match.group(1).strip()
        if candidate.upper() not in {"OTHER", "CHARGES"}:
            return candidate

    rtgs_match = re.search(r"\bRTGS\s+(?:Cr|Dr)[-\s/]+(?:[A-Za-z0-9]+[-\s/]+){1,3}([A-Za-z0-9]{16,24})\b", description, re.I)
    if rtgs_match:
        return rtgs_match.group(1).strip()

    mb_neft = re.search(r"\b(?:mb\s+neft|neft)\s+(?:dr|cr)?\s*([A-Za-z0-9]{10,24})\b", description, re.I)
    if mb_neft:
        candidate = mb_neft.group(1).strip()
        if candidate.upper() not in {"OTHER", "CHARGES", "DR", "CR"}:
            return candidate

    # 4. Dedicated IMPS patterns:
    # Format A: Canara Bank INET-IMPS / MOB-IMPS ending with /<12-digit RRN>
    # e.g. INET-IMPS-CR/LAKSHMI DE/ICICI Bank/749205003098/LAKSHMIDEV/7989518305/11/10/2026 18:59:37/628418647914
    if "IMPS" in description.upper():
        imps_end = re.search(r"/(?:[0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4}\s+[0-9:]+/)?([0-9]{12})$", description)
        if imps_end:
            return imps_end.group(1)
        # Format B: ICICI Bank MMT/IMPS/<12-digit RRN>/...
        imps_mmt = re.search(r"\bMMT/IMPS/([0-9]{12})\b", description, re.I)
        if imps_mmt:
            return imps_mmt.group(1)
        # Format C: General IMPS-<12-digit RRN>...
        imps_gen = re.search(r"\bIMPS[-\s/]+([0-9]{12})\b", description, re.I)
        if imps_gen:
            return imps_gen.group(1)

    direct = clean_text(row.get(reference_col, "")) if reference_col else ""
    # Prefer a populated cheque/reference field. ICICI's internal Tran. Id (S123...) is
    # less useful than the UPI/NEFT identifier embedded in its remarks.
    if direct and direct.lower() not in {"0", "000000000000", "nan"} and not re.fullmatch(r"[SM]\d+", direct, re.I):
        return direct

    # 5. UTRs and UPI references are normally an all-numeric 12+ digit segment, or a
    # bank prefix such as IN followed by digits. Keep the complete matching segment.
    identifiers = re.findall(
        r"(?<![A-Za-z0-9])(?:CNRBH\d{8,}|IN\d{12,}|\d{12,})(?![A-Za-z0-9])",
        description,
        flags=re.I,
    )
    if identifiers:
        valid_ids = [i for i in identifiers if not (len(i) > 24 and i.startswith("202"))]
        if valid_ids:
            return valid_ids[0]
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
        date_value = pd.to_datetime(clean_text(row.get(date_col)), errors="coerce", dayfirst=True)
        debit = parse_amount(row.get(debit_col)) if debit_col else None
        credit = parse_amount(row.get(credit_col)) if credit_col else None
        description = clean_text(row.get(description_col))
        if pd.isna(date_value) or (debit is None and credit is None) or not description:
            continue
        if credit is not None and credit > 0:
            amount = credit
            transaction_type = "Credit"
        elif debit is not None and debit > 0:
            amount = debit
            transaction_type = "Debit"
        elif credit is not None and (debit is None or debit == 0):
            amount = credit
            transaction_type = "Credit"
        elif debit is not None and (credit is None or credit == 0):
            amount = debit
            transaction_type = "Debit"
        else:
            amount = credit if credit is not None else (debit if debit is not None else 0.0)
            transaction_type = "Credit" if (credit is not None and credit >= 0) else "Debit"
        original = {column: clean_text(row.get(column)) for column in data.columns if clean_text(row.get(column))}
        records.append({
            "Date": date_value.strftime("%d-%m-%Y"),
            "Amount": amount,
            "Type": transaction_type,
            "UTR / Reference": extract_identifier(row, reference_col, description_col),
            "Description": description,
            "Party Name": "",
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


def load_stored_party_names(sb_client: Any = None, current_df: pd.DataFrame = None) -> list[str]:
    """Load distinct stored party names from local disk, DataFrame, and Supabase for auto-completion."""
    names_set: set[str] = set()

    # 1. Local stored parties file
    if PARTY_STORE_PATH.exists():
        try:
            data = json.loads(PARTY_STORE_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                names_set.update(data)
            elif isinstance(data, dict):
                names_set.update(data.get("stored_party_names", []))
        except Exception:
            pass

    # 2. Legacy assignments file for seamless migration
    if ASSIGNMENTS_PATH.exists():
        try:
            data = json.loads(ASSIGNMENTS_PATH.read_text(encoding="utf-8"))
            names_set.update(data.get("party_mappings", {}).values())
            names_set.update(data.get("transaction_overrides", {}).values())
        except Exception:
            pass

    # 3. Current transactions DataFrame
    if current_df is not None and not current_df.empty and "Party Name" in current_df.columns:
        vals = current_df["Party Name"].dropna().astype(str).str.strip().tolist()
        names_set.update([v for v in vals if v and v.lower() != "nan"])

    # 4. Supabase
    if sb_client:
        try:
            sb_names = db.fetch_party_names(sb_client)
            names_set.update(sb_names)
        except Exception:
            pass

    clean_list = sorted({re.sub(r"\s+", " ", str(n)).strip() for n in names_set if clean_text(n)})
    return clean_list


def save_stored_party_names(names: list[str], sb_client: Any = None) -> None:
    """Save stored party names to local JSON and sync to Supabase."""
    clean_list = sorted({re.sub(r"\s+", " ", str(n)).strip() for n in names if clean_text(n)})
    PARTY_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    PARTY_STORE_PATH.write_text(json.dumps(clean_list, indent=2, ensure_ascii=False), encoding="utf-8")
    if sb_client:
        for name in clean_list:
            try:
                db.save_party_name(sb_client, name)
            except Exception:
                pass


def assign_party_name_to_transactions(
    fingerprints: list[str],
    new_party_name: str,
    combined_df: pd.DataFrame,
    sb_client: Any = None,
) -> int:
    """Assign party name to given transactions and register party name for auto-completion."""
    new_party_name = clean_text(new_party_name).strip()
    if not fingerprints or combined_df is None or combined_df.empty:
        return 0

    mask = combined_df["fingerprint"].isin(fingerprints)
    combined_df.loc[mask, "Party Name"] = new_party_name

    # Update in Supabase
    if sb_client:
        db.update_transactions_party_bulk(sb_client, fingerprints, new_party_name)

    # Automatically store new party name for autocomplete suggestions
    if new_party_name:
        current_stored = load_stored_party_names(sb_client, combined_df)
        if new_party_name not in current_stored:
            current_stored.append(new_party_name)
            save_stored_party_names(current_stored, sb_client)

    return int(mask.sum())


def apply_assignments(transactions: pd.DataFrame, assignments: dict[str, Any] = None) -> pd.DataFrame:
    transactions = transactions.copy()
    if "Description" in transactions.columns:
        transactions["Party match key"] = transactions["Description"].apply(party_match_key)
        def refine_utr_reference(row: pd.Series) -> str:
            desc = clean_text(row.get("Description", ""))
            curr = clean_text(row.get("UTR / Reference", ""))
            if "CLG" in desc.upper() and (len(curr) > 20 or not curr):
                parts = [p.strip() for p in desc.split("/") if p.strip()]
                if len(parts) >= 3 and parts[0].upper() == "CLG":
                    if parts[2].isdigit():
                        return parts[2]
                    elif parts[1].isdigit():
                        return parts[1]
            if "IMPS" in desc.upper():
                imps_end = re.search(r"/(?:[0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4}\s+[0-9:]+/)?([0-9]{12})$", desc)
                if imps_end:
                    return imps_end.group(1)
                imps_mmt = re.search(r"\bMMT/IMPS/([0-9]{12})\b", desc, re.I)
                if imps_mmt:
                    return imps_mmt.group(1)
                imps_gen = re.search(r"\bIMPS[-\s/]+([0-9]{12})\b", desc, re.I)
                if imps_gen:
                    return imps_gen.group(1)
            if not curr:
                neft_match = re.search(r"\b(?:NEFT|RTGS)\s*(?:Cr|Dr)[-\s/]+([A-Za-z0-9]{10,24})\b", desc, re.I)
                if neft_match:
                    cand = neft_match.group(1).strip()
                    if cand.upper() not in {"OTHER", "CHARGES"}:
                        return cand
                rtgs_match = re.search(r"\bRTGS\s+(?:Cr|Dr)[-\s/]+(?:[A-Za-z0-9]+[-\s/]+){1,3}([A-Za-z0-9]{16,24})\b", desc, re.I)
                if rtgs_match:
                    return rtgs_match.group(1).strip()
            return curr
        if "UTR / Reference" in transactions.columns:
            transactions["UTR / Reference"] = transactions.apply(refine_utr_reference, axis=1)

    if "Party Name" not in transactions.columns:
        transactions["Party Name"] = ""
    else:
        transactions["Party Name"] = transactions["Party Name"].fillna("").astype(str).str.strip()

    # Auto party mapping is completely disabled. Party names are only mapped manually by the user.

    transactions["Assigned Business / Party Name"] = transactions["Party Name"]
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


def sort_dataframe_by_date(df: pd.DataFrame, ascending: bool = False) -> pd.DataFrame:
    """Sort a DataFrame chronologically by its 'Date' column (%d-%m-%Y) in a strictly deterministic, stable manner."""
    if df.empty or "Date" not in df.columns:
        return df.copy() if hasattr(df, "copy") else df

    df_out = df.copy()
    parsed_dates = pd.to_datetime(df_out["Date"], format="%d-%m-%Y", errors="coerce")
    if parsed_dates.isna().any():
        fallback = pd.to_datetime(df_out.loc[parsed_dates.isna(), "Date"], dayfirst=True, errors="coerce")
        parsed_dates = parsed_dates.fillna(fallback)

    df_out["_sort_dt"] = parsed_dates
    sort_cols = ["_sort_dt"]
    sort_asc = [ascending]

    for col, asc in [("Bank", True), ("Amount", False), ("fingerprint", True), ("UTR / Reference", True)]:
        if col in df_out.columns:
            sort_cols.append(col)
            sort_asc.append(asc)

    df_out = df_out.sort_values(by=sort_cols, ascending=sort_asc, kind="mergesort")
    return df_out.drop(columns=["_sort_dt"]).reset_index(drop=True)


def search_transactions(transactions: pd.DataFrame, query: str, mode: str) -> pd.DataFrame:
    """Search by category: universal keywords, description, party name, UTR, date, or amount."""
    query = clean_text(query).strip()
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

    if mode == "description":
        return transactions[transactions["Description"].fillna("").str.contains(re.escape(query), case=False, regex=True)].copy()

    if mode == "party":
        return transactions[transactions["Party Name"].fillna("").str.contains(re.escape(query), case=False, regex=True)].copy()

    if mode == "all":
        mask = (
            transactions["Description"].fillna("").str.contains(re.escape(query), case=False, regex=True)
            | transactions["Party Name"].fillna("").str.contains(re.escape(query), case=False, regex=True)
            | transactions["UTR / Reference"].fillna("").str.contains(re.escape(query), case=False, regex=True)
            | transactions["Bank"].fillna("").str.contains(re.escape(query), case=False, regex=True)
            | transactions["Date"].fillna("").str.contains(re.escape(query), case=False, regex=True)
        )
        return transactions[mask].copy()

    utr_mask = (
        transactions["UTR / Reference"].fillna("").str.contains(re.escape(query), case=False, regex=True)
        | transactions["Description"].fillna("").str.contains(re.escape(query), case=False, regex=True)
    )
    return transactions[utr_mask].copy()


def inject_keyboard_shortcuts() -> None:
    """Inject global keyboard shortcut (Ctrl+F / Cmd+F) to clear previous query and focus StatementHub's in-app search."""
    if components is None:
        return
    components.html(
        """
        <script>
        try {
            function getSearchWidgetAndInput(doc) {
                if (!doc) return { widget: null, input: null };
                // 1. Precise match by label text
                const labels = Array.from(doc.querySelectorAll('label'));
                for (const lbl of labels) {
                    const txt = (lbl.innerText || '').toLowerCase().trim();
                    if (txt.includes('search utr') || txt.includes('search date') || txt.includes('search amount')) {
                        const widget = lbl.closest('[data-testid="stSelectbox"], [data-testid="stTextInput"]');
                        if (widget) {
                            return { widget, input: widget.querySelector('input') };
                        }
                    }
                }

                // 2. Placeholder match
                const byPlaceholder = doc.querySelector(
                    'input[placeholder*="typing a UTR"], input[placeholder*="Enter date"], input[placeholder*="Enter amount"]'
                );
                if (byPlaceholder) {
                    const widget = byPlaceholder.closest('[data-testid="stSelectbox"], [data-testid="stTextInput"]');
                    return { widget, input: byPlaceholder };
                }

                // 3. Fallback scan across all inputs, strictly excluding Bank Account / Period / Date filters
                const inputs = Array.from(doc.querySelectorAll('input'));
                for (const inp of inputs) {
                    const widget = inp.closest('[data-testid="stSelectbox"], [data-testid="stTextInput"]');
                    const lbl = widget ? (widget.querySelector('label')?.innerText || '').toLowerCase() : '';
                    const aria = (inp.getAttribute('aria-label') || '').toLowerCase();
                    const ph = (inp.getAttribute('placeholder') || '').toLowerCase();

                    if (lbl.includes('bank') || aria.includes('bank') || ph.includes('bank')) continue;
                    if (lbl.includes('period') || aria.includes('period') || ph.includes('period')) continue;
                    if (lbl.includes('from date') || lbl.includes('to date')) continue;

                    if (
                        lbl.includes('utr') || lbl.includes('search') ||
                        aria.includes('utr') || aria.includes('search') ||
                        ph.includes('utr') || ph.includes('search')
                    ) {
                        return { widget, input: inp };
                    }
                }
                return { widget: null, input: null };
            }

            function triggerClear(widget, input) {
                // For st.selectbox (UTR search): click BaseWeb's Clear (x) button if a value is selected
                if (widget) {
                    const clearBtn = widget.querySelector(
                        '[aria-label*="Clear" i], [aria-label*="clear" i], [title*="Clear" i], [title*="clear" i]'
                    );
                    if (clearBtn) {
                        const clickTarget = clearBtn.closest('[role="button"]') || clearBtn.parentElement || clearBtn;
                        clickTarget.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
                        clickTarget.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
                        clickTarget.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                        if (clickTarget.click) clickTarget.click();
                    } else {
                        // In BaseWeb selectbox, when an option is selected, there are >= 2 SVGs:
                        // 1st SVG is the Clear (x) icon, 2nd SVG is the dropdown chevron
                        const svgs = widget.querySelectorAll('[data-baseweb="select"] svg');
                        if (svgs.length >= 2) {
                            const clearSvg = svgs[0];
                            const clickTarget = clearSvg.closest('[role="button"]') || clearSvg.parentElement || clearSvg;
                            clickTarget.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
                            clickTarget.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
                            clickTarget.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                            if (clickTarget.click) clickTarget.click();
                        }
                    }
                }

                // Also clear text input directly for st.text_input (Date/Amount mode) or any typed text
                if (input) {
                    input.value = '';
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }

            function handleShortcut(e) {
                if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
                    e.preventDefault();
                    e.stopPropagation();

                    const parentDoc = window.parent.document;
                    const { widget, input } = getSearchWidgetAndInput(parentDoc);
                    if (!widget && !input) return;

                    // Automatically clear any previous UTR / search query
                    triggerClear(widget, input);

                    // Highlight the search bar with accent glow and scroll into view
                    const container = widget || input;
                    if (container) {
                        container.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        container.style.transition = 'box-shadow 0.25s ease';
                        container.style.boxShadow = '0 0 0 3px #ff4b4b';
                        setTimeout(() => { container.style.boxShadow = ''; }, 1200);
                    }

                    // Focus the empty input ready for immediate typing
                    setTimeout(() => {
                        const targetInput = (widget ? widget.querySelector('input') : null) || input;
                        if (targetInput) {
                            targetInput.focus();
                            if (targetInput.select) targetInput.select();
                        }
                    }, 70);
                }
            }

            const parentDoc = window.parent.document;
            if (parentDoc) {
                if (parentDoc._stmtHubKeyHandler) {
                    parentDoc.removeEventListener('keydown', parentDoc._stmtHubKeyHandler, true);
                }
                parentDoc._stmtHubKeyHandler = handleShortcut;
                parentDoc.addEventListener('keydown', parentDoc._stmtHubKeyHandler, true);
            }
            if (document) {
                if (document._stmtHubKeyHandler) {
                    document.removeEventListener('keydown', document._stmtHubKeyHandler, true);
                }
                document._stmtHubKeyHandler = handleShortcut;
                document.addEventListener('keydown', document._stmtHubKeyHandler, true);
            }
        } catch(e) {
            console.error('StatementHub shortcut init error:', e);
        }
        </script>
        """,
        height=0,
        width=0,
    )


def inject_party_autocomplete(stored_party_names: list[str]) -> None:
    """Inject a non-intrusive, precise autocomplete suggestion popup for Party Name editing in Glide Data Grid."""
    if components is None or not stored_party_names:
        return

    clean_names = sorted(list(set(clean_text(n).strip() for n in stored_party_names if clean_text(n).strip())))
    if not clean_names:
        return

    names_json = json.dumps(clean_names)

    html_code = f"""
    <script>
    (function() {{
        try {{
            const parentDoc = window.parent.document;
            if (!parentDoc) return;

            const partyList = {names_json};

            let popup = parentDoc.getElementById('sh-party-ac-popup');
            if (!popup) {{
                popup = parentDoc.createElement('div');
                popup.id = 'sh-party-ac-popup';
                popup.style.cssText = 'position: absolute; display: none; z-index: 99999999; background: #181825; border: 1px solid #45475a; border-radius: 8px; box-shadow: 0 12px 30px rgba(0,0,0,0.65); font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; font-size: 13px; overflow: hidden;';
                parentDoc.body.appendChild(popup);
            }}

            let activeInput = null;
            let currentMatches = [];
            let keyboardSelectedIdx = -1;

            function hidePopup() {{
                if (popup) popup.style.display = 'none';
                currentMatches = [];
                keyboardSelectedIdx = -1;
            }}

            function escapeHtml(str) {{
                return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
            }}

            function applySuggestion(input, val) {{
                if (!input) return;
                try {{
                    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set
                        || Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
                    if (setter) {{
                        setter.call(input, val);
                    }} else {{
                        input.value = val;
                    }}
                    if (input._valueTracker) {{
                        input._valueTracker.setValue('');
                    }}
                    input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }} catch(err) {{
                    input.value = val;
                }}
                hidePopup();
                input.focus();
                // Commit the cell in Glide Data Grid cleanly
                setTimeout(() => {{
                    try {{
                        input.dispatchEvent(new KeyboardEvent('keydown', {{
                            key: 'Enter',
                            code: 'Enter',
                            keyCode: 13,
                            which: 13,
                            bubbles: true,
                            cancelable: true
                        }}));
                    }} catch(e) {{}}
                }}, 25);
            }}

            function updateKeyboardHighlight(newIdx) {{
                keyboardSelectedIdx = newIdx;
                const items = popup.querySelectorAll('.sh-ac-item');
                items.forEach((it, i) => {{
                    if (i === keyboardSelectedIdx) {{
                        it.style.background = '#313244';
                        it.style.color = '#89b4fa';
                        it.scrollIntoView({{ block: 'nearest' }});
                    }} else {{
                        it.style.background = 'transparent';
                        it.style.color = '#cdd6f4';
                    }}
                }});
            }}

            function renderSuggestions(input, matches) {{
                currentMatches = matches;
                keyboardSelectedIdx = -1;
                popup.innerHTML = '';

                const header = parentDoc.createElement('div');
                header.style.cssText = 'padding: 6px 12px; font-size: 11px; color: #a6adc8; border-bottom: 1px solid #313244; background: #11111b; font-weight: 600; display: flex; justify-content: space-between; align-items: center;';
                header.innerHTML = '<span>🏷️ Saved Party Suggestions</span><span style="font-size:10px; opacity:0.7;">(Click or ↓ then Enter)</span>';
                popup.appendChild(header);

                const listContainer = parentDoc.createElement('div');
                listContainer.style.cssText = 'max-height: 220px; overflow-y: auto;';

                matches.forEach((name, idx) => {{
                    const item = parentDoc.createElement('div');
                    item.className = 'sh-ac-item';
                    item.dataset.idx = idx;
                    item.style.cssText = 'padding: 8px 12px; cursor: pointer; color: #cdd6f4; font-size: 13px; border-bottom: 1px solid #27273a; transition: background 0.1s ease; display: flex; align-items: center; justify-content: space-between;';
                    item.innerHTML = `<span><span style="color:#89b4fa;margin-right:8px;">🏷️</span><b>${{escapeHtml(name)}}</b></span><span style="font-size:10px;color:#6c7086;background:#242438;padding:2px 6px;border-radius:4px;">↵ Select</span>`;

                    // Mouse hover ONLY changes visual styling - does NOT alter keyboardSelectedIdx
                    item.addEventListener('mouseenter', () => {{
                        const allItems = popup.querySelectorAll('.sh-ac-item');
                        allItems.forEach(it => {{
                            if (it === item) {{
                                it.style.background = '#313244';
                                it.style.color = '#89b4fa';
                            }} else if (parseInt(it.dataset.idx) !== keyboardSelectedIdx) {{
                                it.style.background = 'transparent';
                                it.style.color = '#cdd6f4';
                            }}
                        }});
                    }});

                    item.addEventListener('mouseleave', () => {{
                        if (parseInt(item.dataset.idx) !== keyboardSelectedIdx) {{
                            item.style.background = 'transparent';
                            item.style.color = '#cdd6f4';
                        }}
                    }});

                    item.addEventListener('mousedown', (e) => {{
                        e.preventDefault();
                        e.stopPropagation();
                        applySuggestion(input, name);
                    }});

                    listContainer.appendChild(item);
                }});

                popup.appendChild(listContainer);

                // Precise positioning below active input
                const rect = input.getBoundingClientRect();
                const parentWin = window.parent;
                const scrollTop = parentWin.pageYOffset || parentDoc.documentElement.scrollTop;
                const scrollLeft = parentWin.pageXOffset || parentDoc.documentElement.scrollLeft;

                popup.style.top = (rect.bottom + scrollTop + 4) + 'px';
                popup.style.left = Math.max(10, rect.left + scrollLeft) + 'px';
                popup.style.minWidth = Math.max(rect.width, 300) + 'px';
                popup.style.maxWidth = '500px';
                popup.style.display = 'block';
            }}

            function onInputChange(e) {{
                const input = e.target;
                if (!input) return;
                const isGdg = input.classList && (input.classList.contains('gdg-input') || input.classList.contains('gdg-md-edit-textarea'));
                const isText = input.tagName === 'TEXTAREA' || (input.tagName === 'INPUT' && input.type === 'text');
                if (!isGdg && !isText) return;

                activeInput = input;
                const q = (input.value || '').trim().toLowerCase();
                if (q.length < 1) {{
                    hidePopup();
                    return;
                }}

                // Prefix matches first, then substring matches
                const exactStart = [];
                const contains = [];
                for (let i = 0; i < partyList.length; i++) {{
                    const p = partyList[i];
                    const pLower = p.toLowerCase();
                    if (pLower.startsWith(q)) {{
                        exactStart.push(p);
                    }} else if (pLower.includes(q)) {{
                        contains.push(p);
                    }}
                    if (exactStart.length + contains.length >= 24) break;
                }}
                const matches = exactStart.concat(contains).slice(0, 8);
                if (!matches.length) {{
                    hidePopup();
                    return;
                }}

                renderSuggestions(input, matches);
            }}

            function onKeyDown(e) {{
                if (!popup || popup.style.display !== 'block' || !currentMatches.length) return;

                if (e.key === 'ArrowDown') {{
                    e.preventDefault();
                    e.stopPropagation();
                    const next = (keyboardSelectedIdx + 1) % currentMatches.length;
                    updateKeyboardHighlight(next);
                }} else if (e.key === 'ArrowUp') {{
                    e.preventDefault();
                    e.stopPropagation();
                    const prev = (keyboardSelectedIdx - 1 + currentMatches.length) % currentMatches.length;
                    updateKeyboardHighlight(prev);
                }} else if ((e.key === 'Enter' || e.key === 'Tab') && keyboardSelectedIdx >= 0) {{
                    // ONLY intercept Enter if user explicitly navigated down using arrow keys
                    e.preventDefault();
                    e.stopPropagation();
                    applySuggestion(activeInput, currentMatches[keyboardSelectedIdx]);
                }} else if (e.key === 'Escape') {{
                    hidePopup();
                }} else if (e.key === 'Enter') {{
                    // User typed free text and pressed Enter without selecting a suggestion:
                    // Hide popup immediately and let Enter pass through cleanly to Glide Data Grid!
                    hidePopup();
                }}
            }}

            // Clean up prior listeners before re-attaching
            if (parentDoc._shPartyAcInputHandler) {{
                parentDoc.removeEventListener('input', parentDoc._shPartyAcInputHandler, true);
                parentDoc.removeEventListener('keydown', parentDoc._shPartyAcKeyHandler, true);
                parentDoc.removeEventListener('click', parentDoc._shPartyAcClickHandler, true);
                parentDoc.removeEventListener('scroll', parentDoc._shPartyAcScrollHandler, true);
            }}

            parentDoc._shPartyAcInputHandler = onInputChange;
            parentDoc._shPartyAcKeyHandler = onKeyDown;
            parentDoc._shPartyAcClickHandler = (e) => {{
                if (popup && popup.style.display === 'block') {{
                    if (!popup.contains(e.target) && e.target !== activeInput) {{
                        hidePopup();
                    }}
                }}
            }};
            parentDoc._shPartyAcScrollHandler = () => {{
                if (popup && popup.style.display === 'block') {{
                    hidePopup();
                }}
            }};

            parentDoc.addEventListener('input', parentDoc._shPartyAcInputHandler, true);
            parentDoc.addEventListener('keydown', parentDoc._shPartyAcKeyHandler, true);
            parentDoc.addEventListener('click', parentDoc._shPartyAcClickHandler, true);
            parentDoc.addEventListener('scroll', parentDoc._shPartyAcScrollHandler, true);

            // Prevent client-side column header click-to-sort on Glide Data Grid canvas,
            // which causes visual-vs-dataframe row index misalignment on edits.
            function preventCanvasHeaderSort(e) {{
                const target = e.target;
                if (!target || target.tagName !== 'CANVAS') return;
                const rect = target.getBoundingClientRect();
                const offsetY = e.clientY - rect.top;
                if (offsetY >= 0 && offsetY <= 40) {{
                    e.stopPropagation();
                    e.stopImmediatePropagation();
                    e.preventDefault();
                }}
            }}

            if (parentDoc._shCanvasHeaderSortHandler) {{
                parentDoc.removeEventListener('click', parentDoc._shCanvasHeaderSortHandler, true);
                parentDoc.removeEventListener('mousedown', parentDoc._shCanvasHeaderSortHandler, true);
            }}
            parentDoc._shCanvasHeaderSortHandler = preventCanvasHeaderSort;
            parentDoc.addEventListener('click', parentDoc._shCanvasHeaderSortHandler, true);
            parentDoc.addEventListener('mousedown', parentDoc._shCanvasHeaderSortHandler, true);

        }} catch(err) {{
            console.error('StatementHub Party Autocomplete Error:', err);
        }}
    }})();
    </script>
    """
    components.html(html_code, height=0, width=0)


def render_single_copy_btn(text: str, label: str = "📋 Copy UTR") -> None:
    """Render a 1-click copy button beside the search box or values."""
    if not text or components is None:
        return
    safe_text = str(text).replace("'", "\\'")
    html_code = f"""
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; align-items: flex-end; height: 100%; margin-bottom: 2px;">
        <button id="copy-utr-btn" onclick="copyUtrSingle('{safe_text}', this)" style="
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            width: 100%;
            height: 40px;
            background: #2a2b3d;
            color: #cdd6f4;
            border: 1px solid #45475a;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.15s ease;
        ">
            {label}
        </button>
    </div>
    <script>
    function copyUtrSingle(val, btn) {{
        navigator.clipboard.writeText(val).then(() => {{
            btn.innerHTML = '✅ Copied!';
            btn.style.background = '#2e7d32';
            btn.style.borderColor = '#81c784';
            btn.style.color = '#ffffff';
            setTimeout(() => {{
                btn.innerHTML = '{label}';
                btn.style.background = '#2a2b3d';
                btn.style.borderColor = '#45475a';
                btn.style.color = '#cdd6f4';
            }}, 2000);
        }}).catch(err => {{
            console.error(err);
        }});
    }}
    </script>
    """
    components.html(html_code, height=45)


def get_bank_color(bank_name: str) -> str:
    """Return a subtle tinted background color for the bank."""
    b = str(bank_name or "").lower().strip()
    if "canara" in b:
        # Canara Bank: Crisp emerald green tint (distinct from ICICI orange & HDFC blue)
        return "rgba(16, 185, 129, 0.16)"
    elif "icici" in b:
        # ICICI Bank: Warm orange/coral tint
        return "rgba(249, 115, 22, 0.16)"
    elif "hdfc" in b:
        # HDFC Bank: Modern sky/navy blue tint
        return "rgba(14, 165, 233, 0.16)"
    elif "sbi" in b or "state bank" in b:
        # State Bank of India: Deep teal tint
        return "rgba(13, 148, 136, 0.16)"
    elif "axis" in b:
        # Axis Bank: Maroon/rose tint
        return "rgba(244, 63, 94, 0.16)"
    elif "kotak" in b:
        # Kotak Mahindra Bank: Crimson red tint
        return "rgba(239, 68, 68, 0.16)"
    elif b:
        # Other bank: Indigo/purple tint
        return "rgba(168, 85, 247, 0.13)"
    return ""


def style_table_by_bank(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """Apply bank-specific background color to rows in data editor."""
    try:
        pd.set_option("styler.render.max_elements", 1_000_000)
    except Exception:
        pass

    if df.empty or "Bank" not in df.columns:
        return df.style

    def bank_styler(d: pd.DataFrame) -> pd.DataFrame:
        styles = pd.DataFrame("", index=d.index, columns=d.columns)
        for b_uniq in d["Bank"].dropna().unique():
            color = get_bank_color(b_uniq)
            if color:
                mask = d["Bank"] == b_uniq
                styles.loc[mask, :] = f"background-color: {color}"
        return styles

    return df.style.apply(bank_styler, axis=None)


def render_bank_color_legend(df: pd.DataFrame) -> None:
    """Render a visual color legend for banks present in the displayed results."""
    if df.empty or "Bank" not in df.columns:
        return

    counts = df["Bank"].value_counts().to_dict()
    if not counts:
        return

    bank_meta = {
        "canara": {"icon": "🟢", "label": "Canara Bank", "bg": "rgba(16, 185, 129, 0.16)", "border": "rgba(16, 185, 129, 0.35)", "color": "#059669"},
        "icici": {"icon": "🟠", "label": "ICICI Bank", "bg": "rgba(249, 115, 22, 0.16)", "border": "rgba(249, 115, 22, 0.35)", "color": "#ea580c"},
        "hdfc": {"icon": "🔵", "label": "HDFC Bank", "bg": "rgba(14, 165, 233, 0.16)", "border": "rgba(14, 165, 233, 0.35)", "color": "#0284c7"},
        "sbi": {"icon": "🩵", "label": "SBI", "bg": "rgba(13, 148, 136, 0.16)", "border": "rgba(13, 148, 136, 0.35)", "color": "#0f766e"},
        "state bank": {"icon": "🩵", "label": "State Bank", "bg": "rgba(13, 148, 136, 0.16)", "border": "rgba(13, 148, 136, 0.35)", "color": "#0f766e"},
        "axis": {"icon": "🟣", "label": "Axis Bank", "bg": "rgba(244, 63, 94, 0.16)", "border": "rgba(244, 63, 94, 0.35)", "color": "#e11d48"},
        "kotak": {"icon": "🔴", "label": "Kotak Bank", "bg": "rgba(239, 68, 68, 0.16)", "border": "rgba(239, 68, 68, 0.35)", "color": "#dc2626"},
    }

    badges_html: list[str] = []
    for bank_name, count in counts.items():
        b_key = str(bank_name).lower()
        matched = None
        for k, meta in bank_meta.items():
            if k in b_key:
                matched = meta
                break
        if not matched:
            matched = {
                "icon": "⚪",
                "label": bank_name,
                "bg": "rgba(168, 85, 247, 0.14)",
                "border": "rgba(168, 85, 247, 0.35)",
                "color": "#9333ea",
            }

        badges_html.append(
            f'<span style="display:inline-flex; align-items:center; gap:5px; padding:3px 10px; border-radius:12px; '
            f'background:{matched["bg"]}; border:1px solid {matched["border"]}; font-size:12px; font-weight:600; '
            f'letter-spacing:0.2px;">'
            f'{matched["icon"]} {bank_name} <span style="opacity:0.75; font-weight:normal; margin-left:2px;">({count:,})</span></span>'
        )

    if badges_html:
        st.markdown(
            f'<div style="display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin: 4px 0 10px 0;">'
            f'<span style="font-size:11.5px; font-weight:600; color:gray; text-transform:uppercase; letter-spacing:0.5px;">Bank Colors:</span>'
            + " ".join(badges_html)
            + "</div>",
            unsafe_allow_html=True,
        )


def render_backup_restore_section(combined: pd.DataFrame, assignments: dict[str, Any], sb_client: Any) -> None:
    with st.expander("💾 Backup & Restore Center"):
        st.markdown("**Create an offline backup or restore your database & stored party names.**")
        stored_parties = load_stored_party_names(sb_client, combined)
        b_col1, b_col2 = st.columns(2)
        with b_col1:
            st.write("**Download Complete Backup**")
            st.caption(f"• **Transactions:** {len(combined):,} records\n• **Stored Party Names:** {len(stored_parties)} names")
            if not combined.empty or stored_parties:
                txns_list = json.loads(combined.to_json(orient="records", date_format="iso")) if not combined.empty else []
                backup_data = {
                    "version": "2.0",
                    "exported_at": datetime.now().isoformat(),
                    "total_transactions": len(combined),
                    "total_stored_party_names": len(stored_parties),
                    "stored_party_names": stored_parties,
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
                    help="Downloads all transactions and stored party names to a secure JSON file",
                )
            else:
                st.caption("No transactions or party names available to backup yet.")

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
                    restored_stored = loaded_backup.get("stored_party_names", [])
                    # Fallback to legacy party_mappings if present
                    if not restored_stored and "party_mappings" in loaded_backup:
                        restored_stored = list(set(loaded_backup.get("party_mappings", {}).values()))

                    st.info(f"Found **{len(restored_txns):,} transactions** and **{len(restored_stored)} stored party names** in backup.")

                    if st.button("🚀 Restore Now", type="primary"):
                        if restored_stored:
                            all_parties = load_stored_party_names(sb_client, combined)
                            all_parties = sorted(set(all_parties) | set(restored_stored))
                            save_stored_party_names(all_parties, sb_client)

                        if sb_client:
                            if restored_txns:
                                r_df = pd.DataFrame(restored_txns)
                                db.sync_transactions_with_dedup(sb_client, r_df)
                            st.session_state["cached_transactions"] = db.fetch_all_transactions(sb_client)
                        else:
                            if restored_txns:
                                st.session_state["cached_transactions"] = pd.DataFrame(restored_txns)

                        st.toast("Database and stored party names restored successfully!", icon="✅")
                except Exception as exc:
                    st.error(f"Invalid backup file format: {exc}")


def render_storage_usage_section(combined: pd.DataFrame, sb_client: Any) -> None:
    """Render database storage usage meter and capacity forecast."""
    storage = db.get_database_storage_stats(sb_client, combined)
    used_mb = storage["used_mb"]
    free_mb = storage["free_mb"]
    total_quota_mb = storage["total_quota_mb"]
    used_pct = storage["used_pct"]
    txn_count = storage["txn_count"]
    remaining_txns = storage["remaining_txns"]
    is_exact = storage["is_exact"]

    with st.expander("💾 Database Storage & Capacity Tracker", expanded=False):
        st.markdown("**Monitor your cloud database storage usage, free space left, and capacity.**")

        # Visual Storage Bar
        progress_val = min(max(used_pct / 100.0, 0.001), 1.0)
        st.progress(progress_val)
        st.caption(
            f"**Used:** {used_mb:.2f} MB of {total_quota_mb:.0f} MB ({used_pct:.2f}%)  |  "
            f"**Free Space Left:** {free_mb:.2f} MB ({100.0 - used_pct:.2f}%)"
        )

        s_col1, s_col2, s_col3, s_col4 = st.columns(4)
        with s_col1:
            st.metric("Storage Used", f"{used_mb:.2f} MB", help="Storage currently occupied by your transaction tables and indexes")
        with s_col2:
            st.metric("Storage Free Left", f"{free_mb:.2f} MB", help="Unused storage remaining on Supabase Free Tier (500 MB)")
        with s_col3:
            st.metric("Saved Transactions", f"{txn_count:,}", help="Total number of bank transactions stored in database")
        with s_col4:
            st.metric("Capacity Remaining", f"~{format_inr(remaining_txns).replace('₹ ', '')}", help="Estimated number of additional transactions that can fit into the remaining free space")

        st.info(
            f"💡 **You have massive storage headroom!** At your current statement volume, your remaining **{free_mb:.2f} MB** can store "
            f"**~{format_inr(remaining_txns).replace('₹ ', '')} more transactions** (over 70+ years of daily bank statements!).",
            icon="🚀",
        )

        b_col, _ = st.columns([2.5, 3.5])
        with b_col:
            if st.button("🔄 Refresh Live Database Storage", key="refresh_db_storage_btn", use_container_width=True):
                if sb_client:
                    refreshed = db.fetch_all_transactions(sb_client)
                    st.session_state["cached_transactions"] = refreshed
                    latest = db.get_database_storage_stats(sb_client, refreshed)
                    st.toast(
                        f"Live database refreshed: {latest['txn_count']:,} transactions | "
                        f"{latest['used_mb']:.2f} MB used | {latest['free_mb']:.2f} MB free.",
                        icon="💾",
                    )
                else:
                    st.toast("Running in local mode.", icon="💾")
                st.rerun()


def export_monthly_report(transactions: pd.DataFrame, exclude_internal: bool = True, exclude_fd: bool = True) -> bytes:
    """Generate a multi-sheet Excel workbook with monthly profit, bank breakdown, FDs, & internal transfers."""
    if transactions.empty or "Date" not in transactions.columns or "Amount" not in transactions.columns:
        return b""

    dates = pd.to_datetime(transactions["Date"], format="%d-%m-%Y", errors="coerce")
    if dates.isna().any():
        fallback = pd.to_datetime(transactions.loc[dates.isna(), "Date"], dayfirst=True, errors="coerce")
        dates.fillna(fallback, inplace=True)

    valid = transactions[~dates.isna()].copy()
    if valid.empty:
        return b""

    valid["DateTime"] = dates[~dates.isna()]
    valid["YearMonth"] = valid["DateTime"].dt.strftime("%Y-%m")
    valid["Month"] = valid["DateTime"].dt.strftime("%B %Y")
    valid["Is_Internal"] = detect_internal_transfers(valid)
    valid["Is_FD"] = detect_fd_investments(valid)

    # Filtered dataset if exclude_internal or exclude_fd is requested
    biz_df = valid.copy()
    if exclude_internal:
        biz_df = biz_df[~biz_df["Is_Internal"]]
    if exclude_fd:
        biz_df = biz_df[~biz_df["Is_FD"]]

    # Monthly Summary
    m_grp = biz_df.groupby(["YearMonth", "Month", "Type"])["Amount"].sum().unstack(fill_value=0.0)
    if "Credit" not in m_grp:
        m_grp["Credit"] = 0.0
    if "Debit" not in m_grp:
        m_grp["Debit"] = 0.0
    counts = valid.groupby(["YearMonth", "Month"]).size().to_dict()

    self_monthly: dict[tuple[str, str], float] = {}
    for (ym, m_name), s_grp in valid[valid["Is_Internal"]].groupby(["YearMonth", "Month"]):
        s_cr = float(s_grp.loc[s_grp["Type"] == "Credit", "Amount"].sum())
        s_dr = float(s_grp.loc[s_grp["Type"] == "Debit", "Amount"].sum())
        self_monthly[(ym, m_name)] = max(s_cr, s_dr) if (s_cr > 0 and s_dr > 0) else (s_cr + s_dr)

    fd_monthly: dict[tuple[str, str], float] = {}
    for (ym, m_name), f_grp in valid[valid["Is_FD"]].groupby(["YearMonth", "Month"]):
        fd_monthly[(ym, m_name)] = float(f_grp["Amount"].sum())

    m_df = m_grp.reset_index().sort_values("YearMonth", ascending=False)
    m_df["Net Profit / Surplus"] = m_df["Credit"] - m_df["Debit"]
    m_df["Status"] = m_df["Net Profit / Surplus"].apply(lambda x: "Profit" if x >= 0 else "Deficit")
    m_df["Margin %"] = m_df.apply(
        lambda r: round((r["Net Profit / Surplus"] / r["Credit"] * 100), 2) if r["Credit"] > 0 else 0.0, axis=1
    )
    m_df["Fixed Deposits (Investments)"] = m_df.apply(lambda r: fd_monthly.get((r["YearMonth"], r["Month"]), 0.0), axis=1)
    m_df["Internal Transfers (Contra)"] = m_df.apply(lambda r: self_monthly.get((r["YearMonth"], r["Month"]), 0.0), axis=1)
    m_df["Transactions"] = m_df.apply(lambda r: counts.get((r["YearMonth"], r["Month"]), 0), axis=1)
    m_df.rename(columns={"Credit": "Business Received", "Debit": "Business Sent"}, inplace=True)
    m_export = m_df.drop(columns=["YearMonth"], errors="ignore")

    # Bank-wise Breakdown
    b_grp = biz_df.groupby(["YearMonth", "Month", "Bank", "Type"])["Amount"].sum().unstack(fill_value=0.0).reset_index()
    if "Credit" not in b_grp:
        b_grp["Credit"] = 0.0
    if "Debit" not in b_grp:
        b_grp["Debit"] = 0.0
    b_grp["Net Profit / Surplus"] = b_grp["Credit"] - b_grp["Debit"]
    b_grp["Status"] = b_grp["Net Profit / Surplus"].apply(lambda x: "Profit" if x >= 0 else "Deficit")
    b_grp.rename(columns={"Credit": "Business Received", "Debit": "Business Sent"}, inplace=True)
    b_export = b_grp.sort_values(["YearMonth", "Bank"], ascending=[False, True]).drop(columns=["YearMonth"], errors="ignore")

    # Fixed Deposits & Investments Sheet
    fd_df = sort_dataframe_by_date(
        valid[valid["Is_FD"]][["Date", "Bank", "Type", "Amount", "UTR / Reference", "Description", "Party Name"]].copy(),
        ascending=False,
    )

    # Internal Transfers Sheet
    internal_df = sort_dataframe_by_date(
        valid[valid["Is_Internal"]][["Date", "Bank", "Type", "Amount", "UTR / Reference", "Description", "Party Name"]].copy(),
        ascending=False,
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        m_export.to_excel(writer, sheet_name="Monthly Operating Profit", index=False)
        b_export.to_excel(writer, sheet_name="Bank Breakdown", index=False)
        if not fd_df.empty:
            fd_df.to_excel(writer, sheet_name="Fixed Deposits & Investments", index=False)
        if not internal_df.empty:
            internal_df.to_excel(writer, sheet_name="Internal Self Transfers", index=False)
    return output.getvalue()


def render_monthly_profit_section(combined: pd.DataFrame) -> None:
    """Render interactive Monthly Profit & Cash Flow analytics dashboard with contra and FD handling."""
    st.subheader("💰 Monthly Profit & Cash Flow Dashboard")
    st.caption("Track how much money you received, how much you spent, and your net monthly profit/deficit across bank accounts.")

    if combined.empty:
        st.info("No transaction data available to calculate monthly profit.")
        return

    # Filter by bank and toggles for self-transfers & FDs
    all_banks = ["All Banks"] + sorted([b for b in combined["Bank"].dropna().unique().tolist() if b])
    p_col_bank, p_col_toggle1, p_col_toggle2 = st.columns([2, 2.5, 2.5], vertical_alignment="bottom")
    with p_col_bank:
        selected_bank = st.selectbox("Filter Analytics by Bank", options=all_banks, index=0, key="analytics_bank_filter")
    with p_col_toggle1:
        exclude_internal = st.toggle(
            "Exclude Self-Transfers (Contra)",
            value=True,
            help="Removes money transferred between your own accounts (e.g. ICICI to Canara) so incoming/outgoing cash reflects true customer revenue and business expenses.",
            key="exclude_internal_toggle",
        )
    with p_col_toggle2:
        exclude_fd = st.toggle(
            "Exclude Fixed Deposits (Investments)",
            value=True,
            help="Separates bank Fixed Deposits / Term Deposits so capital savings are not counted as business losses or expenses.",
            key="exclude_fd_toggle",
        )

    df = combined.copy()
    if selected_bank != "All Banks":
        df = df[df["Bank"] == selected_bank].copy()

    dates = pd.to_datetime(df["Date"], format="%d-%m-%Y", errors="coerce")
    if dates.isna().any():
        fallback = pd.to_datetime(df.loc[dates.isna(), "Date"], dayfirst=True, errors="coerce")
        dates.fillna(fallback, inplace=True)

    valid = df[~dates.isna()].copy()
    if valid.empty:
        st.warning("No transactions with valid dates found for this selection.")
        return

    valid["DateTime"] = dates[~dates.isna()]
    valid["YearMonth"] = valid["DateTime"].dt.strftime("%Y-%m")
    valid["Month"] = valid["DateTime"].dt.strftime("%B %Y")
    if "Is_Internal" not in valid.columns:
        valid["Is_Internal"] = detect_internal_transfers(valid)
    if "Is_FD" not in valid.columns:
        valid["Is_FD"] = detect_fd_investments(valid)

    # Calculate internal transfers sum
    int_dr = float(valid.loc[valid["Is_Internal"] & (valid["Type"] == "Debit"), "Amount"].sum())
    int_cr = float(valid.loc[valid["Is_Internal"] & (valid["Type"] == "Credit"), "Amount"].sum())
    total_internal = max(int_dr, int_cr) if (int_dr > 0 and int_cr > 0) else (int_dr + int_cr)

    # Calculate Fixed Deposits sum
    total_fd = float(valid.loc[valid["Is_FD"], "Amount"].sum())

    # Active dataset based on toggles
    active_df = valid.copy()
    if exclude_internal:
        active_df = active_df[~active_df["Is_Internal"]]
    if exclude_fd:
        active_df = active_df[~active_df["Is_FD"]]

    total_rec = float(active_df.loc[active_df["Type"] == "Credit", "Amount"].sum())
    total_sent = float(active_df.loc[active_df["Type"] == "Debit", "Amount"].sum())
    net_profit = total_rec - total_sent
    margin_pct = (net_profit / total_rec * 100) if total_rec > 0 else 0.0

    # Top KPI Cards - Row 1: Core Operating Performance (4 spacious columns)
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.metric(
            "Business Received" if (exclude_internal or exclude_fd) else "Total Received",
            format_inr(total_rec),
            help="Real customer revenue into your account",
        )
    with kpi2:
        st.metric(
            "Business Sent" if (exclude_internal or exclude_fd) else "Total Sent",
            format_inr(total_sent),
            help="Real vendor & business payments out of your account",
        )
    with kpi3:
        status_label = "🟢 Net Profit" if net_profit >= 0 else "🔴 Net Deficit"
        st.metric(
            "Net Business Profit" if (exclude_internal or exclude_fd) else "Net Surplus / Profit",
            format_inr(net_profit),
            delta=status_label,
            delta_color="normal" if net_profit >= 0 else "inverse",
        )
    with kpi4:
        st.metric(
            "Profit Margin",
            f"{margin_pct:+.1f}%",
            help="Percentage of received money retained as net profit",
        )

    # Top KPI Cards - Row 2: Asset Investments & Internal Movements (dedicated row below for full visibility)
    has_fd = exclude_fd and total_fd > 0
    has_internal = exclude_internal and total_internal > 0

    if has_fd or has_internal:
        if has_fd and has_internal:
            sub_col1, sub_col2 = st.columns(2)
            with sub_col1:
                st.metric(
                    "🏦 Fixed Deposits & Capital Investments",
                    format_inr(total_fd),
                    help="Capital savings deposited into Fixed Deposits / Term Deposits (separated from operating expenses)",
                )
            with sub_col2:
                st.metric(
                    "🔄 Internal Self-Transfers (Contra)",
                    format_inr(total_internal),
                    help="Money transferred between your own bank accounts (excluded to prevent turnover inflation)",
                )
        elif has_fd:
            sub_col1, _ = st.columns([1, 1])
            with sub_col1:
                st.metric(
                    "🏦 Fixed Deposits & Capital Investments",
                    format_inr(total_fd),
                    help="Capital savings deposited into Fixed Deposits / Term Deposits (separated from operating expenses)",
                )
        elif has_internal:
            sub_col1, _ = st.columns([1, 1])
            with sub_col1:
                st.metric(
                    "🔄 Internal Self-Transfers (Contra)",
                    format_inr(total_internal),
                    help="Money transferred between your own bank accounts (excluded to prevent turnover inflation)",
                )

    if exclude_fd and total_fd > 0:
        st.info(f"🏦 **Fixed Deposit Accounting Active:** Identified **{format_inr(total_fd)}** in Bank Fixed Deposits / Capital Investments and separated them from day-to-day business expenses.", icon="🏦")

    if exclude_internal and total_internal > 0:
        st.info(f"💡 **Contra Handling Active:** Excluded **{format_inr(total_internal)}** of internal self-transfers between accounts to display true business revenue and expenses.", icon="🔄")

    st.divider()

    # Monthly Grouping
    m_grp = active_df.groupby(["YearMonth", "Month", "Type"])["Amount"].sum().unstack(fill_value=0.0)
    if "Credit" not in m_grp:
        m_grp["Credit"] = 0.0
    if "Debit" not in m_grp:
        m_grp["Debit"] = 0.0
    counts = valid.groupby(["YearMonth", "Month"]).size().to_dict()

    m_df = m_grp.reset_index().sort_values("YearMonth", ascending=False)
    m_df["Net Profit / Loss"] = m_df["Credit"] - m_df["Debit"]
    m_df["Status"] = m_df["Net Profit / Loss"].apply(lambda x: "🟢 Profit" if x >= 0 else "🔴 Deficit")
    m_df["Margin %"] = m_df.apply(lambda r: round((r["Net Profit / Loss"] / r["Credit"] * 100), 1) if r["Credit"] > 0 else 0.0, axis=1)

    # Monthly Internal & FD amounts
    self_monthly: dict[tuple[str, str], float] = {}
    for (ym, m_name), s_grp in valid[valid["Is_Internal"]].groupby(["YearMonth", "Month"]):
        s_cr = float(s_grp.loc[s_grp["Type"] == "Credit", "Amount"].sum())
        s_dr = float(s_grp.loc[s_grp["Type"] == "Debit", "Amount"].sum())
        self_monthly[(ym, m_name)] = max(s_cr, s_dr) if (s_cr > 0 and s_dr > 0) else (s_cr + s_dr)

    fd_monthly: dict[tuple[str, str], float] = {}
    for (ym, m_name), f_grp in valid[valid["Is_FD"]].groupby(["YearMonth", "Month"]):
        fd_monthly[(ym, m_name)] = float(f_grp["Amount"].sum())

    m_df["Fixed Deposits"] = m_df.apply(lambda r: fd_monthly.get((r["YearMonth"], r["Month"]), 0.0), axis=1)
    m_df["Self Transfers"] = m_df.apply(lambda r: self_monthly.get((r["YearMonth"], r["Month"]), 0.0), axis=1)
    m_df["Transactions"] = m_df.apply(lambda r: counts.get((r["YearMonth"], r["Month"]), 0), axis=1)
    m_df.rename(columns={"Credit": "Received (Credits)", "Debit": "Sent (Debits)"}, inplace=True)

    # 1. Visual Inflows vs Outflows Chart
    st.markdown("#### 📊 Inflows vs. Outflows Comparison (Month-by-Month)")
    chart_df_sorted = m_df.sort_values("YearMonth", ascending=True)
    plot_rows = []
    month_order = []
    for _, r in chart_df_sorted.iterrows():
        m_name = r["Month"]
        if m_name not in month_order:
            month_order.append(m_name)
        plot_rows.append({
            "Month": m_name,
            "Type": "Received (Inflows)",
            "Amount": float(r["Received (Credits)"]),
            "Formatted": format_inr(r["Received (Credits)"]),
        })
        plot_rows.append({
            "Month": m_name,
            "Type": "Sent (Outflows)",
            "Amount": float(r["Sent (Debits)"]),
            "Formatted": format_inr(r["Sent (Debits)"]),
        })
    plot_df = pd.DataFrame(plot_rows)

    if alt is not None and not plot_df.empty:
        # Indian numbering system expression for chart axis (e.g. ₹ 80,00,000, ₹ 20,00,000)
        indian_axis_expr = (
            "datum.value == 0 ? '₹ 0' : "
            "(datum.value < 0 ? '-₹ ' : '₹ ') + ("
            "abs(datum.value) >= 10000000 ? "
                "floor(abs(datum.value) / 10000000) + ',' + "
                "slice('00' + floor((abs(datum.value) % 10000000) / 100000), -2) + ',' + "
                "slice('00' + floor((abs(datum.value) % 100000) / 1000), -2) + ',' + "
                "slice('000' + (abs(datum.value) % 1000), -3) : "
            "abs(datum.value) >= 100000 ? "
                "floor(abs(datum.value) / 100000) + ',' + "
                "slice('00' + floor((abs(datum.value) % 100000) / 1000), -2) + ',' + "
                "slice('000' + (abs(datum.value) % 1000), -3) : "
            "abs(datum.value) >= 1000 ? "
                "floor(abs(datum.value) / 1000) + ',' + "
                "slice('000' + (abs(datum.value) % 1000), -3) : "
            "'' + abs(datum.value)"
            ")"
        )

        chart = (
            alt.Chart(plot_df)
            .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("Month:N", title=None, sort=month_order, axis=alt.Axis(labelAngle=0, labelFontSize=12)),
                xOffset=alt.XOffset("Type:N", sort=["Received (Inflows)", "Sent (Outflows)"]),
                y=alt.Y(
                    "Amount:Q",
                    title="Amount (₹)",
                    axis=alt.Axis(labelExpr=indian_axis_expr, grid=True, titleFontSize=12),
                ),
                color=alt.Color(
                    "Type:N",
                    scale=alt.Scale(
                        domain=["Received (Inflows)", "Sent (Outflows)"],
                        range=["#10b981", "#ef4444"],
                    ),
                    legend=alt.Legend(
                        title=None,
                        orient="top",
                        labelFontSize=12,
                        symbolType="circle",
                    ),
                ),
                tooltip=[
                    alt.Tooltip("Month:N", title="Month"),
                    alt.Tooltip("Type:N", title="Category"),
                    alt.Tooltip("Formatted:N", title="Amount"),
                ],
            )
            .properties(height=360)
        )
        st.altair_chart(chart, use_container_width=True)
    else:
        chart_data = chart_df_sorted[["Month", "Received (Credits)", "Sent (Debits)"]].set_index("Month")
        chart_data.rename(columns={"Received (Credits)": "Received (Inflows)", "Sent (Debits)": "Sent (Outflows)"}, inplace=True)
        st.bar_chart(chart_data, color=["#10b981", "#ef4444"], stack=False)

    # 2. Detailed Month-by-Month Table
    st.markdown("#### 📋 Month-by-Month Profit Breakdown")
    display_m_df = m_df[["Month", "Received (Credits)", "Sent (Debits)", "Net Profit / Loss", "Fixed Deposits", "Self Transfers", "Status", "Margin %", "Transactions"]].copy()
    display_m_df["Received (Credits)"] = display_m_df["Received (Credits)"].apply(format_inr)
    display_m_df["Sent (Debits)"] = display_m_df["Sent (Debits)"].apply(format_inr)
    display_m_df["Net Profit / Loss"] = display_m_df["Net Profit / Loss"].apply(format_inr)
    display_m_df["Fixed Deposits"] = display_m_df["Fixed Deposits"].apply(format_inr)
    display_m_df["Self Transfers"] = display_m_df["Self Transfers"].apply(format_inr)

    st.dataframe(
        display_m_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Margin %": st.column_config.NumberColumn(format="%0.1f %%"),
        },
    )

    # 3. Bank-wise Breakdown Expander
    with st.expander("🏦 View Bank-Wise Breakdown for Each Month"):
        b_grp = active_df.groupby(["YearMonth", "Month", "Bank", "Type"])["Amount"].sum().unstack(fill_value=0.0).reset_index()
        if "Credit" not in b_grp:
            b_grp["Credit"] = 0.0
        if "Debit" not in b_grp:
            b_grp["Debit"] = 0.0
        b_grp["Net Profit / Loss"] = b_grp["Credit"] - b_grp["Debit"]
        b_grp["Status"] = b_grp["Net Profit / Loss"].apply(lambda x: "🟢 Profit" if x >= 0 else "🔴 Deficit")
        b_grp.rename(columns={"Credit": "Received (Credits)", "Debit": "Sent (Debits)"}, inplace=True)
        b_display = b_grp.sort_values(["YearMonth", "Bank"], ascending=[False, True])[
            ["Month", "Bank", "Received (Credits)", "Sent (Debits)", "Net Profit / Loss", "Status"]
        ].copy()
        b_display["Received (Credits)"] = b_display["Received (Credits)"].apply(format_inr)
        b_display["Sent (Debits)"] = b_display["Sent (Debits)"].apply(format_inr)
        b_display["Net Profit / Loss"] = b_display["Net Profit / Loss"].apply(format_inr)
        st.dataframe(
            b_display,
            use_container_width=True,
            hide_index=True,
        )

    # 4. View Detected Fixed Deposits Expander
    if total_fd > 0:
        with st.expander(f"🏦 View All Fixed Deposits & Capital Investments ({len(valid[valid['Is_FD']]):,} entries)"):
            st.markdown("**These transactions represent your capital savings/deposits and are separated from business operating expenses.**")
            fd_display = valid[valid["Is_FD"]][
                ["Date", "Bank", "Type", "Amount", "UTR / Reference", "Description", "Party Name"]
            ].copy()
            fd_display = sort_dataframe_by_date(fd_display, ascending=False)
            parsed_fd_dates = pd.to_datetime(fd_display["Date"], format="%d-%m-%Y", errors="coerce")
            if parsed_fd_dates.isna().any():
                fallback = pd.to_datetime(fd_display.loc[parsed_fd_dates.isna(), "Date"], dayfirst=True, errors="coerce")
                parsed_fd_dates = parsed_fd_dates.fillna(fallback)
            fd_display["Date"] = parsed_fd_dates.dt.date
            st.dataframe(
                fd_display,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Date": st.column_config.DateColumn("Date", format="DD-MM-YYYY", help="Transaction date (click to sort chronologically)"),
                    "Amount": st.column_config.NumberColumn(format="₹ %0.2f"),
                },
            )

    # 5. View Detected Self-Transfers Expander
    if total_internal > 0:
        with st.expander(f"🔄 View All Detected Internal Self-Transfers ({len(valid[valid['Is_Internal']]):,} entries)"):
            st.markdown("**These transactions move money between your own accounts and are excluded from business revenue and expenses.**")
            self_display = valid[valid["Is_Internal"]][
                ["Date", "Bank", "Type", "Amount", "UTR / Reference", "Description", "Party Name"]
            ].copy()
            self_display = sort_dataframe_by_date(self_display, ascending=False)
            parsed_self_dates = pd.to_datetime(self_display["Date"], format="%d-%m-%Y", errors="coerce")
            if parsed_self_dates.isna().any():
                fallback = pd.to_datetime(self_display.loc[parsed_self_dates.isna(), "Date"], dayfirst=True, errors="coerce")
                parsed_self_dates = parsed_self_dates.fillna(fallback)
            self_display["Date"] = parsed_self_dates.dt.date
            st.dataframe(
                self_display,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Date": st.column_config.DateColumn("Date", format="DD-MM-YYYY", help="Transaction date (click to sort chronologically)"),
                    "Amount": st.column_config.NumberColumn(format="₹ %0.2f"),
                },
            )

    # 6. Download Report
    report_bytes = export_monthly_report(df, exclude_internal=exclude_internal, exclude_fd=exclude_fd)
    st.download_button(
        "📥 Download Monthly Profit Report (.xlsx)",
        data=report_bytes,
        file_name="monthly-profit-report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_monthly_profit_btn",
    )


def app() -> None:
    if st is None:
        raise RuntimeError("The browser interface needs Streamlit. Run: pip install -r requirements.txt")
    st.set_page_config(page_title="Unified Bank Search", page_icon="🔎", layout="wide")
    sb_client = db.get_supabase_client()
    inject_keyboard_shortcuts()
    if sb_client and ("cached_transactions" not in st.session_state or st.session_state["cached_transactions"].empty):
        st.session_state["cached_transactions"] = db.fetch_all_transactions(sb_client)

    # Prevent metric numbers from truncating with ellipsis (...)
    st.markdown(
        """
        <style>
        [data-testid="stMetricValue"] {
            font-size: clamp(1.15rem, 1.7vw, 1.85rem) !important;
            white-space: nowrap !important;
            overflow: visible !important;
            text-overflow: clip !important;
        }
        [data-testid="stMetricValue"] > div {
            overflow: visible !important;
            text-overflow: clip !important;
            white-space: nowrap !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col_title, col_status = st.columns([3, 2], vertical_alignment="center")
    with col_title:
        st.title("Unified Bank Search")
    with col_status:
        badges = []
        if sb_client:
            cached_txns = st.session_state.get("cached_transactions", pd.DataFrame())
            storage_quick = db.get_database_storage_stats(sb_client, cached_txns if not cached_txns.empty else None)
            free_space_text = f"{storage_quick['free_mb']:.1f} MB Free / {storage_quick['total_quota_mb']:.0f} MB"
            badges.append(f"🟢 **Supabase Connected** • 💾 **{free_space_text}**")
        else:
            badges.append("💾 **Local Mode**")

        gemini_key = get_gemini_api_key()
        if gemini_key:
            usage_badge = db.get_gemini_api_usage(sb_client)
            badges.append(f"✨ **Gemini API:** `{usage_badge['remaining_today']:,} / {usage_badge['daily_quota']:,} left today`")
        else:
            badges.append("📄 **PDFs need Gemini Key**")

        st.caption(" | ".join(badges))

    gemini_key = get_gemini_api_key()
    with st.expander("✨ Gemini AI Settings & Daily Quota Monitor", expanded=(not bool(gemini_key))):
        st.markdown(
            "**Extract Bank Statements Directly from PDFs:**\n\n"
            "StatementHub uses Google's Gemini multimodal AI to read and extract tabular transactions from bank statement PDFs automatically."
        )
        col_gk_input, col_gk_btn = st.columns([4, 1.2], vertical_alignment="bottom")
        with col_gk_input:
            new_gk = st.text_input(
                "Gemini API Key",
                value=gemini_key if gemini_key else "",
                type="password",
                placeholder="Paste your Gemini API key (e.g. AIzaSy...)",
                help="Get a free API key at https://aistudio.google.com",
            )
        with col_gk_btn:
            if st.button("💾 Save Key", use_container_width=True, key="save_gemini_key_btn"):
                if new_gk.strip():
                    save_gemini_api_key(new_gk.strip())
                    st.toast("Gemini API key saved successfully!", icon="✅")
                    st.rerun()
                else:
                    st.warning("Please enter a valid API key.")
        if gemini_key:
            st.caption("🟢 **Gemini API Key Configured & Ready for PDF Uploads**")
            usage = db.get_gemini_api_usage(sb_client)
            st.divider()
            st.markdown("##### 📊 Daily API Quota & Usage Monitor")

            q_col1, q_col2, q_col3 = st.columns(3)
            with q_col1:
                st.metric(
                    "Requests Used Today",
                    f"{usage['requests_today']:,}",
                    help="API requests consumed today since midnight UTC (5:30 AM IST)",
                )
            with q_col2:
                st.metric(
                    "Requests Remaining Today",
                    f"{usage['remaining_today']:,}",
                    delta=f"{usage['daily_quota']:,} Daily Limit",
                    delta_color="normal",
                )
            with q_col3:
                st.metric(
                    "Tokens Consumed",
                    f"{usage.get('total_tokens', 0):,}",
                    help="Total input & output tokens processed by Gemini",
                )

            pct = min(1.0, usage["requests_today"] / max(1, usage["daily_quota"]))
            st.progress(
                pct,
                text=f"{usage['requests_today']:,} of {usage['daily_quota']:,} daily requests used ({pct * 100:.1f}%) • Free Tier Quota",
            )

            st.caption(
                f"🟢 **Google AI Studio Free Tier:** **1,500 requests/day** • **15 requests/minute**. "
                f"Quota automatically resets daily at 00:00 UTC (5:30 AM IST). "
                + (f"*(Last request: {usage['last_used']})*" if usage.get("last_used") else "")
            )

            col_link, col_reset = st.columns([3, 1])
            with col_link:
                st.link_button(
                    "🔗 Open Google AI Studio Official Usage Dashboard",
                    "https://aistudio.google.com/",
                    use_container_width=True,
                )
            with col_reset:
                if st.button("🔄 Reset Counter", key="reset_gemini_usage_btn", use_container_width=True):
                    db.reset_gemini_api_usage(sb_client)
                    st.toast("API usage counter reset!", icon="✅")
                    st.rerun()
        else:
            st.caption("ℹ️ Get a free API key from [Google AI Studio](https://aistudio.google.com) to enable PDF uploads.")

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

    cached_txns_hint = st.session_state.get("cached_transactions", pd.DataFrame())
    if not cached_txns_hint.empty and "Bank" in cached_txns_hint.columns and "Date" in cached_txns_hint.columns:
        df_hint = cached_txns_hint.copy()
        dt_col = pd.to_datetime(df_hint["Date"], format="%d-%m-%Y", errors="coerce")
        if dt_col.isna().any():
            fallback = pd.to_datetime(df_hint.loc[dt_col.isna(), "Date"], dayfirst=True, errors="coerce")
            dt_col = dt_col.fillna(fallback)
        df_hint["_dt"] = dt_col

        bank_items = []
        for b in sorted([x for x in df_hint["Bank"].dropna().unique() if str(x).strip()]):
            b_max = df_hint.loc[df_hint["Bank"] == b, "_dt"].dropna().max()
            if pd.notna(b_max):
                bank_items.append(f"**{b}:** `{b_max.strftime('%d-%b-%Y')}`")

        if bank_items:
            st.info(
                "**Latest Recorded Dates in Database:** "
                + " &nbsp;•&nbsp; ".join(bank_items)
                + " &nbsp;*(upload statements from these dates onwards)*",
                icon="📅",
            )

    files = st.file_uploader(
        "Add or replace statements", type=["csv", "xls", "xlsx", "pdf"], accept_multiple_files=True,
        help="Select statements to upload. With Supabase, existing transactions are automatically skipped and only new ones are inserted. PDFs are extracted automatically using Gemini AI.",
    )

    current_file_signature = [(f.name, getattr(f, "size", 0)) for f in files] if files else []

    # Process new uploads ONLY when the file set changes
    if files and st.session_state.get("last_uploaded_signature") != current_file_signature:
        with st.spinner("Extracting transactions from statement(s)... (Gemini AI parsing PDF)"):
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
        render_storage_usage_section(combined, sb_client)
        render_backup_restore_section(combined, assignments, sb_client)
        return

    combined = apply_assignments(combined, assignments)
    combined["Is_Internal"] = detect_internal_transfers(combined)
    combined["Is_FD"] = detect_fd_investments(combined)
    combined = sort_dataframe_by_date(combined, ascending=False)

    tab_search, tab_profit = st.tabs(["🔎 Unified Transaction Search", "💰 Monthly Profit & Cash Flow"])

    with tab_search:
        stored_party_names = load_stored_party_names(sb_client, combined)
        # Filters: Bank Account, Time Period & Chronological Sorting
        filter_col1, filter_col2, filter_col3 = st.columns([1.5, 1.5, 1.2])
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
        with filter_col3:
            sort_order = st.selectbox(
                "Sort Order",
                options=[
                    "Date: Newest First (Latest)",
                    "Date: Oldest First",
                    "Amount: Highest First",
                    "Amount: Lowest First",
                ],
                index=0,
                help="Order transactions chronologically or by amount",
            )

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
            if parsed_dates.isna().any():
                fallback = pd.to_datetime(filtered_df.loc[parsed_dates.isna(), "Date"], dayfirst=True, errors="coerce")
                parsed_dates = parsed_dates.fillna(fallback)
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
                    mask &= (parsed_dates <= pd.Timestamp(custom_end_date).replace(hour=23, minute=59, second=59))
                filtered_df = filtered_df[mask].copy()

        # Quick Cash Flow Ribbon on search screen (with contra & FD detection)
        is_internal_filtered = filtered_df["Is_Internal"] if "Is_Internal" in filtered_df.columns else detect_internal_transfers(filtered_df)
        is_fd_filtered = filtered_df["Is_FD"] if "Is_FD" in filtered_df.columns else detect_fd_investments(filtered_df)
        biz_filtered = filtered_df[(~is_internal_filtered) & (~is_fd_filtered)]
        rec_sum = float(biz_filtered.loc[biz_filtered["Type"] == "Credit", "Amount"].sum()) if not biz_filtered.empty else 0.0
        sent_sum = float(biz_filtered.loc[biz_filtered["Type"] == "Debit", "Amount"].sum()) if not biz_filtered.empty else 0.0
        net_sum = rec_sum - sent_sum
        int_filtered_dr = float(filtered_df.loc[is_internal_filtered & (filtered_df["Type"] == "Debit"), "Amount"].sum())
        int_filtered_cr = float(filtered_df.loc[is_internal_filtered & (filtered_df["Type"] == "Credit"), "Amount"].sum())
        internal_sum = max(int_filtered_dr, int_filtered_cr) if (int_filtered_dr > 0 and int_filtered_cr > 0) else (int_filtered_dr + int_filtered_cr)
        fd_sum = float(filtered_df.loc[is_fd_filtered, "Amount"].sum())

        q_col1, q_col2, q_col3 = st.columns(3)
        with q_col1:
            st.metric("Business Received (Inflows)", format_inr(rec_sum), help="Real customer revenue for selected bank & period (excludes self-transfers & FDs)")
        with q_col2:
            st.metric("Business Sent (Outflows)", format_inr(sent_sum), help="Real vendor & business payments for selected bank & period (excludes self-transfers & FDs)")
        with q_col3:
            status_label = "🟢 In Profit" if net_sum >= 0 else "🔴 In Deficit"
            st.metric("Net Business Profit", format_inr(net_sum), delta=status_label, delta_color="normal" if net_sum >= 0 else "inverse")
        info_notes = []
        if internal_sum > 0:
            info_notes.append(f"🔄 Excluded **{format_inr(internal_sum)}** internal self-transfers")
        if fd_sum > 0:
            info_notes.append(f"🏦 Excluded **{format_inr(fd_sum)}** Fixed Deposits / Investments")
        if info_notes:
            st.caption("💡 " + " | ".join(info_notes))

        search_modes = {
            "UTR / Reference": "utr",
            "Date": "date",
            "Amount": "amount",
            "Party Name": "party",
        }
        selected_mode = st.radio("Search by", list(search_modes.keys()), horizontal=True, index=0)
        mode = search_modes[selected_mode]

        if mode == "utr":
            all_ids = sorted(filtered_df.loc[filtered_df["UTR / Reference"] != "", "UTR / Reference"].drop_duplicates().tolist())
            col_search, col_copy = st.columns([5, 1], vertical_alignment="bottom")
            with col_search:
                effective_query = st.selectbox(
                    "Search UTR / Reference",
                    options=all_ids,
                    index=None,
                    placeholder="Start typing a UTR (for example: 3964)",
                    accept_new_options=True,
                ) or ""
            with col_copy:
                if effective_query:
                    render_single_copy_btn(effective_query, label="📋 Copy UTR")
        elif mode == "date":
            effective_query = st.text_input(
                "Search Date",
                placeholder="Enter date (e.g. 12-10-2026, 10-2026, or 2026)",
            ).strip()
        elif mode == "amount":
            effective_query = st.text_input(
                "Search Amount",
                placeholder="Enter amount (e.g. 90, 25000, or ₹25,000)",
            ).strip()
        else:  # party
            effective_query = st.selectbox(
                "Search by Saved Party Name",
                options=stored_party_names,
                index=None,
                placeholder="Select or start typing party name...",
                accept_new_options=True,
            ) or ""

        results = search_transactions(filtered_df, effective_query, mode)
        if sort_order == "Date: Oldest First":
            results = sort_dataframe_by_date(results, ascending=True)
        elif sort_order == "Amount: Highest First":
            results = results.sort_values(by="Amount", ascending=False).reset_index(drop=True)
        elif sort_order == "Amount: Lowest First":
            results = results.sort_values(by="Amount", ascending=True).reset_index(drop=True)
        else:
            results = sort_dataframe_by_date(results, ascending=False)

        result_ids = results[results["UTR / Reference"] != ""]["UTR / Reference"]
        duplicated = result_ids[result_ids.duplicated(keep=False)].unique().tolist()
        if duplicated:
            st.warning(f"Duplicate identifier detected in these results: {', '.join(duplicated[:8])}" + (" …" if len(duplicated) > 8 else ""))
        stored_party_names = load_stored_party_names(sb_client, combined)

        st.subheader(f"Results ({len(results):,})")
        st.caption("💡 Double-click or press Enter on any transaction's **Party Name** cell to enter party name (matching saved names will appear as you type). Use the **Sort Order** filter above to sort.")

        render_bank_color_legend(results)

        inject_party_autocomplete(stored_party_names)

        results = results.reset_index(drop=True)
        results["Self Transfer"] = results["Is_Internal"] if "Is_Internal" in results.columns else detect_internal_transfers(results)
        results["FD / Investment"] = results["Is_FD"] if "Is_FD" in results.columns else detect_fd_investments(results)

        # Keep editor_key stable so editing a party name does NOT destroy/reload the table or reset scroll position
        filter_sig = f"{selected_bank}_{selected_period}_{sort_order}_{mode}_{effective_query}"
        if st.session_state.get("last_editor_filter_sig") != filter_sig:
            st.session_state["last_editor_filter_sig"] = filter_sig
            st.session_state["editor_view_id"] = st.session_state.get("editor_view_id", 0) + 1

        editor_key = f"transactions_editor_{st.session_state.get('editor_view_id', 0)}"

        display_columns = [
            "Bank", "Date", "Description", "Party Name", "Amount", "Type",
            "UTR / Reference", "Balance", "Self Transfer", "FD / Investment", "Duplicate UTR",
        ]

        display_df = results[display_columns].copy()
        parsed_disp_dates = pd.to_datetime(display_df["Date"], format="%d-%m-%Y", errors="coerce")
        if parsed_disp_dates.isna().any():
            fallback = pd.to_datetime(display_df.loc[parsed_disp_dates.isna(), "Date"], dayfirst=True, errors="coerce")
            parsed_disp_dates = parsed_disp_dates.fillna(fallback)
        display_df["Date"] = parsed_disp_dates.dt.date
        display_df["_fingerprint"] = results["fingerprint"].values

        styled_results = style_table_by_bank(display_df)

        edited_results = st.data_editor(
            styled_results,
            disabled=[col for col in display_df.columns if col != "Party Name"],
            use_container_width=True,
            hide_index=True,
            key=editor_key,
            column_config={
                "_fingerprint": None,
                "Bank": st.column_config.TextColumn("Bank", help="Transaction rows are color-coded by Bank"),
                "Date": st.column_config.DateColumn("Date", format="DD-MM-YYYY"),
                "Amount": st.column_config.NumberColumn(format="₹ %0.2f"),
                "Balance": st.column_config.NumberColumn(format="₹ %0.2f"),
                "Self Transfer": st.column_config.CheckboxColumn("Self Transfer", help="Detected internal transfer between your own accounts"),
                "FD / Investment": st.column_config.CheckboxColumn("FD / Investment", help="Detected Fixed Deposit or capital investment"),
                "Party Name": st.column_config.TextColumn(
                    "Party Name",
                    help="Double-click or press Enter to edit party name",
                ),
            },
        )

        orig_party_map = dict(
            zip(
                results["fingerprint"].astype(str),
                results["Party Name"].fillna("").astype(str).str.strip(),
            )
        )

        changes_by_fp: dict[str, str] = {}
        if not edited_results.empty and "_fingerprint" in edited_results.columns and "Party Name" in edited_results.columns:
            for _, r in edited_results.iterrows():
                fp = str(r.get("_fingerprint", "")).strip()
                if not fp:
                    continue
                new_val = clean_text(r.get("Party Name", "")).strip()
                old_val = orig_party_map.get(fp, "")
                if new_val != old_val:
                    changes_by_fp[fp] = new_val

        editor_state = st.session_state.get(editor_key, {})
        edited_rows = editor_state.get("edited_rows", {})
        if not changes_by_fp and edited_rows and "_fingerprint" in display_df.columns:
            for idx_raw, col_edits in edited_rows.items():
                if "Party Name" in col_edits:
                    try:
                        row_idx = int(idx_raw)
                        if 0 <= row_idx < len(display_df):
                            fp = str(display_df.iloc[row_idx]["_fingerprint"]).strip()
                            new_val = clean_text(col_edits["Party Name"]).strip()
                            old_val = orig_party_map.get(fp, "")
                            if fp and new_val != old_val:
                                changes_by_fp[fp] = new_val
                    except Exception:
                        pass

        if changes_by_fp:
            updated_count = 0
            for fp, new_val in changes_by_fp.items():
                assign_party_name_to_transactions([fp], new_val, combined, sb_client)
                results.loc[results["fingerprint"] == fp, "Party Name"] = new_val
                updated_count += 1

            if updated_count > 0:
                st.session_state["cached_transactions"] = combined
                st.session_state["editor_view_id"] = st.session_state.get("editor_view_id", 0) + 1
                st.toast(f"Saved party name for {updated_count} transaction{'s' if updated_count > 1 else ''}.", icon="✅")
                st.rerun()


        with st.expander(f"📁 Stored Party Names ({len(stored_party_names)})"):
            st.markdown("**Directory of stored party names used for auto-completion suggestions.**")
            col_new_p, col_add_btn = st.columns([4, 1.2], vertical_alignment="bottom")
            with col_new_p:
                new_party_to_add = st.text_input(
                    "Add a new party name to suggestions list",
                    placeholder="e.g. SRI RAMA MEDICALS",
                    key="add_party_input",
                )
            with col_add_btn:
                if st.button("➕ Add Name", use_container_width=True, key="add_party_btn"):
                    clean_new = clean_text(new_party_to_add).strip()
                    if clean_new:
                        if clean_new not in stored_party_names:
                            stored_party_names.append(clean_new)
                            save_stored_party_names(stored_party_names, sb_client)
                            st.toast(f"Added '{clean_new}' to stored party names.", icon="✅")
                            st.rerun()
                        else:
                            st.info(f"'{clean_new}' is already in stored party names.")

            if stored_party_names:
                counts = combined["Party Name"].value_counts().to_dict() if not combined.empty and "Party Name" in combined.columns else {}
                p_data = [
                    {"Party Name": p, "Assigned Transactions": counts.get(p, 0)}
                    for p in stored_party_names
                ]
                st.dataframe(pd.DataFrame(p_data), use_container_width=True, hide_index=True)

        export = export_excel(results)
        st.download_button(
            "Download displayed results as Excel",
            export,
            "unified-search-results.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with tab_profit:
        render_monthly_profit_section(combined)

    render_storage_usage_section(combined, sb_client)
    render_backup_restore_section(combined, assignments, sb_client)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--verify", nargs="+")
    args, _ = parser.parse_known_args()
    if args.verify:
        import_paths(args.verify)
    else:
        app()
