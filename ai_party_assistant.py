"""AI-powered Party Name Assistant for StatementHub.

Uses Google Gemini to analyze unassigned bank transactions, match them
against the user's stored party directory, and present an interactive
preview table for human review and one-click approval.
"""

from __future__ import annotations

import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

import db


def get_gemini_client(api_key: str | None = None) -> Any:
    """Initialize Google GenAI client with provided or saved API key."""
    if not api_key:
        # Check Streamlit session state, secrets, or env
        if "gemini_api_key" in st.session_state and st.session_state["gemini_api_key"]:
            api_key = str(st.session_state["gemini_api_key"]).strip()
        elif hasattr(st, "secrets") and "gemini" in st.secrets and "api_key" in st.secrets["gemini"]:
            api_key = str(st.secrets["gemini"]["api_key"]).strip()
        elif hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
            api_key = str(st.secrets["GEMINI_API_KEY"]).strip()
        else:
            sec_file = Path(".streamlit/secrets.toml")
            if sec_file.exists():
                try:
                    import toml
                    data = toml.load(sec_file)
                    api_key = data.get("gemini", {}).get("api_key") or data.get("GEMINI_API_KEY")
                except Exception:
                    pass
        if not api_key:
            api_key = os.environ.get("GEMINI_API_KEY", "").strip()

    if not api_key:
        return None

    try:
        from google import genai
        return genai.Client(api_key=api_key)
    except Exception as e:
        st.error(f"Failed to initialize Gemini GenAI client: {e}")
        return None


def generate_ai_party_suggestions(
    unassigned_records: list[dict[str, Any]],
    stored_party_names: list[str],
    api_key: str | None = None,
    sb_client: Any = None,
) -> list[dict[str, Any]]:
    """Analyze unassigned transactions using Gemini and suggest party names."""
    if not unassigned_records:
        return []

    client = get_gemini_client(api_key)
    if client is None:
        raise ValueError("Gemini API key is not configured. Please add it to .streamlit/secrets.toml or Gemini AI Settings.")

    from google.genai import types

    # Prepare compact transaction payload
    tx_payload = [
        {
            "fingerprint": r.get("fingerprint", ""),
            "date": str(r.get("Date", "")),
            "amount": float(r.get("Amount", 0.0)),
            "type": str(r.get("Type", "")),
            "bank": str(r.get("Bank", "")),
            "description": str(r.get("Description", "")),
        }
        for r in unassigned_records
    ]

    clean_stored = sorted(list(set(str(p).strip() for p in stored_party_names if str(p).strip())))
    clean_stored_map = {p.upper().strip(): p.strip() for p in clean_stored if p.strip()}

    prompt = f"""You are an expert Indian bank statement transaction and party classifier.
Analyze the following bank transactions and match them STRICTLY against the Known Party Directory.

Known Party Directory (MANDATORY CLOSED SET):
{json.dumps(clean_stored)}

CRITICAL INSTRUCTIONS:
1. STRICT CLOSED-SET ONLY: You must ONLY assign a party name if it exists in the Known Party Directory above.
2. DO NOT INVENT OR CREATE ANY NEW PARTY NAMES. If a transaction does not clearly match any name in the Known Party Directory, set "suggested_party": "" and "confidence": "Low".
3. Use exact string matching from the Known Party Directory:
   - Utility or electricity bills (e.g. UPI Bharat Connect Util, power bills) -> map to 'ELECTRICITY BILLS' (only if present in directory).
   - Beneficiary payments with vendor names (e.g. Lupin, Dr Reddys, Mankind, pharmacy stores) -> map to their exact matching name in the directory.
   - Bank charges, NEFT charges, or SMS charges -> map to 'NEFT CHARGES' or 'SMS CHARGES' (only if present in directory).
   - Cash deposits -> map to 'CASH DEPOSIT' (only if present in directory).
   - Home loan EMI -> map to 'HOME LOAN EMI' (only if present in directory).
4. If you are unsure or the vendor/party is not in the directory, return "suggested_party": "" and "reason": "Not in stored party directory".

Transactions to classify:
{json.dumps(tx_payload)}

Required JSON schema:
[
  {{
    "fingerprint": "string matching input fingerprint",
    "suggested_party": "EXACT string from Known Party Directory, or empty string '' if no match",
    "confidence": "High" | "Medium" | "Low",
    "reason": "short explanation (under 12 words)"
  }}
]
"""

    models_to_try = ["gemini-3.8-flash", "gemini-3.5-flash"]
    last_err = None
    response = None

    for model_name in models_to_try:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1,
                ),
            )
            # Log usage
            try:
                tokens_used = getattr(getattr(response, "usage_metadata", None), "total_token_count", 0) or 0
                db.record_gemini_api_call(sb_client, tokens=tokens_used)
            except Exception:
                pass
            break
        except Exception as exc:
            last_err = exc
            continue

    if response is None or not response.text:
        raise RuntimeError(f"Gemini API failed to return response: {last_err}")

    raw_text = response.text.strip()
    # Strip markdown code blocks if wrapped
    if raw_text.startswith("```"):
        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
        raw_text = re.sub(r"\s*```$", "", raw_text)

    try:
        results = json.loads(raw_text)
    except Exception as e:
        raise ValueError(f"Failed to parse Gemini JSON output: {e}\nRaw output: {raw_text[:200]}")

    results_by_fp = {r.get("fingerprint"): r for r in results if isinstance(r, dict)}

    # Merge suggestions with original record metadata
    enriched = []
    for orig in unassigned_records:
        fp = orig.get("fingerprint", "")
        ai_res = results_by_fp.get(fp, {})
        raw_sugg = (ai_res.get("suggested_party") or "").strip()
        conf = (ai_res.get("confidence") or "Medium").capitalize()
        reason = ai_res.get("reason") or "AI analyzed narration"

        # STRICT ENFORCEMENT: Only allow names from the stored party directory
        canonical = clean_stored_map.get(raw_sugg.upper(), "")
        if canonical:
            sugg = canonical
            approved_default = bool(conf in ["High", "Medium"])
        else:
            # If not in the stored directory, leave blank for manual assignment
            sugg = ""
            conf = "Low"
            reason = "Not in stored party directory (Manual mapping needed)"
            approved_default = False

        enriched.append({
            "Approve": approved_default,
            "Date": orig.get("Date", ""),
            "Amount": float(orig.get("Amount", 0.0)),
            "Type": orig.get("Type", ""),
            "Bank": orig.get("Bank", ""),
            "Suggested Party": sugg,
            "Confidence": f"🟢 High" if conf == "High" else (f"🟡 Medium" if conf == "Medium" else "🔴 Low"),
            "Reason": reason,
            "Description": orig.get("Description", ""),
            "_fingerprint": fp,
        })

    return enriched


def render_ai_party_assistant_section(
    combined_df: pd.DataFrame,
    stored_party_names: list[str],
    sb_client: Any = None,
    save_stored_party_names_fn: Any = None,
) -> None:
    """Render the AI Party Name Assistant banner, generator, and review table."""
    if combined_df.empty or "Party Name" not in combined_df.columns:
        return

    # Find unassigned transactions
    unassigned_mask = combined_df["Party Name"].fillna("").astype(str).str.strip().eq("")
    unassigned_total = int(unassigned_mask.sum())

    # Cycle counter forces Streamlit expander to collapse shut when user applies or cancels
    cycle = st.session_state.get("ai_expander_cycle", 0)
    has_suggestions = bool(st.session_state.get("ai_party_suggestions"))
    zw_space = "\u200b" * (cycle % 10)
    expander_label = f"✨ AI Party Name Assistant ({unassigned_total:,} Unassigned Transactions){zw_space}"

    with st.expander(expander_label, expanded=has_suggestions):
        st.markdown(
            "**Use Google Gemini AI to analyze bank narrations and suggest vendor / party names with full review and approval control.**"
        )

        if unassigned_total == 0:
            st.success("🎉 All transactions in your statement currently have assigned party names! No unassigned transactions found.")
            return

        usage = db.get_gemini_api_usage(sb_client)
        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        with col_stat1:
            st.metric("Unassigned Transactions", f"{unassigned_total:,}")
        with col_stat2:
            st.metric("Saved Party Directory", f"{len(stored_party_names):,} names")
        with col_stat3:
            # Check for internal transfers/FDs
            internal_count = int((unassigned_mask & (combined_df.get("Is_Internal", False) | combined_df.get("Is_FD", False))).sum())
            st.metric("Internal Transfers / FDs", f"{internal_count:,}", help="Excluded from operating vendor mapping by default")
        with col_stat4:
            st.metric(
                "Gemini Requests Left",
                f"{usage['remaining_today']:,}",
                delta=f"{usage['requests_today']:,} used today",
                delta_color="normal",
                help="Google AI Studio free tier: 1,500 requests/day",
            )

        col_opt1, col_opt2 = st.columns([2, 1], vertical_alignment="bottom")
        with col_opt1:
            batch_limit = st.slider(
                "Batch size to analyze with AI",
                min_value=min(5, unassigned_total),
                max_value=min(100, unassigned_total),
                value=min(30, unassigned_total),
                step=5 if unassigned_total >= 10 else 1,
                help="Process transactions in quick, manageable batches for easy human review.",
            )
        with col_opt2:
            skip_internal = st.checkbox("Skip internal transfers & FDs", value=True)

        # Filter candidate records
        candidate_df = combined_df[unassigned_mask].copy()
        if skip_internal:
            if "Is_Internal" in candidate_df.columns:
                candidate_df = candidate_df[~candidate_df["Is_Internal"]].copy()
            if "Is_FD" in candidate_df.columns:
                candidate_df = candidate_df[~candidate_df["Is_FD"]].copy()

        if candidate_df.empty:
            st.info("All remaining unassigned transactions are marked as internal transfers or FDs.")
            return

        batch_records = candidate_df.head(batch_limit).to_dict(orient="records")

        btn_col, clear_col = st.columns([3, 1], vertical_alignment="bottom")
        with btn_col:
            generate_clicked = st.button(
                f"✨ Generate AI Suggestions for Next {len(batch_records)} Transactions",
                type="primary",
                use_container_width=True,
                key="btn_generate_ai_suggestions",
            )
        with clear_col:
            if st.session_state.get("ai_party_suggestions"):
                if st.button("🗑️ Clear Suggestions", use_container_width=True, key="btn_clear_ai_suggestions"):
                    st.session_state["ai_party_suggestions"] = None
                    st.rerun()

        if generate_clicked:
            with st.spinner("🤖 Google Gemini is analyzing statement narrations against your party directory..."):
                try:
                    suggestions = generate_ai_party_suggestions(
                        unassigned_records=batch_records,
                        stored_party_names=stored_party_names,
                        sb_client=sb_client,
                    )
                    st.session_state["ai_party_suggestions"] = suggestions
                    st.toast(f"Generated {len(suggestions)} AI party suggestions ready for review!", icon="✨")
                    st.rerun()
                except Exception as e:
                    st.error(f"AI Suggestion Error: {e}")

        # Render Review & Approve Interface if suggestions exist
        suggestions_data = st.session_state.get("ai_party_suggestions")
        if suggestions_data:
            st.divider()
            st.markdown("#### 📋 Review & Approve AI Suggestions")
            st.caption(
                "💡 **Strict Closed-Set Matching:** AI strictly matches against your **Stored Party Directory** and never creates new party names. "
                "Any transaction not matching your saved list is left blank for you to map manually. "
                "Review the suggestions below, toggle checkboxes, or edit names before clicking **Apply Approved**."
            )

            preview_df = pd.DataFrame(suggestions_data)

            # Bulk toggle controls
            t_col1, t_col2, t_col3 = st.columns([1.5, 1.5, 4])
            with t_col1:
                if st.button("☑️ Select All", key="ai_select_all_btn", use_container_width=True):
                    for item in suggestions_data:
                        item["Approve"] = True
                    st.session_state["ai_party_suggestions"] = suggestions_data
                    st.rerun()
            with t_col2:
                if st.button("⬜ Deselect All", key="ai_deselect_all_btn", use_container_width=True):
                    for item in suggestions_data:
                        item["Approve"] = False
                    st.session_state["ai_party_suggestions"] = suggestions_data
                    st.rerun()

            edited_preview = st.data_editor(
                preview_df,
                disabled=["Date", "Amount", "Type", "Bank", "Confidence", "Reason", "Description", "_fingerprint"],
                hide_index=True,
                use_container_width=True,
                key="ai_preview_editor",
                column_config={
                    "Approve": st.column_config.CheckboxColumn("Approve?", help="Check to save this party name"),
                    "Date": st.column_config.TextColumn("Date", width="small"),
                    "Amount": st.column_config.NumberColumn("Amount", format="₹ %0.2f", width="small"),
                    "Type": st.column_config.TextColumn("Type", width="small"),
                    "Bank": st.column_config.TextColumn("Bank", width="small"),
                    "Suggested Party": st.column_config.TextColumn(
                        "Suggested Party (Editable)",
                        help="Double-click to edit name before saving",
                        width="medium",
                    ),
                    "Confidence": st.column_config.TextColumn("Confidence", width="small"),
                    "Reason": st.column_config.TextColumn("AI Reasoning", width="medium"),
                    "Description": st.column_config.TextColumn("Statement Narration", width="large"),
                    "_fingerprint": None,
                },
            )

            # Count approved rows
            approved_count = int(edited_preview["Approve"].sum()) if not edited_preview.empty and "Approve" in edited_preview.columns else 0

            save_col1, save_col2 = st.columns([3, 1], vertical_alignment="bottom")
            with save_col1:
                apply_clicked = st.button(
                    f"✅ Apply {approved_count} Approved Party Names to Database",
                    type="primary",
                    disabled=(approved_count == 0),
                    use_container_width=True,
                    key="btn_apply_ai_approved",
                )
            with save_col2:
                if st.button("Cancel", use_container_width=True, key="btn_cancel_ai_approved"):
                    st.session_state["ai_expander_cycle"] = cycle + 1
                    st.session_state["ai_party_suggestions"] = None
                    st.rerun()

            if apply_clicked and approved_count > 0:
                approved_rows = edited_preview[edited_preview["Approve"]].copy()
                saved_count = 0
                new_parties_introduced: set[str] = set()

                for _, row in approved_rows.iterrows():
                    fp = str(row.get("_fingerprint", "")).strip()
                    party_name = str(row.get("Suggested Party", "")).strip().upper()
                    if fp and party_name:
                        # Update in combined_df
                        mask = combined_df["fingerprint"] == fp
                        combined_df.loc[mask, "Party Name"] = party_name
                        # Update in Supabase
                        if sb_client:
                            db.update_transactions_party_bulk(sb_client, [fp], party_name)
                        new_parties_introduced.add(party_name)
                        saved_count += 1

                # Update stored party names directory if any new names were added
                if new_parties_introduced:
                    updated_stored = sorted(set(stored_party_names) | new_parties_introduced)
                    if save_stored_party_names_fn:
                        save_stored_party_names_fn(updated_stored, sb_client)
                    elif sb_client:
                        for p in new_parties_introduced:
                            try:
                                db.save_party_name(sb_client, p)
                            except Exception:
                                pass

                # Clear AI suggestions, increment expander cycle to collapse shut, and refresh view
                st.session_state["ai_expander_cycle"] = cycle + 1
                st.session_state["ai_party_suggestions"] = None
                st.session_state["cached_transactions"] = combined_df
                st.session_state["editor_view_id"] = st.session_state.get("editor_view_id", 0) + 1

                st.toast(f"Successfully applied {saved_count} party names to database!", icon="✅")
                st.rerun()
