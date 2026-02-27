# -*- coding: utf-8 -*-
"""
Streamlit-only Inspector dashboard
- Upload CSV or Excel (.xlsx), up to 500 MB (see .streamlit/config.toml)
- Optional owners.json (account_id -> {owner, email})
- Header/no-header support + column auto-detect + manual mapping
- If no AccountId column, can extract 12-digit account_id from FindingArn
- Per-account XLSX download
- "Compose in Outlook (web)" link with To/Subject/Body prefilled
  NOTE: Attachments cannot be added via URL; user must attach the XLSX manually.
"""

from __future__ import annotations

import io
import json
import re
import urllib.parse as up
from dataclasses import dataclass
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

# ---------- Models ----------

@dataclass
class AccountRow:
    account_id: str
    account_name: str
    total_findings: int
    high_pct: float
    owner: str
    email: str
    xlsx_bytes: bytes
    filename: str


# ---------- Utilities ----------

RE_ACCOUNT_12 = re.compile(r"\b(\d{12})\b")

def _norm(s: str) -> str:
    """Normalize a column name for matching."""
    return str(s).strip().lower().replace(" ", "").replace("-", "").replace("_", "")

def owners_from_uploaded(uploaded_file) -> Dict[str, dict]:
    """Load owners.json if provided; else {}."""
    if not uploaded_file:
        return {}
    try:
        return json.load(uploaded_file)
    except Exception:
        uploaded_file.seek(0)
        return json.loads(uploaded_file.read().decode("utf-8"))

def suggest_column_names(df: pd.DataFrame) -> dict:
    """Try to auto-map dataset columns to logical fields (account_id, account_name, severity, finding_arn)."""
    if df.empty:
        return {}

    candidates = {c: _norm(c) for c in df.columns}

    # Synonyms for each field (normalized)
    account_id_syn   = {"accountid", "account", "account_id", "acctid", "acct", "awsaccountid"}
    account_name_syn = {"accountname", "account_name", "acctname", "name"}
    severity_syn     = {"severity", "sev"}
    arn_syn          = {"findingarn", "arn"}

    pick = {"account_id": None, "account_name": None, "severity": None, "finding_arn": None}

    for col, norm in candidates.items():
        if pick["account_id"] is None and norm in account_id_syn:
            pick["account_id"] = col
        if pick["account_name"] is None and norm in account_name_syn:
            pick["account_name"] = col
        if pick["severity"] is None and norm in severity_syn:
            pick["severity"] = col
        if pick["finding_arn"] is None and norm in arn_syn:
            pick["finding_arn"] = col

    return pick

def extract_account_from_arn(arn: str) -> Optional[str]:
    if not isinstance(arn, str):
        return None
    m = RE_ACCOUNT_12.search(arn)
    return m.group(1) if m else None

def outlook_web_compose_link(to: str, subject: str, body: str) -> str:
    """
    Generate Outlook on the web (OWA) compose deeplink with To/Subject/Body.
    NOTE: Attachments cannot be added via URL; user must attach manually.
    """
    base = "https://outlook.office.com/mail/deeplink/compose"
    params = {
        "to": to or "",
        "subject": subject or "",
        "body": body or "",
    }
    q = "&".join(f"{k}={up.quote(v)}" for k, v in params.items())
    return f"{base}?{q}"


# ---------- Core processing ----------

def compute_rows(
    df_norm: pd.DataFrame,
    owners: Dict[str, dict],
    count_critical_as_high: bool = False,
) -> List[AccountRow]:
    """
    df_norm must contain columns: account_id, account_name, severity (uppercased)
    """
    missing = [c for c in ["account_id", "severity", "account_name"] if c not in df_norm.columns]
    if missing:
        raise ValueError(f"Internal error: normalized frame missing columns: {missing}")

    rows: List[AccountRow] = []
    for acct, group in df_norm.groupby("account_id"):
        name = group["account_name"].iloc[0]
        total = len(group)
        if count_critical_as_high:
            is_high = group["severity"].isin(["HIGH", "CRITICAL"])
        else:
            is_high = group["severity"].isin(["HIGH"])
        high_pct = round(is_high.mean() * 100.0, 2)

        info = owners.get(str(acct), {}) if isinstance(owners, dict) else {}
        owner = info.get("owner", "unknown")
        email = info.get("email", "")

        # Build XLSX in memory for this account
        bio = io.BytesIO()
        group.drop(columns=[], errors="ignore").to_excel(bio, index=False)
        bio.seek(0)

        rows.append(
            AccountRow(
                account_id=str(acct),
                account_name=str(name),
                total_findings=total,
                high_pct=high_pct,
                owner=owner,
                email=email,
                xlsx_bytes=bio.getvalue(),
                filename=f"{acct}.xlsx",
            )
        )
    return rows


# ---------- Streamlit App ----------

def main():
    st.set_page_config(page_title="AWS Inspector Dashboard", layout="wide")
    st.title("AWS Inspector Findings by Account (Streamlit only)")

    with st.sidebar:
        st.header("Owners (optional)")
        owners_file = st.file_uploader("owners.json", type=["json"], key="owners")
        owners = owners_from_uploaded(owners_file)
        st.caption("owners.json maps account_id → { owner, email }.")

        st.divider()
        st.header("Input options")
        has_header = st.checkbox("First row contains column names (header)", value=True, help="Uncheck if your file has no header row.")
        sev_mode = st.checkbox("Count CRITICAL as High", value=False)

    st.subheader("Upload data")
    up_file = st.file_uploader("Upload Inspector findings (.csv or .xlsx)", type=["csv", "xlsx"], key="data")

    if not up_file:
        st.info("Upload a file exported from Inspector (CSV or Excel).")
        st.stop()

    # Show file size info
    if hasattr(up_file, "size"):
        size_mb = up_file.size / (1024 * 1024)
        st.caption(f"Uploaded file size: ~{size_mb:.1f} MB (limit: 500 MB)")

    # Read file
    try:
        if up_file.name.lower().endswith(".xlsx"):
            df = pd.read_excel(up_file, engine="openpyxl", header=0 if has_header else None)
        else:
            df = pd.read_csv(up_file, header=0 if has_header else None, low_memory=False)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    # If no header, assign generic names
    if not has_header:
        df.columns = [f"col_{i}" for i in range(len(df.columns))]

    st.write("Preview (first 50 rows):")
    st.dataframe(df.head(50), use_container_width=True)

    # Column mapping
    st.subheader("Map columns")

    suggest = suggest_column_names(df)
    account_id_col = suggest.get("account_id")
    account_name_col = suggest.get("account_name")
    severity_col = suggest.get("severity")
    arn_col = suggest.get("finding_arn")

    c1, c2, c3 = st.columns(3)
    with c1:
        account_id_sel = st.selectbox(
            "Account ID column",
            options=["<None>"] + list(df.columns),
            index=(df.columns.tolist().index(account_id_col) + 1) if account_id_col in df.columns else 0,
            help="If not present, select FindingArn below and we'll extract the 12‑digit account ID.",
        )
    with c2:
        account_name_sel = st.selectbox(
            "Account Name column (optional)",
            options=["<None>"] + list(df.columns),
            index=(df.columns.tolist().index(account_name_col) + 1) if account_name_col in df.columns else 0,
        )
    with c3:
        severity_sel = st.selectbox(
            "Severity column",
            options=["<None>"] + list(df.columns),
            index=(df.columns.tolist().index(severity_col) + 1) if severity_col in df.columns else 0,
        )

    arn_sel = st.selectbox(
        "FindingArn / ARN column (optional, used to extract Account ID)",
        options=["<None>"] + list(df.columns),
        index=(df.columns.tolist().index(arn_col) + 1) if arn_col in df.columns else 0,
    )

    # Validate mapping and build a normalized working frame
    if account_id_sel == "<None>":
        if arn_sel == "<None>":
            st.error("Select an Account ID column, or select a FindingArn column so we can extract it.")
            st.stop()
        df["_account_id"] = df[arn_sel].apply(extract_account_from_arn)
        if df["_account_id"].isna().all():
            st.error("Could not extract any 12‑digit account IDs from the selected ARN column.")
            st.stop()
        account_id_sel = "_account_id"

    if severity_sel == "<None>":
        st.error("Please select a Severity column.")
        st.stop()

    work = pd.DataFrame()
    work["account_id"] = df[account_id_sel].astype(str)
    work["severity"] = df[severity_sel].astype(str).str.strip().str.upper()
    if account_name_sel != "<None>":
        work["account_name"] = df[account_name_sel].astype(str)
    else:
        work["account_name"] = work["account_id"]

    # Compute per-account rows
    try:
        rows = compute_rows(work, owners, count_critical_as_high=sev_mode)
    except Exception as e:
        st.error(str(e))
        st.stop()

    if not rows:
        st.warning("No accounts found after mapping.")
        st.stop()

    # Summary table
    st.subheader("Summary")
    display_df = pd.DataFrame(
        [
            {
                "Account ID": r.account_id,
                "Account Name": r.account_name,
                "Total Findings": r.total_findings,
                "High %": r.high_pct,
                "Owner": r.owner,
                "Email": r.email,
            }
            for r in rows
        ]
    )
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # Email compose defaults
    st.subheader("Per‑account actions")
    subject_prefix = st.text_input("Subject prefix", value="AWS Inspector Findings for Account")
    body_template = st.text_area(
        "Email body",
        value=(
            "Hello {owner},\n\n"
            "Please find attached the latest findings for AWS account "
            "{account_name} ({account_id}).\n\nRegards,\nSecurity Team"
        ),
        height=140,
    )
    st.caption("Note: Outlook on the web cannot pre‑attach files from a URL. Download the XLSX, then attach it in the draft.")

    # Per-account UI
    for r in rows:
        st.markdown("---")
        cA, cB, cC = st.columns([3, 2, 3])
        with cA:
            st.write(f"**{r.account_id}** — {r.account_name}")
            st.caption(f"Owner: {r.owner or 'unknown'} | Email: {r.email or '(enter below)'}")
        with cB:
            st.metric("Total Findings", r.total_findings)
            st.metric("High %", f"{r.high_pct}%")
        with cC:
            st.download_button(
                "⬇️ Download XLSX",
                data=r.xlsx_bytes,
                file_name=r.filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{r.account_id}",
            )

        cD, cE = st.columns([3, 3])
        with cD:
            to_email = r.email or st.text_input(f"Recipient for {r.account_id}", key=f"email_{r.account_id}")
        with cE:
            subject = f"{subject_prefix} {r.account_name}"
            body = body_template.format(
                owner=(r.owner or ""),
                account_name=r.account_name,
                account_id=r.account_id,
            )
            owa_link = outlook_web_compose_link(to_email, subject, body)
            st.link_button("✉️ Compose in Outlook (web)", owa_link, use_container_width=True)


if __name__ == "__main__":
    main()
