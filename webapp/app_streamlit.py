# -*- coding: utf-8 -*-
"""
Streamlit app (no Azure, no SES):
- Upload Inspector CSV
- Optional owners.json (account_id -> {owner, email})
- Per-account stats + XLSX download
- "Compose in Outlook (web)" link per account (To/Subject/Body prefilled)
  NOTE: Attachments cannot be pre-added via URL; user must attach XLSX manually.
"""

from __future__ import annotations
import io
import json
import urllib.parse as up
from dataclasses import dataclass
from typing import Dict, List

import pandas as pd
import streamlit as st


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


def owners_from_uploaded(uploaded_file) -> Dict[str, dict]:
    """Load owners.json if provided; else {}."""
    if not uploaded_file:
        return {}
    try:
        return json.load(uploaded_file)
    except Exception:
        uploaded_file.seek(0)
        return json.loads(uploaded_file.read().decode("utf-8"))


def compute_rows(df: pd.DataFrame, owners: Dict[str, dict]) -> List[AccountRow]:
    required = ["account_id", "severity"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"CSV missing required columns: {missing}")

    rows: List[AccountRow] = []
    for acct, group in df.groupby("account_id"):
        name = group["account_name"].iloc[0] if "account_name" in group.columns else str(acct)
        total = len(group)
        high_pct = round(((group["severity"] == "High").mean() * 100.0), 2)

        info = owners.get(str(acct), {}) if isinstance(owners, dict) else {}
        owner = info.get("owner", "unknown")
        email = info.get("email", "")

        # Build XLSX in memory
        bio = io.BytesIO()
        group.to_excel(bio, index=False)
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


def outlook_web_compose_link(to: str, subject: str, body: str) -> str:
    """
    Generate Outlook on the web (OWA) compose deeplink with To/Subject/Body.
    NOTE: Attachments cannot be added via URL; user must attach manually.
    """
    base = "https://outlook.office.com/mail/deeplink/compose"
    # URL-encode values
    params = {
        "to": to or "",
        "subject": subject or "",
        "body": body or "",
    }
    # Use quote to preserve newlines as %0D%0A
    q = "&".join(f"{k}={up.quote(v)}" for k, v in params.items())
    return f"{base}?{q}"


def main():
    st.set_page_config(page_title="AWS Inspector Dashboard", layout="wide")
    st.title("AWS Inspector Findings by Account (Streamlit only)")

    with st.sidebar:
        st.header("Owners (optional)")
        owners_file = st.file_uploader("owners.json", type=["json"])
        owners = owners_from_uploaded(owners_file)

        st.caption("owners.json maps account_id → { owner, email }.")

    uploaded_csv = st.file_uploader("Upload Inspector CSV", type=["csv"])
    if not uploaded_csv:
        st.info("CSV must include: account_id, severity. Optional: account_name.")
        return

    try:
        df = pd.read_csv(uploaded_csv)
    except Exception as e:
        st.error(f"Failed to read CSV: {e}")
        return

    try:
        rows = compute_rows(df, owners)
    except Exception as e:
        st.error(str(e))
        return

    if not rows:
        st.warning("No accounts found.")
        return

    # Summary
    st.subheader("Summary")
    st.dataframe(
        pd.DataFrame(
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
        ),
        use_container_width=True,
        hide_index=True,
    )

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

    for r in rows:
        st.markdown("---")
        c1, c2, c3 = st.columns([3, 2, 3])
        with c1:
            st.write(f"**{r.account_id}** — {r.account_name}")
            st.caption(f"Owner: {r.owner or 'unknown'} | Email: {r.email or '(enter below)'}")
        with c2:
            st.metric("Total Findings", r.total_findings)
            st.metric("High %", f"{r.high_pct}%")
        with c3:
            st.download_button(
                "⬇️ Download XLSX",
                data=r.xlsx_bytes,
                file_name=r.filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{r.account_id}",
            )

        # Compose row: recipient + link
        c4, c5 = st.columns([3, 3])
        with c4:
            to_email = r.email or st.text_input(f"Recipient for {r.account_id}", key=f"email_{r.account_id}")
        with c5:
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
