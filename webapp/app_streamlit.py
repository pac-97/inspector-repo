# -*- coding: utf-8 -*-
"""
Streamlit version of the AWS Inspector dashboard.
- Upload CSV
- Compute per-account summary
- Generate per-account XLSX in memory
- Show downloads per account
- (Windows only) Open Outlook compose window using an optional .oft template and attach that account's XLSX

Expected CSV columns (minimum):
  - account_id
  - severity
Optional columns:
  - account_name

Optional owners.json format:
{
  "123456789012": {"owner": "Team A", "email": "teama@example.com"},
  "210987654321": {"owner": "Team B", "email": "teamb@example.com"}
}
"""

from __future__ import annotations
import io
import os
import platform
import tempfile
from dataclasses import dataclass
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

# Try Windows Outlook (pywin32) import lazily
IS_WINDOWS = platform.system() == "Windows"
try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover
    win32com = None  # type: ignore

UPLOAD_ROOT = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_ROOT, exist_ok=True)

@dataclass
class AccountSummary:
    account_id: str
    account_name: str
    total_findings: int
    high_pct: float
    owner: str
    email: str
    xlsx_bytes: bytes
    filename: str


def load_owners_from_json_bytes(raw: Optional[bytes]) -> Dict[str, dict]:
    if not raw:
        return {}
    try:
        return pd.read_json(io.BytesIO(raw), typ='dict').to_dict()  # robust load
    except Exception:
        # Fallback to std json
        import json
        return json.loads(raw.decode('utf-8'))


def process_df(df: pd.DataFrame, owners: Dict[str, dict]) -> List[AccountSummary]:
    if 'account_id' not in df.columns:
        raise ValueError("CSV must include 'account_id' column")
    if 'severity' not in df.columns:
        raise ValueError("CSV must include 'severity' column")

    rows: List[AccountSummary] = []
    for acct, group in df.groupby('account_id'):
        name = (
            group['account_name'].iloc[0]
            if 'account_name' in group.columns and pd.notna(group['account_name'].iloc[0])
            else str(acct)
        )
        total = len(group)
        high_pct = float((group['severity'] == 'High').mean() * 100.0)

        # owners lookup
        oinfo = owners.get(str(acct), {}) if isinstance(owners, dict) else {}
        owner = oinfo.get('owner', 'unknown') if isinstance(oinfo, dict) else 'unknown'
        email = oinfo.get('email', '') if isinstance(oinfo, dict) else ''

        # Prepare XLSX in memory
        bio = io.BytesIO()
        group.to_excel(bio, index=False)
        bio.seek(0)

        rows.append(
            AccountSummary(
                account_id=str(acct),
                account_name=str(name),
                total_findings=int(total),
                high_pct=round(high_pct, 2),
                owner=owner,
                email=email,
                xlsx_bytes=bio.getvalue(),
                filename=f"{acct}.xlsx",
            )
        )
    return rows


def compose_outlook_email(
    to_email: str,
    subject: str,
    body_text: str,
    xlsx_bytes: bytes,
    filename: str,
    oft_template_path: Optional[str] = None,
):
    """Open Outlook compose window with attachment. Windows + Outlook only.
    - Writes the XLSX to a NamedTemporaryFile and attaches it.
    - If 'oft_template_path' is provided and exists, uses it; else creates a blank item.
    - Displays the compose window for manual review/send.
    """
    if not (IS_WINDOWS and win32com):
        raise RuntimeError("Outlook automation is only available on Windows with pywin32 installed.")

    outlook = win32com.client.Dispatch('Outlook.Application')
    if oft_template_path and os.path.exists(oft_template_path):
        mail = outlook.CreateItemFromTemplate(oft_template_path)
    else:
        mail = outlook.CreateItem(0)

    # Keep existing template body if present; then prepend our message
    try:
        # Prefer HTMLBody if the template is HTML; fall back to Body
        existing_html = getattr(mail, 'HTMLBody', None)
        if existing_html:
            # Insert our text at the top
            mail.HTMLBody = f"<p>{body_text.replace('\n', '<br>')}</p>" + existing_html
        else:
            mail.Body = f"{body_text}\n\n" + getattr(mail, 'Body', '')
    except Exception:
        mail.Body = body_text

    mail.To = to_email
    mail.Subject = subject

    # Write attachment to a temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(xlsx_bytes)
        tmp.flush()
        attach_path = tmp.name

    mail.Attachments.Add(attach_path)
    mail.Display()  # open compose window


def main():
    st.set_page_config(page_title="AWS Inspector Dashboard", layout="wide")
    st.title("AWS Inspector Findings by Account")

    with st.sidebar:
        st.header("Options")
        st.write("Upload optional owners.json to map account → owner/email.")
        owners_file = st.file_uploader("owners.json (optional)", type=["json"], key="owners")
        owners = load_owners_from_json_bytes(owners_file.read() if owners_file else None)

        st.divider()
        st.write("Outlook (.oft) template (Windows only, optional). If not provided, a blank email is used.")
        oft_file = st.file_uploader("Outlook template .oft (optional)", type=["oft"], key="oft")
        oft_path: Optional[str] = None
        if oft_file is not None:
            # Persist to a temp file so Outlook can read it from disk
            with tempfile.NamedTemporaryFile(delete=False, suffix=".oft", dir=UPLOAD_ROOT) as of:
                of.write(oft_file.read())
                of.flush()
                oft_path = of.name
        st.caption("Outlook compose works only when you run this app locally on Windows with Outlook installed.")

    uploaded = st.file_uploader("Select Inspector CSV", type=["csv"], key="csv")
    if not uploaded:
        st.info("Upload a CSV to begin. Required columns: 'account_id', 'severity'. Optional: 'account_name'.")
        return

    # Read CSV
    try:
        df = pd.read_csv(uploaded)
    except Exception as e:
        st.error(f"Failed to read CSV: {e}")
        return

    # Validate & process
    try:
        rows = process_df(df, owners)
    except Exception as e:
        st.error(str(e))
        return

    if not rows:
        st.warning("No accounts found in the CSV.")
        return

    # Summary grid
    st.subheader("Summary")
    display_df = pd.DataFrame([
        {
            "Account ID": r.account_id,
            "Account Name": r.account_name,
            "Total Findings": r.total_findings,
            "High %": r.high_pct,
            "Owner": r.owner,
            "Email": r.email,
        }
        for r in rows
    ])
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    st.subheader("Per-Account Actions")
    default_subject_prefix = st.text_input("Email subject prefix", value="AWS Inspector Findings for Account")
    default_greeting = st.text_area(
        "Email greeting/body prefix",
        value=(
            "Hello {owner},\n\n"
            "Please find attached the latest findings for AWS account "
            "{account_name} ({account_id}).\n\nRegards,\nSecurity Team"
        ),
    )

    for r in rows:
        st.markdown("---")
        c1, c2, c3, c4 = st.columns([2, 3, 2, 3])
        with c1:
            st.write(f"**{r.account_id}**")
            st.caption(r.account_name)
        with c2:
            st.metric("Total Findings", r.total_findings)
            st.metric("High %", f"{r.high_pct}%")
        with c3:
            st.download_button(
                label="⬇️ Download XLSX",
                data=r.xlsx_bytes,
                file_name=r.filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{r.account_id}",
            )
        with c4:
            # Compose in Outlook (Windows only)
            btn_label = "✉️ Compose in Outlook (Windows only)"
            if IS_WINDOWS and win32com:
                if st.button(btn_label, key=f"ol_{r.account_id}"):
                    to_email = r.email or ""
                    subject = f"{default_subject_prefix} {r.account_name}"
                    body_text = default_greeting.format(
                        owner=r.owner or "",
                        account_name=r.account_name,
                        account_id=r.account_id,
                    )
                    try:
                        compose_outlook_email(
                            to_email=to_email,
                            subject=subject,
                            body_text=body_text,
                            xlsx_bytes=r.xlsx_bytes,
                            filename=r.filename,
                            oft_template_path=oft_path,
                        )
                        st.success("Opened Outlook compose window.")
                    except Exception as e:
                        st.error(f"Failed to open Outlook: {e}")
            else:
                st.button(btn_label, key=f"ol_{r.account_id}", disabled=True)
                st.caption("Run locally on Windows with Outlook for this button to work.")


if __name__ == '__main__':
    main()
