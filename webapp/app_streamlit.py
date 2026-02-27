# app_streamlit.py
# Streamlit-native AWS Inspector dashboard with per-account XLSX and SES emailing (attachments)

from __future__ import annotations
import io
import os
import json
import base64
import mimetypes
from typing import Dict, List, Optional
from dataclasses import dataclass

import pandas as pd
import streamlit as st

# --- Choose email backend: "ses" or "smtp"
EMAIL_BACKEND = "ses"

# --- If using AWS SES ---
# Configure via Streamlit secrets (Settings -> Secrets):
# [aws]
# region = "ap-south-1"
# access_key = "AKIA..."
# secret_key = "..."
import boto3
from botocore.exceptions import BotoCoreError, ClientError

# --- If using SMTP (optional alternative) ---
# Configure via secrets like:
# [smtp]
# host = "email-smtp.ap-south-1.amazonaws.com"
# port = 587
# user = "SMTP_USER"
# password = "SMTP_PASS"
import smtplib
from email.message import EmailMessage


@dataclass
class AccountRow:
    account_id: str
    account_name: str
    total_findings: int
    high_pct: float
    owner: str
    email: str
    xlsx_bytes: bytes  # attachment content
    filename: str


def load_owners(uploaded_file) -> Dict[str, dict]:
    """Load owners.json if provided; else return {}"""
    if not uploaded_file:
        return {}
    try:
        # First try strict JSON
        return json.load(uploaded_file)
    except Exception:
        # If user re-uploads, Streamlit may pass a BytesIO-like; handle bytes
        uploaded_file.seek(0)
        return json.loads(uploaded_file.read().decode("utf-8"))


def compute_rows(df: pd.DataFrame, owners: Dict[str, dict]) -> List[AccountRow]:
    req = ["account_id", "severity"]
    missing = [c for c in req if c not in df.columns]
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

        # Build per-account XLSX in memory
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


# ---------------------------
# Email helpers
# ---------------------------
def send_email_ses(
    region: str,
    access_key: str,
    secret_key: str,
    sender: str,
    recipient: str,
    subject: str,
    body_text: str,
    attachment_bytes: bytes,
    attachment_filename: str,
):
    """Send email with attachment via AWS SES (Raw email)."""
    client = boto3.client(
        "ses",
        region_name=region,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
    )

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content(body_text)

    # Guess MIME
    mime_type, _ = mimetypes.guess_type(attachment_filename)
    maintype, subtype = (mime_type.split("/", 1) if mime_type else ("application", "octet-stream"))

    msg.add_attachment(
        attachment_bytes,
        maintype=maintype,
        subtype=subtype,
        filename=attachment_filename,
    )

    # SES requires raw email sending for attachments
    try:
        response = client.send_raw_email(
            Source=sender,
            Destinations=[recipient],
            RawMessage={"Data": msg.as_bytes()},
        )
        return response.get("MessageId")
    except (BotoCoreError, ClientError) as e:
        raise RuntimeError(f"SES send_raw_email failed: {e}")


def send_email_smtp(
    host: str, port: int, user: str, password: str,
    sender: str, recipient: str, subject: str, body_text: str,
    attachment_bytes: bytes, attachment_filename: str,
    use_tls: bool = True
):
    """Send email with attachment via SMTP."""
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content(body_text)

    mime_type, _ = mimetypes.guess_type(attachment_filename)
    maintype, subtype = (mime_type.split("/", 1) if mime_type else ("application", "octet-stream"))

    msg.add_attachment(
        attachment_bytes,
        maintype=maintype,
        subtype=subtype,
        filename=attachment_filename,
    )

    with smtplib.SMTP(host, port, timeout=30) as server:
        if use_tls:
            server.starttls()
        server.login(user, password)
        server.send_message(msg)


# ---------------------------
# Streamlit UI
# ---------------------------
def main():
    st.set_page_config(page_title="AWS Inspector Dashboard", layout="wide")
    st.title("AWS Inspector Findings by Account (Streamlit/Cloud)")

    with st.sidebar:
        st.header("Owners & Email")
        owners_file = st.file_uploader("owners.json (optional)", type=["json"])
        owners = load_owners(owners_file)

        st.markdown("**Email Settings**")

        backend = st.selectbox("Email backend", options=["ses", "smtp"], index=(0 if EMAIL_BACKEND == "ses" else 1))

        if backend == "ses":
            st.caption("Using AWS SES (store credentials in Streamlit Secrets)")
            sender = st.text_input("Sender email (verified in SES)", value="")
            # Use Secrets for region/creds
            region = st.text_input("AWS Region", value=st.secrets.get("aws", {}).get("region", "ap-south-1"))
        else:
            st.caption("Using SMTP (store credentials in Streamlit Secrets)")
            sender = st.text_input("Sender email", value="")
            smtp_host = st.text_input("SMTP Host", value=st.secrets.get("smtp", {}).get("host", ""))
            smtp_port = st.number_input("SMTP Port", value=int(st.secrets.get("smtp", {}).get("port", 587)))
            smtp_tls = st.checkbox("Use STARTTLS", value=True)

        st.divider()
        st.header("Compose Defaults")
        subject_prefix = st.text_input("Subject prefix", value="AWS Inspector Findings for Account")
        body_template = st.text_area(
            "Email body template",
            value=(
                "Hello {owner},\n\n"
                "Please find attached the latest findings for AWS account "
                "{account_name} ({account_id}).\n\nRegards,\nSecurity Team"
            ),
            height=140,
        )

    uploaded_csv = st.file_uploader("Upload Inspector CSV", type=["csv"])
    if not uploaded_csv:
        st.info("Expected columns: 'account_id', 'severity' (required), 'account_name' (optional).")
        return

    # Read CSV
    try:
        df = pd.read_csv(uploaded_csv)
    except Exception as e:
        st.error(f"Failed to read CSV: {e}")
        return

    # Compute rows
    try:
        rows = compute_rows(df, owners)
    except Exception as e:
        st.error(str(e))
        return

    if not rows:
        st.warning("No accounts found in CSV.")
        return

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

    st.subheader("Per-Account Actions")
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
                "⬇️ Download XLSX",
                data=r.xlsx_bytes,
                file_name=r.filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{r.account_id}",
            )

        with c4:
            to_email = r.email or st.text_input(f"Recipient for {r.account_id}", key=f"email_{r.account_id}")
            subject = f"{subject_prefix} {r.account_name}"
            body = body_template.format(
                owner=(r.owner or ""),
                account_name=r.account_name,
                account_id=r.account_id,
            )

            send_key = f"send_{r.account_id}"
            if st.button("📧 Send Email with Attachment", key=send_key):
                if not sender:
                    st.error("Sender email is required.")
                elif not to_email:
                    st.error("Recipient email is required.")
                else:
                    try:
                        if backend == "ses":
                            aws = st.secrets.get("aws", {})
                            message_id = send_email_ses(
                                region=aws.get("region", "ap-south-1"),
                                access_key=aws["access_key"],
                                secret_key=aws["secret_key"],
                                sender=sender,
                                recipient=to_email,
                                subject=subject,
                                body_text=body,
                                attachment_bytes=r.xlsx_bytes,
                                attachment_filename=r.filename,
                            )
                            st.success(f"Email sent via SES (MessageId: {message_id}).")
                        else:
                            smtp = st.secrets.get("smtp", {})
                            send_email_smtp(
                                host=smtp_host or smtp.get("host", ""),
                                port=int(smtp_port or smtp.get("port", 587)),
                                user=smtp.get("user", ""),
                                password=smtp.get("password", ""),
                                sender=sender,
                                recipient=to_email,
                                subject=subject,
                                body_text=body,
                                attachment_bytes=r.xlsx_bytes,
                                attachment_filename=r.filename,
                                use_tls=bool(smtp_tls),
                            )
                            st.success("Email sent via SMTP.")
                    except KeyError as ke:
                        st.error(f"Missing secret for {backend.upper()}: {ke}")
                    except Exception as e:
                        st.error(f"Failed to send email: {e}")


if __name__ == "__main__":
    main()
