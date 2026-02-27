# app_streamlit.py
# Streamlit app that:
# - uploads Inspector CSV
# - computes per-account summary
# - generates per-account XLSX in memory
# - lets users download per account
# - signs in with Microsoft (MSAL, PKCE)
# - creates Outlook-on-the-web draft with the XLSX attached via Microsoft Graph

from __future__ import annotations
import base64
import io
import json
import time
from dataclasses import dataclass
from typing import Dict, List, Optional

import pandas as pd
import requests
import streamlit as st
from msal import PublicClientApplication

# ------------------ Config ------------------
SCOPES = [
    "Mail.ReadWrite",  # create draft + attach
    "offline_access",
    "openid", "email", "profile",
]

# The redirect URI MUST exactly match one of the redirect URIs you configure in Azure App Registration.
# Best is to set it as your deployed Streamlit app URL, e.g., "https://<app>.streamlit.app".
DEFAULT_REDIRECT_URI = ""  # optionally set in secrets

# ------------------ Data model ------------------
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

# ------------------ Helpers ------------------
def get_query_params():
    try:
        return st.experimental_get_query_params()
    except Exception:
        # fallback (older/newer API differences)
        return {}

def clear_query_params():
    try:
        st.experimental_set_query_params()
    except Exception:
        pass


def owners_from_uploaded(uploaded_file) -> Dict[str, dict]:
    if not uploaded_file:
        return {}
    try:
        return json.load(uploaded_file)
    except Exception:
        uploaded_file.seek(0)
        return json.loads(uploaded_file.read().decode("utf-8"))


def compute_rows(df: pd.DataFrame, owners: Dict[str, dict]) -> List[AccountRow]:
    missing = [c for c in ["account_id", "severity"] if c not in df.columns]
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

# ---------- Microsoft identity (MSAL) ----------
def authority():
    tenant_id = st.secrets.get('azure', {}).get('tenant_id', '')
    if not tenant_id:
        raise RuntimeError("Missing azure.tenant_id in Streamlit secrets")
    return f"https://login.microsoftonline.com/{tenant_id}"

def client_id():
    cid = st.secrets.get('azure', {}).get('client_id', '')
    if not cid:
        raise RuntimeError("Missing azure.client_id in Streamlit secrets")
    return cid

def redirect_uri():
    # Prefer explicit secret; fallback to current app URL if provided as secret
    return st.secrets.get('azure', {}).get('redirect_uri', DEFAULT_REDIRECT_URI)

def build_auth_url(state: str) -> str:
    app = PublicClientApplication(client_id(), authority=authority())
    return app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=redirect_uri() or st.secrets.get('app_url', ''),
        state=state,
        prompt="select_account",
    )

def acquire_token_by_code(code: str) -> dict:
    app = PublicClientApplication(client_id(), authority=authority())
    result = app.acquire_token_by_authorization_code(
        code=code,
        scopes=SCOPES,
        redirect_uri=redirect_uri() or st.secrets.get('app_url', ''),
    )
    if 'access_token' not in result:
        raise RuntimeError(f"Auth failed: {result}")
    return result

# ---------- Microsoft Graph (draft + attachment) ----------
def create_draft_and_get_weblink(
    access_token: str,
    to_email: str,
    subject: str,
    body_text: str,
    xlsx_bytes: bytes,
    filename: str,
) -> str:
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

    # 1) Create draft
    create_url = "https://graph.microsoft.com/v1.0/me/messages"
    payload = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body_text},
        "toRecipients": [{"emailAddress": {"address": to_email}}],
    }
    r = requests.post(create_url, headers=headers, json=payload, timeout=30)
    r.raise_for_status()
    message_id = r.json()["id"]

    # 2) Attach file (base64 bytes; suitable for small/medium files)
    attach_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments"
    attach_payload = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": filename,
        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "contentBytes": base64.b64encode(xlsx_bytes).decode("utf-8"),
    }
    r2 = requests.post(attach_url, headers=headers, json=attach_payload, timeout=30)
    r2.raise_for_status()

    # 3) Fetch webLink of the draft
    get_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}?$select=webLink"
    r3 = requests.get(get_url, headers=headers, timeout=30)
    r3.raise_for_status()
    return r3.json()["webLink"]

# ------------------ Streamlit UI ------------------
def main():
    st.set_page_config(page_title="AWS Inspector Dashboard", layout="wide")
    st.title("AWS Inspector Findings by Account — Outlook (web) drafts")

    # Handle auth redirect
    params = get_query_params()
    if 'code' in params:
        try:
            token = acquire_token_by_code(params['code'][0])
            st.session_state['graph_token'] = token['access_token']
            st.success("Signed in with Microsoft.")
            clear_query_params()  # clean URL
        except Exception as e:
            st.error(f"Sign-in error: {e}")

    with st.sidebar:
        st.header("Owners")
        owners_file = st.file_uploader("owners.json (optional)", type=["json"])
        owners = owners_from_uploaded(owners_file)

        st.divider()
        st.header("Microsoft 365 Sign-in")
        if 'graph_token' in st.session_state:
            st.success("Authenticated with Microsoft Graph.")
            if st.button("Sign out"):
                st.session_state.pop('graph_token', None)
                st.rerun()
        else:
            state = f"st_{int(time.time())}"
            try:
                auth_url = build_auth_url(state)
                st.link_button("🔐 Sign in with Microsoft", auth_url)
                st.caption("After signing in, you will be redirected back here.")
            except Exception as e:
                st.error(f"Auth configuration issue: {e}")
                st.write("Ensure azure.tenant_id, azure.client_id, and azure.redirect_uri are set in Streamlit secrets.")

    uploaded_csv = st.file_uploader("Upload Inspector CSV", type=["csv"])
    if not uploaded_csv:
        st.info("Required CSV columns: account_id, severity. Optional: account_name.")
        return

    # Read CSV
    try:
        df = pd.read_csv(uploaded_csv)
    except Exception as e:
        st.error(f"Failed to read CSV: {e}")
        return

    # Process
    try:
        rows = compute_rows(df, owners)
    except Exception as e:
        st.error(str(e))
        return

    if not rows:
        st.warning("No accounts found in CSV.")
        return

    # Summary
    st.subheader("Summary")
    st.dataframe(pd.DataFrame([
        {
            "Account ID": r.account_id,
            "Account Name": r.account_name,
            "Total Findings": r.total_findings,
            "High %": r.high_pct,
            "Owner": r.owner,
            "Email": r.email,
        } for r in rows
    ]), use_container_width=True, hide_index=True)

    st.subheader("Per-account actions")
    subject_prefix = st.text_input("Subject prefix", value="AWS Inspector Findings for Account")
    body_tmpl = st.text_area(
        "Email body template",
        value=(
            "Hello {owner},\n\n"
            "Please find attached the latest findings for AWS account "
            "{account_name} ({account_id}).\n\nRegards,\nSecurity Team"
        ),
        height=140,
    )

    has_token = 'graph_token' in st.session_state
    if not has_token:
        st.info("Sign in with Microsoft to enable 'Create draft in Outlook (web) with attachment'.")

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

        c4, c5 = st.columns([3, 3])
        with c4:
            to_email = r.email or st.text_input(f"Recipient for {r.account_id}", key=f"email_{r.account_id}")
        with c5:
            subject = f"{subject_prefix} {r.account_name}"
            body = body_tmpl.format(owner=r.owner or "", account_name=r.account_name, account_id=r.account_id)
            if st.button("✉️ Create draft in Outlook (web) with attachment", key=f"owa_{r.account_id}", disabled=not has_token):
                try:
                    link = create_draft_and_get_weblink(
                        st.session_state['graph_token'],
                        to_email, subject, body,
                        r.xlsx_bytes, r.filename,
                    )
                    st.success("Draft created. Click below to open in Outlook on the web.")
                    st.link_button("Open draft in Outlook", link)
                except Exception as e:
                    st.error(f"Failed to create draft: {e}")


if __name__ == "__main__":
    main()
