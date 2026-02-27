import os
from flask import Flask, request, redirect, url_for, render_template, send_file
import pandas as pd
import json
import tempfile

# Windows Outlook integration
try:
    import win32com.client
except ImportError:
    win32com = None

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ACCOUNT_OWNER_FILE'] = 'owners.json'  # simple lookup file

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# helper to load account owner info from JSON

def load_owners():
    if os.path.exists(app.config['ACCOUNT_OWNER_FILE']):
        with open(app.config['ACCOUNT_OWNER_FILE'], 'r') as f:
            return json.load(f)
    return {}


@app.route('/', methods=['GET'])
def index():
    # look for preprocessed data file
    data_file = os.path.join(app.config['UPLOAD_FOLDER'], 'summary.json')
    if os.path.exists(data_file):
        with open(data_file) as f:
            data = json.load(f)
    else:
        data = None
    return render_template('dashboard.html', accounts=data, owners=load_owners())


@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file:
        return redirect(url_for('index'))

    path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(path)
    process_csv(path)
    return redirect(url_for('index'))


def process_csv(path):
    # assume columns have account_id, account_name, severity etc.
    df = pd.read_csv(path)
    # group and compute statistics
    summary = []
    for acct, group in df.groupby('account_id'):
        name = group['account_name'].iloc[0]
        total = len(group)
        high_pct = (group['severity'] == 'High').mean() * 100
        owner = load_owners().get(str(acct), {}).get('owner')
        # save account-specific excel
        outpath = os.path.join(app.config['UPLOAD_FOLDER'], f"{acct}.xlsx")
        group.to_excel(outpath, index=False)
        summary.append({
            'account_id': acct,
            'account_name': name,
            'total_findings': total,
            'high_pct': round(high_pct, 2),
            'owner': owner,
            'file': outpath,
        })
    with open(os.path.join(app.config['UPLOAD_FOLDER'], 'summary.json'), 'w') as f:
        json.dump(summary, f)


@app.route('/send/<acct_id>')
def send(acct_id):
    owners = load_owners()
    acct_info = None
    data_file = os.path.join(app.config['UPLOAD_FOLDER'], 'summary.json')
    if os.path.exists(data_file):
        with open(data_file) as f:
            for a in json.load(f):
                if str(a['account_id']) == str(acct_id):
                    acct_info = a
                    break
    if not acct_info:
        return "Account not found", 404

    recipient = owners.get(str(acct_id), {}).get('email', '')
    subject = f"AWS Inspector Findings for Account {acct_info['account_name']}"
    body = f"Hello {owners.get(str(acct_id), {}).get('owner','')},\n\n" \
           f"Please find attached the latest findings for AWS account {acct_info['account_name']} ({acct_id}).\n\n" \
           "Regards,\nSecurity Team"

    filepath = acct_info['file']

    # using Outlook automation if available
    if win32com:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(os.path.abspath(filepath))
        mail.Display()  # open in Outlook for manual send
        return redirect(url_for('index'))

    # fallback: mailto link (attachments not supported)
    mailto = f"mailto:{recipient}?subject={subject}&body={body}"
    return redirect(mailto)


if __name__ == '__main__':
    app.run(debug=True)