"""
Email Server for QC Dashboard
==============================
A tiny Flask API that runs alongside the HTML dashboard on EC2.
Handles the "Send Email" feature — sends the dashboard HTML as email via AWS SES.

Features:
- Admin-only access (PIN-protected)
- Human-in-the-loop (preview before send)
- Sends full HTML dashboard as email body
- Configurable recipients

Runs on port 8504 on EC2 (CloudFront routes /api/* to this)

Usage:
    python3 email_server.py
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import boto3
import os
import json
from pathlib import Path
from datetime import datetime

app = Flask(__name__)
CORS(app)

# ── Config ──────────────────────────────────────────────────────────────────
ADMIN_PINS = os.environ.get('QC_ADMIN_PINS', 'prathima2026,teammate2026').split(',')
SES_SENDER = os.environ.get('QC_SES_SENDER', 'callaudit-noreply@amazon.com')
SES_REGION = os.environ.get('QC_SES_REGION', 'us-east-1')
WEB_DIR = Path('/home/ec2-user/web')
DEFAULT_RECIPIENTS = os.environ.get('QC_EMAIL_RECIPIENTS', '').split(',')

# Allowed admin users (logins)
ADMIN_USERS = os.environ.get('QC_ADMIN_USERS', 'pratpk,teammate').split(',')


def get_ses_client():
    """Get SES client using EC2 instance profile."""
    session = boto3.Session()
    return session.client('ses', region_name=SES_REGION)


def get_dashboard_html():
    """Read the current dashboard HTML file."""
    html_path = WEB_DIR / 'index.html'
    if html_path.exists():
        return html_path.read_text(encoding='utf-8')
    return None


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'time': datetime.now().isoformat()})


@app.route('/api/verify-admin', methods=['POST'])
def verify_admin():
    """Verify admin PIN before showing send options."""
    data = request.json or {}
    pin = data.get('pin', '').strip()
    
    if pin in ADMIN_PINS:
        return jsonify({'authorized': True})
    return jsonify({'authorized': False, 'error': 'Invalid PIN'}), 401


@app.route('/api/preview-email', methods=['POST'])
def preview_email():
    """Generate email preview — human-in-the-loop step."""
    data = request.json or {}
    pin = data.get('pin', '').strip()
    
    if pin not in ADMIN_PINS:
        return jsonify({'error': 'Unauthorized'}), 401
    
    html = get_dashboard_html()
    if not html:
        return jsonify({'error': 'No dashboard HTML found'}), 404
    
    # Get recipients
    recipients = data.get('recipients', DEFAULT_RECIPIENTS)
    cc = data.get('cc', [])
    subject = data.get('subject', f'QC Dashboard — General {datetime.now().strftime("%b %Y")}')
    
    # Return preview info (don't send yet — human-in-the-loop)
    return jsonify({
        'preview': {
            'subject': subject,
            'from': SES_SENDER,
            'to': recipients,
            'cc': cc,
            'html_size': len(html),
            'html_snippet': html[:500] + '...',
        },
        'ready_to_send': True
    })


@app.route('/api/send-email', methods=['POST'])
def send_email():
    """Actually send the email after human confirmation."""
    data = request.json or {}
    pin = data.get('pin', '').strip()
    
    if pin not in ADMIN_PINS:
        return jsonify({'error': 'Unauthorized'}), 401
    
    html = get_dashboard_html()
    if not html:
        return jsonify({'error': 'No dashboard HTML found'}), 404
    
    recipients = data.get('recipients', DEFAULT_RECIPIENTS)
    cc = data.get('cc', [])
    subject = data.get('subject', f'QC Dashboard — General {datetime.now().strftime("%b %Y")}')
    
    if not recipients or all(not r.strip() for r in recipients):
        return jsonify({'error': 'No recipients specified'}), 400
    
    # Build SES request
    destination = {
        'ToAddresses': [r.strip() for r in recipients if r.strip()]
    }
    if cc:
        destination['CcAddresses'] = [c.strip() for c in cc if c.strip()]
    
    try:
        ses = get_ses_client()
        response = ses.send_email(
            Source=SES_SENDER,
            Destination=destination,
            Message={
                'Subject': {'Charset': 'UTF-8', 'Data': subject},
                'Body': {
                    'Html': {'Charset': 'UTF-8', 'Data': html}
                }
            },
            Tags=[
                {'Name': 'Application', 'Value': 'QC-Dashboard'},
                {'Name': 'SentBy', 'Value': 'admin'},
            ]
        )
        return jsonify({
            'success': True,
            'message_id': response['MessageId'],
            'sent_to': recipients,
            'sent_at': datetime.now().isoformat()
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/config', methods=['GET'])
def get_config():
    """Return public config for the frontend (no secrets)."""
    return jsonify({
        'default_recipients': DEFAULT_RECIPIENTS,
        'sender': SES_SENDER,
        'admin_users': ADMIN_USERS,
    })


if __name__ == '__main__':
    print("QC Dashboard Email Server starting on port 8504...")
    app.run(host='0.0.0.0', port=8504, debug=False)
