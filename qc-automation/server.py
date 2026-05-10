"""
QC Dashboard Server — Combined Web + Email API (v2)
====================================================
Single server on port 8503 that:
1. Serves the HTML dashboard (GET /)
2. Handles email sending via SES with selectable sections (POST /api/*)
3. Generates clean email from CSVs (not raw dashboard HTML)

Usage on EC2:
    export QC_ADMIN_PINS="prathima2026,teammate2026"
    export QC_SES_SENDER="finopsapqc@amazon.com"
    python3 server.py
"""

from http.server import HTTPServer, SimpleHTTPRequestHandler
import json
import os
import sys
import boto3
from pathlib import Path
from datetime import datetime

# Add current directory to path for imports
sys.path.insert(0, '/home/ec2-user/qc-dashboard')
from email_template import generate_email_html, AVAILABLE_SECTIONS

# ── Config ──────────────────────────────────────────────────────────────────
PORT = 8503
WEB_DIR = Path('/home/ec2-user/web')
DATA_DIR = Path('/home/ec2-user/qc-dashboard/data/current')
PREV_DIR = Path('/home/ec2-user/qc-dashboard/data/previous')
ADMIN_PINS = os.environ.get('QC_ADMIN_PINS', 'prathima2026,teammate2026').split(',')
SES_SENDER = os.environ.get('QC_SES_SENDER', 'finopsapqc@amazon.com')
SES_REGION = os.environ.get('QC_SES_REGION', 'us-east-1')
MONTH_LABEL = os.environ.get('QC_MONTH_LABEL', 'Apr 2026')


def get_ses_client():
    session = boto3.Session()
    return session.client('ses', region_name=SES_REGION)


class DashboardHandler(SimpleHTTPRequestHandler):
    """Serves static files + handles /api/* endpoints."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(WEB_DIR), **kwargs)

    def do_GET(self):
        if self.path == '/api/health':
            self._json_response({'status': 'ok', 'time': datetime.now().isoformat()})
        elif self.path == '/api/config':
            self._json_response({
                'sender': SES_SENDER,
                'month': MONTH_LABEL,
                'sections': AVAILABLE_SECTIONS,
            })
        elif self.path == '/api/sections':
            self._json_response({'sections': AVAILABLE_SECTIONS})
        else:
            super().do_GET()

    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length).decode('utf-8') if content_length else '{}'
        
        try:
            data = json.loads(body)
        except:
            data = {}

        if self.path == '/api/verify-admin':
            self._handle_verify(data)
        elif self.path == '/api/preview-email':
            self._handle_preview(data)
        elif self.path == '/api/send-email':
            self._handle_send(data)
        else:
            self._json_response({'error': 'Not found'}, 404)

    def _handle_verify(self, data):
        pin = data.get('pin', '').strip()
        if pin in ADMIN_PINS:
            self._json_response({'authorized': True})
        else:
            self._json_response({'authorized': False, 'error': 'Invalid PIN'}, 401)

    def _handle_preview(self, data):
        """Generate email preview with selected sections."""
        pin = data.get('pin', '').strip()
        if pin not in ADMIN_PINS:
            self._json_response({'error': 'Unauthorized'}, 401)
            return
        
        sections = data.get('sections', [s['id'] for s in AVAILABLE_SECTIONS])
        
        try:
            html = generate_email_html(
                data_dir=str(DATA_DIR),
                prev_dir=str(PREV_DIR),
                sections=sections,
                month_label=MONTH_LABEL
            )
            self._json_response({
                'preview_html': html,
                'sections_included': sections,
                'size_kb': round(len(html) / 1024, 1)
            })
        except Exception as e:
            self._json_response({'error': f'Template error: {str(e)}'}, 500)

    def _handle_send(self, data):
        pin = data.get('pin', '').strip()
        if pin not in ADMIN_PINS:
            self._json_response({'error': 'Unauthorized'}, 401)
            return

        recipients = data.get('recipients', [])
        cc = data.get('cc', [])
        subject = data.get('subject', f'QC Snapshot — General {MONTH_LABEL}')
        sections = data.get('sections', [s['id'] for s in AVAILABLE_SECTIONS])

        if not recipients or all(not r.strip() for r in recipients):
            self._json_response({'error': 'No recipients specified'}, 400)
            return

        # Generate clean email HTML from CSVs
        try:
            html = generate_email_html(
                data_dir=str(DATA_DIR),
                prev_dir=str(PREV_DIR),
                sections=sections,
                month_label=MONTH_LABEL
            )
        except Exception as e:
            self._json_response({'error': f'Template error: {str(e)}'}, 500)
            return

        destination = {'ToAddresses': [r.strip() for r in recipients if r.strip()]}
        if cc:
            destination['CcAddresses'] = [c.strip() for c in cc if c.strip()]

        try:
            ses = get_ses_client()
            response = ses.send_email(
                Source=SES_SENDER,
                Destination=destination,
                Message={
                    'Subject': {'Charset': 'UTF-8', 'Data': subject},
                    'Body': {'Html': {'Charset': 'UTF-8', 'Data': html}}
                },
                Tags=[
                    {'Name': 'Application', 'Value': 'QC-Dashboard'},
                    {'Name': 'Sections', 'Value': ','.join(sections)[:256]},
                ]
            )
            self._json_response({
                'success': True,
                'message_id': response['MessageId'],
                'sent_to': [r.strip() for r in recipients if r.strip()],
                'sections': sections,
                'sent_at': datetime.now().isoformat()
            })
        except Exception as e:
            self._json_response({'error': str(e)}, 500)

    def _json_response(self, data, status=200):
        response = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', len(response))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        self.wfile.write(response)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, format, *args):
        pass


if __name__ == '__main__':
    os.chdir(str(WEB_DIR))
    server = HTTPServer(('0.0.0.0', PORT), DashboardHandler)
    print(f"QC Dashboard Server v2 running on port {PORT}")
    print(f"  Dashboard: http://0.0.0.0:{PORT}/")
    print(f"  Email API: http://0.0.0.0:{PORT}/api/send-email")
    print(f"  Sender: {SES_SENDER}")
    print(f"  Month: {MONTH_LABEL}")
    print(f"  Admin PINs: {len(ADMIN_PINS)} configured")
    print(f"  Data: {DATA_DIR}")
    server.serve_forever()
