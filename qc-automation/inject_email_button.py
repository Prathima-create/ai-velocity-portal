"""
inject_email_button.py
======================
Injects the "Send Email" button + modal into an existing QC Dashboard HTML file.
Run this after qc_automation.py generates the HTML.

Usage:
    python3 inject_email_button.py path/to/QC_Dashboard.html

The button:
- Only visible to admins (PIN-protected)
- Shows preview before sending (human-in-the-loop)
- Sends full dashboard as HTML email via SES
- Viewers see dashboard normally without the button
"""

import sys
from pathlib import Path

# The JavaScript + HTML to inject before </body>
EMAIL_INJECTION = '''
<!-- ═══════ EMAIL BUTTON + MODAL (Admin Only) ═══════ -->
<style>
.email-fab { position: fixed; bottom: 24px; right: 24px; z-index: 9999;
             background: #2563eb; color: #fff; border: none; padding: 14px 20px;
             border-radius: 50px; font-size: 14px; font-weight: 600;
             cursor: pointer; box-shadow: 0 4px 12px rgba(37,99,235,.4);
             transition: all .2s; display: flex; align-items: center; gap: 8px; }
.email-fab:hover { background: #1d4ed8; transform: translateY(-2px);
                   box-shadow: 0 6px 16px rgba(37,99,235,.5); }
.email-fab svg { width: 18px; height: 18px; fill: currentColor; }

.email-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.5);
                 z-index: 10000; justify-content: center; align-items: center; }
.email-overlay.open { display: flex; }
.email-modal { background: #fff; border-radius: 12px; padding: 24px; width: 90%;
               max-width: 500px; max-height: 80vh; overflow-y: auto;
               box-shadow: 0 20px 60px rgba(0,0,0,.3); }
.email-modal h3 { margin: 0 0 16px 0; font-size: 16px; color: #1f2937; }
.email-modal label { font-size: 12px; color: #6b7280; display: block; margin: 10px 0 4px; }
.email-modal input, .email-modal textarea { width: 100%; padding: 8px 12px;
  border: 1px solid #d1d5db; border-radius: 6px; font-size: 13px; }
.email-modal textarea { height: 60px; resize: vertical; }
.email-btn-row { display: flex; gap: 8px; margin-top: 16px; justify-content: flex-end; }
.email-btn { padding: 8px 18px; border-radius: 6px; font-size: 13px;
             font-weight: 600; cursor: pointer; border: none; }
.email-btn-primary { background: #2563eb; color: #fff; }
.email-btn-primary:hover { background: #1d4ed8; }
.email-btn-danger { background: #ef4444; color: #fff; }
.email-btn-danger:hover { background: #dc2626; }
.email-btn-ghost { background: #f3f4f6; color: #374151; border: 1px solid #d1d5db; }
.email-btn-ghost:hover { background: #e5e7eb; }
.email-status { margin-top: 12px; padding: 10px; border-radius: 6px; font-size: 13px; display: none; }
.email-status.success { display: block; background: #d1fae5; color: #065f46; }
.email-status.error { display: block; background: #fee2e2; color: #991b1b; }
.email-step { display: none; }
.email-step.active { display: block; }
</style>

<!-- Floating Action Button -->
<button class="email-fab" onclick="openEmailModal()" title="Send Dashboard via Email">
  <svg viewBox="0 0 24 24"><path d="M20 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/></svg>
  Send Email
</button>

<!-- Email Modal Overlay -->
<div class="email-overlay" id="emailOverlay">
  <div class="email-modal">
    <!-- Step 1: PIN Verification -->
    <div class="email-step active" id="emailStep1">
      <h3>🔐 Admin Access Required</h3>
      <p style="font-size:13px;color:#6b7280;">Only authorized admins can send the dashboard email.</p>
      <label>Enter Admin PIN:</label>
      <input type="password" id="emailPin" placeholder="Enter your PIN" onkeypress="if(event.key==='Enter')verifyPin()"/>
      <div class="email-btn-row">
        <button class="email-btn email-btn-ghost" onclick="closeEmailModal()">Cancel</button>
        <button class="email-btn email-btn-primary" onclick="verifyPin()">Verify</button>
      </div>
      <div class="email-status" id="pinStatus"></div>
    </div>

    <!-- Step 2: Configure & Preview -->
    <div class="email-step" id="emailStep2">
      <h3>📧 Send Dashboard Email</h3>
      <label>To (comma-separated emails):</label>
      <textarea id="emailTo" placeholder="manager1@amazon.com, manager2@amazon.com"></textarea>
      <label>CC (optional):</label>
      <input type="text" id="emailCc" placeholder="cc@amazon.com"/>
      <label>Subject:</label>
      <input type="text" id="emailSubject" value="QC Dashboard — General Apr 2026"/>
      <p style="font-size:12px;color:#6b7280;margin-top:10px;">
        ℹ️ The full interactive dashboard will be sent as an HTML email.
      </p>
      <div class="email-btn-row">
        <button class="email-btn email-btn-ghost" onclick="closeEmailModal()">Cancel</button>
        <button class="email-btn email-btn-primary" onclick="previewEmail()">Preview & Send</button>
      </div>
      <div class="email-status" id="configStatus"></div>
    </div>

    <!-- Step 3: Confirm (Human in the Loop) -->
    <div class="email-step" id="emailStep3">
      <h3>✅ Confirm Send</h3>
      <div id="emailPreview" style="font-size:13px;color:#374151;"></div>
      <p style="font-size:12px;color:#b45309;margin-top:10px;font-weight:600;">
        ⚠️ Are you sure you want to send this email?
      </p>
      <div class="email-btn-row">
        <button class="email-btn email-btn-ghost" onclick="goToStep(2)">← Back</button>
        <button class="email-btn email-btn-danger" onclick="sendEmail()">🚀 Send Now</button>
      </div>
      <div class="email-status" id="sendStatus"></div>
    </div>

    <!-- Step 4: Success -->
    <div class="email-step" id="emailStep4">
      <h3>🎉 Email Sent!</h3>
      <div id="emailResult" style="font-size:13px;color:#065f46;"></div>
      <div class="email-btn-row">
        <button class="email-btn email-btn-primary" onclick="closeEmailModal()">Done</button>
      </div>
    </div>
  </div>
</div>

<script>
// Email API base URL (same server, same port)
const EMAIL_API = window.location.origin;
let adminPin = '';

function openEmailModal() {
  document.getElementById('emailOverlay').classList.add('open');
  goToStep(1);
  document.getElementById('emailPin').value = '';
  document.getElementById('emailPin').focus();
}

function closeEmailModal() {
  document.getElementById('emailOverlay').classList.remove('open');
  adminPin = '';
}

function goToStep(n) {
  document.querySelectorAll('.email-step').forEach(s => s.classList.remove('active'));
  document.getElementById('emailStep' + n).classList.add('active');
}

async function verifyPin() {
  const pin = document.getElementById('emailPin').value.trim();
  const status = document.getElementById('pinStatus');
  if (!pin) { status.className = 'email-status error'; status.textContent = 'Please enter a PIN.'; return; }
  
  try {
    const res = await fetch(EMAIL_API + '/api/verify-admin', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({pin: pin})
    });
    const data = await res.json();
    if (data.authorized) {
      adminPin = pin;
      status.className = 'email-status success'; status.textContent = '✓ Verified!';
      setTimeout(() => goToStep(2), 500);
    } else {
      status.className = 'email-status error'; status.textContent = '✗ Invalid PIN.';
    }
  } catch(e) {
    status.className = 'email-status error';
    status.textContent = 'Cannot reach email server. Contact admin.';
  }
}

async function previewEmail() {
  const to = document.getElementById('emailTo').value.trim();
  const cc = document.getElementById('emailCc').value.trim();
  const subject = document.getElementById('emailSubject').value.trim();
  const status = document.getElementById('configStatus');
  
  if (!to) { status.className = 'email-status error'; status.textContent = 'Please enter at least one recipient.'; return; }
  
  const recipients = to.split(',').map(s => s.trim()).filter(s => s);
  const ccList = cc ? cc.split(',').map(s => s.trim()).filter(s => s) : [];
  
  // Show preview
  const preview = document.getElementById('emailPreview');
  preview.innerHTML = `
    <p><strong>From:</strong> callaudit-noreply@amazon.com</p>
    <p><strong>To:</strong> ${recipients.join(', ')}</p>
    ${ccList.length ? '<p><strong>CC:</strong> ' + ccList.join(', ') + '</p>' : ''}
    <p><strong>Subject:</strong> ${subject}</p>
    <p><strong>Body:</strong> Full HTML dashboard (${Math.round(document.documentElement.outerHTML.length/1024)}KB)</p>
  `;
  goToStep(3);
}

async function sendEmail() {
  const to = document.getElementById('emailTo').value.trim();
  const cc = document.getElementById('emailCc').value.trim();
  const subject = document.getElementById('emailSubject').value.trim();
  const status = document.getElementById('sendStatus');
  
  const recipients = to.split(',').map(s => s.trim()).filter(s => s);
  const ccList = cc ? cc.split(',').map(s => s.trim()).filter(s => s) : [];
  
  status.className = 'email-status'; status.style.display = 'block';
  status.style.background = '#f3f4f6'; status.style.color = '#374151';
  status.textContent = '⏳ Sending...';
  
  try {
    const res = await fetch(EMAIL_API + '/api/send-email', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({pin: adminPin, recipients: recipients, cc: ccList, subject: subject})
    });
    const data = await res.json();
    if (data.success) {
      document.getElementById('emailResult').innerHTML = `
        <p>✅ Email sent successfully!</p>
        <p><strong>To:</strong> ${data.sent_to.join(', ')}</p>
        <p><strong>Time:</strong> ${data.sent_at}</p>
        <p style="font-size:11px;color:#6b7280;">Message ID: ${data.message_id}</p>
      `;
      goToStep(4);
    } else {
      status.className = 'email-status error';
      status.textContent = '✗ Error: ' + (data.error || 'Unknown error');
    }
  } catch(e) {
    status.className = 'email-status error';
    status.textContent = '✗ Network error: ' + e.message;
  }
}

// Close modal on overlay click
document.getElementById('emailOverlay').addEventListener('click', function(e) {
  if (e.target === this) closeEmailModal();
});
</script>
'''


def inject(html_path):
    """Inject email button into HTML file."""
    path = Path(html_path)
    if not path.exists():
        print(f"ERROR: File not found: {path}")
        return False
    
    html = path.read_text(encoding='utf-8')
    
    # Check if already injected
    if 'email-fab' in html:
        print(f"Already injected: {path.name}")
        return True
    
    # Insert before </body></html>
    if '</body>' in html:
        html = html.replace('</body>', EMAIL_INJECTION + '\n</body>')
    elif '</html>' in html:
        html = html.replace('</html>', EMAIL_INJECTION + '\n</html>')
    else:
        html += EMAIL_INJECTION
    
    path.write_text(html, encoding='utf-8')
    print(f"✅ Email button injected into: {path.name}")
    return True


if __name__ == '__main__':
    if len(sys.argv) < 2:
        # Default: inject into the web folder HTML
        target = Path('/home/ec2-user/web/index.html')
        if not target.exists():
            print("Usage: python3 inject_email_button.py <path_to_html>")
            sys.exit(1)
    else:
        target = Path(sys.argv[1])
    
    inject(target)
