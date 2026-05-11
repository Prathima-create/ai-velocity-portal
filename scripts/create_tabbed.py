"""Create a tabbed HTML wrapper combining General + SHT dashboards."""
import sys, io, json
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

base = Path(r"C:\Users\pratpk\amazon.com\Automation hosting - Documents\Fincom_QC")

general_path = base / "General_Apr_2026" / "QC_Dashboard_General_Apr_2026.html"
sht_path = base / "SHT_Apr_2026" / "QC_Dashboard_SHT_Apr_2026.html"

print(f"Reading General: {general_path} ({general_path.stat().st_size//1024}KB)")
print(f"Reading SHT: {sht_path} ({sht_path.stat().st_size//1024}KB)")

general_html = general_path.read_text(encoding='utf-8')
sht_html = sht_path.read_text(encoding='utf-8')

# Escape for embedding in JavaScript
general_escaped = json.dumps(general_html)
sht_escaped = json.dumps(sht_html)

wrapper = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>FinCom QC Dashboard - Apr 2026</title>
<style>
body {{ margin:0; font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; }}
.tab-bar {{ display:flex; background:#1a1a2e; padding:0 20px; position:sticky; top:0; z-index:9999; }}
.tab-btn {{ padding:14px 28px; color:#aaa; cursor:pointer; border:none; background:none; font-size:15px; font-weight:600; border-bottom:3px solid transparent; transition:all 0.2s; }}
.tab-btn:hover {{ color:#fff; }}
.tab-btn.active {{ color:#fff; border-bottom-color:#4285f4; background:rgba(66,133,244,0.1); }}
.tab-content {{ display:none; }}
.tab-content.active {{ display:block; }}
iframe {{ width:100%; border:none; height:calc(100vh - 52px); }}
</style>
</head>
<body>
<div class="tab-bar">
  <button class="tab-btn active" onclick="switchTab(this, 'general')">General</button>
  <button class="tab-btn" onclick="switchTab(this, 'sht')">SHT</button>
</div>
<div id="general" class="tab-content active">
  <iframe id="general-frame"></iframe>
</div>
<div id="sht" class="tab-content">
  <iframe id="sht-frame"></iframe>
</div>
<script>
var dashboards = {{
  general: {general_escaped},
  sht: {sht_escaped}
}};
function switchTab(btn, tab) {{
  document.querySelectorAll('.tab-content').forEach(function(el) {{ el.classList.remove('active'); }});
  document.querySelectorAll('.tab-btn').forEach(function(el) {{ el.classList.remove('active'); }});
  document.getElementById(tab).classList.add('active');
  btn.classList.add('active');
}}
// Initialize iframes
document.getElementById('general-frame').srcdoc = dashboards.general;
document.getElementById('sht-frame').srcdoc = dashboards.sht;
</script>
</body>
</html>'''

out_path = base / "index.html"
out_path.write_text(wrapper, encoding='utf-8')
print(f"\nTabbed dashboard saved to: {out_path}")
print(f"File size: {out_path.stat().st_size // 1024} KB")
