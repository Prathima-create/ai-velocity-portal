#!/bin/bash
# ============================================================
# EC2 Setup Script — FinCom QC Automation Dashboard
# ============================================================
# Run this on a fresh Amazon Linux 2023 / Ubuntu EC2 instance
# to set up the Streamlit dashboard with S3 sync.
#
# Prerequisites:
#   - EC2 instance (t3.micro or larger)
#   - IAM role attached with S3 read/write access to your bucket
#   - Security group allowing port 8501 from VPN (10.0.0.0/8)
#   - Git installed
#
# Usage:
#   chmod +x setup_ec2.sh
#   ./setup_ec2.sh
# ============================================================

set -e

echo "============================================"
echo "  FinCom QC Dashboard — EC2 Setup"
echo "============================================"

# ─── System Dependencies ──────────────────────────────────────
echo ""
echo "📦 Installing system dependencies..."

if command -v yum &> /dev/null; then
    # Amazon Linux
    sudo yum update -y
    sudo yum install -y python3 python3-pip git
elif command -v apt &> /dev/null; then
    # Ubuntu/Debian
    sudo apt update -y
    sudo apt install -y python3 python3-pip python3-venv git
fi

# ─── Clone/Update Repository ─────────────────────────────────
echo ""
echo "📂 Setting up application..."

APP_DIR="/opt/qc-dashboard"
if [ -d "$APP_DIR" ]; then
    echo "  Updating existing installation..."
    cd "$APP_DIR"
    git pull origin main 2>/dev/null || true
else
    echo "  Cloning repository..."
    sudo mkdir -p "$APP_DIR"
    sudo chown $(whoami):$(whoami) "$APP_DIR"
    git clone https://github.com/Prathima-create/ai-velocity-portal.git "$APP_DIR"
fi

cd "$APP_DIR/qc-automation"

# ─── Python Virtual Environment ──────────────────────────────
echo ""
echo "🐍 Setting up Python environment..."

python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

# ─── Environment Configuration ────────────────────────────────
echo ""
echo "⚙️  Setting up configuration..."

if [ ! -f .env ]; then
    cp .env.template .env
    echo "  Created .env from template — EDIT THIS FILE with your values!"
    echo "  nano $APP_DIR/qc-automation/.env"
fi

# Create data directory
mkdir -p data

# ─── Systemd Service (Streamlit App) ─────────────────────────
echo ""
echo "🚀 Setting up systemd service..."

sudo tee /etc/systemd/system/qc-dashboard.service > /dev/null << EOF
[Unit]
Description=FinCom QC Automation Dashboard (Streamlit)
After=network.target

[Service]
Type=simple
User=$(whoami)
WorkingDirectory=$APP_DIR/qc-automation
EnvironmentFile=$APP_DIR/qc-automation/.env
ExecStart=$APP_DIR/qc-automation/venv/bin/streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

# ─── Systemd Timer (S3 Sync — every 30 minutes) ──────────────
echo ""
echo "⏰ Setting up sync timer..."

sudo tee /etc/systemd/system/qc-sync.service > /dev/null << EOF
[Unit]
Description=QC Dashboard SharePoint → S3 Sync
After=network.target

[Service]
Type=oneshot
User=$(whoami)
WorkingDirectory=$APP_DIR/qc-automation
EnvironmentFile=$APP_DIR/qc-automation/.env
ExecStart=$APP_DIR/qc-automation/venv/bin/python sync_sharepoint_to_s3.py
EOF

sudo tee /etc/systemd/system/qc-sync.timer > /dev/null << EOF
[Unit]
Description=Run QC Sync every 30 minutes

[Timer]
OnBootSec=2min
OnUnitActiveSec=30min
Persistent=true

[Install]
WantedBy=timers.target
EOF

# ─── Enable and Start Services ────────────────────────────────
echo ""
echo "✅ Enabling services..."

sudo systemctl daemon-reload
sudo systemctl enable qc-dashboard.service
sudo systemctl enable qc-sync.timer

echo ""
echo "============================================"
echo "  SETUP COMPLETE!"
echo "============================================"
echo ""
echo "📝 Next steps:"
echo "  1. Edit your environment config:"
echo "     nano $APP_DIR/qc-automation/.env"
echo ""
echo "  2. (Optional) Place initial CSV data files in:"
echo "     $APP_DIR/qc-automation/data/"
echo ""
echo "  3. Start the dashboard:"
echo "     sudo systemctl start qc-dashboard"
echo ""
echo "  4. Start the sync timer:"
echo "     sudo systemctl start qc-sync.timer"
echo ""
echo "  5. Access the dashboard at:"
echo "     http://$(hostname -I | awk '{print $1}'):8501"
echo ""
echo "📊 Management commands:"
echo "  sudo systemctl status qc-dashboard    # Check dashboard status"
echo "  sudo systemctl restart qc-dashboard   # Restart dashboard"
echo "  sudo systemctl status qc-sync.timer   # Check sync timer"
echo "  sudo journalctl -u qc-dashboard -f    # View dashboard logs"
echo "  sudo journalctl -u qc-sync -f         # View sync logs"
echo ""
