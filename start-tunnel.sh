#!/bin/bash
# Cloudflare Tunnel helper script
# This will tunnel your localhost:3000 to a public HTTPS URL

# Use local cloudflared if available, otherwise try system one
if [ -f "./cloudflared" ]; then
    CLOUDFLARED="./cloudflared"
elif command -v cloudflared &> /dev/null; then
    CLOUDFLARED="cloudflared"
else
    echo "❌ cloudflared not found!"
    echo "Download it with: wget https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64 -O cloudflared && chmod +x cloudflared"
    exit 1
fi

echo "Starting Cloudflare Tunnel for https://localhost:3000..."
echo "Make sure your server is running with: npm run serve:ssl"
echo ""
echo "Copy the 'https://' URL that appears below and use it in your manifest"
echo "Press Ctrl+C to stop the tunnel"
echo ""

$CLOUDFLARED tunnel --url https://localhost:3000
