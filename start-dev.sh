#!/bin/bash
# Start both the HTTPS server and Cloudflare tunnel
# This script will start the server in the background and show the tunnel URL

echo "🚀 Starting development server and Cloudflare tunnel..."
echo ""

# Check if cloudflared exists
if [ ! -f "./cloudflared" ]; then
    echo "❌ cloudflared not found! Download it first:"
    echo "   wget https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64 -O cloudflared"
    echo "   chmod +x cloudflared"
    exit 1
fi

# Start server in background
echo "📦 Starting HTTPS server on port 3000..."
npm run serve:ssl > /tmp/serve.log 2>&1 &
SERVER_PID=$!
echo "   Server PID: $SERVER_PID"
echo "   (Logs: /tmp/serve.log)"
echo ""

# Wait a moment for server to start
sleep 2

# Start tunnel
echo "🌐 Starting Cloudflare tunnel..."
echo "   ⚠️  Copy the HTTPS URL that appears below!"
echo "   ⚠️  Then run: ./update-manifest-url.sh <URL>"
echo ""
echo "   Press Ctrl+C to stop both server and tunnel"
echo ""

# Trap to kill server when tunnel stops
trap "kill $SERVER_PID 2>/dev/null; exit" INT TERM

./cloudflared tunnel --url https://localhost:3000
