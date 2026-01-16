#!/bin/bash
# Script to update manifest with Cloudflare Tunnel URL
# Usage: ./update-manifest-url.sh https://your-url.trycloudflare.com

if [ -z "$1" ]; then
    echo "Usage: ./update-manifest-url.sh https://your-url.trycloudflare.com"
    exit 1
fi

TUNNEL_URL="$1"
# Remove trailing slash if present
TUNNEL_URL="${TUNNEL_URL%/}"

echo "Updating manifest.xml with URL: $TUNNEL_URL"

# Update manifest.xml
sed -i "s|https://localhost:3000|$TUNNEL_URL|g" manifest.xml
sed -i "s|http://localhost:3000|$TUNNEL_URL|g" manifest.xml

# Update dist/manifest.xml
sed -i "s|https://localhost:3000|$TUNNEL_URL|g" dist/manifest.xml
sed -i "s|http://localhost:3000|$TUNNEL_URL|g" dist/manifest.xml

echo "✅ Manifest updated! Now upload dist/manifest.xml to Word Online"
