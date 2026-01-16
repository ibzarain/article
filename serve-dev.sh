#!/bin/bash
# Serve the dist folder without SSL for testing
cd "$(dirname "$0")"
serve dist -l 3000 --no-clipboard --no-redirect
