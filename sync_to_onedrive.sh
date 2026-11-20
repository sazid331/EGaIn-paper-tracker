#!/bin/bash
# Pull latest from GitHub and sync Excel reports to OneDrive
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
ONEDRIVE_DIR="$HOME/Library/CloudStorage/OneDrive-UniversityofTennessee/EGaIn Paper Tracker"

cd "$REPO_DIR" || exit 1
git pull --quiet origin main 2>/dev/null

mkdir -p "$ONEDRIVE_DIR"
for f in reports/egain_papers_*.xlsx; do
    [ -f "$f" ] && cp "$f" "$ONEDRIVE_DIR/"
done
