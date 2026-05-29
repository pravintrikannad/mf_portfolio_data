#!/bin/bash
# Installs the Downloads watcher as a macOS LaunchAgent (auto-starts on login).

REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON="$(which python3)"
PLIST="$HOME/Library/LaunchAgents/com.mfscanner.watch_downloads.plist"

pip3 install watchdog openpyxl requests -q

cat > "$PLIST" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.mfscanner.watch_downloads</string>
    <key>ProgramArguments</key>
    <array>
        <string>$PYTHON</string>
        <string>$REPO_DIR/watch_downloads.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>$REPO_DIR</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>$REPO_DIR/logs/watcher.log</string>
    <key>StandardErrorPath</key>
    <string>$REPO_DIR/logs/watcher.log</string>
</dict>
</plist>
EOF

mkdir -p "$REPO_DIR/logs"
launchctl unload "$PLIST" 2>/dev/null
launchctl load "$PLIST"
echo "✅ Watcher installed and running."
echo "   Logs: $REPO_DIR/logs/watcher.log"
echo "   To stop:  launchctl unload $PLIST"
echo "   To start: launchctl load $PLIST"
