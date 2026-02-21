#!/bin/bash
# Launcher script for the Docx Duplicate Finder GUI
# It ensures dependencies are installed via uv and runs the GUI wrapper script

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

echo "Checking environment and starting GUI... please wait."

# Ensure uv is available
if ! command -v uv &> /dev/null; then
    echo "錯誤: 'uv' 命令找不到。請確認已正確安裝 uv 且系統環境變數設定正確。"
    exit 1
fi

# Run the GUI application using uv run to leverage the virtual env automatically
uv run gui_app.py
