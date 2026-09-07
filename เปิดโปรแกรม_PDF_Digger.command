#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
if [ ! -d "$DIR/.venv" ]; then
    DIR="/Users/kritsmacbook/Desktop/Keng Docs/PDF digger"
fi
cd "$DIR"
echo "กำลังเปิดโปรแกรม PDF Digger..."
"$DIR/.venv/bin/python" "$DIR/gui.py"
