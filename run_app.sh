#!/bin/bash
cd "$(dirname "$0")"
echo "กำลังเปิดโปรแกรม PDF Digger..."
./.venv/bin/python gui.py
