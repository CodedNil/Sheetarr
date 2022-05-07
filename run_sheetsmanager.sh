#!/bin/sh
cd "$(dirname "$0")";
CWD="$(pwd)"
echo $CWD
python3 sheetarr.py
