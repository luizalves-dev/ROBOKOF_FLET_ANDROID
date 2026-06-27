#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
flet build apk . --yes
printf '\nAPK gerado em: %s/build/apk\n' "$PWD"
