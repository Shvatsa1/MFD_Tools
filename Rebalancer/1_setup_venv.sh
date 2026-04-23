#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv
fi
# shellcheck disable=SC1091
. .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
[ -f mfd_pack.ini ] || cp mfd_pack.example.ini mfd_pack.ini
echo
echo "Next: add data/latestNAV_Reports.xlsx, put client files under data/clients/by_client/, edit mfd_pack.ini, then:"
echo "  python build_mfd_pack.py"
echo
echo "Optional bootstrap (age/glide):"
echo "  python build_mfd_pack.py --bootstrap-client-risk"
echo
