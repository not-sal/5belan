#!/usr/bin/env bash
# Build the three installers from the repo files and (optionally) publish.
# Usage: installers/build_release.sh <out_dir> [--publish "release notes in Arabic or English"]
# Needs: makensis (apt install nsis), zip, tar, python3, git. Run from the repo root.
set -euo pipefail
OUT="${1:?out dir}"; shift || true
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
VER="$(grep -m1 '^APP_VERSION' "$ROOT/deepseek_chat.py" | sed 's/.*"\(.*\)".*/\1/')"
echo "version $VER"
rm -rf "$OUT"; mkdir -p "$OUT/win" "$OUT/linux/5belan" "$OUT/mac/5belan-macOS"
APPFILES="deepseek_chat.py guard.py 5belan.png 5belan.ico requirements.txt"
for f in $APPFILES; do cp "$ROOT/$f" "$OUT/win/"; done
cp "$ROOT/installers/win/5belan.nsi" "$OUT/win/"
sed -i "s/^!define APPVER .*/!define APPVER \"$VER\"/" "$OUT/win/5belan.nsi"
(cd "$OUT/win" && makensis -V1 5belan.nsi)
for f in deepseek_chat.py guard.py 5belan.png 5belan.ico requirements.txt; do cp "$ROOT/$f" "$OUT/linux/5belan/"; done
cp "$ROOT"/installers/linux/* "$OUT/linux/5belan/"
for f in deepseek_chat.py guard.py 5belan.png requirements.txt; do cp "$ROOT/$f" "$OUT/mac/5belan-macOS/"; done
cp "$ROOT"/installers/mac/* "$OUT/mac/5belan-macOS/"
(cd "$OUT" && tar czf 5belan-linux.tar.gz -C linux 5belan && (cd mac && zip -qr ../5belan-macOS.zip 5belan-macOS))
cp "$OUT/win/5belan-Setup.exe" "$OUT/"
ls -la "$OUT"/5belan-Setup.exe "$OUT"/5belan-linux.tar.gz "$OUT"/5belan-macOS.zip
if [ "${1:-}" = "--publish" ]; then
  NOTES="${2:?notes}"
  (cd "$ROOT" && python3 publish.py --repo saleh9688/5belan --pip PySide6-WebEngine --notes "$NOTES")   # writes update.json, does NOT push
  (cd "$ROOT" && git add -A && git -c user.name=saleh9688 -c user.email=saleh96881@gmail.com commit -q -m "$VER: $NOTES" && git push -q origin HEAD:main)
  echo "pushed. verify: curl -s https://raw.githubusercontent.com/saleh9688/5belan/main/update.json | grep version"
fi
