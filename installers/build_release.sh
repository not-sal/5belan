#!/usr/bin/env bash
# Build the three installers from the repo files and (optionally) publish.
# Usage: installers/build_release.sh <out_dir> [--publish "release notes in Arabic or English"]
# Needs: makensis (apt install nsis / winget NSIS.NSIS), tar, python, git; zip optional. Run from the repo root.
set -euo pipefail
OUT="${1:?out dir}"; shift || true
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
# Tools: also works from Git Bash on Windows (NSIS from winget, "py" launcher, no zip).
MAKENSIS="$(command -v makensis || true)"
[ -z "$MAKENSIS" ] && [ -x "/c/Program Files (x86)/NSIS/makensis.exe" ] && MAKENSIS="/c/Program Files (x86)/NSIS/makensis.exe"
[ -n "$MAKENSIS" ] || { echo "makensis not found (install NSIS)"; exit 1; }
PY=""
for c in python3 py python; do "$c" -c "import sys" >/dev/null 2>&1 && { PY="$c"; break; }; done
[ -n "$PY" ] || { echo "python not found"; exit 1; }
# zip a folder; marks .command/.sh as executable (Python fallback when zip is missing)
zipdir() {
  if command -v zip >/dev/null; then zip -qr "$1" "$2"; return; fi
  "$PY" - "$1" "$2" <<'EOF'
import os, sys, zipfile
out, src = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for base, _, files in os.walk(src):
        for f in sorted(files):
            p = os.path.join(base, f)
            info = zipfile.ZipInfo.from_file(p, p.replace(os.sep, "/"))
            mode = 0o755 if f.endswith((".command", ".sh")) else 0o644
            info.external_attr = (0o100000 | mode) << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            with open(p, "rb") as fh:
                z.writestr(info, fh.read())
EOF
}
VER="$(grep -m1 '^APP_VERSION' "$ROOT/deepseek_chat.py" | sed 's/.*"\(.*\)".*/\1/')"
echo "version $VER"
rm -rf "$OUT"; mkdir -p "$OUT/win" "$OUT/linux/5belan" "$OUT/mac/5belan-macOS"
APPFILES="deepseek_chat.py guard.py 5belan.png 5belan.ico requirements.txt"
for f in $APPFILES; do cp "$ROOT/$f" "$OUT/win/"; done
cp "$ROOT/installers/win/5belan.nsi" "$OUT/win/"
sed -i "s/^!define APPVER .*/!define APPVER \"$VER\"/" "$OUT/win/5belan.nsi"
(cd "$OUT/win" && "$MAKENSIS" -V1 5belan.nsi)
for f in deepseek_chat.py guard.py 5belan.png 5belan.ico requirements.txt; do cp "$ROOT/$f" "$OUT/linux/5belan/"; done
cp "$ROOT"/installers/linux/* "$OUT/linux/5belan/"
for f in deepseek_chat.py guard.py 5belan.png requirements.txt; do cp "$ROOT/$f" "$OUT/mac/5belan-macOS/"; done
cp "$ROOT"/installers/mac/* "$OUT/mac/5belan-macOS/"
(cd "$OUT" && tar czf 5belan-linux.tar.gz -C linux 5belan && (cd mac && zipdir ../5belan-macOS.zip 5belan-macOS))
cp "$OUT/win/5belan-Setup.exe" "$OUT/"
ls -la "$OUT"/5belan-Setup.exe "$OUT"/5belan-linux.tar.gz "$OUT"/5belan-macOS.zip
if [ "${1:-}" = "--publish" ]; then
  NOTES="${2:?notes}"
  (cd "$ROOT" && "$PY" publish.py --repo saleh9688/5belan --pip PySide6-WebEngine --notes "$NOTES")   # writes update.json, does NOT push
  (cd "$ROOT" && git add -A && git -c user.name=saleh9688 -c user.email=saleh96881@gmail.com commit -q -m "$VER: $NOTES" && git push -q origin HEAD:main)
  echo "pushed. verify: curl -s https://raw.githubusercontent.com/saleh9688/5belan/main/update.json | grep version"
fi
