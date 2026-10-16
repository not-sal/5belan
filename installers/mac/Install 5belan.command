#!/bin/bash
# 5belan installer for macOS. Double-click this file (first time: right-click > Open).
set -u
SRC="$(cd "$(dirname "$0")" && pwd)"
APPDIR="$HOME/Library/Application Support/5belan"
APP="$HOME/Applications/5belan.app"
echo "== 5belan installer for macOS =="

# 1. Python 3.9+ (Homebrew or python.org)
PYBIN=""
for c in /opt/homebrew/bin/python3 /usr/local/bin/python3 python3; do
  if command -v "$c" >/dev/null 2>&1 && "$c" -c "import sys; sys.exit(0 if (3,9)<=sys.version_info[:2]<=(3,14) else 1)" 2>/dev/null; then PYBIN="$(command -v "$c")"; break; fi
done
if [ -z "$PYBIN" ]; then
  echo "Python 3.9 - 3.14 was not found."
  echo "Install it from https://www.python.org/downloads/macos/ (or: brew install python), then run this installer again."
  read -n 1 -s -r -p "Press any key to close..."; exit 1
fi
echo "Using Python: $PYBIN"

# 2. copy the program files (existing chats and settings are kept)
mkdir -p "$APPDIR"
for f in deepseek_chat.py guard.py 5belan.png requirements.txt; do cp "$SRC/$f" "$APPDIR/$f"; done

# 3. private environment + packages
if [ ! -x "$APPDIR/.venv/bin/python" ]; then "$PYBIN" -m venv "$APPDIR/.venv" || { echo "Could not create the environment."; read -n 1 -s -r -p "Press any key..."; exit 1; }; fi
PY="$APPDIR/.venv/bin/python"
"$PY" -m pip install --disable-pip-version-check --upgrade pip >/dev/null 2>&1
echo "Installing packages (a few minutes)..."
"$PY" -m pip install --disable-pip-version-check -r "$APPDIR/requirements.txt" || { echo "Package installation failed. Check your internet connection and run again."; read -n 1 -s -r -p "Press any key..."; exit 1; }
"$PY" -m pip install --disable-pip-version-check PySide6-WebEngine >/dev/null 2>&1 || echo "Note: the built-in browser engine could not be installed. The app works without it."
QT_QPA_PLATFORM=offscreen "$PY" -c "import PySide6.QtWidgets, pypdf, docx, openpyxl, matplotlib, arabic_reshaper, bidi; print('Packages OK')" || { echo "Packages failed to load."; read -n 1 -s -r -p "Press any key..."; exit 1; }

# 4. the 5belan.app launcher (no terminal window)
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
cat > "$APP/Contents/MacOS/5belan" <<L
#!/bin/bash
cd "$APPDIR" || exit 1
exec "$APPDIR/.venv/bin/python" guard.py deepseek_chat.py > "$APPDIR/launch.log" 2>&1
L
chmod +x "$APP/Contents/MacOS/5belan"
cat > "$APP/Contents/Info.plist" <<'L'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
<key>CFBundleName</key><string>5belan</string>
<key>CFBundleDisplayName</key><string>5belan</string>
<key>CFBundleIdentifier</key><string>chat.belan.app</string>
<key>CFBundleExecutable</key><string>5belan</string>
<key>CFBundleIconFile</key><string>5belan</string>
<key>CFBundlePackageType</key><string>APPL</string>
<key>CFBundleVersion</key><string>1</string>
<key>LSMinimumSystemVersion</key><string>11.0</string>
<key>NSHighResolutionCapable</key><true/>
</dict></plist>
L
ICONSET="$(mktemp -d)/5belan.iconset"; mkdir -p "$ICONSET"
for s in 16 32 64 128 256 512; do
  sips -z $s $s "$SRC/5belan.png" --out "$ICONSET/icon_${s}x${s}.png" >/dev/null 2>&1
  sips -z $((s*2)) $((s*2)) "$SRC/5belan.png" --out "$ICONSET/icon_${s}x${s}@2x.png" >/dev/null 2>&1
done
iconutil -c icns "$ICONSET" -o "$APP/Contents/Resources/5belan.icns" >/dev/null 2>&1 || true
xattr -dr com.apple.quarantine "$APP" "$APPDIR" 2>/dev/null || true
echo
echo "Done. 5belan is in ~/Applications (open Finder > Go > Applications in your home folder). Drag it to the Dock if you like."
open "$APP" 2>/dev/null || true
read -n 1 -s -r -p "Press any key to close..."
