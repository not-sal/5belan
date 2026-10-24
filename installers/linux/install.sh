#!/usr/bin/env bash
# 5belan installer for Linux (x86_64 and ARM64: Ubuntu, Debian, Fedora, Arch, Orange Pi...). Run once:  bash install.sh
set -u
cd "$(dirname "$(readlink -f "$0")")"
HERE="$PWD"
ARCH="$(uname -m)"
echo "== 5belan installer =="
echo "Folder: $HERE"
echo "Arch:   $ARCH"

SUDO=""
if [ "$(id -u)" -ne 0 ]; then SUDO="sudo"; fi

echo
echo "[1/4] System packages (needs internet, may ask for your password)..."
if command -v apt-get >/dev/null 2>&1; then
  $SUDO apt-get update -y
  $SUDO apt-get install -y python3 python3-venv python3-pip fonts-noto-core fonts-noto-ui-core fonts-dejavu-core xdg-utils || true
  for p in libegl1 libgl1 libxkbcommon0 libxkbcommon-x11-0 libdbus-1-3 libfontconfig1 libglib2.0-0 \
           libxcb-cursor0 libxcb-icccm4 libxcb-image0 libxcb-keysyms1 libxcb-render-util0 libxcb-shape0 \
           libxcb-xinerama0 libxcb-xkb1 libnss3 libasound2t64 libasound2 libxcomposite1 libxdamage1 libxrandr2 libxtst6 libxshmfence1 libgbm1 libdrm2 libxfixes3 libxkbfile1 libcups2t64 libcups2; do
    $SUDO apt-get install -y "$p" >/dev/null 2>&1 || true
  done
elif command -v dnf >/dev/null 2>&1; then
  $SUDO dnf install -y python3 python3-pip xdg-utils mesa-libEGL mesa-libGL libxkbcommon-x11 xcb-util-cursor \
     xcb-util-wm xcb-util-image xcb-util-keysyms xcb-util-renderutil fontconfig dbus-libs google-noto-sans-arabic-fonts || true
elif command -v pacman >/dev/null 2>&1; then
  $SUDO pacman -S --noconfirm --needed python python-pip xdg-utils libxkbcommon-x11 xcb-util-cursor xcb-util-wm \
     xcb-util-image xcb-util-keysyms xcb-util-renderutil noto-fonts || true
else
  echo "NOTE: unknown package manager. Make sure python3 (3.9+), venv, and the Qt xcb libraries are installed."
fi

echo
echo "[2/4] Python environment..."
python3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" || { echo "ERROR: Python 3.9 or newer is required."; exit 1; }
python3 -m venv .venv || { echo "ERROR: could not create the virtual environment (install python3-venv)."; exit 1; }
PY="$HERE/.venv/bin/python"
ln -sf python "$HERE/.venv/bin/5belan" 2>/dev/null || true   # process name in the system monitor
"$PY" -m pip install --upgrade pip >/dev/null 2>&1

NEWOK=1
if [ "$ARCH" = "aarch64" ] || [ "$ARCH" = "arm64" ]; then
  # ARM64 wheels of PySide6 6.9+ need glibc 2.39; older systems use 6.8.0.2
  GLIBC="$(ldd --version 2>/dev/null | head -n1 | grep -oE '[0-9]+\.[0-9]+$')"
  echo "glibc: ${GLIBC:-unknown}"
  NEWOK="$(python3 -c "
v='${GLIBC:-0.0}'.split('.')
print(1 if (int(v[0]),int(v[1]))>=(2,39) else 0)" 2>/dev/null)"
fi

echo
echo "[3/4] Installing the Qt (PySide6) and document packages. This can take several minutes..."
if [ "$NEWOK" = "1" ]; then
  "$PY" -m pip install PySide6-Essentials PySide6-Addons
else
  "$PY" -m pip install "PySide6-Essentials==6.8.0.2" "PySide6-Addons==6.8.0.2" "shiboken6==6.8.0.2"
fi
"$PY" -m pip install pypdf python-docx openpyxl matplotlib arabic-reshaper python-bidi
"$PY" -c "import PySide6.QtWebEngineWidgets" >/dev/null 2>&1 || "$PY" -m pip install PySide6-WebEngine >/dev/null 2>&1 || true

echo
echo "Checking..."
QT_QPA_PLATFORM=offscreen "$PY" -c "import PySide6.QtWidgets, pypdf, docx, openpyxl, matplotlib, arabic_reshaper, bidi; print('Core packages: OK')" || { echo "ERROR: packages failed to load. Read the message above."; exit 1; }
QT_QPA_PLATFORM=offscreen "$PY" -c "import PySide6.QtWebEngineWidgets; print('Built-in browser engine: OK')" 2>/dev/null || echo "NOTE: built-in browser engine not available here. The app still works; only the in-app browser stays off."

echo
echo "[4/4] Launcher and menu icon..."
chmod +x run.sh
cat > 5belan.desktop <<D
[Desktop Entry]
Type=Application
Name=5belan
Comment=5belan AI assistant
Exec=$HERE/run.sh
Icon=$HERE/5belan.png
Terminal=false
Categories=Utility;
StartupWMClass=deepseek_chat.py
D
chmod +x 5belan.desktop
mkdir -p "$HOME/.local/share/applications"
cp 5belan.desktop "$HOME/.local/share/applications/5belan.desktop"
DESK="$(xdg-user-dir DESKTOP 2>/dev/null || echo "$HOME/Desktop")"
if [ -d "$DESK" ]; then
  cp 5belan.desktop "$DESK/5belan.desktop" && chmod +x "$DESK/5belan.desktop"
  command -v gio >/dev/null 2>&1 && gio set "$DESK/5belan.desktop" metadata::trusted true 2>/dev/null || true
  echo "Desktop icon: $DESK/5belan.desktop (right-click > Allow Launching if it shows a warning)"
fi
echo
echo "Done. Start it from the app menu / desktop icon, or run:  ./run.sh"
