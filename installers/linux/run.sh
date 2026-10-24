#!/usr/bin/env bash
# Starts 5belan with no terminal window. Errors go to launch.log in this folder.
cd "$(dirname "$(readlink -f "$0")")" || exit 1
export QT_AUTO_SCREEN_SCALE_FACTOR=1
export QTWEBENGINE_CHROMIUM_FLAGS="${QTWEBENGINE_CHROMIUM_FLAGS:---no-sandbox}"
if [ -x .venv/bin/5belan ]; then PY=.venv/bin/5belan; elif [ -x .venv/bin/python ]; then PY=.venv/bin/python; else PY=python3; fi
exec "$PY" guard.py deepseek_chat.py > launch.log 2>&1
