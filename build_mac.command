#!/bin/bash
# Double-click this file on a Mac to build  "KeithsQuotationGenerator.app"  (needs Python 3.9+ from python.org).
cd "$(dirname "$0")" || exit 1
echo "=== Keith's Quotation Generator - Mac build ==="

PY="$(command -v python3)"
if [ -z "$PY" ] || ! "$PY" -c 'import sys, tkinter; sys.exit(0 if sys.version_info >= (3, 9) and tkinter.TkVersion >= 8.6 else 1)' 2>/dev/null; then
  echo
  echo "Python 3.9+ with Tk 8.6 was not found."
  echo "1. Download and install Python from  https://www.python.org/downloads/macos/"
  echo "   (the python.org installer includes Tk; the Apple/Xcode python does not work well)"
  echo "2. Double-click this file again."
  open "https://www.python.org/downloads/macos/"
  read -n 1 -s -r -p "Press any key to close..."
  exit 1
fi

echo "Using: $PY ($("$PY" --version))  -  $(uname -m) Mac"
"$PY" -m venv .venv-build || exit 1
source .venv-build/bin/activate
python -m pip install --upgrade pip >/dev/null
python -m pip install -r requirements.txt || { echo "Installing requirements failed."; read -n 1 -s -r -p "Press any key..."; exit 1; }
python build_exe.py || { echo "Build failed - see the messages above."; read -n 1 -s -r -p "Press any key..."; exit 1; }

echo
echo "Finished. Opening the 'release' folder:"
echo "  KeithsQuotationGenerator.app  - drag it to Applications"
echo "  KeithsQuotationGenerator-mac-*.zip  - send this to other Macs of the same type"
open release
read -n 1 -s -r -p "Press any key to close..."
