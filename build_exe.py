#!/usr/bin/env python3
"""
build_exe.py -- packages the app with PyInstaller.   Run it ON the kind of computer you are building for
(PyInstaller cannot cross-compile):

    python build_exe.py

  Windows : release/KeithsQuotationGenerator.exe   + products.csv + images/ + templates.json
  Mac     : release/KeithsQuotationGenerator.app   + KeithsQuotationGenerator-mac-<arm64|x86_64>.zip
            (a Mac build only runs on the same kind of Mac: Apple Silicon "arm64" or older Intel "x86_64")
  Linux   : release/KeithsQuotationGenerator       + products.csv + images/ + templates.json
"""
import importlib.util
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import PyInstaller.__main__

HERE = Path(__file__).parent.resolve()
os.chdir(HERE)
SEP = os.pathsep                       # ';' on Windows, ':' elsewhere
NAME = "KeithsQuotationGenerator"
BUNDLE_ID = "com.keith.quotationgenerator"
IS_MAC = sys.platform == "darwin"
ONEDIR = IS_MAC or os.environ.get("QG_ONEDIR") == "1"      # a Mac app must be a folder-type bundle


def pyinstaller_args(placeholder):
    args = [
        "quotation_app.py",
        "--name", NAME,
        "--windowed", "--noconfirm", "--clean",
        "--onedir" if ONEDIR else "--onefile",
        "--add-data", f"assets{SEP}assets",                  # logo + signature (read-only, inside the app)
        "--add-data", f"products.csv{SEP}defaults",          # sample files copied to the user's folder on first run
        "--add-data", f"images{SEP}defaults/images",
        "--add-data", f"templates.json{SEP}defaults",        # the sample "Matrice 4 Thermal Package" template
        "--add-data", f"{placeholder}{SEP}docx/parts",       # python-docx needs this folder to exist
        "--collect-data", "docx",                            # python-docx's built-in templates
    ]
    if IS_MAC:
        args += ["--osx-bundle-identifier", BUNDLE_ID]
    for module in ("docx2pdf", "tqdm", "pythoncom", "win32com.client"):   # Word -> PDF
        if importlib.util.find_spec(module.split(".")[0]):
            args += ["--hidden-import", module]
    return args


def assemble_release():
    release = HERE / "release"
    shutil.rmtree(release, ignore_errors=True)
    release.mkdir()
    dist = HERE / "dist"
    if IS_MAC:
        app = dist / f"{NAME}.app"
        target = release / app.name
        shutil.copytree(app, target, symlinks=True)
        subprocess.run(["codesign", "--force", "--deep", "--sign", "-", str(target)], check=False)   # ad-hoc signature
        arch = platform.machine() or "mac"
        zip_path = release / f"{NAME}-mac-{arch}.zip"
        subprocess.run(["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent", str(target), str(zip_path)], check=True)
        print(f"\nDone -> {target}\n        {zip_path}  (share this zip)")
        return
    if ONEDIR:
        shutil.copytree(dist / NAME, release / NAME)
        shown = f"{NAME}/"
    else:
        exe = next(dist.glob(NAME + "*"))
        shutil.copy(exe, release / exe.name)
        shown = exe.name
    shutil.copy(HERE / "products.csv", release / "products.csv")
    shutil.copytree(HERE / "images", release / "images")
    shutil.copy(HERE / "templates.json", release / "templates.json")
    print(f"\nDone -> {release}\n  {shown}, products.csv, images/, templates.json")


if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as tmp:
        placeholder = Path(tmp) / "keep.txt"
        placeholder.write_text("python-docx looks for templates through this folder - do not delete.\n")
        PyInstaller.__main__.run(pyinstaller_args(placeholder))
    assemble_release()
