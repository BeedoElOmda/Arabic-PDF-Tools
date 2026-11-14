#!/usr/bin/env bash
# Builds every Windows package of Arabic PDF Tools from this repository:
#   dist/ArabicPDFTools-Setup-<ver>.exe     installer (per user, no admin)
#   dist/ArabicPDFTools-Portable-<ver>.zip  portable folder (no installation)
#   dist/Arabic-PDF-Tools-Script-<ver>.zip  Install.cmd version (needs Python/internet)
# Runs on Linux (GitHub Actions ubuntu-latest or locally) with: python3, pip, curl, zip, makensis.
# Everything bundled is pinned, so every build of a version contains the same files.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="$(tr -d '[:space:]' < "$ROOT/VERSION")"
WORK="$ROOT/build"
STAGE="$WORK/stage"
DIST="$ROOT/dist"

# --- pinned inputs ---------------------------------------------------------
PY_URL="https://github.com/astral-sh/python-build-standalone/releases/download/20241206/cpython-3.12.8+20241206-x86_64-pc-windows-msvc-install_only.tar.gz"
PY_SHA256="${PY_SHA256:-}"   # optional: set to verify the download
WHEELS=(pymupdf==2.0 python-pptx==1.0.2 python-docx==1.2.0 pdf2docx==0.5.13 pikepdf==10.16.0
        pillow==12.3.0 fonttools==4.66.1 numpy==2.5.3 opencv-python-headless==5.0.0.93 lxml==6.1.3
        typing_extensions==4.16.0 XlsxWriter==3.2.9 Deprecated==3.0.1 wrapt==2.5.2 packaging==26.3
        fire==0.7.1 termcolor==3.3.0)

echo "== Arabic PDF Tools $VERSION"
rm -rf "$WORK" "$DIST"
mkdir -p "$STAGE" "$DIST" "$WORK/dl"

# --- 1. portable Windows Python ----------------------------------------------
echo "== Python runtime"
curl -sSL -o "$WORK/dl/python.tar.gz" "$PY_URL"
if [ -n "$PY_SHA256" ]; then echo "$PY_SHA256  $WORK/dl/python.tar.gz" | sha256sum -c -; fi
tar -xzf "$WORK/dl/python.tar.gz" -C "$STAGE"          # -> $STAGE/python

# --- 2. libraries (Windows builds, core only, no AI extras) -------------------
echo "== Libraries"
SP="$STAGE/python/Lib/site-packages"
python3 -m pip install -q --no-compile --no-deps --target "$SP" \
  --platform win_amd64 --python-version 3.12 --implementation cp --only-binary=:all: "${WHEELS[@]}"

# PyMuPDF needs the C++ runtime; numpy ships the same Microsoft DLL: put it next to python.exe
cp "$(ls "$SP"/numpy.libs/msvcp140-*.dll | head -1)" "$STAGE/python/msvcp140.dll"

# trim files Windows users never need
find "$STAGE/python" -name "*.pdb" -delete
find "$STAGE/python" -name "*.lib" -delete
rm -rf "$STAGE/python/include" "$STAGE/python/libs" "$STAGE/python/Lib/test" \
       "$STAGE/python/Lib/idlelib/idle_test" "$STAGE/python/Lib/ensurepip" "$STAGE/python/Scripts" \
       "$SP/bin" "$SP/pymupdf/mupdf-devel" "$SP/wrapt-stubs" "$SP"/pip "$SP"/pip-*.dist-info
find "$STAGE/python" -name "__pycache__" -type d -prune -exec rm -rf {} +

# --- 3. the app -------------------------------------------------------------
echo "== App"
mkdir -p "$STAGE/app"
cp "$ROOT"/src/*.py "$ROOT"/src/*.pyw "$ROOT"/src/icon.ico "$STAGE/app/"
echo "$VERSION" > "$STAGE/app/version.txt"

# --- 4. installer + portable launcher ------------------------------------------
echo "== Installer"
cp "$ROOT"/packaging/*.nsi "$WORK/"
( cd "$WORK" && makensis -V2 -DAPPVERSION="$VERSION" installer.nsi && makensis -V2 -DAPPVERSION="$VERSION" launcher.nsi )
mv "$WORK/ArabicPDFTools-Setup-$VERSION.exe" "$DIST/"

echo "== Portable"
P="$WORK/portable/Arabic PDF Tools (Portable)"
mkdir -p "$P/data"
cp "$WORK/ArabicPDFTools.exe" "$P/"
cp -r "$STAGE/python" "$STAGE/app" "$P/data/"
cp "$ROOT/packaging/portable-README.txt" "$P/README.txt"
( cd "$WORK/portable" && zip -q -r -9 "$DIST/ArabicPDFTools-Portable-$VERSION.zip" "Arabic PDF Tools (Portable)" )

# --- 5. script version (Install.cmd) -------------------------------------------
echo "== Script version"
S="$WORK/script/Arabic-PDF-Tools"
mkdir -p "$S"
cp "$ROOT"/src/*.py "$ROOT"/script-version/* "$S/"
( cd "$WORK/script" && zip -q -r "$DIST/Arabic-PDF-Tools-Script-$VERSION.zip" Arabic-PDF-Tools )

# the launcher exe on its own too (so it can be signed separately)
cp "$WORK/ArabicPDFTools.exe" "$DIST/ArabicPDFTools.exe"

echo "== Done"
ls -la "$DIST"
