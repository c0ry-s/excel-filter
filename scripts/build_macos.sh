#!/usr/bin/env bash
set -euo pipefail

APP_NAME="Rex Excel Filter"
ENTRYPOINT="excelfilter.py"

# Keep this aligned with your app version
VERSION="1.0.1"

ICON_ICNS="Assets/rexexcelfilter.icns"

# Each entry is "source:destination" (macOS uses : for add-data)
ADD_DATA=(
  "Assets/Rexie.png:Assets"
)

# If tkinterdnd2 ever fails to import in the frozen build, keep this
HIDDEN_IMPORTS=(
  "tkinterdnd2"
)

OUT_DIR="$HOME/dev/builds/macos/RexExcelFilter"
ZIP_BASENAME="RexExcelFilter"

# Always run from project root
cd "$(dirname "$0")/.."

# Clean old outputs
rm -rf build dist *.spec

# Build
PYI_ARGS=(
  --noconfirm
  --clean
  --windowed
  --name "$APP_NAME"
  --icon "$ICON_ICNS"
)

for item in "${ADD_DATA[@]}"; do
  PYI_ARGS+=( --add-data "$item" )
done

for hi in "${HIDDEN_IMPORTS[@]}"; do
  PYI_ARGS+=( --hidden-import "$hi" )
done

python3 -m PyInstaller "${PYI_ARGS[@]}" "$ENTRYPOINT"

# Stage artifact
mkdir -p "$OUT_DIR"
ZIP_NAME="${ZIP_BASENAME}-${VERSION}-macOS.zip"

cd dist
zip -r "$ZIP_NAME" "${APP_NAME}.app"
mv "$ZIP_NAME" "$OUT_DIR/"

echo "✅ Built: $OUT_DIR/$ZIP_NAME"
