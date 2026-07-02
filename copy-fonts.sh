#!/bin/bash
# copy-fonts.sh – Copies David font from macOS to the fonts/ directory.
# Run this once before `docker compose build`.
# If running on a different Mac, run this script on that Mac.

set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
FONT_DIR="$SCRIPT_DIR/fonts"
mkdir -p "$FONT_DIR"

SOURCES=(
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/david.ttf"
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/davidbd.ttf"
    "/Library/Fonts/David.ttf"
    "/Library/Fonts/David Bold.ttf"
    "$HOME/Library/Fonts/David.ttf"
    "$HOME/Library/Fonts/David Bold.ttf"
)

found=0
for src in "${SOURCES[@]}"; do
    if [ -f "$src" ]; then
        cp "$src" "$FONT_DIR/"
        echo "✓ $(basename "$src")"
        found=1
    fi
done

if [ "$found" -eq 0 ]; then
    echo "David font not found. Install Microsoft Office or copy david.ttf manually to:"
    echo "  $FONT_DIR/"
    exit 1
fi

echo ""
echo "Done. Now run: docker compose build gotenberg"
