#!/bin/bash
# assets/icon.png -> assets/icon.icns  (yalnızca macOS'ta gerekir)
set -euo pipefail
cd "$(dirname "$0")/.."
ICONSET=$(mktemp -d)/icon.iconset
mkdir -p "$ICONSET"
for size in 16 32 128 256 512; do
  sips -z $size $size assets/icon.png --out "$ICONSET/icon_${size}x${size}.png" >/dev/null
  sips -z $((size*2)) $((size*2)) assets/icon.png --out "$ICONSET/icon_${size}x${size}@2x.png" >/dev/null
done
iconutil -c icns "$ICONSET" -o assets/icon.icns
echo "✓ assets/icon.icns"
