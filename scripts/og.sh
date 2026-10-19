#!/bin/sh
# Renders site/og.png (1200x630) and app icons with headless Chrome. Run after build.mjs.
set -e
cd "$(dirname "$0")/.."
CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
"$CHROME" --headless=new --disable-gpu --hide-scrollbars --virtual-time-budget=4000 --window-size=1200,630 --screenshot="$PWD/site/og.png" "file://$PWD/scripts/og.html" >/dev/null 2>&1
for size in 180 512; do
  printf '<html><body style="margin:0;background:transparent"><img src="file://%s/site/favicon.svg" width="%s" height="%s"></body></html>' "$PWD" $size $size > /tmp/icon-$size.html
  out=site/icon-$size.png; [ $size = 180 ] && out=site/apple-touch-icon.png
  "$CHROME" --headless=new --disable-gpu --hide-scrollbars --default-background-color=00000000 --window-size=$size,$size --screenshot="$PWD/$out" "file:///tmp/icon-$size.html" >/dev/null 2>&1
done
ls -la site/og.png site/apple-touch-icon.png site/icon-512.png
