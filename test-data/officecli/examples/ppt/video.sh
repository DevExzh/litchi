#!/bin/bash
# Video Showcase — generates video.pptx with an embedded MP4 (poster, sizing,
# volume, autoplay, loop, trim) across four slides.
# CLI twin of video.py (officecli Python SDK). Both produce an equivalent video.pptx.
#
# Requirements (for the generated media): pip install imageio imageio-ffmpeg numpy
# Usage: ./video.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
DIR="$(dirname "$0")"
FILE="$DIR/video.pptx"

# --- Generate the demo video + poster into a temp dir (mirrors video.py) ------
TMP="$(mktemp -d "${TMPDIR:-/tmp}/officecli_video_XXXXXX")"
VIDEO="$TMP/demo.mp4"
COVER="$TMP/cover.png"
trap 'rm -rf "$TMP"' EXIT

echo "[1/3] Generating video and cover image..."
python3 - "$VIDEO" "$COVER" <<'PY'
import sys
try:
    import imageio.v3 as iio
    import numpy as np
except ImportError:
    print("ERROR: imageio not installed. Run: pip install imageio imageio-ffmpeg numpy")
    sys.exit(1)

video_path, cover_path = sys.argv[1], sys.argv[2]
W, H, FPS, DURATION = 640, 360, 30, 3
total_frames = FPS * DURATION
frames = []
for i in range(total_frames):
    t = i / (total_frames - 1)
    frame = np.zeros((H, W, 3), dtype=np.uint8)
    for y in range(H):
        yf = y / H
        r = int(20 + 60 * t + 40 * yf)
        g = int(30 + 80 * (1 - abs(t - 0.5) * 2) * (1 - yf))
        b = int(80 + 120 * (1 - t) + 50 * yf)
        frame[y, :, 0] = min(r, 255)
        frame[y, :, 1] = min(g, 255)
        frame[y, :, 2] = min(b, 255)
    cx = int(100 + t * (W - 200)); cy = H // 2; radius = 40
    yy, xx = np.ogrid[:H, :W]
    mask = (xx - cx) ** 2 + (yy - cy) ** 2 < radius ** 2
    frame[mask, 0] = 255; frame[mask, 1] = 200; frame[mask, 2] = 50
    for row in range(3):
        bar_y = 60 + row * 50
        bar_w = int(200 + 100 * (1 - abs(t - 0.5) * 2))
        frame[bar_y:bar_y + 12, 50:50 + bar_w, :] = [200, 200, 220]
    frames.append(frame)
iio.imwrite(video_path, frames, fps=FPS)
iio.imwrite(cover_path, frames[0])
PY

echo "[2/3] Building presentation: $FILE"
rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ---- Slide 1: Title slide with gradient background ----
$CLI add "$FILE" / --type slide --prop layout=title
$CLI set "$FILE" /slide[1] --prop background=radial:1B2838-4472C4-bl
$CLI set "$FILE" /slide[1]/placeholder[ctrTitle] --prop text="Video Demo" --prop color=FFFFFF --prop size=44
$CLI set "$FILE" /slide[1]/placeholder[subTitle] --prop text="Embedded video with officecli" --prop color=B4C7E7 --prop size=20

# ---- Slide 2: Video slide ----
$CLI add "$FILE" / --type slide --prop title="Animated Video"
$CLI set "$FILE" /slide[2] --prop background=0D1B2A
$CLI set "$FILE" /slide[2]/shape[1] --prop color=FFFFFF
$CLI add "$FILE" /slide[2] --type video \
    --prop src="$VIDEO" \
    --prop poster="$COVER" \
    --prop x=2cm --prop y=4cm --prop width=22cm --prop height=12.5cm \
    --prop volume=80 --prop autoplay=true

# ---- Slide 3: Video info with chart ----
$CLI add "$FILE" / --type slide --prop title="Video Properties"
$CLI set "$FILE" /slide[3] --prop background=1B2838
$CLI set "$FILE" /slide[3]/shape[1] --prop color=FFFFFF
$CLI add "$FILE" /slide[3] --type shape \
    --prop text="Resolution: 640x360\nFPS: 30\nDuration: 3s\nFormat: MP4" \
    --prop font=Consolas --prop size=16 --prop color=B4C7E7 \
    --prop x=1cm --prop y=4cm --prop width=10cm --prop height=6cm \
    --prop fill=0D1B2A --prop line=4472C4 --prop linewidth=1pt
$CLI add "$FILE" /slide[3] --type chart \
    --prop chartType=bar --prop title="Frame Colors" \
    --prop categories="Red,Green,Blue" \
    --prop "series1=Start:20,30,200" \
    --prop "series2=End:80,30,80" \
    --prop colors=E74C3C,27AE60 \
    --prop x=13cm --prop y=4cm --prop width=12cm --prop height=8cm

# ---- Slide 4: loop / trimStart / trimEnd ----
$CLI add "$FILE" / --type slide --prop title="loop / trimStart / trimEnd"
$CLI set "$FILE" /slide[4] --prop background=0D1B2A
$CLI set "$FILE" /slide[4]/shape[1] --prop color=FFFFFF
# loop=true — video restarts after it reaches the end
# trimStart / trimEnd — play only a sub-range of the video (seconds)
$CLI add "$FILE" /slide[4] --type video \
    --prop src="$VIDEO" \
    --prop poster="$COVER" \
    --prop x=2cm --prop y=4cm --prop width=22cm --prop height=12.5cm \
    --prop volume=60 --prop autoplay=true \
    --prop loop=true --prop trimStart=0 --prop trimEnd=2
$CLI add "$FILE" /slide[4] --type shape \
    --prop text="loop=true  trimStart=0  trimEnd=2\nVideo loops continuously; playback is clipped to the 0–2s range." \
    --prop size=14 --prop color=B4C7E7 \
    --prop x=1cm --prop y=17cm --prop width=24cm --prop height=2cm

$CLI close "$FILE"

echo "[3/3] Validating..."
$CLI validate "$FILE"
echo "Generated: $FILE"
