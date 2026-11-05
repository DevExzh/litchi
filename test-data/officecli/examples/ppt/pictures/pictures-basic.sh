#!/bin/bash
# Basic PowerPoint pictures — embed images, position/resize, crop, rotate, hyperlink.
#
# CLI twin of pictures-basic.py (officecli Python SDK). Both produce an
# equivalent pictures-basic.pptx. This one drives the officecli binary directly:
# one `officecli ... --prop k=v` invocation per element.
#
#   - slide 1: src= file vs data-URI (ways to supply an image)
#   - slide 2: crop variants — symmetric, vertical/horizontal, per-edge
#   - slide 3: rotation
#   - slide 4: hyperlinks (click-to-open URL / jump to slide / next-slide action)
#   - slide 5: Set-only effects — brightness / contrast / glow / shadow
#
# Requirements: Pillow (pip install Pillow) to synthesize the sample PNGs.
# Usage: ./pictures-basic.sh [officecli path]

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/pictures-basic.pptx"

WORKDIR="$(mktemp -d -t ocli-pics-XXXXXX)"
trap 'rm -rf "$WORKDIR"' EXIT

GRAD="$WORKDIR/gradient.png"
GEO="$WORKDIR/geometric.png"
PHOTO="$WORKDIR/photo.png"

# ── Synthesize the three sample PNGs (gradient, geometric, photo-like) ─────────
python3 - "$GRAD" "$GEO" "$PHOTO" <<'PY'
import sys
from PIL import Image, ImageDraw

grad, geo, photo = sys.argv[1:4]

def make_gradient(path, w=400, h=300, c1=(231, 76, 60), c2=(52, 152, 219)):
    img = Image.new("RGB", (w, h)); pix = img.load()
    for y in range(h):
        t = y / (h - 1)
        r = int(c1[0] * (1 - t) + c2[0] * t)
        g = int(c1[1] * (1 - t) + c2[1] * t)
        b = int(c1[2] * (1 - t) + c2[2] * t)
        for x in range(w):
            pix[x, y] = (r, g, b)
    ImageDraw.Draw(img).text((20, 20), "gradient.png", fill=(255, 255, 255))
    img.save(path)

def make_geometric(path, w=400, h=300):
    img = Image.new("RGB", (w, h), (245, 245, 220)); d = ImageDraw.Draw(img)
    d.ellipse((50, 50, 180, 180), fill=(231, 76, 60), outline=(0, 0, 0), width=3)
    d.rectangle((200, 80, 350, 220), fill=(52, 152, 219), outline=(0, 0, 0), width=3)
    d.polygon([(120, 200), (60, 270), (180, 270)], fill=(241, 196, 15), outline=(0, 0, 0))
    d.text((10, 10), "geometric.png", fill=(0, 0, 0))
    img.save(path)

def make_photo(path, w=400, h=300):
    img = Image.new("RGB", (w, h)); cx, cy = w / 2, h / 2
    maxd = (cx ** 2 + cy ** 2) ** 0.5; pix = img.load()
    for y in range(h):
        for x in range(w):
            dd = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5 / maxd
            pix[x, y] = (int(255 * (1 - dd * 0.7)), int(180 * (1 - dd * 0.5)), int(80 * (1 - dd * 0.3)))
    ImageDraw.Draw(img).text((10, 10), "photo.png", fill=(255, 255, 255))
    img.save(path)

make_gradient(grad); make_geometric(geo); make_photo(photo)
PY

# data-URI form for the geometric image (slide 1b)
GEO_URI="data:image/png;base64,$(base64 < "$GEO" | tr -d '\n')"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ══════════════════════════════════════════════════════════════════════════════
# Slide 1: three src= forms (file path / data-URI)
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" / --type slide
$CLI add "$FILE" "/slide[1]" --type textbox \
    --prop text="Three ways to supply src= (file path / data-URI)" \
    --prop size=24 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# 1a. File path
$CLI add "$FILE" "/slide[1]" --type picture \
    --prop src="$GRAD" \
    --prop x=0.5in --prop y=1.3in --prop width=3.5in --prop height=2.6in \
    --prop alt="gradient image from disk"
$CLI add "$FILE" "/slide[1]" --type textbox \
    --prop text="src=<file path>" \
    --prop size=12 --prop italic=true \
    --prop x=0.5in --prop y=4in --prop width=3.5in --prop height=0.4in

# 1b. data-URI
$CLI add "$FILE" "/slide[1]" --type picture \
    --prop src="$GEO_URI" \
    --prop x=4.5in --prop y=1.3in --prop width=3.5in --prop height=2.6in \
    --prop alt="geometric shapes embedded as data-URI"
$CLI add "$FILE" "/slide[1]" --type textbox \
    --prop text="src=data:image/png;base64,..." \
    --prop size=12 --prop italic=true \
    --prop x=4.5in --prop y=4in --prop width=3.5in --prop height=0.4in

# 1c. Another file (use the photo)
$CLI add "$FILE" "/slide[1]" --type picture \
    --prop src="$PHOTO" \
    --prop x=8.5in --prop y=1.3in --prop width=3.5in --prop height=2.6in \
    --prop alt="pseudo-photo gradient" \
    --prop name="hero-photo" \
    --prop compressionState=print
$CLI add "$FILE" "/slide[1]" --type textbox \
    --prop text='src=<file> + name="hero-photo" + compressionState=print' \
    --prop size=12 --prop italic=true \
    --prop x=8.5in --prop y=4in --prop width=3.5in --prop height=0.4in

# ══════════════════════════════════════════════════════════════════════════════
# Slide 2: crop variants
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" / --type slide
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="Crop — symmetric / vertical,horizontal / per-edge" \
    --prop size=24 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# Original (uncropped reference)
$CLI add "$FILE" "/slide[2]" --type picture \
    --prop src="$GEO" \
    --prop x=0.5in --prop y=1.3in --prop width=3in --prop height=2.2in
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="original (no crop)" --prop size=12 \
    --prop x=0.5in --prop y=3.6in --prop width=3in --prop height=0.4in

# crop=20 — symmetric all edges
$CLI add "$FILE" "/slide[2]" --type picture \
    --prop src="$GEO" --prop crop=20 \
    --prop x=4in --prop y=1.3in --prop width=3in --prop height=2.2in
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="crop=20  (20% off each edge)" --prop size=12 \
    --prop x=4in --prop y=3.6in --prop width=3in --prop height=0.4in

# crop=10,30 — vertical 10%, horizontal 30%
$CLI add "$FILE" "/slide[2]" --type picture \
    --prop src="$GEO" --prop crop=10,30 \
    --prop x=7.5in --prop y=1.3in --prop width=3in --prop height=2.2in
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="crop=10,30  (10% top/bot, 30% left/right)" --prop size=12 \
    --prop x=7.5in --prop y=3.6in --prop width=3.5in --prop height=0.4in

# Per-edge: cropLeft + cropTop
$CLI add "$FILE" "/slide[2]" --type picture \
    --prop src="$GEO" \
    --prop cropLeft=25 --prop cropTop=25 \
    --prop x=0.5in --prop y=4.3in --prop width=3in --prop height=2.2in
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="cropLeft=25 + cropTop=25" --prop size=12 \
    --prop x=0.5in --prop y=6.6in --prop width=3in --prop height=0.4in

# 4-value crop: left,top,right,bottom
$CLI add "$FILE" "/slide[2]" --type picture \
    --prop src="$GEO" --prop crop=5,10,40,20 \
    --prop x=4in --prop y=4.3in --prop width=3in --prop height=2.2in
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="crop=5,10,40,20  (L,T,R,B)" --prop size=12 \
    --prop x=4in --prop y=6.6in --prop width=3in --prop height=0.4in

# ══════════════════════════════════════════════════════════════════════════════
# Slide 3: rotation
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" / --type slide
$CLI add "$FILE" "/slide[3]" --type textbox \
    --prop text="Rotation — degrees clockwise" \
    --prop size=24 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# positions: x y degrees ; ylabel = y + 2.3
for spec in "0.5 1.5 0 3.8" "4.5 1.5 30 3.8" "8.5 1.5 90 3.8" \
            "0.5 4.5 180 6.8" "4.5 4.5 270 6.8" "8.5 4.5 -45 6.8"; do
    set -- $spec
    X="$1"; Y="$2"; DEG="$3"; YLAB="$4"
    $CLI add "$FILE" "/slide[3]" --type picture \
        --prop src="$GEO" \
        --prop x="${X}in" --prop y="${Y}in" --prop width=3in --prop height=2.2in \
        --prop rotation="$DEG"
    $CLI add "$FILE" "/slide[3]" --type textbox \
        --prop text="rotation=$DEG" --prop size=12 \
        --prop x="${X}in" --prop y="${YLAB}in" --prop width=3in --prop height=0.4in
done

# ══════════════════════════════════════════════════════════════════════════════
# Slide 4: clickable hyperlinks on pictures
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" / --type slide
$CLI add "$FILE" "/slide[4]" --type textbox \
    --prop text="Clickable Pictures — link= and tooltip=" \
    --prop size=24 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# External URL
$CLI add "$FILE" "/slide[4]" --type picture \
    --prop src="$GRAD" \
    --prop x=0.5in --prop y=1.5in --prop width=3.5in --prop height=2.6in \
    --prop link="https://example.com" \
    --prop tooltip="Open example.com"
$CLI add "$FILE" "/slide[4]" --type textbox \
    --prop text="link=https://example.com" --prop size=12 \
    --prop x=0.5in --prop y=4.2in --prop width=3.5in --prop height=0.4in

# In-deck slide jump
$CLI add "$FILE" "/slide[4]" --type picture \
    --prop src="$GEO" \
    --prop x=4.5in --prop y=1.5in --prop width=3.5in --prop height=2.6in \
    --prop link="slide[1]" \
    --prop tooltip="Back to slide 1"
$CLI add "$FILE" "/slide[4]" --type textbox \
    --prop text="link=slide[1]  (jump to slide 1)" --prop size=12 \
    --prop x=4.5in --prop y=4.2in --prop width=3.5in --prop height=0.4in

# Named action: nextslide
$CLI add "$FILE" "/slide[4]" --type picture \
    --prop src="$PHOTO" \
    --prop x=8.5in --prop y=1.5in --prop width=3.5in --prop height=2.6in \
    --prop link="nextslide" \
    --prop tooltip="Advance one slide"
$CLI add "$FILE" "/slide[4]" --type textbox \
    --prop text="link=nextslide  (named action)" --prop size=12 \
    --prop x=8.5in --prop y=4.2in --prop width=3.5in --prop height=0.4in

# ══════════════════════════════════════════════════════════════════════════════
# Slide 5: Set-only effects — brightness / contrast / glow / shadow
# These four props are schema-declared add:false / set:true. Pattern: Add the
# picture, capture its DOM path from the "Added picture at ..." message, then
# Set the effect. Also exercises cropBottom / cropRight by their named form.
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" / --type slide
$CLI add "$FILE" "/slide[5]" --type textbox \
    --prop text="Picture effects (Set-only) — brightness / contrast / glow / shadow" \
    --prop size=24 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=13in --prop height=0.6in

# add_pic <x> <y> [extra --prop args...] ; echoes the picture's DOM path
add_pic() {
    local x="$1" y="$2"; shift 2
    local out
    out=$($CLI add "$FILE" "/slide[5]" --type picture \
        --prop src="$PHOTO" \
        --prop x="${x}in" --prop y="${y}in" --prop width=2.8in --prop height=2.1in \
        "$@")
    # "Added picture at /slide[5]/shape[@id=...]" → last token
    echo "$out" | grep "Added picture at" | tail -1 | awk '{print $NF}'
}

label() {  # label <x> <y> <text>
    $CLI add "$FILE" "/slide[5]" --type textbox \
        --prop text="$3" --prop size=11 --prop italic=true \
        --prop x="${1}in" --prop y="${2}in" --prop width=2.8in --prop height=0.4in
}

# Reference (untouched)
add_pic 0.5 1.2 >/dev/null
label 0.5 3.4 "(reference)"

# brightness +40 — lifts mid-tones
P_BRIGHT=$(add_pic 3.6 1.2)
$CLI set "$FILE" "$P_BRIGHT" --prop brightness=40
label 3.6 3.4 "brightness=40"

# contrast -30 — flattens
P_CON=$(add_pic 6.7 1.2)
$CLI set "$FILE" "$P_CON" --prop contrast=-30
label 6.7 3.4 "contrast=-30"

# brightness + contrast together
P_COMBO=$(add_pic 9.8 1.2)
$CLI set "$FILE" "$P_COMBO" --prop brightness=-20 --prop contrast=40
label 9.8 3.4 "brightness=-20 + contrast=40"

# glow — color-radius-opacity
P_GLOW=$(add_pic 0.5 4.2)
$CLI set "$FILE" "$P_GLOW" --prop glow=FFD700-12-75
label 0.5 6.4 "glow=FFD700-12-75"

# shadow — color-blur-angle-dist-opacity
P_SHADOW=$(add_pic 3.6 4.2)
$CLI set "$FILE" "$P_SHADOW" --prop shadow=000000-10-45-6-50
label 3.6 6.4 "shadow=000000-10-45-6-50"

# cropRight + cropBottom — by-name form (vs the 4-value crop=)
add_pic 6.7 4.2 --prop cropRight=25 --prop cropBottom=15 >/dev/null
label 6.7 6.4 "cropRight=25 + cropBottom=15"

# Everything together: trim corners + brightness + contrast + glow + shadow
P_ALL=$(add_pic 9.8 4.2 --prop cropLeft=10 --prop cropTop=10 --prop cropRight=10 --prop cropBottom=10)
$CLI set "$FILE" "$P_ALL" \
    --prop brightness=15 \
    --prop contrast=20 \
    --prop glow=4472C4-8-60 \
    --prop shadow=000000-6-135-3-40
label 9.8 6.4 "trimmed + bright + contrast + glow + shadow"

$CLI close "$FILE"

$CLI validate "$FILE"
echo "Generated: $FILE"
