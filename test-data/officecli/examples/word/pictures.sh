#!/bin/bash
# pictures.sh — embed and lay out images in a Word document (.docx).
#
# Exercises the docx `picture` element (schemas/help/docx/picture.json): inline
# vs floating (anchored) pictures, crop, alt text, text-wrap modes, behind-text
# watermark, relative alignment, absolute position, sizing, and click links.
# CLI twin of pictures.py (officecli Python SDK); both produce an equivalent
# pictures.docx.
#
# In docx a picture is a run inside a paragraph, so every picture is added to a
# paragraph path (/body/p[@paraId=X]); the picture's own path is that paragraph
# plus a run index (/body/p[@paraId=X]/r[N]). Floating layout (wrap / behindText
# / hAlign / vAlign / hPosition / vPosition) requires --prop anchor=true; inline
# pictures sit in the text flow like a big character.
#
# Requirements: Pillow (pip install Pillow) to synthesize the sample PNGs.
# Usage: ./pictures.sh [officecli path]
#
# NOTE: intentionally NO `set -e`. Like the SDK twin's batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
DIR="$(dirname "$0")"
FILE="$DIR/pictures.docx"
LOGO="$DIR/pictures-logo.png"
BANNER="$DIR/pictures-banner.png"

echo "Building $FILE ..."

# ── Synthesize two sample PNGs (a square logo + a wide banner) in-dir ──────────
python3 - "$LOGO" "$BANNER" <<'PY'
import sys
from PIL import Image, ImageDraw

logo, banner = sys.argv[1:3]

def make_logo(path, w=300, h=300):
    img = Image.new("RGB", (w, h), (245, 245, 220)); d = ImageDraw.Draw(img)
    d.ellipse((40, 40, 260, 260), fill=(52, 152, 219), outline=(0, 0, 0), width=4)
    d.polygon([(150, 80), (210, 220), (90, 220)], fill=(241, 196, 15), outline=(0, 0, 0))
    d.text((110, 140), "LOGO", fill=(0, 0, 0))
    img.save(path)

def make_banner(path, w=800, h=200, c1=(231, 76, 60), c2=(142, 68, 173)):
    img = Image.new("RGB", (w, h)); pix = img.load()
    for x in range(w):
        t = x / (w - 1)
        col = tuple(int(c1[i] * (1 - t) + c2[i] * t) for i in range(3))
        for y in range(h):
            pix[x, y] = col
    ImageDraw.Draw(img).text((20, 20), "banner.png", fill=(255, 255, 255))
    img.save(path)

make_logo(logo); make_banner(banner)
PY

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ══════════════════════════════════════════════════════════════════════════════
# 1. Inline picture — sits in the text flow like a large character
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" /body --type paragraph --prop text="1. Inline Picture" --prop style=Heading1
$CLI add "$FILE" /body --type paragraph \
  --prop text="An inline picture flows with the paragraph text — no anchor, no wrap; it occupies its own line box like an oversized glyph:"
$CLI add "$FILE" /body --type paragraph --prop text=""
# Features: inline picture (default, no anchor); width / height sizing (unit-qualified)
$CLI add "$FILE" "/body/p[3]" --type picture \
  --prop src="$LOGO" \
  --prop width=3cm --prop height=3cm

# ══════════════════════════════════════════════════════════════════════════════
# 2. Cropped picture — trim edges via crop=L,T,R,B (percent)
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" /body --type paragraph --prop text="2. Cropped Picture" --prop style=Heading1
$CLI add "$FILE" /body --type paragraph \
  --prop text="crop=L,T,R,B trims each edge by a percentage of the source. Here 10% left, 5% top, 15% right, 8% bottom (per-edge cropLeft/cropTop/cropRight/cropBottom also accepted on add):"
$CLI add "$FILE" /body --type paragraph --prop text=""
# Features: crop=L,T,R,B four-value form (percent of original per edge)
$CLI add "$FILE" "/body/p[6]" --type picture \
  --prop src="$BANNER" \
  --prop crop=10,5,15,8 \
  --prop width=10cm --prop height=2.5cm

# ══════════════════════════════════════════════════════════════════════════════
# 3. Picture with alt text — accessibility / screen readers
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" /body --type paragraph --prop text="3. Alt Text (Accessibility)" --prop style=Heading1
$CLI add "$FILE" /body --type paragraph \
  --prop text="alt= writes the DocProperties description read aloud by screen readers. Aliases: altText, description."
$CLI add "$FILE" /body --type paragraph --prop text=""
# Features: alt (alternative text for accessibility)
$CLI add "$FILE" "/body/p[9]" --type picture \
  --prop src="$LOGO" \
  --prop width=3cm --prop height=3cm \
  --prop alt="Company logo: a blue circle enclosing a yellow triangle"

# ══════════════════════════════════════════════════════════════════════════════
# 4. Behind-text watermark — floating picture behind the paragraph text
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" /body --type paragraph --prop text="4. Behind-Text Watermark" --prop style=Heading1
$CLI add "$FILE" /body --type paragraph \
  --prop text="A floating picture with anchor=true, wrap=none and behindText=true sits behind the text like a watermark. It is centered on the page margins via hAlign=center + vAlign=center. This paragraph text should render on top of the faint image behind it, demonstrating the behind-text z-order stacking that a plain inline picture cannot achieve."
# Features: anchor=true (floating), wrap=none + behindText=true (watermark), hAlign/vAlign=center
$CLI add "$FILE" "/body/p[11]" --type picture \
  --prop src="$BANNER" \
  --prop anchor=true --prop wrap=none --prop behindText=true \
  --prop hAlign=center --prop vAlign=center \
  --prop hRelative=margin --prop vRelative=margin \
  --prop width=12cm --prop height=3cm \
  --prop alt="Decorative watermark banner"

# ══════════════════════════════════════════════════════════════════════════════
# 5. Square text-wrap — body text flows around a floating picture
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" /body --type paragraph --prop text="5. Square Text Wrap" --prop style=Heading1
$CLI add "$FILE" /body --type paragraph \
  --prop text="With anchor=true and wrap=square, the surrounding paragraph text flows around the picture's bounding box. The picture below is right-aligned to the margin, so this long paragraph wraps down its left side. Keep reading to see the text reflow around the floated image on the right — square wrap uses a rectangular boundary regardless of the image's own shape, so text keeps a clean vertical edge against the picture. The remaining lines continue underneath once the text clears the bottom of the anchored picture's bounding rectangle."
# Features: wrap=square (text flows around bounding box), hAlign=right relative to margin
$CLI add "$FILE" "/body/p[13]" --type picture \
  --prop src="$LOGO" \
  --prop anchor=true --prop wrap=square \
  --prop hAlign=right --prop hRelative=margin --prop vRelative=paragraph \
  --prop width=3.5cm --prop height=3.5cm \
  --prop alt="Logo floated right with square wrap"

# ══════════════════════════════════════════════════════════════════════════════
# 6. Tight wrap + absolute position — hPosition / vPosition offsets
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" /body --type paragraph --prop text="6. Absolute Position (hPosition / vPosition)" --prop style=Heading1
$CLI add "$FILE" /body --type paragraph \
  --prop text="Instead of relative alignment, a floating picture can be pinned to an absolute offset from its reference frame. Here hPosition=2cm and vPosition=1cm place the picture 2cm from the left margin and 1cm down, with wrap=tight so text hugs the boundary. This paragraph provides enough text for the wrap to be visible against the absolutely-positioned image."
# Features: anchor=true, wrap=tight, hPosition/vPosition (absolute EMU offsets, unit-qualified)
$CLI add "$FILE" "/body/p[15]" --type picture \
  --prop src="$LOGO" \
  --prop anchor=true --prop wrap=tight \
  --prop hPosition=2cm --prop vPosition=1cm \
  --prop hRelative=margin --prop vRelative=paragraph \
  --prop width=3cm --prop height=3cm \
  --prop alt="Logo at absolute 2cm,1cm offset with tight wrap"

# ══════════════════════════════════════════════════════════════════════════════
# 7. Clickable picture — link= makes the image a hyperlink
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" /body --type paragraph --prop text="7. Clickable Picture (link)" --prop style=Heading1
$CLI add "$FILE" /body --type paragraph \
  --prop text="link= wraps the picture in a click hyperlink. An absolute URL round-trips as an external relationship; a #anchor or bookmark name becomes an internal jump."
$CLI add "$FILE" /body --type paragraph --prop text=""
# Features: link (external URL hyperlink on the image); alt on a clickable image
$CLI add "$FILE" "/body/p[18]" --type picture \
  --prop src="$BANNER" \
  --prop width=10cm --prop height=2.5cm \
  --prop link="https://example.com" \
  --prop alt="Banner linking to example.com"

# ══════════════════════════════════════════════════════════════════════════════
# 8. Decorative picture — mark as decorative for accessibility
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" /body --type paragraph --prop text="8. Decorative Picture (accessibility)" --prop style=Heading1
$CLI add "$FILE" /body --type paragraph \
  --prop text="decorative=true marks the image as decorative: screen readers skip it entirely (no alt text is announced). Stored as an adec:decorative extension under the picture's docPr. Use it for purely ornamental images that carry no information."
$CLI add "$FILE" /body --type paragraph --prop text=""
# Features: decorative=true (accessibility — screen readers skip the image)
$CLI add "$FILE" "/body/p[21]" --type picture \
  --prop src="$BANNER" \
  --prop width=10cm --prop height=2.5cm \
  --prop decorative=true

$CLI close "$FILE"

$CLI validate "$FILE"
echo "Generated: $FILE"
