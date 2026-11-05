#!/bin/bash
# Generate complex textbox test document — 10 textbox scenarios.
#
# Scenarios 1, 4, 5, 6, 7, 8, 9, 10 use the HIGH-LEVEL `add --type textbox`
# command (fill, border, gradient, rotation, vertical text, geometry, corner
# radius, shadow, no-fill/no-line, wrap, positioning, z-order, plus per-paragraph
# run formatting via set on the inner <textbox>/p[N] paths).
#
# Only scenarios 2 and 3 stay on `raw-set` with pre-authored DrawingML, because
# they exercise surface the high-level command does not expose: per-run mixed
# formatting inside one paragraph (2) and a nested table (3). See textbox.md.

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
OUT="$(dirname "$0")/textbox.docx"

echo "Using CLI: officecli"
echo "Output file: $OUT"

# ==================== Create base document ====================
rm -f "$OUT"
officecli create "$OUT"
officecli add "$OUT" /body --type paragraph --prop text="Complex Textbox Examples" --prop style=Heading1 --prop align=center
officecli add "$OUT" /body --type paragraph --prop text="The following contains multiple complex textbox scenarios for testing textbox behavior under various conditions."

# ==================== Scenario 1: Basic Textbox (with border and background) ====================
# HIGH-LEVEL API. `add --type textbox` creates the floating shape + its first
# paragraph in one call; inner paragraphs are then addressable at
# <textbox-path>/p[N] for run formatting, and more paragraphs are appended with
# `add <textbox-path> --type paragraph`. Paragraph-level format keys are the bare
# forms (bold/italic/color/size/align) — they apply to every run in the paragraph.
officecli add "$OUT" /body --type paragraph --prop text="Scenario 1: Basic Textbox (with border and background)" --prop style=Heading2

TB=$(officecli add "$OUT" /body --type textbox \
  --prop text="Basic Textbox" \
  --prop width=15cm --prop height=3.33cm \
  --prop fill=E6F3FF --prop line.color=0070C0 --prop line.width=2pt \
  --prop wrap=topAndBottom --prop textAnchor=top | grep -oE '/body/textbox\[[0-9]+\]')
officecli set "$OUT" "$TB/p[1]" --prop align=center --prop bold=true --prop color=0070C0 --prop size=14
officecli add "$OUT" "$TB" --type paragraph --prop text="This is a textbox with a blue border and light blue background. It contains a centered title and a normal paragraph."

echo "Done: Scenario 1: Basic Textbox"

# ==================== Scenario 2: Multi-paragraph Rich Text Textbox ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 2: Multi-paragraph Rich Text Textbox" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body/w:sectPr" --action insertbefore --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251660288" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="5400000" cy="2400000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="2" name="TextBox 2"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="5400000" cy="2400000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="FFFDE7"/></a:solidFill>
                    <a:ln w="19050"><a:solidFill><a:srgbClr val="FF8C00"/></a:solidFill><a:prstDash val="dash"/></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="FF8C00"/></w:rPr><w:t>Rich Text Content</w:t></w:r></w:p>
                      <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> | </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r><w:r><w:t xml:space="preserve"> | </w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Underline</w:t></w:r><w:r><w:t xml:space="preserve"> | </w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t>Strikethrough</w:t></w:r></w:p>
                      <w:p><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="20"/></w:rPr><w:t>Red small</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:color w:val="00B050"/><w:sz w:val="36"/></w:rPr><w:t>Green large</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:color w:val="0000FF"/><w:sz w:val="28"/><w:b/><w:i/></w:rPr><w:t>Blue bold italic</w:t></w:r></w:p>
                      <w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Yellow highlight</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:highlight w:val="green"/><w:color w:val="FFFFFF"/></w:rPr><w:t>Green highlight white</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/></w:rPr><w:t>-- Right-aligned quote</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 2: Rich Text Textbox"

# ==================== Scenario 3: Textbox with Nested Table ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 3: Textbox with Nested Table" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body/w:sectPr" --action insertbefore --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251661312" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="5400000" cy="2000000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="3" name="TextBox 3"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="5400000" cy="2000000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>Table inside textbox:</w:t></w:r></w:p>
                      <w:tbl>
                        <w:tblPr>
                          <w:tblStyle w:val="TableGrid"/>
                          <w:tblW w:w="5000" w:type="pct"/>
                          <w:tblBorders>
                            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                          </w:tblBorders>
                        </w:tblPr>
                        <w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>
                        <w:tr>
                          <w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Name</w:t></w:r></w:p></w:tc>
                          <w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Department</w:t></w:r></w:p></w:tc>
                          <w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Score</w:t></w:r></w:p></w:tc>
                        </w:tr>
                        <w:tr>
                          <w:tc><w:p><w:r><w:t>John</w:t></w:r></w:p></w:tc>
                          <w:tc><w:p><w:r><w:t>Engineering</w:t></w:r></w:p></w:tc>
                          <w:tc><w:p><w:r><w:rPr><w:color w:val="00B050"/><w:b/></w:rPr><w:t>95</w:t></w:r></w:p></w:tc>
                        </w:tr>
                        <w:tr>
                          <w:tc><w:p><w:r><w:t>Sarah</w:t></w:r></w:p></w:tc>
                          <w:tc><w:p><w:r><w:t>Marketing</w:t></w:r></w:p></w:tc>
                          <w:tc><w:p><w:r><w:rPr><w:color w:val="FF0000"/><w:b/></w:rPr><w:t>78</w:t></w:r></w:p></w:tc>
                        </w:tr>
                      </w:tbl>
                      <w:p><w:r><w:rPr><w:i/><w:sz w:val="18"/><w:color w:val="888888"/></w:rPr><w:t>* Table nested inside a textbox</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 3: Nested Table"

# ==================== Scenario 4: Rotated Textbox (45 degrees + gradient background) ====================
# HIGH-LEVEL API. rotation= takes degrees (emits a:xfrm rot); fill.gradient= takes
# a comma-separated stop list (NOT the cChart `C1-C2:angle` form). textAnchor=center
# centres the text vertically despite the rotation.
officecli add "$OUT" /body --type paragraph --prop text="Scenario 4: Rotated Textbox (45 degrees)" --prop style=Heading2

TB=$(officecli add "$OUT" /body --type textbox \
  --prop text="Rotated 45" \
  --prop width=6.67cm --prop height=3.33cm \
  --prop rotation=45 \
  --prop fill.gradient=FF6B6B,FFE66D --prop line.color=C0392B --prop line.width=1.5pt \
  --prop wrap=topAndBottom --prop textAnchor=center \
  --prop hRelative=column --prop anchor.x=4.17cm | grep -oE '/body/textbox\[[0-9]+\]')
officecli set "$OUT" "$TB/p[1]" --prop align=center --prop bold=true --prop color=FFFFFF --prop size=14
officecli add "$OUT" "$TB" --type paragraph --prop text="Gradient + Rotation"
officecli set "$OUT" "$TB/p[2]" --prop align=center --prop color=FFFFFF

echo "Done: Scenario 4: Rotated Textbox"

# ==================== Scenario 5: Vertical Text Textbox ====================
# HIGH-LEVEL API. textDirection=eaVert (alias: vert) rotates the text flow to
# East-Asian top-to-bottom. Other values: horz, vert, vert270, wordArtVert.
officecli add "$OUT" /body --type paragraph --prop text="Scenario 5: Vertical Text Textbox" --prop style=Heading2

TB=$(officecli add "$OUT" /body --type textbox \
  --prop text="Vertical text content" \
  --prop width=2.22cm --prop height=6.67cm \
  --prop fill=FFF0F5 --prop line.color=8B0000 --prop line.width=1pt \
  --prop textDirection=eaVert --prop wrap=topAndBottom --prop textAnchor=top | grep -oE '/body/textbox\[[0-9]+\]')
officecli set "$OUT" "$TB/p[1]" --prop align=center --prop bold=true --prop color=8B0000 --prop size=18

echo "Done: Scenario 5: Vertical Textbox"

# ==================== Scenario 6: Rounded Rectangle + Shadow ====================
# HIGH-LEVEL API. cornerRadius sets the roundRect adjust handle (raw guide value
# 16667 ≈ 16.7%); shadow=true emits the standard outer drop shadow (its default
# blur/dist/dir/color/alpha match Word's preset — a compact "blur;dist;dir;color;alpha"
# form is also accepted for custom shadows).
officecli add "$OUT" /body --type paragraph --prop text="Scenario 6: Rounded Rectangle Textbox" --prop style=Heading2

TB=$(officecli add "$OUT" /body --type textbox \
  --prop text="Rounded Rectangle + Shadow" \
  --prop width=15cm --prop height=4.17cm \
  --prop geometry=roundRect --prop cornerRadius=16667 \
  --prop fill=E8F5E9 --prop line.color=2E7D32 --prop line.width=2.25pt \
  --prop shadow=true --prop wrap=topAndBottom --prop textAnchor=center | grep -oE '/body/textbox\[[0-9]+\]')
officecli set "$OUT" "$TB/p[1]" --prop align=center --prop bold=true --prop color=2E7D32 --prop size=15
officecli add "$OUT" "$TB" --type paragraph --prop text="This is a rounded rectangle textbox with an outer shadow effect."
officecli set "$OUT" "$TB/p[2]" --prop align=center
officecli add "$OUT" "$TB" --type paragraph --prop text="Uses geometry=roundRect + cornerRadius for rounded corners"
officecli set "$OUT" "$TB/p[3]" --prop align=center --prop italic=true --prop color=666666

echo "Done: Scenario 6: Rounded Rectangle"

# ==================== Scenario 7: Side-by-side Textboxes (Card Layout) ====================
# HIGH-LEVEL API. Three independent floating boxes with wrap=none, each pinned at
# an absolute horizontal offset (anchor.x) relative to the text column. Note: the
# high-level `add` places each textbox in its own host paragraph, so the three
# cards sit at a slight vertical stagger rather than a single shared baseline (the
# raw-XML original packed all three wp:anchor into one paragraph). For pixel-exact
# co-baseline cards, fall back to raw-set (see the raw scenarios below).
officecli add "$OUT" /body --type paragraph --prop text="Scenario 7: Side-by-side Textboxes (Card Layout)" --prop style=Heading2

add_card() {   # $1=title $2=big $3=label $4=fill $5=accent $6=offset
  local tb
  tb=$(officecli add "$OUT" /body --type textbox \
    --prop text="$1" \
    --prop geometry=roundRect --prop width=4.72cm --prop height=3.89cm \
    --prop fill="$4" --prop line.color="$5" --prop line.width=1pt \
    --prop wrap=none --prop hRelative=column --prop anchor.x="$6" --prop textAnchor=center | grep -oE '/body/textbox\[[0-9]+\]')
  officecli set "$OUT" "$tb/p[1]" --prop align=center --prop bold=true --prop color="$5" --prop size=14
  officecli add "$OUT" "$tb" --type paragraph --prop text="$2"
  officecli set "$OUT" "$tb/p[2]" --prop align=center --prop size=24
  officecli add "$OUT" "$tb" --type paragraph --prop text="$3"
  officecli set "$OUT" "$tb/p[3]" --prop align=center --prop color=888888 --prop size=9
}
add_card "Card A" "128"   "Daily Visits" E3F2FD 1565C0 0cm
add_card "Card B" "56"    "New Orders"   FFF3E0 E65100 5.28cm
add_card "Card C" "99.8%" "Uptime"       E8F5E9 2E7D32 10.56cm

echo "Done: Scenario 7: Side-by-side Cards"

# ==================== Scenario 8: Borderless Transparent Textbox ====================
# HIGH-LEVEL API. fill=none + line.color=none make a fully invisible container —
# only the text shows. (Both sentinels emit a:noFill; before that fix a literal
# "none" was rejected by the color parser and this box needed raw-set.)
officecli add "$OUT" /body --type paragraph --prop text="Scenario 8: Borderless Transparent Textbox" --prop style=Heading2

TB=$(officecli add "$OUT" /body --type textbox \
  --prop text="Borderless transparent text" \
  --prop width=11.11cm --prop height=2.22cm \
  --prop fill=none --prop line.color=none \
  --prop hRelative=column --prop anchor.x=1.39cm \
  --prop wrap=topAndBottom --prop textAnchor=center | grep -oE '/body/textbox\[[0-9]+\]')
officecli set "$OUT" "$TB/p[1]" --prop align=center --prop italic=true --prop size=22 --prop color=AAAAAA

echo "Done: Scenario 8: Transparent Textbox"

# ==================== Scenario 9: Text Overflow Textbox ====================
# HIGH-LEVEL API. A short fixed height with more text than fits — autoFit is
# deliberately omitted so the box does NOT grow (omitting autoFit = fixed height;
# --prop autoFit=true would emit a:spAutoFit and resize to content instead).
officecli add "$OUT" /body --type paragraph --prop text="Scenario 9: Text Overflow Textbox" --prop style=Heading2

TB=$(officecli add "$OUT" /body --type textbox \
  --prop text="Line 1: This is a fixed-height textbox with too much text to test overflow behavior." \
  --prop width=15cm --prop height=1.67cm \
  --prop fill=FCE4EC --prop line.color=C62828 --prop line.width=1pt \
  --prop wrap=topAndBottom --prop textAnchor=top | grep -oE '/body/textbox\[[0-9]+\]')
officecli set "$OUT" "$TB/p[1]" --prop bold=true --prop color=C62828
officecli add "$OUT" "$TB" --type paragraph --prop text="Line 2: In real usage, the textbox height is limited but content can be long."
officecli add "$OUT" "$TB" --type paragraph --prop text="Line 3: Word usually auto-expands the textbox height, but fixed height may truncate."
officecli add "$OUT" "$TB" --type paragraph --prop text="Line 4: This line may be truncated or overflow, depending on bodyPr settings."
officecli add "$OUT" "$TB" --type paragraph --prop text="Line 5: Continuing to test more overflow content..."
officecli add "$OUT" "$TB" --type paragraph --prop text="Line 6: Final overflow line."

echo "Done: Scenario 9: Overflow Textbox"

# ==================== Scenario 10: Textbox Stacking (Z-order) ====================
# HIGH-LEVEL API. behindDoc=true pushes the bottom box behind the body text;
# relativeHeight sets the stacking order (higher = front). The top box uses
# fill.opacity for its translucent (80%) fill so the bottom box shows through.
officecli add "$OUT" /body --type paragraph --prop text="Scenario 10: Textbox Stacking (Z-order)" --prop style=Heading2

# Bottom layer — behind the document text, lower relativeHeight.
TB=$(officecli add "$OUT" /body --type textbox \
  --prop text="Bottom layer (behindDoc)" \
  --prop width=8.33cm --prop height=4.17cm \
  --prop fill=BBDEFB --prop line.color=1565C0 --prop line.width=1.5pt \
  --prop wrap=none --prop hRelative=column --prop anchor.x=0.56cm \
  --prop behindDoc=true --prop relativeHeight=251670528 --prop textAnchor=top | grep -oE '/body/textbox\[[0-9]+\]')
officecli set "$OUT" "$TB/p[1]" --prop bold=true --prop color=1565C0 --prop size=14
officecli add "$OUT" "$TB" --type paragraph --prop text="This textbox is behind the document content."
officecli add "$OUT" "$TB" --type paragraph --prop text="It should be partially obscured by the top layer textbox."

# Top layer — in front (higher relativeHeight), translucent 80% fill so the bottom shows through.
TB=$(officecli add "$OUT" /body --type textbox \
  --prop text="Top layer (translucent)" \
  --prop width=8.33cm --prop height=3.33cm \
  --prop fill=FFCDD2 --prop fill.opacity=80 --prop line.color=C62828 --prop line.width=1.5pt \
  --prop wrap=none --prop hRelative=column --prop anchor.x=3.33cm \
  --prop vRelative=paragraph --prop anchor.y=1.11cm \
  --prop relativeHeight=251671552 --prop textAnchor=top | grep -oE '/body/textbox\[[0-9]+\]')
officecli set "$OUT" "$TB/p[1]" --prop bold=true --prop color=C62828 --prop size=14
officecli add "$OUT" "$TB" --type paragraph --prop text="This textbox is on top, 80% opacity."
officecli add "$OUT" "$TB" --type paragraph --prop text="It partially obscures the bottom blue textbox."

echo "Done: Scenario 10: Z-order Stacking"

# ==================== Verification ====================
officecli close "$OUT"
echo ""
echo "=========================================="
echo "Document generated: $OUT"
echo "=========================================="
officecli view "$OUT" outline
echo ""
officecli validate "$OUT"
