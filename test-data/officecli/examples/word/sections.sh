#!/bin/bash
# sections.sh — exercise the docx `section` property surface
# (schemas/help/docx/section.json) using the officecli CLI directly.
#
# A section is a run of content that shares one page layout: page size &
# orientation, margins, columns, header/footer refs, line numbering,
# footnote/endnote behaviour and vertical alignment. A .docx always has ONE
# trailing "final" section (the body-level sectPr addressed at path /). Every
# `add / --type section` inserts a section BREAK: it closes off the content
# added so far into a new /section[N] carrying its own SectionProperties, and
# the still-open trailing section shifts down to hold whatever comes next.
#
# So the pattern is:  add paragraphs  ->  add section (break)  ->  repeat.
# The section you just added owns the paragraphs ABOVE it. Sections are
# addressed /section[N] (query/get/set/remove); the final trailing section is
# addressed / (and Set rejects a break `type` there — it has no break).
#
# CLI twin of sections.py (officecli SDK); both produce an equivalent
# sections.docx demonstrating three sections with different layouts:
#   1. two-column, page-bottom footnotes (lowerRoman, restart each page)
#   2. single-column landscape, vertically centered, line numbering
#   3. two-column continuous with endnotes collected at document end
#
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/sections.docx"
echo "Building $FILE ..."
rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

body() { officecli add "$FILE" /body --type paragraph --prop "text=$1"; }
head1() { officecli add "$FILE" /body --type paragraph --prop "text=$1" --prop style=Heading1; }

# =====================================================================
# SECTION 1 content — two-column newsletter body with footnotes
# Paragraph indices: p[1] heading, p[2]..p[9] body (enough copy that
# column 1 fills top-to-bottom and text wraps into column 2).
# =====================================================================
head1 "1. Two-Column Layout with Footnotes"
body "Multi-column layout flows text down the first column, then wraps to the top of the next. This section uses two balanced columns with a one-centimetre gutter between them, and enough running text follows that the wrap from the foot of the left column to the head of the right column is plainly visible on the printed page."
body "Newspapers and newsletters have used the two-column measure for centuries because a narrower column is easier for the eye to track: the reader's gaze travels a shorter distance before returning to the start of the next line, so long passages of body text feel less tiring than the same words set across the full page width."
body "Footnotes in this section are numbered with lower-case Roman numerals and the counter restarts on every page, mirroring a printed periodical. The reference marker sits inline in the running text, while the note itself is anchored at the bottom of the page column, beneath a short separator rule."
body "Word measures the column width from the page width, minus the left and right margins, minus the inter-column spacing, divided across the requested column count. Because equal-width columns are enabled here, both columns share exactly the same measure, and the gutter between them is held at the one-centimetre value we set."
body "When a column is filled to the bottom margin, the text engine breaks the flow and resumes typesetting at the top of the next column on the same page. Only after the last column on a page is full does the flow move to a new page, so a two-column section packs roughly twice as many lines onto each sheet as a single-column one."
body "This behaviour is why the amount of body copy matters for a demonstration: with only two or three short paragraphs the first column never fills, and the layout reads on the page as if it were a single column. Several paragraphs of steady prose, like these, are needed before the column break actually occurs."
body "Balancing is the final refinement. On the last page of a multi-column section Word tries to even out the columns so they end at roughly the same height, rather than leaving one long column beside a short stub. The result is the tidy, squared-off block of text that readers expect from a professionally set page."
body "With that, the first section has enough material to spill from column one into column two and, depending on the font and page size, perhaps onto a second page — exactly the wrapping behaviour a two-column layout is meant to show."

# Footnotes attach to a body paragraph (/body/p[N]); they render at the foot
# of the page per this section's footnotePr.pos.
officecli add "$FILE" '/body/p[4]' --type footnote --prop text="Column width = (page width - margins - column spacing) / column count."
officecli add "$FILE" '/body/p[8]' --type footnote --prop text="Balanced columns keep the two measures visually equal on the final page."

# --- Section 1 settings: 2 columns, footnotes, A4 portrait ---
# Features: type=nextPage break; columns=2 + columnSpace gutter;
#           footnotePr.* number format / per-page restart / start / position;
#           titlePage distinct first-page header/footer; per-section page size.
officecli add "$FILE" / --type section \
  --prop type=nextPage \
  --prop pageWidth=21cm --prop pageHeight=29.7cm --prop orientation=portrait \
  --prop marginTop=2.54cm --prop marginBottom=2.54cm \
  --prop marginLeft=2.54cm --prop marginRight=2.54cm \
  --prop marginHeader=1.25cm --prop marginFooter=1.25cm \
  --prop columns=2 --prop columnSpace=1cm \
  --prop titlePage=true \
  --prop footnotePr.numFmt=lowerRoman \
  --prop footnotePr.numRestart=eachPage \
  --prop footnotePr.numStart=1 \
  --prop footnotePr.pos=pageBottom

# =====================================================================
# SECTION 2 content — single-column landscape, vertically centered
# =====================================================================
head1 "2. Landscape, Single Column, Vertically Centered"
body "This section switches to landscape orientation with a single column and asymmetric margins. The vertical alignment is set to center, so a short block of content floats in the middle of the page height instead of hugging the top margin."
body "Landscape sections are common for wide tables, timelines and figures. Because page setup is a per-section property, this page can be landscape while its neighbours stay portrait, all within one document."
body "Line numbering is enabled here in continuous mode, numbering every fifth line, with a gutter between the number column and the body text — the layout used for legal and manuscript review."
body "A single-column section does not wrap, so it needs no extra copy to demonstrate; the point here is the interaction of landscape geometry, centered vertical alignment, and the margin line numbers rather than any column flow."

# --- Section 2 settings: landscape, 1 column, vAlign, line numbering ---
# Features: orientation=landscape with swapped pageWidth/pageHeight;
#           columns=1 (single column); vAlign=center vertical alignment;
#           lineNumbers=continuous + lineNumberCountBy interval + distance;
#           pageNumFmt page-number format + pageStart starting number.
officecli add "$FILE" / --type section \
  --prop type=nextPage \
  --prop orientation=landscape \
  --prop pageWidth=29.7cm --prop pageHeight=21cm \
  --prop marginTop=2cm --prop marginBottom=2cm \
  --prop marginLeft=3cm --prop marginRight=1.5cm \
  --prop columns=1 \
  --prop vAlign=center \
  --prop lineNumbers=continuous \
  --prop lineNumberCountBy=5 \
  --prop lineNumberDistance=288 \
  --prop pageNumFmt=decimal \
  --prop pageStart=1

# =====================================================================
# SECTION 3 content — continuous two-column with endnotes
# Paragraph indices: p[15] heading, p[16]..p[22] body (enough copy for the
# two continuous columns to fill and wrap on the page).
# =====================================================================
head1 "3. Continuous Two-Column with Endnotes"
body "A continuous section break changes the layout without starting a new page. This section returns to portrait, splits into two columns again, and collects its references as endnotes gathered at the end of the document rather than at the foot of each page."
body "Because the break is continuous rather than next-page, the two-column layout begins immediately below the heading on whatever page the previous section left off, instead of ejecting to a fresh sheet. This is the classic magazine construction: a full-width headline followed by columned body text on the same page."
body "Endnotes here use upper-case Roman numerals and restart per section. Unlike footnotes, endnote bodies are not printed at the foot of the page; they live in a separate store and are rendered together where endnotePr.pos points them — in this document, at the very end."
body "The choice between footnotes and endnotes is editorial. Footnotes keep the reference in the reader's eye on the same page, which suits explanatory asides, whereas endnotes keep the body text clean and gather citations in one place, which suits scholarly bibliographies and long reference lists."
body "As with the first section, the wrap only becomes visible once the first column fills to the bottom margin. These middle paragraphs exist to supply that volume of copy, so the reader can watch the text leave the foot of the left column and continue at the top of the right column without a page break intervening."
body "Continuous sections are also how a document mixes column counts on a single page: a full-width introduction, then a continuous break into two or three columns, then another continuous break back to full width for a closing note — all flowing down the same sheet of paper."
body "This closing paragraph rounds out the section with enough text that the two continuous columns fill and balance, completing the demonstration of a continuous multi-column layout with document-end endnotes."

# Endnotes attach to a body paragraph (/body/p[N]), same as footnotes.
officecli add "$FILE" '/body/p[18]' --type endnote --prop text="Endnotes are collected per endnotePr.pos; here they gather at the document end."
officecli add "$FILE" '/body/p[22]' --type endnote --prop text="Upper-Roman numbering restarts each section under endnotePr.numRestart=eachSect."

# --- Section 3 settings: continuous, 2 columns, endnotes ---
# Features: type=continuous break (no page eject); columns=2 + columnSpace;
#           endnotePr.* number format / per-section restart / start / docEnd position.
officecli add "$FILE" / --type section \
  --prop type=continuous \
  --prop orientation=portrait \
  --prop pageWidth=21cm --prop pageHeight=29.7cm \
  --prop columns=2 --prop columnSpace=0.8cm \
  --prop endnotePr.numFmt=upperRoman \
  --prop endnotePr.numRestart=eachSect \
  --prop endnotePr.numStart=1 \
  --prop endnotePr.pos=docEnd

# =====================================================================
# FINAL trailing section — addressed / (no break type; it is the last one).
# Set page setup on it so the tail of the document has a defined layout.
# =====================================================================
# Features: body-level sectPr at path / ; vAlign top; single column.
officecli set "$FILE" / \
  --prop orientation=portrait \
  --prop pageWidth=21cm --prop pageHeight=29.7cm \
  --prop marginTop=2.54cm --prop marginBottom=2.54cm \
  --prop columns=1 --prop vAlign=top

officecli close "$FILE"

echo ""
echo "--- Sections in the finished document ---"
officecli query "$FILE" section

echo ""
officecli validate "$FILE"
echo "Created: $FILE"
