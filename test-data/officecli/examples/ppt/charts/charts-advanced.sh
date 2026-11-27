#!/bin/bash
# Advanced Charts Showcase — properties not covered by the per-type decks.
# Generates: charts-advanced.pptx
#
# CLI twin of charts-advanced.py (officecli Python SDK). Both produce an
# equivalent charts-advanced.pptx.
#
#   Slide 1  RTL & anchor          direction=rtl, anchor cm-form
#   Slide 2  Axis-level shortcuts  axisvisible / valaxisvisible / catAxisVisible, orientation, position, lines
#   Slide 3  Crossings             crossBetween / crosses / crossesAt
#   Slide 4  Categories axis       labeloffset, ticklabelskip
#   Slide 5  Marker size & fills   markersize, areafill, chartFill, plotvisonly
#   Slide 6  Built-in style + blanks  style=1..48, dispBlanksAs, dataRange
#   Slide 7  chart-axis Set        dispUnits, logBase, minorUnit, visible, labelRotation
#   Slide 8  chart-series mutation values=, categories= per-series
#
# Usage: ./charts-advanced.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-advanced.pptx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ===========================================================================
# Slide 1 — RTL + anchor variants
# ===========================================================================
officecli add "$FILE" / --type slide
officecli add "$FILE" "/slide[1]" --type shape \
  --prop text="RTL + anchor — direction=rtl, named-token anchor, cm-form anchor" \
  --prop size=24 --prop bold=true --prop autoFit=normal \
  --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in

officecli add "$FILE" "/slide[1]" --type chart \
  --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="default (LTR)" --prop legend=bottom \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180;B:50,75,110,150"

# RTL must be Set after Add (direction is set-only)
officecli add "$FILE" "/slide[1]" --type chart \
  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="direction=rtl (Set after Add)" --prop legend=bottom \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180;B:50,75,110,150"
officecli set "$FILE" "/slide[1]/chart[2]" --prop direction=rtl

# Anchor cm-form: x,y,w,h
officecli add "$FILE" "/slide[1]" --type chart \
  --prop anchor=0.3cm,11cm,15.5cm,7cm \
  --prop chartType=column --prop title="anchor=0.3cm,11cm,15.5cm,7cm" --prop legend=bottom \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

# ===========================================================================
# Slide 2 — axis-level shortcuts
# ===========================================================================
officecli add "$FILE" / --type slide
officecli add "$FILE" "/slide[2]" --type shape \
  --prop text="Axis shortcuts — axisvisible / valaxisvisible / catAxisVisible, orientation, position, lines" \
  --prop size=24 --prop bold=true --prop autoFit=normal \
  --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in

officecli add "$FILE" "/slide[2]" --type chart \
  --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="axisvisible=false (both axes hidden)" \
  --prop legend=none --prop axisvisible=false --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

officecli add "$FILE" "/slide[2]" --type chart \
  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="valaxisvisible=false (Y hidden, X shown)" \
  --prop legend=none --prop valaxisvisible=false --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

officecli add "$FILE" "/slide[2]" --type chart \
  --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="catAxisVisible=false (X hidden)" \
  --prop legend=none --prop catAxisVisible=false --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

officecli add "$FILE" "/slide[2]" --type chart \
  --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="axisorientation=true (reversed) + axisposition=top" \
  --prop legend=none --prop axisorientation=true --prop axisposition=top \
  --prop cataxisline=333333:1 --prop valaxisline=333333:1 \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

# ===========================================================================
# Slide 3 — Crossings
# ===========================================================================
officecli add "$FILE" / --type slide
officecli add "$FILE" "/slide[3]" --type shape \
  --prop text="Crossings — crossBetween / crosses / crossesAt" \
  --prop size=24 --prop bold=true --prop autoFit=normal \
  --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in

officecli add "$FILE" "/slide[3]" --type chart \
  --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="crossBetween=between (default)" \
  --prop legend=none --prop crossBetween=between --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

officecli add "$FILE" "/slide[3]" --type chart \
  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="crossBetween=midCat" \
  --prop legend=none --prop crossBetween=midCat --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

officecli add "$FILE" "/slide[3]" --type chart \
  --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="crosses=max (Y crosses at top)" \
  --prop legend=none --prop crosses=max --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

# crossesAt is the overriding form of crosses in CT_ValAx (mutually-exclusive
# schema choice). crosses=autoZero is the OOXML default crossesAt supersedes,
# so we set crossesAt alone — the axis crosses the category axis at value 100.
officecli add "$FILE" "/slide[3]" --type chart \
  --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="crossesAt=100 + crosses=autoZero" \
  --prop legend=none --prop crossesAt=100 \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,-30,140,180"

# ===========================================================================
# Slide 4 — Category axis layout
# ===========================================================================
officecli add "$FILE" / --type slide
officecli add "$FILE" "/slide[4]" --type shape \
  --prop text="Category axis — labeloffset, ticklabelskip" \
  --prop size=24 --prop bold=true --prop autoFit=normal \
  --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in

officecli add "$FILE" "/slide[4]" --type chart \
  --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="labeloffset=100 (default)" \
  --prop labeloffset=100 --prop legend=none \
  --prop categories=January,February,March,April,May,June \
  --prop "data=A:60,90,140,180,160,210"

officecli add "$FILE" "/slide[4]" --type chart \
  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="labeloffset=300 (push labels down)" \
  --prop labeloffset=300 --prop legend=none \
  --prop categories=January,February,March,April,May,June \
  --prop "data=A:60,90,140,180,160,210"

officecli add "$FILE" "/slide[4]" --type chart \
  --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="ticklabelskip=2 (every other label)" \
  --prop ticklabelskip=2 --prop legend=none \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec \
  --prop "data=A:60,90,140,180,160,210,200,190,170,150,130,170"

officecli add "$FILE" "/slide[4]" --type chart \
  --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="ticklabelskip=3" \
  --prop ticklabelskip=3 --prop legend=none \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec \
  --prop "data=A:60,90,140,180,160,210,200,190,170,150,130,170"

# ===========================================================================
# Slide 5 — Marker size, area/chart fills, plotvisonly
# ===========================================================================
officecli add "$FILE" / --type slide
officecli add "$FILE" "/slide[5]" --type shape \
  --prop text="Marker size & fills — markersize (standalone), areafill, chartFill, plotvisonly" \
  --prop size=24 --prop bold=true --prop autoFit=normal \
  --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in

officecli add "$FILE" "/slide[5]" --type chart \
  --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=line --prop title="markersize=12 (standalone key)" \
  --prop showMarker=true --prop markersize=12 --prop legend=none \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

officecli add "$FILE" "/slide[5]" --type chart \
  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="areafill (applies to every series shape)" \
  --prop areafill=4472C4-A5C8FF:90 --prop legend=none \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180;B:50,75,110,150"

officecli add "$FILE" "/slide[5]" --type chart \
  --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="chartFill=#FFF8E7 (chart-level fill)" \
  --prop chartFill=#FFF8E7 --prop legend=none \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

officecli add "$FILE" "/slide[5]" --type chart \
  --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="plotvisonly=true (skip hidden rows when bound to a sheet)" \
  --prop plotvisonly=true --prop legend=none \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"

# ===========================================================================
# Slide 6 — Built-in style id + dispBlanksAs
# ===========================================================================
officecli add "$FILE" / --type slide
officecli add "$FILE" "/slide[6]" --type shape \
  --prop text="Built-in style & blank handling — style=1..48, dispBlanksAs, dataRange" \
  --prop size=24 --prop bold=true --prop autoFit=normal \
  --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in

officecli add "$FILE" "/slide[6]" --type chart \
  --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop style=2 --prop title="style=2" --prop legend=bottom \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180;B:50,75,110,150"

officecli add "$FILE" "/slide[6]" --type chart \
  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop style=42 --prop title="style=42" --prop legend=bottom \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180;B:50,75,110,150"

# dispBlanksAs is Set/Get only — Add first, then Set.
officecli add "$FILE" "/slide[6]" --type chart \
  --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=line --prop title="dispBlanksAs=gap (Set after Add)" --prop showMarker=true \
  --prop legend=bottom --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"
officecli set "$FILE" "/slide[6]/chart[3]" --prop dispBlanksAs=gap

# dataRange is Add-time alternative to data= for sheet-backed sources;
# in a standalone pptx this is largely symbolic — demonstrate the syntax,
# then fall back to inline data so the chart renders.
officecli add "$FILE" "/slide[6]" --type chart \
  --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="dataRange syntax demo (fallback inline)" \
  --prop dataRange=Sheet1!A1:D5 --prop legend=bottom --prop catTitle=Quarter \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180;B:50,75,110,150"

# ===========================================================================
# Slide 7 — chart-axis Set (per-axis post-Add)
# ===========================================================================
officecli add "$FILE" / --type slide
officecli add "$FILE" "/slide[7]" --type shape \
  --prop text="chart-axis Set — dispUnits, logBase, minorUnit, visible, labelRotation per-axis" \
  --prop size=24 --prop bold=true --prop autoFit=normal \
  --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in

officecli add "$FILE" "/slide[7]" --type chart \
  --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="after: dispUnits=thousands (Set on value axis)" \
  --prop legend=none --prop categories=Q1,Q2,Q3,Q4 --prop "data=Rev:120000,135000,148000,162000"
officecli set "$FILE" "/slide[7]/chart[1]/axis[@role=value]" \
  --prop dispUnits=thousands --prop format=#,##0 --prop minorUnit=10000 \
  --prop labelRotation=0 --prop visible=true

officecli add "$FILE" "/slide[7]" --type chart \
  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=line --prop title="after: logBase=10 (Set on value axis)" \
  --prop legend=none --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:5,50,500,5000"
officecli set "$FILE" "/slide[7]/chart[2]/axis[@role=value]" \
  --prop logBase=10 --prop min=1 --prop max=10000 --prop majorGridlines=true

officecli add "$FILE" "/slide[7]" --type chart \
  --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="after: visible=false on value axis" \
  --prop legend=none --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"
officecli set "$FILE" "/slide[7]/chart[3]/axis[@role=value]" --prop visible=false

officecli add "$FILE" "/slide[7]" --type chart \
  --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="after: labelRotation=-45 on category axis" \
  --prop legend=none --prop categories=January,February,March,April --prop "data=A:60,90,140,180"
officecli set "$FILE" "/slide[7]/chart[4]/axis[@role=category]" \
  --prop labelRotation=-45 --prop title=Month --prop visible=true

# ===========================================================================
# Slide 8 — chart-series values=/categories= Set + Get readback round-trip
# ===========================================================================
officecli add "$FILE" / --type slide
officecli add "$FILE" "/slide[8]" --type shape \
  --prop text="chart-series mutation — values=, categories= + get-readback round-trip" \
  --prop size=24 --prop bold=true --prop autoFit=normal \
  --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in

officecli add "$FILE" "/slide[8]" --type chart \
  --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="before: A=60,90,140,180" --prop legend=bottom \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"
# Mutate the values after add
officecli set "$FILE" "/slide[8]/chart[1]/series[1]" --prop "values=200,150,100,80"
officecli add "$FILE" "/slide[8]" --type shape \
  --prop text="After Set values=200,150,100,80 the series flips downward." \
  --prop size=10 --prop italic=true --prop color=666666 \
  --prop x=0.3in --prop y=4in --prop width=6in --prop height=0.4in

officecli add "$FILE" "/slide[8]" --type chart \
  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in \
  --prop chartType=column --prop title="per-series categories= (range)" --prop legend=bottom \
  --prop categories=Q1,Q2,Q3,Q4 --prop "data=A:60,90,140,180"
# Per-series category override is range-only — requires sheet backing, so this is
# a demonstration of the syntax only; effective result depends on workbook.

# Round-trip: change one series, read it back, stamp the JSON onto the slide.
officecli set "$FILE" "/slide[8]/chart[1]/series[1]" --prop name="Readback Demo" --prop color=C00000
officecli get "$FILE" "/slide[8]/chart[1]/series[1]" --json
officecli add "$FILE" "/slide[8]" --type shape \
  --prop text="chart-series get --json: readback fields alpha/outlineColor/scatterStyle/..." \
  --prop size=9 --prop color=222222 \
  --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in

# chart-axis get-readback — surfaces axisFont/axisMax/axisMin/axisNumFmt/
# axisOrientation/axisTitle/labelOffset/tickLabelSkip read-only fields.
officecli set "$FILE" "/slide[8]/chart[1]/axis[@role=value]" \
  --prop title="Readback Y" --prop format=$#,##0 --prop min=0 --prop max=300 --prop majorUnit=75
officecli get "$FILE" "/slide[8]/chart[1]/axis[@role=value]" --json
officecli add "$FILE" "/slide[8]" --type shape \
  --prop text="chart-axis get --json: readback axisFont/axisMax/axisMin/axisNumFmt/axisOrientation/axisTitle/labelOffset/tickLabelSkip" \
  --prop size=9 --prop color=222222 \
  --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
