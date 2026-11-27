#!/bin/bash
# Generate a showcase document with beautiful charts
# Contains 8 chart types: combo chart, 3D bar, scatter+trendline, 3D pie, bubble, stock OHLC, filled radar, multi-ring doughnut
# 4 Sheets: monthly sales, analysis data, stock data, capability assessment

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
XLSX="$(dirname "$0")/charts.xlsx"
echo ""
echo "=========================================="
echo "Generating beautiful charts document: $XLSX"
echo "=========================================="

rm -f "$XLSX"
officecli create "$XLSX"
officecli open "$XLSX"

###############################################################################
# Sheet1: Monthly sales data
###############################################################################
echo "  -> Populating Sheet1: Monthly sales data"

officecli set "$XLSX" '/Sheet1/A1' --prop value="Month" --prop font.bold=true --prop fill=1F4E79 --prop font.color=FFFFFF --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/B1' --prop value="East Sales" --prop font.bold=true --prop fill=2E75B6 --prop font.color=FFFFFF --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/C1' --prop value="South Sales" --prop font.bold=true --prop fill=9DC3E6 --prop font.color=1F4E79 --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/D1' --prop value="North Sales" --prop font.bold=true --prop fill=BDD7EE --prop font.color=1F4E79 --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/E1' --prop value="Total" --prop font.bold=true --prop fill=C55A11 --prop font.color=FFFFFF --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/F1' --prop value="YoY Growth %" --prop font.bold=true --prop fill=548235 --prop font.color=FFFFFF --prop font.size=11 --prop alignment.horizontal=center

MONTHS=("Jan" "Feb" "Mar" "Apr" "May" "Jun" "Jul" "Aug" "Sep" "Oct" "Nov" "Dec")
EAST=(120 135 148 162 155 178 195 210 188 172 165 198)
SOUTH=(95 108 115 128 142 155 168 175 160 148 135 158)
NORTH=(88 92 105 118 125 138 145 152 140 130 122 142)
TOTAL=(303 335 368 408 422 471 508 537 488 450 422 498)
GROWTH=(5.2 8.1 12.3 15.6 10.2 18.5 22.1 25.3 16.8 11.2 7.5 19.8)

for i in $(seq 0 11); do
    row=$((i + 2))
    officecli set "$XLSX" "/Sheet1/A${row}" --prop "value=${MONTHS[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/B${row}" --prop "value=${EAST[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/C${row}" --prop "value=${SOUTH[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/D${row}" --prop "value=${NORTH[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/E${row}" --prop "value=${TOTAL[$i]}" --prop 'numFmt=#,##0' --prop font.bold=true --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/F${row}" --prop "value=${GROWTH[$i]}" --prop 'numFmt=0.0"%"' --prop alignment.horizontal=center
done

echo "  Done: Sheet1 data"

###############################################################################
# Sheet2: Scatter/bubble chart data
###############################################################################
echo "  -> Populating Sheet2: Analysis data"

officecli add "$XLSX" / --type sheet --prop name=Analysis

officecli set "$XLSX" '/Analysis/A1' --prop value="Ad Spend (10K)" --prop font.bold=true --prop fill=7030A0 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Analysis/B1' --prop value="Sales (10K)" --prop font.bold=true --prop fill=7030A0 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Analysis/C1' --prop value="Margin %" --prop font.bold=true --prop fill=7030A0 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Analysis/D1' --prop value="Market Share %" --prop font.bold=true --prop fill=7030A0 --prop font.color=FFFFFF --prop alignment.horizontal=center

AD_SPEND=(10 15 22 28 35 42 50 58 65 72 80 88 95 105 115)
SALES_REV=(45 68 95 120 155 180 220 260 290 335 370 410 445 500 550)
PROFIT=(8.5 10.2 12.1 14.5 16.8 15.2 18.3 20.1 19.5 22.3 21.8 24.5 23.1 26.8 28.2)
MKT_SHARE=(2.1 3.2 4.5 5.8 7.2 8.5 10.1 11.8 12.5 14.2 15.8 17.5 18.2 20.5 22.1)

for i in $(seq 0 14); do
    row=$((i + 2))
    officecli set "$XLSX" "/Analysis/A${row}" --prop "value=${AD_SPEND[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Analysis/B${row}" --prop "value=${SALES_REV[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Analysis/C${row}" --prop "value=${PROFIT[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Analysis/D${row}" --prop "value=${MKT_SHARE[$i]}" --prop alignment.horizontal=center
done

echo "  Done: Sheet2 data"

###############################################################################
# Sheet3: Stock data (with red/green coloring)
###############################################################################
echo "  -> Populating Sheet3: Stock data"

officecli add "$XLSX" / --type sheet --prop name=StockData

officecli set "$XLSX" '/StockData/A1' --prop value="Date" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/B1' --prop value="Open" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/C1' --prop value="High" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/D1' --prop value="Low" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/E1' --prop value="Close" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/F1' --prop value="Volume (10K)" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center

DATES=("3/1" "3/2" "3/3" "3/4" "3/5" "3/6" "3/7" "3/8" "3/9" "3/10" "3/11" "3/12" "3/13" "3/14" "3/15" "3/16" "3/17" "3/18" "3/19" "3/20")
OPEN=(52.3 53.1 52.8 54.2 55.1 54.5 56.2 57.8 58.5 57.2 56.8 58.3 59.5 60.2 59.8 61.5 62.3 61.8 63.5 64.2)
HIGH=(53.8 54.2 54.5 55.8 56.3 56.8 58.1 59.2 59.8 58.5 58.2 59.8 61.2 61.5 61.8 63.2 63.8 63.5 65.2 65.8)
LOW=(51.5 52.2 51.8 53.5 54.2 53.8 55.5 56.8 57.2 56.1 55.8 57.5 58.8 59.2 58.5 60.8 61.2 60.5 62.8 63.5)
CLOSE=(53.1 52.8 54.2 55.1 54.5 56.2 57.8 58.5 57.2 56.8 58.3 59.5 60.2 59.8 61.5 62.3 61.8 63.5 64.2 65.1)
VOLUME=(285 312 268 345 298 378 425 468 395 310 352 415 485 442 368 512 548 478 562 598)

for i in $(seq 0 19); do
    row=$((i + 2))

    open=${OPEN[$i]}
    close=${CLOSE[$i]}
    if (( $(echo "$close > $open" | bc -l) )); then
        COLOR="FF0000"; BG="FFF2F2"  # Up: red
    elif (( $(echo "$close < $open" | bc -l) )); then
        COLOR="008000"; BG="F2FFF2"  # Down: green
    else
        COLOR="666666"; BG="F5F5F5"  # Flat: gray
    fi

    officecli set "$XLSX" "/StockData/A${row}" --prop "value=${DATES[$i]}" --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/B${row}" --prop "value=${OPEN[$i]}" --prop 'numFmt=0.00' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/C${row}" --prop "value=${HIGH[$i]}" --prop 'numFmt=0.00' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/D${row}" --prop "value=${LOW[$i]}" --prop 'numFmt=0.00' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/E${row}" --prop "value=${CLOSE[$i]}" --prop 'numFmt=0.00' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/F${row}" --prop "value=${VOLUME[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
done

echo "  Done: Sheet3 stock data (with red/green coloring)"

###############################################################################
# Sheet4: Capability radar chart data
###############################################################################
echo "  -> Populating Sheet4: Capability assessment"

officecli add "$XLSX" / --type sheet --prop name=Assessment

officecli set "$XLSX" '/Assessment/A1' --prop value="Dimension" --prop font.bold=true --prop fill=002060 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Assessment/B1' --prop value="Product A" --prop font.bold=true --prop fill=0070C0 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Assessment/C1' --prop value="Product B" --prop font.bold=true --prop fill=00B050 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Assessment/D1' --prop value="Product C" --prop font.bold=true --prop fill=FFC000 --prop font.color=000000 --prop alignment.horizontal=center

DIMS=("Performance" "Stability" "Usability" "Security" "Scalability" "Value" "Ecosystem" "Docs")
PA=(92 88 75 95 82 70 85 78)
PB=(78 92 88 80 90 85 72 82)
PC=(85 76 92 72 78 92 88 70)

for i in $(seq 0 7); do
    row=$((i + 2))
    officecli set "$XLSX" "/Assessment/A${row}" --prop "value=${DIMS[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Assessment/B${row}" --prop "value=${PA[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Assessment/C${row}" --prop "value=${PB[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Assessment/D${row}" --prop "value=${PC[$i]}" --prop alignment.horizontal=center
done

echo "  Done: Sheet4 data"

###############################################################################
# Charts 1-8 — HIGH-LEVEL API (`officecli add --type chart`).
# Each chart is one `add` call: chartType + a cell dataRange (or inline data for
# the pie/doughnut, whose source cells aren't contiguous) + styling props
# (title, colors, legend, 3D view, trendline, secondary axis, radar fill, stock
# up/down bars). Positioned with x/y/width/height in cell units.
# Exception: Chart 5 (bubble) stays on raw-set — the high-level command can't yet
# map a dataRange to a single x/y/size series (see its note below).
###############################################################################

# Chart 1: Combo — regional sales columns + YoY-growth line on a secondary axis.
echo "  -> Chart 1: Combo chart (columns + secondary-axis line)"
officecli add "$XLSX" /Sheet1 --type chart \
  --prop chartType=combo \
  --prop title="Monthly Sales and YoY Growth Trend" \
  --prop dataRange=Sheet1!A1:F13 \
  --prop combotypes=column,column,column,column,line \
  --prop secondaryaxis=5 \
  --prop colors=2E75B6,9DC3E6,BDD7EE,C55A11,FF0000 \
  --prop legend=b --prop axisTitle="Sales (10K)" \
  --prop x=7 --prop y=0 --prop width=11 --prop height=18

# Chart 2: 3D clustered column — regional comparison.
echo "  -> Chart 2: 3D bar chart"
officecli add "$XLSX" /Sheet1 --type chart \
  --prop chartType=column3d \
  --prop title="3D Regional Sales Comparison" \
  --prop dataRange=Sheet1!A1:D13 \
  --prop view3d=15,20,30 \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop legend=b \
  --prop x=7 --prop y=19 --prop width=11 --prop height=18

# Chart 3: Scatter + linear trendline — ad spend vs sales.
echo "  -> Chart 3: Scatter plot + trendline"
officecli add "$XLSX" /Analysis --type chart \
  --prop chartType=scatter \
  --prop title="Ad Spend vs Sales Correlation" \
  --prop dataRange=Analysis!A1:B16 \
  --prop trendline=linear \
  --prop colors=7030A0 \
  --prop catTitle="Ad Spend (10K)" --prop axisTitle="Sales (10K)" \
  --prop legend=b \
  --prop x=5 --prop y=0 --prop width=11 --prop height=18

# Chart 4: Exploded 3D pie — July regional share. Values live in one row
# (B8:D8) with category labels in another (B1:D1), so they're passed inline.
echo "  -> Chart 4: 3D pie chart (exploded)"
officecli add "$XLSX" /Sheet1 --type chart \
  --prop chartType=pie3d \
  --prop title="July Regional Sales Share (3D)" \
  --prop categories="East Sales,South Sales,North Sales" \
  --prop series1="Jul:195,168,145" \
  --prop explosion=10 --prop view3d=30,70,30 \
  --prop dataLabels=percent \
  --prop colors=1F4E79,C55A11,548235 \
  --prop legend=b \
  --prop x=19 --prop y=0 --prop width=9 --prop height=18

# Chart 5: Bubble — ad spend (x) vs sales (y), bubble size = market share.
# KEPT ON raw-set: the high-level `add --type chart --prop chartType=bubble`
# reads a multi-column dataRange as several y-series sharing column A as x — it
# cannot map three columns to a single x / y / size series (multi-point bubble).
# Until that mapping exists, the faithful single-series bubble needs raw XML.
echo "  -> Chart 5: Bubble chart (raw-set — see note)"
CHART5_REL=$(officecli add-part "$XLSX" /Analysis --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)
officecli raw-set "$XLSX" '/Analysis/chart[2]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Spend-Revenue-Market Share Bubble</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:bubbleChart>
        <c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Analysis!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="7030A0"><a:alpha val="60000" /></a:srgbClr></a:solidFill>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:ln>
          </c:spPr>
          <c:xVal><c:numRef><c:f>Analysis!$A$2:$A$16</c:f></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>Analysis!$B$2:$B$16</c:f></c:numRef></c:yVal>
          <c:bubbleSize><c:numRef><c:f>Analysis!$D$2:$D$16</c:f></c:numRef></c:bubbleSize>
          <c:bubble3D val="1" />
        </c:ser>
        <c:axId val="300" /><c:axId val="400" />
      </c:bubbleChart>
      <c:valAx>
        <c:axId val="300" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Ad Spend (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="400" />
      </c:valAx>
      <c:valAx>
        <c:axId val="400" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Sales (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="300" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'
officecli raw-set "$XLSX" '/Analysis/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>16</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>37</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"3\" name=\"Chart 5\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART5_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

# Chart 6: Stock OHLC candlestick — hi-low lines + up/down bars (red up, green down).
echo "  -> Chart 6: Stock OHLC chart"
officecli add "$XLSX" /StockData --type chart \
  --prop chartType=stock \
  --prop title="Stock Candlestick Chart (OHLC)" \
  --prop dataRange=StockData!A1:E21 \
  --prop hilowlines=true \
  --prop updownbars=100:FF0000:00B050 \
  --prop legend=b \
  --prop x=7 --prop y=0 --prop width=13 --prop height=22

# Chart 7: Filled radar — product capability comparison.
echo "  -> Chart 7: Filled radar chart"
officecli add "$XLSX" /Assessment --type chart \
  --prop chartType=radar --prop radarStyle=filled \
  --prop title="Product Capability Radar Comparison" \
  --prop dataRange=Assessment!A1:D9 \
  --prop colors=4472C4,00B050,FFC000 \
  --prop legend=b \
  --prop x=5 --prop y=0 --prop width=11 --prop height=20

# Chart 8: Multi-ring doughnut — Aug vs Dec regional share (two rings). The two
# source rows aren't adjacent, so the ring values are passed inline.
echo "  -> Chart 8: Multi-ring doughnut chart"
officecli add "$XLSX" /Sheet1 --type chart \
  --prop chartType=doughnut \
  --prop title="Aug vs Dec Regional Sales Multi-Ring" \
  --prop categories="East,South,North" \
  --prop series1="Aug:210,175,152" \
  --prop series2="Dec:198,158,142" \
  --prop dataLabels=percent \
  --prop colors=1F4E79,C55A11,548235 \
  --prop legend=b \
  --prop x=19 --prop y=19 --prop width=9 --prop height=18

###############################################################################
# Validation
###############################################################################
officecli close "$XLSX"

echo ""
echo "=========================================="
echo "Validating file"
echo "=========================================="
officecli validate "$XLSX"
officecli view "$XLSX" outline
echo ""
ls -lh "$XLSX"
echo ""
echo "All done! 8 chart types generated"