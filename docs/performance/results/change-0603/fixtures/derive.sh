#!/bin/bash
# Derive change 0603's measurement fixtures.
#
# No real .xlsx reaches the source-backed value editor (change 0602 measured
# 95 of 95 refused at the package-root relationship allow-list), so the shapes
# measured here are change 0602's projections, regenerated verbatim with that
# record's own two scripts and its `plain` variant. `degate.py` opens the three
# relationship gates and folds shared strings back in as inline strings;
# `project.py` projects the workbook and every worksheet onto the value-only
# element and attribute allow-lists while leaving <sheetData> byte-untouched,
# which is what leaves `xmlns:mc` and `xmlns:x14ac` declared and every
# `mc:`/`x14ac:` attribute removed -- the declaration-only shape this change
# admits.
set -eu
R=${1:?repository root}
D=${2:?output directory}
F=$R/docs/performance/results/change-0602/fixtures
mkdir -p "$D"
derive() { # source-path stem
  python3 "$F/degate.py"  "$R/$1"            "$D/$2-deg-plain.xlsx"  plain
  python3 "$F/project.py" "$D/$2-deg-plain.xlsx" "$D/$2-proj-plain.xlsx" plain
}
derive test-data/ooxml/xlsx/FormatConditionTests.xlsx                FormatConditionTests
derive test-data/poi/test-data/spreadsheet/dataValidationTableRange.xlsx dataValidationTableRange
derive test-data/ooxml/xlsx/sheet-state-show.xlsx                    sheet-state-show
derive test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx no_drawing_patriarch
derive test-data/ooxml/xlsx/MatrixFormulaEvalTestData.xlsx           MatrixFormulaEvalTestData
