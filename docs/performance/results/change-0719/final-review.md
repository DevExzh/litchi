# 0719 final review

Review status: accepted for the current generated producer-shape scope, with
the stated ZIP-order limitation. The source tree remained frozen during the
release build (source-frozen build process 59828); this review introduces no
Rust edits.

The strengthened `xlsx_producer_*_source_one_edit_save` oracle keeps the
existing sheet-0 midpoint target. That worksheet is generated as an all-numeric
grid. The source target fragment is required to occur exactly once, and the
candidate `sheet1.xml` payload must equal the source payload with that one
known `<c r="..."><v>old</v></c>` fragment replaced by the expected numeric
value. This is a generated-corpus oracle, rather than a generic XLSX rewrite
claim; it covers dimension, root, row, namespace, lexical, and untouched-cell
damage in the edited worksheet.

The semantic oracle reopens source and output through the source-backed
workbook and reads `Rect::ALL` on every worksheet. It checks sheet names, the
unique target, the expected replacement, and equality of every other semantic
cell. The full-grid read prevents an added stored cell outside the generated
rectangle from being hidden by a shape-sized range. The workbook closure then
allows only the required `calcPr` invalidation: workbook XML must match after
removing the direct `calcPr`, and the parsed output calculation properties must
match the required invalidated specification.

The package oracle requires the same normalized member set and requires the
selected worksheet member to change. The workbook member is allowed to remain
byte-identical: a source workbook can already carry the required calculation
flags, which the first source-bound capture exposed. The calculation closure
still checks the outer workbook bytes after removing the direct `calcPr` and
checks the typed output properties, while the source archive's outer hash
remains unchanged. Every other member's complete local record and central
record is compared. The central record's relocatable local-header offset is
zeroed before comparison; compressed data, local headers, central metadata,
names, and extras remain compared. The helper stores records by normalized
member name and therefore makes no ZIP member ordering claim. The exact
worksheet splice and the raw-member checks together cover the generated
producer corpus; this review does not generalize that proof to arbitrary real
files.

The timer starts after sink reservation and source-backed editor construction.
It covers worksheet planning, staging, commit, and sequential publication. The
publication return value is dropped before elapsed time is sampled. Source
semantic setup, output reopening, semantic and splice checks, package and
`calcPr` checks, output hashing, retained sink accounting, and remaining commit
or sink destruction are outside the reported interval. The selector remains a
plan/edit/publication measurement, not an open-through-teardown lifecycle
measurement.

No speedup or production performance claim follows from this harness oracle.
