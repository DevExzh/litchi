# Log sections for change 0757

Ready-to-paste paragraphs for the coordinator, one per shared log. This change
does not edit `HOTSPOTS.md`, `REPORT.md` or `GOAL_AUDIT.md` itself; the numbers
are the record's.

---

## For `HOTSPOTS.md`

## 0757 — The fresh XLS writer's shared-string order is a function of the cells, and its SST, formula, number-format, sheet-name and defined-name fields refuse what they cannot hold

[0757](0757-xls-fresh-writer-sst-determinism.md) closes the two findings 0753
left open. The SST now lists each string at its first occurrence in worksheet,
row and column order (the order of the cell records), found by sorting each
worksheet's string cells by one packed `row << 16 | column` key; the keyed
SipHash map only deduplicates. These fields refuse what BIFF8 cannot hold with
`Error::StringTooLong` before any output — when the string is given to the
writer, not when the workbook is written: SST entries (65,535 UTF-16 units; the
old cut through a surrogate pair made litchi's reader refuse the workbook),
formula string constants (255), number formats (1–255; 70,000 units used to
overflow a `u16`, and an empty one was written and refused on open), worksheet
names (31, now counted in UTF-16 units) and defined names and their comments
(255, now encoded correctly for Latin-1 and supplementary characters);
data-validation strings with Latin-1 characters no longer make the worksheet
unreadable. Cost, on a `write_to`-only probe of 80,000 string cells:
+11.4–15.2% instructions and 0.4–12.2% wall time depending on glibc heap state;
the harness's numeric and single-string shapes are within noise in
instructions, with `xls_fresh_write_to/large` 1.05–1.08 in wall time from a
heap-trim effect (0.983 with malloc thresholds pinned). Found, not fixed (with
the review): internal hyperlinks and AutoFilter strings are written with the
wrong length for non-ASCII text and litchi's reader then drops them (with the
worksheet's cells, for hyperlinks); font names are cut to 31 units; fonts added
through `add_cell_style` get the invalid index 4; forbidden sheet-name
characters make the reader refuse the workbook; more than 210 custom number
formats make the reader refuse it; PivotTable writers count characters instead
of UTF-16 units; XFEXT/STYLEEXT/CRN lengths may wrap. Left: the table and the
emission sort a string worksheet's cells twice (sharing one sort costs +7.5%
peak on the probe). A multi-string XLS harness selector would now measure the
SST path.

---

## For `REPORT.md`

## 0757 — Deterministic shared strings and typed string refusals in the fresh XLS writer

[0757](0757-xls-fresh-writer-sst-determinism.md) (base `9ff78bbf1c`, measured
`921787e2c0`, final `a400341bdf`) makes the fresh XLS writer's output
reproducible: eight base processes writing the same 80,000-string workbook wrote
eight different byte streams, the branch writes one, and a pinned multi-string
golden passes in twelve separate processes (it fails on the base as "not
deterministic"). It adds `Error::StringTooLong` and refuses, before any output,
SST strings over 65,535 UTF-16 units, formula string constants and number
formats over 255, worksheet names over 31 and defined names over 255, where the
base cut the string, wrapped a length, or wrote a record its own reader refused;
`encode_ptg_tokens`, `register_number_format` and `add_cell_style` become
fallible (breaking, owner decision 1), and every refusal happens when the string
is registered, so a refused call leaves the writer writing the same bytes as
before it. Output bytes
are unchanged for worksheets with at most one distinct string (pre-0753 goldens,
all harness corpora). Paired ABBA on CPU 16, two windows: `xls_fresh_write_to`
tiny 1.015/1.016, payload-heavy 1.013/1.012, large 1.079/1.053 (3.4% fewer
instructions; a glibc heap-trim effect, 0.983 with thresholds pinned); control
`xls_semantic_one_edit_save` 0.991–0.997. Probe (`write_to` only): 80,000
distinct strings 1.122 (1.004 pinned), 64 repeated labels 1.057 (1.061 pinned),
+11.4–15.2% instructions. It also carries the 0753 review follow-ups: fallible,
capped DOC/PPT stream reservations made after the document-wide checks,
`InPlaceRecord` reachable only through a closure that always patches its header,
and exhaustive tests of `utf16_units` and `contains_field_character`; DOC and
PPT output is byte-identical. `performance_claim: none`.

---

## For `GOAL_AUDIT.md`

## 0757 — ADR 0006 determinism restored for the fresh XLS writer, and GOAL rule 3 enforced on its string fields

[0757](0757-xls-fresh-writer-sst-determinism.md) removes a process-dependent
output from the fresh XLS writer (the SST order followed a per-process hash
seed; ADR 0006 requires deterministic serialization) and replaces silent
truncation, wrapped lengths and misencoded strings in its SST, formula-string,
number-format, worksheet-name, defined-name and data-validation encoders with a
typed refusal before any output (GOAL.md rule 3: never trade a typed refusal for
a partial edit). Refusals leave the writer unchanged (ADR 0016); no limit is
relaxed, and the worksheet-name rule now counts the UTF-16 units MS-XLS counts.
The public API changes (`Error::StringTooLong`, a fallible `encode_ptg_tokens`)
rely on owner decision 1 of change 0652; the source-backed editor's
`insert_formula`, which shares the encoder, now refuses instead of truncating.
Refusals happen when a string is registered, never only at write time, so a
refused call cannot leave the writer unable to write (the review's finding,
fixed). Recorded for follow-up: internal-hyperlink, AutoFilter, font-name and
PivotTable string lengths, the font index 4 of added cell styles, forbidden
sheet-name characters, the custom number-format count, and possible
XFEXT/STYLEEXT/CRN length wraps. The non-iWork goal remains active.
