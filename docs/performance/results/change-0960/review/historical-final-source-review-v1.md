# 0960 final source review

Reviewed 2026-10-07 against the current isolated candidate source. No source,
test, build, or workload files were changed by this review.

## Exact source

Candidate file:
`/home/zhuhe/litchi-performance-0960/source-candidate-v1/crates/litchi-xlsx/src/raw/worksheet/codec.rs`

- bytes: 88,079
- SHA-256: `d26067ac170890fc46e7e91ff9bbb9565aaa40fe9260e2bbe853aee851633e3b`

The corrected event-boundary code at lines 1740-1765 keeps raw
`reader.buffer_position()` values for `lane::Entry::locate` and separately uses
`ReaderOrigin::offset` for public event spans. This fixes the earlier BOM
coordinate bug. The focused test at lines 2008-2024 covers an actual BOM body
and requires one admitted hint with three cells.

## Review result

The candidate mechanism is semantically bounded and the corrected ordering is
sound:

1. The candidate entry is captured without moving the event.
2. MCE equivalence and the caller observer run before parser transition.
3. Parser transition remains authoritative for worksheet namespace, context,
   duplicate `sheetData`, and all existing refusal behavior.
4. Hint recognition/reservation runs only after `Context::SheetData` and the
   SpreadsheetML child-namespace check.
5. `lane::recognize` is pure; failed recognition declines without parser state.
6. Cell/event caps and checked subtraction bound the optional request.
7. Only the optional `try_reserve_exact` failure is swallowed; ordinary
   `finish_cell` reservation/error handling remains unchanged.

The focused tests cover observer refusal, malformed-tail fallback, duplicate
`sheetData`, non-strict bodies, event-cap decline, BOM input, and injected hint
reserve failure. The injected failure is a private test hook that exercises the
same no-error fallback branch; it does not claim to be an allocator fault
injection.

External route census passed for eight actual worksheets: four dense/sparse
worksheets and four medium worksheets, 27,008 cells total, with eight attempts,
eight admissions, zero declines, zero cap declines, and eight successful exact
reservations. This establishes route coverage for the supplied corpus only; it
does not establish all XLSX producers use the strict route.

The source/event cap remains 8 MiB and 131,072 provisional events; the hint also
caps summary cells at 131,072. A strict body can be fully reserved before a
malformed tail or later MCE-equivalence failure is discovered. That is bounded
by the existing source/event admission and changes resource timing only; the
candidate must retain this caveat in performance interpretation.

## Quality binding

- Full `litchi-xlsx --lib` quality run: 1,404 passed, zero failed (`rc=0`),
  receipt `0960-full-xlsx-lib-v2-manifest.txt.result`.
- Focused final v6 run: 7 passed, 1 ignored external census test (`rc=0`),
  receipt `0960-focused-strict-hint-v6-manifest.txt.result`.
- Route census: `rc=0`, receipt `0960-managed-route-census-v1-manifest.txt.result`.
- Earlier focused v5 failure was the test-only moved-`entries` compile error;
  v6 contains the `&entries` correction and is the passing focused source.
- Current Clippy run is not clean: `0960-clippy-xlsx-lib-v1` returned `rc=101`
  for `clippy::needless_borrow` at line 1811 (`&resolver` should be
  `resolver`). This is a one-token quality fix, but the current hash above is
  therefore **not final quality-approved** until that correction is applied and
  Clippy is rerun against the resulting new source hash.

No performance or adoption conclusion is made by this source review. The
recognizer's second bounded body scan and front-loaded fallback allocation must
be measured separately from parser allocation savings.
