# 0266 — CSV/TSV import and formula evaluation (plan step 18, M4)

Status: implemented on `uo2/s18-csv-eval` (cut from `feat/unified-ops-api` at
4d1afe92bc); awaiting coordinator review and merge.

Plan: [0248](0248-coordinator-restart-plan.md) §6 Phase E step 18. Design:
§13 P0 "CSV/TSV `import`" and "formula write + evaluation (via
`litchi-eval`)". Glossary: [API_GLOSSARY.md](../API_GLOSSARY.md), section
"Unified CSV/TSV import and formula evaluation (record 0266)".

## 1. What landed

### Import (`litchi_ops::import`)

- `File::import(target, source, &import::Options)` and
  `File::import_with_limits(.., &SemanticLimits)` return
  `transaction::Commit<import::Summary>`. `transaction::Edit::import` is the
  same mutation inside a closure transaction (it composes with `set` and other
  imports). The target is a sheet (`/Sheet1`, origin A1) or an origin cell
  (`/Sheet1/B2`).
- Formats: XLSX, XLSB and XLS. Every other format refuses with
  `unsupported_capability`; ODS has no unified write path yet (M6), so it is a
  gap, as planned.
- Atomicity: one import is one transaction operation. XLSX and XLSB stage
  every field in one native edit and publish one native commit. The BIFF8
  editor admits only 4,096 staged changes per edit (a leaf constant), so XLS
  commits chunks of 1,024 cells, each on the previous result. The chunks are
  retained steps of the same unified transaction, so the caller still gets
  one atomic commit, one reversible patch and an unchanged input on failure.
- Exact no-op: empty input, a lone BOM, or a lone `sep=` line produce no
  native commit; `patch.is_noop()` holds and the bytes are unchanged. The
  target is still resolved, so a missing sheet is still `path_not_found`.

| Aspect | Behaviour |
|---|---|
| Encoding | `Encoding::Auto` (default): UTF-8, UTF-16 LE or UTF-16 BE by byte-order mark, else UTF-8. Explicit encodings strip their own mark. Malformed bytes: `invalid_input` with the byte offset. |
| Delimiter | One ASCII character; default from the source's `sep=X` first line, else `,`. `Options::tsv()` uses a tab. An explicit delimiter overrides `sep=`. The `sep=` line is never data. |
| Quote | `"` by default; doubled quotes escape; quoted fields may hold delimiters and line breaks; text after a closing quote is kept (`"a"b` → `ab`). `quote(None)` reads fields verbatim. An unterminated quote is `invalid_input` with its line. |
| Records | `\r\n`, `\n`, `\r`. A blank line is a record with one empty field, so later rows keep their line positions (OfficeCLI's rule). Ragged records are allowed. |
| Header | `header(true)`: the first record is written as text, never inferred. |
| Types | Numbers (`[+-]digits[.digits][e±digits]`, finite, normal or zero; canonical spellings kept verbatim in XLSX), ISO dates/date-times (`yyyy-mm-dd`, `T` or space, `hh:mm:ss`, fraction, `Z`) as serials in the workbook's date system with a `yyyy-mm-dd` or `yyyy-mm-dd hh:mm:ss` number format, `TRUE`/`FALSE` in any case, else text. `infer(false)` and `infer_numbers/booleans/dates` switch inference off. |
| Formulas | Off by default (an untrusted CSV cannot plant formulas); `formulas(true)` writes `=…` fields as formulas. |
| Empty fields | Nothing on an absent cell; the value of an existing cell is cleared and its formatting kept. |
| Refused characters | C0 controls other than tab and line breaks, and U+FFFE/U+FFFF: no cell format stores them (`invalid_input`). |
| Cell length | More than 32,767 characters: `invalid_value`. |
| Grid | The rectangle must fit the format's grid from the origin (XLSX/XLSB 1,048,576 × 16,384; XLS 65,536 × 256): `out_of_range`. |
| Limits | The transaction's `SemanticLimits`: the source and staged payload (field bytes plus 64 per field) ≤ `max_staged_bytes` (16 MiB), fields ≤ `max_expanded_targets` (65,536), retained snapshots ≤ `max_snapshot_bytes`. Checked while parsing, before any native work. |
| XLS | Dates stay ISO text (`Summary::dates_as_text`): the unified XLS write has no number-format change yet. Importing over an existing XLS formula cell is refused (`unsupported_capability`), as the XLS cell setter refuses it. |

Errors are typed: the focused `import::Error`
(`InvalidOption`/`Encoding`/`UnterminatedQuote`/`InvalidCharacter`/`FieldTooLong`,
content-free, classified `invalid_argument`/`invalid_input`/`invalid_value`)
travels as `error::Source::Import`.

### Evaluation (`litchi_ops::calc`, feature `eval`)

- `File::evaluate(path, &calc::Options) -> calc::Evaluated` (value and
  `Origin::{Literal, Cached, Computed}`); cached values are preferred unless
  `Options::recalculate(true)` or the formula has no cache.
- `File::recalculate(&calc::Options) -> calc::Report`: every formula cell in
  sheet, row, column order with its text, cached value, computed value and
  `Status::{Match, Mismatch, Uncached, Unsupported, Volatile}`. It is the
  optional recalculation verb; it is a read and never writes caches.
- The engine runs over a snapshot loaded per worksheet through typed
  per-format loaders (`format/{xlsx,xlsb,xls}/calc.rs`): the generic
  `WorkbookTrait` of the XLS handle reports empty values and wrong sheet
  indexes, so it is not used. An XLSB sheet whose source owner only has the
  cached-value fallback grid refuses (`unsupported_capability`) instead of
  reporting no formulas.
- No ambient clock or randomness: formulas calling `NOW`, `TODAY`, `RAND`,
  `RANDBETWEEN` or `RANDARRAY` keep their cached value in every mode; without
  one, `evaluate` refuses. The engine's futures are polled once with a no-op
  waker (no runtime); network functions are not built.
- `_xlfn.`/`_xlws.` prefixes are removed before evaluation; a top-level blank
  result reads 0, as in Excel. Engine errors that are not spreadsheet error
  values are reported as unsupported, never guessed.
- Bounds: `Options::max_cells` (1,048,576) bounds loaded cells and engine
  reads; `Options::max_depth` (64) bounds a formula chain. The engine
  recurses per chain link (about 3 KiB of stack per link optimized, about
  19 KiB in a debug build, measured), so the bound protects the stack.
  `evaluate` resolves forward chains by first computing the sheet's earlier
  formulas in row-major order, so a 399-link running total evaluates under
  the default bound; backward chains longer than the bound are
  `limit_exceeded`.

### Capabilities

`capabilities::Verb` gains `Evaluate` (a read) and `Import` (an edit).
Both are supported for XLSX, XLSB and XLS adapters (import also needs a
retained source; evaluate needs the `eval` feature, otherwise it refuses with
"build litchi-ops with the eval feature"); other formats refuse with "this
format has no worksheets". The `computedValue` property gap now points at
the evaluation verbs and keeps the follow-up "project computed values in cell
views". Goldens re-blessed (`LITCHI_BLESS=1`, all-features and the
non-encryption feature set): only the two new verbs and the gap text changed.

### Surfaces

- Facade: `litchi::import` and (feature `eval`) `litchi::calc`; the facade
  `eval` feature turns on `litchi-ops?/eval`. `litchi::eval` still re-exports
  the engine crate, which is why the unified module is named `calc`.
- CLI: `litchi import FILE PATH [SOURCE] [--file F | --stdin] [--format
  csv|tsv] [--delimiter C] [--header] [--start-cell A1] [--encoding E]
  [--no-infer] [--formulas] [--max-cells N] [--max-bytes N] [--out F]`
  (OfficeCLI's shape; delimiter precedence `--delimiter` > `sep=` >
  `--format` > `.tsv`/`.tab` > CSV), `litchi evaluate FILE PATH
  [--recalculate]`, `litchi recalculate FILE [--all]`, all with `--json`.
- Python: `File.import_csv(target, data, *, delimiter, quote, header,
  encoding, infer, formulas, max_cells, max_bytes) -> Commit` (`value` is the
  summary dict; `import` is a Python keyword), `File.evaluate(path, *,
  recalculate=False) -> dict`, `File.recalculate() -> dict`, GIL released,
  stubs with `_ImportSummary`, `_Evaluated`, `_Recalculation` TypedDicts.

### Leaf changes in `litchi-eval` (non-iWork, not otherwise owned)

1. Panic fix: `parse_criteria` sliced `&s[..2]` and panicked on non-ASCII
   criteria (`COUNTIF(.., "Köln")`, found by the corpus run). Now prefix
   stripping; regression test added.
2. Excel arithmetic and comparison coercion in `bin_op`: error operands
   propagate, a blank is 0/""/FALSE, booleans are 1/0, numeric text is a
   number, non-numeric text is `#VALUE!`, division by zero is `#DIV/0!`
   (was the text "Division by zero"), comparisons order numbers < text <
   booleans and compare text without case.
3. `FormulaEvaluator::with_max_depth` and the typed `DepthLimit` error.

### Topology

New optional edge `litchi-ops -> litchi-eval` (feature `eval`):
`tools/crate_boundaries.json`, policy count 269 → 270 in
`tools/test_check_crate_boundaries.py`, ADR 0024 amendment. `litchi-eval` is
a runtime-neutral bulk package, so the non-iWork gate's partition is
unchanged.

## 2. Parity against cached values

`cargo run --release -p litchi-ops --features xlsx,xlsb,xls,eval --example
eval_parity` over `test-data/` (326 workbooks; 38 byte-identical copies
counted once; 81 with formulas). Evidence:
[results/0266-csv-eval/eval-parity.json](results/0266-csv-eval/eval-parity.json)
(totals, per format, per function, per file, every mismatch, up to three
unsupported samples per function, failures).

Tolerance: numbers agree when `|a-b| <= 1e-9 * max(1, |a|, |b|)`; text,
booleans and error values exactly; an empty value equals empty text.

| Format | Formulas | Match | Mismatch | Uncached | Unsupported | Volatile | Pass rate of compared |
|---|---|---|---|---|---|---|---|
| XLS | 1,884 | 1,139 | 97 | 115 | 498 | 35 | 92.2% |
| XLSB | 317 | 314 | 1 | 0 | 2 | 0 | 99.7% |
| XLSX | 9,881 | 1,699 | 721 | 0 | 7,461 | 0 | 70.2% |
| **All** | 12,082 | 3,152 | 819 | 115 | 7,961 | 35 | 79.4% (26.3% of all cached) |

Per function (functions with at least 15 formulas; a formula counts for every
function it calls; 184 distinct functions in the JSON):

| Function | Formulas | Match | Mismatch | Unsupported | Pass rate of compared |
|---|---|---|---|---|---|
| IF | 7,474 | 251 | 297 | 6,890 | 45.8% |
| ROW | 5,477 | 9 | 0 | 5,468 | 100% |
| INDIRECT | 3,801 | 0 | 0 | 3,801 | — |
| (operators only) | 2,301 | 1,838 | 66 | 382 | 96.5% |
| VLOOKUP | 1,778 | 4 | 0 | 1,774 | 100% |
| ISNA | 1,722 | 4 | 2 | 1,716 | 66.7% |
| FORMULATEXT | 202 | 35 | 167 | 0 | 17.3% |
| ROUND | 195 | 175 | 6 | 14 | 96.7% |
| TEXT | 154 | 50 | 64 | 40 | 43.9% |
| SUM | 149 | 105 | 16 | 28 | 86.8% |
| CONCATENATE | 93 | 90 | 1 | 0 | 98.9% |
| LAMBDA | 86 | 1 | 3 | 82 | 25.0% |
| TRIM | 70 | 13 | 54 | 0 | 19.4% |
| COLUMN | 60 | 6 | 54 | 0 | 10.0% |
| ANCHORARRAY | 56 | 2 | 9 | 45 | 18.2% |
| SPLITSITEMS | 54 | 0 | 54 | 0 | 0% |
| ISERROR | 51 | 10 | 3 | 38 | 76.9% |
| ERROR.TYPE | 48 | 4 | 41 | 2 | 8.9% |
| SUBTOTAL | 45 | 32 | 9 | 4 | 78.0% |
| MONTH | 45 | 5 | 3 | 1 | 62.5% |
| ABS | 36 | 20 | 0 | 16 | 100% |
| TODAY | 35 | 0 | 0 | 0 | — (volatile, cache kept) |
| AVERAGE | 26 | 8 | 2 | 16 | 80.0% |
| AND | 25 | 9 | 12 | 0 | 42.9% |
| NOT, COUNTIFS, INT, FV, LARGE, ABS, WEEKNUM, SUBSTITUTE, AVEDEV, FACT | 15–23 each | | 0 mismatches | | 100% |
| COUNTIF, MEDIAN, MODE, ISNUMBER, STDEV, VAR, UPPER | 15–17 each | | 1–2 each | | 82–93% |

What drives the numbers (the mismatches are all listed in the JSON):

- One XLSX workbook (`StructuredRefs-lots-with-lookups.xlsx`, 7,274 formulas)
  uses whole-column references (`Query!$A:$N`), structured references
  (`Table[#All]`), `INDIRECT` and `&`; it accounts for most unsupported
  formulas. `DynamicArrayFixture.xlsx` and `LambdaAndRelatedFunctions.xlsx`
  account for most XLSX mismatches (spilled arrays, implicit intersection
  `@`, `LAMBDA`, `FORMULATEXT` rendering).
- Engine-scope causes: no `&`, `^`, `%`, unions or array arithmetic;
  no whole-row/column references; no defined names or tables (a name such as
  `name1` even parses as a cell reference); many functions return descriptive
  text instead of Excel's error values (reported as unsupported); `DATE`/`DAY`/
  `MONTH` are off by one before 1900-03-01; text-to-number coercion inside
  aggregates (`SUM("1")`, `MIN("-100")`) and `DOLLAR`'s currency symbol
  differ.
- Not evaluated (20 files): 10 XLS workbooks whose worksheet grid the unified
  XLS read already refuses (`corrupt_file`; `view text` fails the same way),
  2 XLSB workbooks with only the cached-value fallback grid, 4 encrypted
  workbooks, and 4 files that do not open.

The test `recalculation_matches_cached_results_in_each_format` pins three
fully matching workbooks (XLSX 40, XLSB 30, XLS 7 of 12 cached formulas).

## 3. Measurements

Release build, `taskset -c 5`, ABBA (leaf, unified, unified, leaf) × 7 rounds
after one warm-up per leg; medians. `cargo run --release -p litchi-ops
--features xlsx,xlsb,xls,eval --example m4_bench -- 10000 7`. Evidence:
[results/0266-csv-eval/m4-bench.json](results/0266-csv-eval/m4-bench.json)
(every sample).

| Scenario | Leaf leg | Unified leg | Unified / leaf |
|---|---|---|---|
| import XLSX, 10,000 rows × 5 (50,000 cells, 454 KB CSV) | 91.2 ms (109.6k rows/s) | 92.6 ms (108.0k rows/s) | 1.015 |
| import XLSB, 10,000 × 5 | 96.5 ms (103.6k rows/s) | 98.6 ms (101.5k rows/s) | 1.021 |
| import XLS, 1,000 × 5 | 2,722 ms (367 rows/s) | 2,553 ms (392 rows/s) | 0.938 |
| recalculate XLSX, 2,000 rows × 3 uncached formulas | 6.10 ms | 5.80 ms | 0.951 |

- The leaf import leg stages the already-typed values in one native edit (XLS:
  the same 1,024-cell chunks) and commits; the unified leg also decodes,
  parses and infers the CSV. The whole CSV front end costs 1.5–2% on XLSX and
  XLSB.
- XLS is dominated by the leaf: BIFF8 cell insertion grows with the
  worksheet (about 2.7 s for 5,000 cells on both legs). That is a
  `litchi-xls` cost, not an import cost; follow-up F5.
- The evaluation leaf leg is `litchi_eval::FormulaEvaluator` directly over the
  XLSX handle's `WorkbookTrait`; the unified leg loads a snapshot, evaluates
  every formula and compares with caches, and is not slower.

## 4. Deltas vs OfficeCLI

| OfficeCLI | Litchi | Disposition |
|---|---|---|
| `--header` also sets an AutoFilter and a frozen pane | header row written as text only | deferred: AutoFilter/freeze are P1 surfaces without unified writes (F3) |
| `--decimal ,` and thousands grouping (`1,234`) | plain numbers only; such fields stay text | deferred (F4) |
| `=…` fields always become formulas (and are evaluated and cached) | opt-in `formulas`; no cache written | wontfix default (security); cache write is F2 |
| date-times use `yyyy-mm-dd` | `yyyy-mm-dd hh:mm:ss` for date-times | intentional: no hidden time |
| `.xlsx` only | XLSX, XLSB, XLS | superset |
| BOM stripped, UTF-8 only | UTF-8/UTF-16 LE/BE by mark or option | superset |
| `computedValue` on `get`/`query` | `File::evaluate`/`recalculate`; not in node views | deferred (F1) |
| batch `import` command | not integrated in batch results (`batch/execute.rs`) | deferred (F6) |

## 5. Gaps and follow-ups

- F1 `computedValue` projected on demand in cell node views (per-node
  evaluation needs a shared snapshot to stay within the read budget).
- F2 recalculation that writes formula caches back (needs leaf cached-value
  writes for XLSX/XLSB/XLS).
- F3 import `header` AutoFilter/freeze pane once those P1 writes exist.
- F4 locale numbers (decimal comma, grouping).
- F5 `litchi-xls` cell insertion cost (quadratic-looking growth) and its
  4,096 staged-change ceiling.
- F6 batch `import`/`evaluate` command items.
- F7 engine scope in `litchi-eval`: `&`/`^`/`%` operators, whole-row/column
  references, defined names and tables, Excel error values instead of
  descriptive text, pre-1900-03-01 date serials, dynamic arrays.
- F8 ODS import and evaluation (ODS has no unified write path; its OpenFormula
  syntax also needs translation).
- F9 the XLS worksheet grids the unified read refuses (10 corpus files).
- The `litchi-eval` engine API is async with boxed futures and a boxed
  `WorkbookTrait`; the unified layer drives it synchronously. Removing the
  boxing is engine work, not facade work.

## 6. Verification

Commands and exit codes are in the step report; all ran on this branch:
`cargo fmt --all --check`; `cargo clippy -p litchi-eval -p litchi-cli -p
litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D
warnings`; `cargo test` for `litchi-ops`, `litchi-eval`, `litchi-cli` and
`litchi` with `--all-features`; `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps
--all-features` for the changed crates; `python3
tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`;
`tools/check_api_glossary.py`; `crates/pyo3-litchi/tests/run_tests.sh`;
the capability delta regeneration.

## 7. Not claimed

- Engine fidelity beyond the measured corpus rates; the per-function rates
  come from 81 corpus workbooks, some functions from a single file.
- That import is faster than the leaf: the legs differ by the CSV front end,
  within noise of each other.
- ODS import or evaluation.
