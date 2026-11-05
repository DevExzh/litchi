# 0354 — `litchi-sheet-text` leaf and structured metadata keywords

- Status: implemented on `uo2/sheet-text-meta`, for coordinator review.
- Date: 2026-10-20.
- Decision: owner decision [0350](0350-owner-api-debt-decision.md), memo
  ([0331](0331-owner-decision-memo.md)) items 26 and 20 (0343 §4 D7). Both are
  required work with every consumer migrated in the same change.

## 1. Item 26: `litchi::sheet::text` becomes the `litchi-sheet-text` leaf

### 1.1 Which crate owns delimited text

The CSV, TSV, PRN, SYLK, DIF and fixed-width readers and writers (about
2,300 lines) lived in the facade (`crates/litchi/src/sheet/text`). They move
into a **new** leaf, `crates/litchi-sheet-text`, not into `litchi-sheet`:

- `litchi-sheet` is the format-neutral vocabulary crate ADR 0002 places
  under every spreadsheet format ("concrete workbook, worksheet, value and
  formula types belong to XLS, XLSB, XLSX, ODS, or other format crates").
  XLSX, XLSB and DrawingML depend on it; putting four concrete text formats
  and their readers there would make every OOXML spreadsheet build carry
  them and break that rule.
- No other leaf owns delimited text. `File::import` (record 0266) had its
  own private CSV/TSV tokenizer in `litchi-ops`, which is not a leaf.

The new crate depends on `litchi-core` (cell values, the sheet traits, BOM
vocabulary) and `litchi-sheet` (grid bounds); `litchi` (behind the new
`sheet-text` feature, re-exported as `litchi::sheet::text` beside
`litchi::sheet::{view, eval}`) and `litchi-ops` depend on it. ADR 0024 gains
an amendment; `tools/crate_boundaries.json` records the four edges (68
packages, 286 internal dependency declarations); the non-iWork gate counts
the crate among its bulk packages (46 to 47); CI tests it and the
perf-baseline workflow watches it; `docs/FEATURE_MATRIX.md` lists it.

### 1.2 Names (ADR 0034)

Every reader and writer is kept; no old path is aliased. The `sheet-text`
and `facade-sheet-text` glossary profiles ban each old name.

| Old (`litchi::sheet::text`) | New (`litchi_sheet_text`, `litchi::sheet::text`) |
|---|---|
| `TextWorkbook::{open, from_path_with_config, from_reader, from_bytes}` | `Workbook::{open, open_with_options, from_reader, from_reader_with_options, from_bytes, from_bytes_with_options}` |
| `TextWorkbook::{config, worksheet_name, set_worksheet_name, set_data, detected_bom, write_to}` | `Workbook::{options, name, rows, bom, worksheet, write_to, to_bytes}`, built with `Workbook::new(rows, options).with_name(..)`; the workbook has no `&mut self` method (record 0351 §2.1) |
| `TextWorksheet`, `TextCell`, `TextWorksheetIterator`, `TextRowIterator`, `TextCellIterator` | `Worksheet<'a>` (borrows the rows; the old one copied every row per `active_worksheet` call), `Cell`, `Worksheets`, `Cells`, `Rows`: the associated types of record 0353's statically dispatched sheet traits (no `Box<dyn>`; ported in the merge of trunk `482f0a18eb`) |
| `TextConfig`, `formats::DelimitedConfig` | one `delimited::Options` (`csv`, `tsv`, `prn`, `pipe`; infallible consuming setters) over a `delimited::Dialect` |
| `TextParser` | `delimited::Tokenizer` |
| `formats::{read,write}_{delimited,sylk,dif,fixed_width}`, `{Sylk,Dif,FixedWidth}Config`, `{Sylk,Dif,FixedWidth}Parser` | `{delimited,sylk,dif,fixed_width}::{read, write, Options}` |
| (none: no limits, boxed string errors) | `SourceLimits` (`max_input_bytes`, 512 MiB), `SemanticLimits` (`max_cells`, 2^24), `Encoding`, a content-free typed `Error` |

### 1.3 One delimited tokenizer

`File::import` and the text workbook were two CSV parsers with different
grammars. `delimited::Tokenizer` is now the only one: it is import's
borrowing, RFC 4180-style tokenizer moved into the leaf, with the dialect
switches the text workbook needs (`comment`, backslash `escapes`) and the
one import needs (`refuse_controls`). `File::import` decodes with the leaf's
`Encoding` (`litchi_ops::import::Encoding` is that type) and splits the
`sep=` line with `delimited::declared_delimiter`; its inference, limits and
error codes are unchanged (its 19 unit and 13 integration tests pass
unchanged).

The tokenizer classifies each byte with one table lookup (two 256-byte
tables built per tokenizer) instead of comparing it with the delimiter, the
line breaks, the quote and the refused characters, so the import route got
faster, not slower. Parse-only ABBA, release (`lto = true`), `taskset -c 5`,
10,000 records × 5 fields, 41 parses per run, six runs per binary
(medians of the per-run medians):

| Input | Before (import's own parser) | After (`delimited::Tokenizer`) | After / before |
|---|---|---|---|
| plain, 383,378 bytes | 882.2 µs | 796.3 µs | 0.903 |
| quoted (a `,` inside one quoted field per record), 413,378 bytes | 953.1 µs | 841.7 µs | 0.883 |

The harness is a scratch `#[ignore]` test calling `import::parse`
([results/0354/parse_bench.rs.txt](results/0354/parse_bench.rs.txt), placed
at `crates/litchi-ops/src/import/parse_bench.rs` for the run, not built by
the tree), built with `cargo test --release -p litchi-ops --features xlsx
--lib --no-run`; the two binaries' SHA-256 are
`d97e0be8861a4d03460f6feda6778d12c5c87a47bb6c061aee106d5eef7104a6`
(before, trunk `5ff605692e`) and
`20dec77681001dcaf22911f18fb5596b498f22ad0839f89de1d243ecb6cf0274`
(after). Every sample: [results/0354/parse-abba.tsv](results/0354/parse-abba.tsv).
The parse is about 1.5–2% of an XLSX import (record 0266), so the end-to-end
import change is within noise; it was not re-measured.

### 1.4 Behaviour changes (all toward strictness or fidelity)

| Area | Before | After |
|---|---|---|
| Unterminated quote | field kept silently | `Error::UnterminatedQuote { line }` |
| Lone `\r` | dropped (`a\rb` read `ab`) | ends a record, as in import |
| Quote inside an unquoted field | toggled quoting (`a"b,c"d` read one field) | literal, as spreadsheet applications read it |
| Invalid UTF-8 | replaced with U+FFFD | `Error::Encoding { offset }`; UTF-16 sources (by BOM or `Encoding`) now read |
| `strip_bom: false` | kept the BOM bytes as text | removed (the BOM is reported by `Workbook::bom`) |
| Writer quoting | a leading `#` (comment) or a backslash (with escapes) was not quoted, so they did not read back | quoted, so every written row reads back (test) |
| `write_bom(Some(Utf16*))` | a UTF-16 mark before UTF-8 bytes | the output is UTF-16; UTF-32 is `Error::InvalidOption { Bom }` |
| `DelimitedConfig` defaults | no BOM stripping on read, a UTF-8 BOM on write | merged into `delimited::Options`: BOM detected on read, none written |
| SYLK | `;` in a string split it; `""` stayed doubled; `C;…;K3;E1+2` lost `3`; `Y99999999` allocated rows until memory ran out | `;;` and `""` escape and read back; `K` is the formula's cached value; out-of-grid coordinates are `Error::OutOfGrid`; padding cells count against `max_cells` |
| Fixed width | byte slicing (panicked inside multi-byte text), runs of one position | columns cut by characters; detected boundaries fall in the middle of whitespace runs |
| Errors | `Box<dyn Error>` strings, some naming a worksheet | typed, content-free `Error` (`MissingWorksheet` names none) |

## 2. Item 20: structured `Metadata.keywords`

### 2.1 The type and the split rule

`litchi_core::Metadata::keywords` is `Vec<String>`: empty means none,
`#[serde(default, skip_serializing_if = "Vec::is_empty")]`, so JSON views,
YAML front matter and the `litchi-py` `metadata()` dictionary write an array
(and omit it when empty). The rule for formats that store one string lives
in `litchi_core::metadata::keywords`:

- **Read** (`split`, `Keywords::list`): split at every `;` and `,`, trim
  each piece, drop empty pieces. An OOXML `cp:value` child is one keyword
  whatever it contains.
- **Write** (`join`, `Keywords::from_list`): join with `, ` (the separator
  ODF's former projection, OfficeCLI's examples and python-docx users write).
  A keyword the split would not give back (it holds `;` or `,`, or has
  surrounding whitespace) cannot live in a plain string: `join` refuses it
  (`keywords::Error::Separator`), while OOXML writes it as a `cp:value` child.
- **Round trip**: a writer given the list its source already reads as keeps
  the source text byte for byte, whatever separators the producer used.

### 2.2 Per format

| Format | Read | Write |
|---|---|---|
| OOXML (DOCX, XLSX, PPTX, XLSB) `cp:keywords` | `Keywords::list` of the retained mixed content | unchanged list → the source `Keywords` (exact bytes; `properties::Core::merged`); changed → `Keywords::from_list`; empty → element removed |
| ODF (all ten) `meta:keyword` | one entry per element | one element per keyword in the shared `metadata::Edit::set_keywords`, the `MetaXmlPatch` mutable-save path (which also stops leaving duplicate elements behind), the ODT/OTH builders and ODM |
| OLE SummaryInformation (DOC, XLS, PPT) PIDSI `KEYWORDS` | `split` | unchanged list → kept; changed → `join` or a typed refusal; empty → removed |
| RTF `\keywords` | `split` | — (RTF metadata has no unified write) |
| Pages `Keywords` plist key | `split` | `metadata::Edit::set_keywords(&[String])`: unchanged list → kept; changed → `join` or `Error::InvalidKeyword` |
| Keynote, Numbers | `split` | refused as before (0343) |

### 2.3 The unified surface

- `litchi_ops::value::ValueType::List` (schema type `list`),
  `Scalar::List`, `node::Value::List` (a JSON array in node views; a list of
  `str` in Python node dictionaries), and `props::PropValue::List` (a JSON
  array of strings in batches and props objects, a Python `list`/`tuple` of
  `str`, `From<Vec<String>>`).
- Every root/metadata `keywords` key (DOCX, XLSX, XLSB, XLS, PPTX, ODT, ODS,
  ODP, ODI and the ODT-family formats) is `ValueType::List`. `set` takes a
  list, or text in the separator form (the CLI's `--prop keywords=a,b`), or
  `null` to remove. Query comparison holds when one keyword does
  (`[keywords=review]`); `!=` when none equals; `exists` when there is one.
- The capability and registry goldens changed only in the 26 `keywords`
  type entries (`string` → `list`), re-blessed in both feature sets.
- Field refresh: DOCX `KEYWORDS`/`DOCPROPERTY Keywords` and ODF
  `text:keywords` show the list joined with `, `.
- Markdown front matter (`with_metadata`) writes `keywords:` as a YAML
  sequence. The facade Markdown golden (`crates/litchi/tests/golden/
  markdown_corpus.tsv`) re-records the metadata+html columns of 43
  fixtures: 11 now list their keywords, and 32 whose keyword text was empty
  (`keywords: ''`) omit the key, 13 bytes shorter; the regeneration checked
  both conditions per fixture, and every other column is unchanged.
- The OfficeCLI corpus outcomes are unchanged (the three keyword `set` lines
  stay `ok`); the capability delta changed only in its golden hashes.

## 3. Evidence

- `litchi-sheet-text`: 21 unit tests and 11 doc tests (tokenizer grammar,
  dialect validation, comments/escapes/trim, writer round trip, BOMs and
  UTF-16, limits, SYLK escapes/formulas/out-of-grid, DIF, fixed width with
  non-ASCII text).
- `litchi-core`: list/split/join/from_list tests.
- `litchi-ooxml-common`: an unchanged list commits nothing and keeps
  `a; b,c <cp:value>d</cp:value>` byte for byte; a changed list writes
  `a, b, <cp:value>x, y</cp:value>`.
- `litchi-ole-common`: the unchanged list keeps `q3; revenue`, a change
  writes `q4, revenue`, and `a,b` is refused atomically.
- `litchi-odf-common`: one element per keyword, unchanged lists untouched,
  empty list removes every keyword.
- `litchi-ops/tests/metadata_keywords.rs`: list and separated-text round
  trips in DOCX, XLSX, PPTX, XLSB, DOC, XLS, PPT, ODT, ODS and ODP; JSON
  arrays in node views and `Metadata` serde; batch JSON arrays; byte-identical
  no-op for the unchanged list; typed refusals in the OLE formats.
- `litchi-py`: `MetadataKeywordsTests` (ten formats, root format list,
  `str`-only lists); `.pyi` stubs (`keywords: List[str]`, `_PropValue` and
  `_FormatValue` include `List[str]`).

## 4. Deltas vs OfficeCLI

| OfficeCLI | Litchi | Disposition |
|---|---|---|
| `get /` writes `"keywords": "q4,review,strategy"` (the stored string) | a JSON array `["q4", "review", "strategy"]` | wontfix: owner decision 0350 (structured metadata) |
| `--prop keywords="q4,review,strategy"` stores the text verbatim | stored as `q4, review, strategy` (split, then joined) | wontfix: one rule in every format; a list equal to the stored one keeps the stored text |
| `KEYWORDS` field shows the stored string | the list joined with `, ` | wontfix, same reason |

## 5. Not claimed, follow-ups

- The import route was measured at the parse level only (§1.3).
- RTF has no unified metadata write, so `\keywords` is read-only (unchanged).
- The design document's §16.4 wording is now closed: the text tooling is an
  independent `litchi-sheet-text` leaf, re-exported by `litchi::sheet::text`,
  and unified `import` consumes that leaf. The historical "re-homed behind
  `import`" wording is superseded by the implemented design text.

## 6. Gates

After merging trunk `e652864de8` (record 0351), each with exit code 0:
`cargo fmt --all --check`; warning-denied clippy of the 15 changed leaf
crates, `litchi-ops`, `litchi` and `litchi-py` (`--all-targets
--all-features`); `tools/check_feature_matrix.py` (31 facade steps) and
`--ops --clippy --all-targets --only ops` (22 steps); `cargo check
--workspace --exclude litchi-ods --all-features --all-targets`; `cargo test`
of the 15 changed crates, `-p litchi --all-features`, `-p litchi-ops
--all-features --lib`, `--doc` and every `--test` in chunks; `litchi-ods`
lib clippy only (its suite needs a purged fixture); rustdoc with
`-Dwarnings`; `check_crate_boundaries.py`; `non_iwork_gate.py verify`;
`check_api_glossary.py` (all profiles, with the new `sheet-text` and
`facade-sheet-text`); `crates/pyo3-litchi/tests/run_tests.sh`;
`tools/perf-baseline` `cargo check --locked --all-targets`; the tool unit
tests; `officecli_capability_delta.py --check`; `officecli_corpus.py
--check` after the corpus bless; `unified_bench_matrix.py --check` (0354 is
listed under `evidence_without_rows`).
