# 0274 — Template merge (plan step 20b, M4 engine 2 of 3)

Status: implemented on `uo2/s20-merge` (cut from `feat/unified-ops-api` at
395d0ffa9f); awaiting coordinator review and merge.

Plan: [0248](0248-coordinator-restart-plan.md) Phase E step 20. Design: §13
P0 "template `merge` (the docx `mail_merge` module owns settings/recipient
data only — it generates nothing)" and the naming note (`merge` always means
template/mail merge, OfficeCLI parity). Goal: verb-name decision "`merge`
means template merge". Glossary: [API_GLOSSARY.md](../API_GLOSSARY.md),
section "Unified template merge (record 0274)".

## 1. What landed

### API (`litchi_ops::merge`, re-exported as `litchi::merge`)

- `File::merge(&merge::Record, &merge::Options)` and
  `File::merge_with_limits(.., &SemanticLimits)` return
  `transaction::Commit<merge::Summary>`; `transaction::Edit::merge` is the
  same mutation inside a closure transaction (it composes with `set` and
  other edits). The template handle is never changed, so one opened
  template fills any number of records.
- Data: `merge::Data::{from_json, from_json_with_limits, from_csv,
  from_csv_with}` give the records of a source in order; `merge::Record` is
  one record (`insert`, `get`, `iter`, `FromIterator`).
- `merge::Options` (`missing(Missing::{Keep, Empty})`, `fields(bool)`),
  `merge::Summary` (`replacements`, `fields`, `skipped_fields`, `keys`,
  `unresolved`; `Debug` prints counts only), and the focused, content-free
  `merge::Error` (`Json { line, column }`, `TooDeep`, `NotAnObject`,
  `NotARecord { record }`) travelling as `error::Source::Merge`
  (`invalid_json`; `TooDeep` is `limit_exceeded`).
- Formats: DOCX, PPTX and XLSX. Every other format refuses with
  `unsupported_capability` naming the format.
- Atomicity: one merge is one transaction operation and at most one native
  commit. An exact no-op (no placeholder or field matches, or every match is
  a missing key under `Missing::Keep`) makes no native commit: the result is
  the template handle and `patch.is_noop()` holds.

| Aspect | Behaviour |
|---|---|
| Placeholder | OfficeCLI's `\{\{\s*(\w[\w.\-\[\] ]*?)\s*\}\}`; the key keeps inner spaces; one left-to-right pass, values never re-matched. `\w` is `char::is_alphanumeric` or `_`. |
| DOCX scope | Every paragraph of the body's paragraphs, tables and block content controls (text boxes included), and of the headers, footers, footnotes, endnotes and comments the main part relates. |
| DOCX formatting | A placeholder may span runs. Its value takes the place of its first character and that run's formatting; the rest is removed from the runs it spans, which keep their own formatting and every other child. Tabs and line feeds in values become `w:tab`/`w:br`; `xml:space="preserve"` is added where needed. |
| DOCX fields | `MERGEFIELD` (simple `w:fldSimple` and complex `w:fldChar` forms) whose name the record has (exact, then the single ASCII-case-insensitive match) becomes one run with the value, formatted as the field's first result run. `\b`/`\f` apply to non-empty values; `\* Upper/Lower/Caps/FirstCap` apply. Fields whose markup holds other content (bookmarks, nested fields, drawings) or spans paragraphs stay and are counted in `skipped_fields`. Other fields are untouched. |
| PPTX | Every paragraph of every slide shape (group members included), run by run through the native text edits, so each run keeps its formatting. Line breaks and fields are boundaries. A value with a line break is `invalid_value`. |
| XLSX | Every text-value cell of every worksheet; formulas are never touched; a cell keeps its style. |
| Missing keys | Listed in `Summary::unresolved`; `Missing::Keep` (default, OfficeCLI) keeps the placeholder or field, `Missing::Empty` removes it (Word's behaviour). |
| JSON data | An object is one record; an array of objects is several. Strings as themselves, numbers as written (`5200.00` stays), `true`/`false`, `null` as nothing, objects/arrays as indented JSON; nested members also as `a.b` / `a[0]` keys with literal keys winning (OfficeCLI's flattening). A small RFC 8259 reader keeps number spellings; depth ≤ 128. |
| CSV/TSV data | Header row of keys; one record per later row; short rows give empty values; blank rows are skipped; parsed by the `import` front end without inference. |
| Values | Characters XML cannot hold (C0 controls other than tab/LF/CR, U+FFFE/U+FFFF, U+FFFC) are `invalid_value`; CR and CRLF become LF. |
| Limits | The transaction's `SemanticLimits`: replacements ≤ `max_expanded_targets` (65,536); their text plus 64 bytes each ≤ `max_staged_bytes` (16 MiB); data sources and their rendered records ≤ `max_staged_bytes`. The DOCX leaf also bounds a substitution by its 65,536-operation ledger and 16 MiB replacement charge. |
| Policies | Signed and locked files refuse as every write does (`document_protected`); a signature-removal-authorized DOCX merges into its unsigned copy; DOCX protection and tracked revisions refuse a changed publication in the leaf; `retain_source(false)` refuses. |

### DOCX leaf: story substitution (`litchi_docx::document::substitute`)

The leaf had no way to rewrite text across runs; the main-document edit
offers per-element text replacement (one full rewrite and rescan per
call), which is quadratic for thousands of placeholders, and nothing for
headers or fields as a whole. Added (non-iWork leaf, small surface):

- `document::Edit::story()` reads the main document's `Story`: one
  `Paragraph` per paragraph with its text (`w:t` text; `\t`, `\n` and
  `OBJECT` = U+FFFC for fixed content and field markers) and its `Field`s
  (range, instruction, `FieldKind`, `is_replaceable`).
- `document::Edit::substitute(&Story, &[Replacement])` applies byte-range
  replacements: text ranges are rewritten in the `w:t` elements they cover
  (formatting-preserving as above); a range equal to a replaceable field's
  range replaces the field's markup with one run. All changed body blocks are
  spliced in one rewrite and one rescan, recorded as exact
  `Operation::ReplaceElement` operations (so the patch inverts and replays
  exactly), with an exact byte readback per block and a semantic readback of
  every paragraph's text.
- `document::Edit<Package>::{story_parts, part_story, substitute_part}` do
  the same for header, footer, footnotes, endnotes and comments parts, staged
  as exact part replacements (`replace_part`) that publish atomically with
  the main-document edit.
- Source-backed edits refuse `substitute` (unmanaged-only, as every
  element-path edit); the unified layer materializes the owning package of a
  source-backed handle once and keeps it on the handle (`Docx::promoted`,
  also used by the other promoting writes).

The unified `changed()` check now also treats a DOCX package-graph change
as a change (a header-only merge changes no main-document XML).

### Capabilities

`capabilities::Verb::Merge` (an edit): supported for DOCX, PPTX and XLSX
adapters (with a retained source); other formats refuse with "template
merge fills DOCX, PPTX and XLSX templates", follow-up F1. Goldens
re-blessed (`LITCHI_BLESS=1`, all features and the non-encryption feature
set): only the new verb changed. The OfficeCLI capability delta table is
unchanged (it compares the property registry, which merge does not touch).

### Surfaces

- CLI: `litchi merge TEMPLATE DATA --out OUT [--format json|csv|tsv]
  [--missing keep|empty] [--no-fields] [--max-replacements N]
  [--max-bytes N] [--force]`, and OfficeCLI's `litchi merge TEMPLATE OUTPUT
  --data JSON|FILE [--force]` (a `.json`/`.csv`/`.tsv` path that exists is
  read, anything else is inline JSON). Several records need `{n}` (the
  one-based record number) in the output name. Existing outputs are refused
  without `--force`, all checked before any is written; each output is an
  atomic save. `--json` prints OfficeCLI's `output`, `replacedKeys`,
  `unresolvedPlaceholders` plus `replacements`, `fields`, `skippedFields`;
  several records print `{"records": n, "outputs": [...]}`. Text output and
  the unresolved-placeholder warnings follow OfficeCLI.
- Python: `File.merge(data, *, missing="keep", fields=True,
  max_replacements=None, max_bytes=None) -> Commit` (`data` is a mapping,
  serialized by the standard `json` module, or JSON text as str/bytes; one
  record), GIL released; `Commit.value` is the `_MergeSummary` TypedDict.
  Stubs updated; `test_merge.py`.

## 2. Measurements

Release build, `taskset -c 5`, ABBA (leaf, unified, unified, leaf) × 11
rounds after one warm-up per leg; medians. Machine shared with other agents
(load average about 10 on 32 cores). `cargo build --release -p litchi-ops
--features docx --example merge_bench; taskset -c 5
target/release/examples/merge_bench 1000 11` (binary SHA-256
`a670118b…dbdcf0`, built from this branch). Evidence:
[results/0274-merge/merge-bench.json](results/0274-merge/merge-bench.json)
(every sample).

| Scenario | Leaf leg | Unified leg | Unified / leaf |
|---|---|---|---|
| 1,000 records into a 6-paragraph, 8-placeholder DOCX (one placeholder split across runs), each output serialized; template opened source-backed (`auto`) | 37.3 ms (26.8k records/s) | 38.5 ms (26.0k records/s) | 1.033 |
| the same, template opened eager | 37.6 ms (26.6k records/s) | 39.0 ms (25.7k records/s) | 1.035 |
| one record into a 2,000-paragraph DOCX with 10,000 placeholders (2,000 split across runs), no serialization | 9.25 ms (1.08M placeholders/s) | 9.66 ms (1.04M placeholders/s) | 1.044 |

- The leaf leg is the hand-written equivalent over the new leaf API: open a
  `litchi_docx::Package` once, then per record `edit`, `story`, a plain
  `{{key}}` scan with a `HashMap` lookup, `substitute`, `commit` (and
  `to_bytes`). The unified leg is `File::merge` (+ `File::to_bytes`), which
  adds OfficeCLI's placeholder grammar, the `Summary` key sets, value
  checks, the transaction patch and the handle adoption.
- Serialization (deflate) dominates the per-record cost of both legs.
- Three changes brought the unified leg from 1.13–1.50× (first
  measurement) to the numbers above: a `File::merge` that builds its DOCX
  editor once instead of an admission editor plus a merge editor; the
  promoted owning package of a source-backed DOCX kept on the handle
  (`auto` was 1.50× when every merge re-materialized it); and planner
  allocations (borrowed key sets, borrowed values) plus a binary search for
  a range's pieces in the leaf.
- All three scenarios are within the goal's 10% single-write budget; the
  1k-op batch budget (2× of one hand-written edit) does not apply, a merge
  being one operation.

## 3. Deltas vs OfficeCLI

OfficeCLI reference: `src/officecli/Core/TemplateMerger.cs` and
`CommandBuilder.Import.cs` (`BuildMergeCommand`), read-only. No corpus script
uses `merge`, so the corpus outcomes are unchanged.

| OfficeCLI | Litchi | Disposition |
|---|---|---|
| `merge TEMPLATE OUTPUT --data JSON\|FILE.json [--force]` | the same form, plus `merge TEMPLATE DATA --out OUT` | superset |
| one JSON object; an array is `invalid_json` | an object or an array of objects; CSV/TSV with a header row | superset |
| one output | one output per record (`{n}` in the name) | superset |
| DOCX: replacement inside each `w:t` only; a placeholder split across runs is not replaced (but reported unresolved) | split placeholders are replaced, each run keeps its formatting | deliberate |
| DOCX unresolved scan: body, headers, footers | every merged story | superset |
| DOCX `MERGEFIELD` not handled | simple and complex `MERGEFIELD`s filled, `\b`/`\f`/`\*` case switches | superset |
| DOCX tabs/line feeds written as characters in `w:t` | `w:tab`/`w:br` | deliberate |
| PPTX: top-level `p:sp` of slides and their notes; a paragraph's runs collapsed into the first run (later runs' formatting lost) | every slide shape paragraph (group members included), per-run formatting kept; speaker notes not merged | deliberate; notes deferred (F2) |
| PPTX `a:br`/`a:fld` invisible to matching | boundaries; a value with a line break is refused | deliberate |
| XLSX: every cell's text, including a formula cell's cached string (the formula is replaced by an inline string) | text-value cells only; formulas untouched | deliberate |
| XLSX rich-text runs flattened (inline string) | flattened too (the leaf has no rich-text cell write) | parity; F4 |
| README mentions chart titles; the code does not merge charts | no chart merge | parity with the code; F5 |
| `\w` = .NET Unicode word characters (letters, marks, digits, connector punctuation) | letters, digits (`char::is_alphanumeric`) and `_` | delta: keys with combining marks or connector punctuation other than `_` |
| nested objects/arrays rendered by `JsonNode.ToString()` (indented, .NET default escaping: non-ASCII and HTML-sensitive characters as `\uXXXX`) | indented, only quotes, backslashes and controls escaped | delta (nested values only) |
| unresolved: placeholders left in the output (including ones a value inserted) | template keys the record lacks | deliberate |
| missing keys always kept | `Missing::Keep` default, `Missing::Empty` option | superset |
| no repeat regions | none | parity (OfficeCLI has none) |
| no limits | `SemanticLimits` | Litchi policy |

## 4. Gaps and follow-ups

- F1 merge for XLSB, DOC, XLS, PPT, RTF and ODF (reported by `capabilities()`).
- F2 PPTX speaker notes: the leaf's notes edit replaces a notes page's text
  as a whole; run-level notes edits are needed to keep formatting.
- F3 PPTX table cells: the native cell text edit replaces the whole cell.
- F4 XLSX rich-text cells keep their runs (needs a leaf rich-text cell write).
- F5 chart titles and series names (DOCX/PPTX/XLSX chart parts).
- F6 DOCX paragraphs inside body children the element paths do not address
  (custom XML blocks, block-level alternate content) are not merged; fields
  spanning paragraphs or holding other content are skipped.
- F7 `MERGEFIELD` numeric (`\#`) and date (`\@`) pictures, `\m`, `\v`.
- F8 a batch `merge` command item; data-source helpers in Python (CSV).
- F9 consolidation with the find/replace engine (`ops/src/find/`, built
  concurrently): the leaf `document::substitute` primitive (span
  replacement keeping the first run's formatting) is what a formatting-aware
  DOCX replace needs too, and the PPTX run rewriting
  (`format/pptx/write/merge.rs::rewrite`) is a cross-run matcher that should
  be shared rather than duplicated.
- F10 output names from a record key (`{name}`) in the CLI.

## 5. Verification

All on this branch, exit status 0 unless noted:

- `cargo fmt --all --check`;
- `cargo clippy -p litchi-docx -p litchi-cli -p litchi-ops -p litchi -p
  litchi-py --all-targets --all-features -- -D warnings`, and `cargo clippy
  -p litchi-ops --no-default-features [--features docx|pptx|xlsx|xlsb|
  doc,ppt,xls,rtf] --all-targets -- -D warnings`;
- `cargo test -p litchi-ops --all-features` (lib 866, doc tests, 927
  integration tests in batches to keep the target under 20 GB), `cargo test
  -p litchi --all-features`, `cargo test -p litchi-docx --all-features`,
  `cargo test -p litchi-cli --all-features`;
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p
  litchi-docx -p litchi-ops -p litchi-cli -p litchi -p litchi-py`;
- `python3 tools/check_crate_boundaries.py`; `python3
  tools/non_iwork_gate.py verify`;
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir
  target/api-glossary` (26 profiles);
- `crates/pyo3-litchi/tests/run_tests.sh` (255 tests);
- `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
  officecli_corpus` (outcomes unchanged), `python3 tools/officecli_corpus.py`
  and `--check` (no diff);
- capability goldens re-blessed (all features and
  `doc,docx,ppt,pptx,xls,xlsx,xlsb,eval`), `python3
  tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` (no
  diff).

## 6. Not claimed

- Parity with Word's own mail merge beyond the listed switches.
- Merge of formats other than DOCX, PPTX and XLSX.
- That merge is faster than a hand-written leaf merge: the legs are within
  the measured overhead above.
