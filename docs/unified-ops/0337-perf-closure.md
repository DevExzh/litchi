# 0337 — Performance closure: 1k-batch evidence for every format, follow-ups for every open miss, the `Path` size

- Status: implemented on `uo2/perf-closure`, cut from `feat/unified-ops-api`
  at `d165759e4a`.
- Date: 2026-10-17
- Input: the definition-of-done audit of 2026-10-17 (§5, WP7, WP8),
  `docs/UNIFIED_OPS_GOAL.md` (PERFORMANCE BAR: a 1k-op homogeneous batch
  within 2× of one hand-written leaf edit, linear scaling), records
  [0209](0209-xls-publication-work.md), [0310](0310-miss-sweep.md),
  [0315](0315-query-walker-read-path-misses.md) and
  [0326](0326-pptx-regression-path-parser-typed-coercion-read-misses.md).
- Scope: the new harness `tools/perf-baseline/unified-batch-closure`
  (new route names only), `tools/unified_bench_matrix.py` and its curated
  file, `docs/unified-ops/results/0337/`, and one fix in the XLS write run
  (`crates/litchi-ops/src/format/xls/write.rs`) with a regression test.
- `performance_claim`: paired measurements in §2–§4. Release build of the
  harness (standalone workspace, `debug = 1`), one pinned CPU
  (`taskset -c 30`) of the shared 32-core host at load 3.5–14.5 (other
  agents were building), three trials (processes) per set, each case an
  in-process ABBA of 5 rounds (10 samples a leg) after an untimed
  equivalence check. Instruction counts are callgrind (`perf` is blocked,
  `perf_event_paranoid=4`). Evidence: [results/0337/](results/0337/):
  `build.sh`, `measure.sh`, `binaries.sha256`, `scaling.py` →
  `scaling.md`, and the run set `abba/` (40 JSON documents and the log,
  316 KB).

## 1. Summary

| item | before | after |
|---|---|---|
| formats with a 1k-batch measurement on current trunk | 16 of 21, none in the matrix for docx/xlsx/pptx/xlsb/xls (audit §1 7a, 7c) | **21 of 21**, all in the matrix (§2) |
| 1k batches within 2× of one leaf edit | — | **20 of 21**; Keynote misses (22×, §3.3) |
| linear scaling of the unified layer's added cost (100 → 1k → 10k) | not checked | linear for 15 of 17 formats measured at 10k; Keynote (quadratic) and ODC (superlinear) are not (§3) |
| XLS file-backed 1k removal (record 0209, 2.093×, outside the matrix) | 1.362–1.376 on trunk | **1.148–1.162** (fix, §4); 0209's row is in the matrix, superseded |
| bench-matrix open misses | **101** | **103** (+2: Keynote 1k and 10k batches) |
| open misses with neither a follow-up nor an explained residual | **19** by the audit's count, **26** by the matrix's new check (§5) | **0**; `--check` now fails on such a miss |
| parsed `Path` size (616 B) | — | design note (§6): 248 B possible, not implemented (public names change) |

Results are byte-identical: no read path, registry or capability changed;
the XLS change removes source checks inside a write run only (§4). The
query-corpus differential, corpus fidelity, the CRUD checklist goldens and
the capability and registry goldens pass unchanged; nothing was re-blessed
(§8).

## 2. The batch harness

`unified-batch-closure` builds one binary for all 21 formats (each a
feature). For every format it pairs:

- **unified leg**: one atomic `File::batch_with_limits` of N `set` (or
  `remove`) items on distinct targets, items cloned outside the clock;
- **leaf leg**: one hand-written leaf edit doing the same work, split only
  where the leaf's own bound forces it (XLS `MAX_STAGED_CHANGES` 4,096: three
  commits at 10,000 in both legs), with the limits the unified adapter
  passes the leaf (DOC, RTF);
- **check** (outside the clock, before timing): the two outputs are equal
  — exact bytes in every case — and every item reports `changed`; the
  unified leg's leaf-commit count is reported (1, or 3 for XLS at 10,000,
  equal to the leaf's).

Owners: eager for the OOXML and OLE formats, plus a file-backed XLS owner
(record 0209's shape); the default open for RTF, ODF and iWork. Every count
of a set runs on one fixture, so 100 / 1k / 10k compare at a fixed document
size.

| format | scenario (one leaf call per op unless noted) | leaf edit |
|---|---|---|
| docx | run `bold` on `/body/p[i]/r[1]` | `Edit::format_run` |
| xlsx | cell `value` on `/Sheet1/A{i}` | `WorksheetEdit::set` |
| pptx | shape `text` on `/slide[1]/shape[i]` | one `set_shape_text_chunks` (256 a chunk) |
| xlsb | cell `value` and cell `remove` (32 columns) | `cell_values::Edit::{set_scalar, remove}` |
| doc | paragraph `text` | `Edit::replace_paragraph` |
| xls | cell `value` and cell `remove`, eager and file owners | `Edit::{set_value, remove_cell}` |
| ppt | shape `text` | one `set_shape_text_chunks` |
| rtf | paragraph `text` | one `replace_body_paragraph_texts` |
| odt, oth, odm | paragraph `text` (OTH/ODM hold the ODT handle, record 0282; fixtures from their own leaf builders) | `litchi_odt::Edit::set_text` |
| ods | cell `value` (10 columns) | `set_cells` in 4,096-change chunks |
| odp | paragraph `text` (100-paragraph boxes) | one `splice_scene` |
| odg | shape `text` | `Edit::set_shape_text` |
| odc | series `label` | one `replace_chart` of the edited definition |
| odi | the frame's 6 independent keys (`name`, `title`, `description`, `anchorType`, `zIndex`, `layer`) | the matching setters; ODI holds one frame, so no batch has 1,000 distinct targets, and a new image has no `meta.xml` for the metadata keys |
| odb | query `command` | `Edit::set_query_command` |
| odf (formula) | token `text` | `Edit::set_text` |
| pages | paragraph `text` | `edit_section_paragraphs` + `set` |
| keynote | body paragraph `text` (slide 2 of `test-data/iwork/keynote/multi-slide.key`, body set to N lines outside the clock) | the slide-text edit stages one operation: the one `replace` that yields the batch's text |
| numbers | cell `value` (10 body columns of a new table) | `edit_table_cells().extend(..)` |

Structural batches: the unified layer coalesces a structural family only
for XLS and XLSB cell removal (measured) and RTF picture removal (not
measured: no fixture). Every other `add`/`remove` item is its own leaf
commit (`transaction.rs` `coalesce_remove`), so a 1k structural batch is
1,000 commits against one leaf edit in docx, pptx, doc, ppt, rtf, odt,
ods, odp, odg, odc, odb, odf and iWork; that is a coverage gap of the
coalescer, recorded as follow-up **F-0337-STRUCT** (§7), not measured as a
row.

Fixture sizes: set A uses 10,000 targets; set B (DOC, ODG, ODB, ODF
formula) uses 1,000, because those leaves' per-operation cost grows with the
document (DOC 1,000 sets take 5.6 s, ODG 4.0 s, ODB and ODF 0.93 s in
**both** legs; 10,000 would take 1–10 minutes a sample); set D repeats XLS
with exactly N cells (record 0209's shape).

## 3. Results

### 3.1 1,000-operation batches (budget 2.0)

Min–max of the three trials' ratios; µs are medians of the trials'
medians.

| format | owner | fixture | leaf µs | unified µs | ratio |
|---|---|---:|---:|---:|---|
| docx | eager | 10,000 | 24,848 | 30,350 | 1.216–1.267 |
| xlsx | eager | 10,000 | 1,166 | 1,995 | 1.701–1.717 |
| pptx | eager | 10,000 | 110,458 | 125,448 | 1.136–1.149 |
| xlsb set / remove | eager | 10,000 | 13,353 / 12,948 | 14,211 / 13,680 | 1.006–1.064 / 0.965–1.056 |
| doc | eager | 1,000 | 5,671,778 | 5,585,987 | 0.984–0.993 |
| xls set / remove | eager | 1,000 | 4,310 / 3,381 | 4,364 / 3,444 | 1.009–1.013 / 1.014–1.021 |
| xls set / remove | file | 1,000 | 4,107 / 3,328 | 4,610 / 3,850 | 1.121–1.123 / 1.148–1.162 (after §4) |
| ppt | eager | 10,000 | 5,341 | 5,863 | 1.070–1.098 |
| rtf | default | 10,000 | 11,663 | 12,663 | 1.059–1.086 |
| odt / oth / odm | default | 10,000 | 3,161 / 3,031 / 3,018 | 3,829 / 3,709 / 3,695 | 1.210–1.233 / 1.216–1.224 / 1.218–1.225 |
| ods | default | 10,000 | 64,422 | 64,595 | 1.003–1.017 |
| odp | default | 10,000 | 12,736 | 13,652 | 1.072–1.076 |
| odg | default | 1,000 | 4,023,170 | 4,052,545 | 0.995–1.016 |
| odc | default | 10,000 | 45,483 | 46,799 | 1.029–1.030 |
| odi (6 ops) | default | one frame | 181 | 189 | 1.038–1.044 |
| odb | default | 1,000 | 969,108 | 967,892 | 0.990–1.004 |
| odf | default | 1,000 | 935,499 | 935,318 | 0.997–1.005 |
| pages | default | 10,000 | 2,417 | 2,618 | 1.082–1.089 |
| keynote | default | 10,000 | 5,595 | 124,172 | **21.88–22.39** |
| numbers | default | 10,000 | 834 | 1,138 | 1.337–1.387 |

At 10,000 operations (also budget 2.0, the goal's "≥ 1,000"): every format
measured there meets except Keynote (387–390×); the highest others are
odt/oth/odm 1.74–1.86, xlsx 1.72–1.84, xlsb remove 1.64–1.83, odp 1.58–1.59,
odc 1.57–1.59, pages 1.62–1.65. All rows: the matrix (`bench-matrix.md`,
record 0337) and `abba/`.

### 3.2 Linear scaling (`scaling.md`)

For each case measured at 100, 1,000 and 10,000 operations on one fixture,
`scaling.py` takes each leg's added cost per added operation between the
counts and checks the unified layer's own share (unified marginal minus
leaf marginal): linear when the 1k→10k value is within 1.5× of the 100→1k
value or under 0.5 µs an operation. Leaves whose own cost grows with the
document (RTF 10,000: 0.79 s in both legs; XLS: three commits at 10,000)
show it in both legs and cancel.

| result | formats (added µs/op, 100→1k → 1k→10k) |
|---|---|
| linear | docx 3.70 → 4.73, xlsx 0.83 → 0.97, xlsb 0.66 → 0.53 / 1.04 → 0.78, xls 0.04–0.49 → 0.01–0.28 (after §4), ppt 0.47 → 0.44, rtf 0.31 → −0.25, odt 0.67 → 0.76, oth 0.69 → 0.76, odm 0.68 → 0.77, ods 0.84 → 1.11, odp 0.92 → 1.02, pages 0.20 → 0.20, numbers 0.30 → 0.29 |
| borderline | pptx 0.07 → 0.52: a flat 13–19 % over a 110 ms leaf at every count; 0.52 µs an operation is at this host's noise level for a 125 ms sample |
| **not linear** | **keynote** 119.9 → 248.0 (quadratic, §3.3); **odc** 0.75 → 3.68 (F-0337-ODC, §7) |

DOC, ODG, ODB and ODF formula were measured at 100 and 1,000 only (their
leaf cost per operation grows with the document, §2); unified/leaf stays
0.98–1.06 at both counts. ODI has one frame.

### 3.3 Keynote: a quadratic paragraph-set run

Callgrind of three 1,000-set batches (`--toggle-collect` on
`File::batch_with_limits`): 6.32 G instructions, 94 % in
`Edit::coalesce_set`, 57 % in `CharSearcher::next_match` (splitting text
on `\n`). Per item, `format/iwork/keynote/write.rs`:

- `Pending::joins` and `Pending::touch` split the base and the staged text
  (`split('\n').nth(line)`), and the cancel check zips every line of both;
- `hull` re-scans the base, and `leaf_admits` opens a slide-text edit;
- `Pending::stage_set` splits the whole staged text into a `Vec` and joins
  it again.

Each is O(text) per item, so N sets on an N-line body cost O(N²): 100 /
1,000 / 10,000 sets on a 10,000-line body take 16 ms / 124 ms / 2.36 s
against the leaf's one replace (5.5–6.1 ms). This is the iWork adapter's
code (record 0314), outside this record's scope; follow-up
**F-0337-KEYNOTE**: keep a staged text group as lines with UTF-16 prefix
offsets and a count of lines that differ from the base, so a paragraph set
costs its own line. Pages and Numbers runs are linear (0.20 and 0.30 µs an
operation added).

## 4. XLS file-backed batches (record 0209)

Record 0209 measured the file-backed 1,000-cell removal at 2.093× and left
it as a facade follow-up; the matrix did not list it (it starts at 0249).
It is now a curated row (0209's table, verbatim) on the series of this
record's same-shape case, which supersedes it.

On trunk the case reads 1.362–1.376 (set 1.287–1.305; eager owner
1.01–1.02): later records halved the gap, and what remains is the source
checks. Per coalesced item the run made about eight source observations on
a file owner — two in `remove_plan_with_source`/`cell_plan_with_source`,
one in `direct_target`, two in `Run::stage`, the resolver's closing check
and the two of the leaf's `Edit::source_cell` — each an `fstat`.

Fix (`format/xls/write.rs`): a plan made inside a run (`source` is the
run's leaf edit) and `Run::stage` add no source check of their own. The
run checks the source when it opens and before and after its commit, and
each item's resolution and leaf `source_cell` read still observe it, so a
source that changes during the run is still refused and nothing is
published. Plans outside a run (the run's first item, the ordinary
setters) check as before. The new test
`stale_files_are_rejected_inside_coalesced_set_and_remove_runs` changes the
file between two coalesced sets and two coalesced removals and expects
`io_error` with the file left as changed; the existing
`stale_files_are_rejected_while_a_native_run_is_pending` and every other
XLS test pass.

| case (file owner) | trunk | after |
|---|---|---|
| remove, 1,000 on 1,000 cells | 1.362–1.376 | **1.148–1.162** |
| set, 1,000 on 1,000 cells | 1.287–1.305 | 1.121–1.123 |
| remove / set, 100 on 100 | 1.347–1.353 / 1.273–1.308 | 1.150–1.160 / 1.124–1.141 |
| remove / set, 1,000 on 10,000 | 1.056–1.134 / 1.102–1.149 | 1.052–1.056 / 1.040–1.061 |
| remove / set, 10,000 on 10,000 | 0.876–1.151 / 0.901–1.123 | 1.025–1.076 / 0.978–1.052 |

The eager owner is unchanged (1.00–1.04 before and after). The remaining
0.5 µs an item on a file owner is the three checks that stay (resolver
and leaf).

## 5. Follow-ups for the open misses

`tools/unified_bench_matrix.py` now gives every open miss a disposition:
`follow-up` (a named follow-up) or `residual` (inherent to the comparison,
or kept by an owner decision), from an explicit `disposition` in the
curated row or annotation, else from the explanation's wording. The
Markdown states the counts, and `--check` fails on an open miss with
neither. The check found 26 such misses (the audit counted 19 of them; it
also counted the 6 RTF typed single sets of record 0326 and the XLSB note
get as having one). Each now has one:

| misses | disposition (curated annotation) |
|---|---|
| doc `doc_text_set` 1.27 (record 0311) | **residual**: the comparison leaf publishes no readable package; against the `Package::edit` leaf the set wraps it is 1.03–1.04 (0311 §4.2), and this record's batches of 100 and 1,000 sets read 0.98–1.06 against it. The difference is the commit's parse of its own output, its validation, which stays. |
| docx `saut` `r[bold=true]` bytes/file 1.19–1.20 (0315) | **F-0337-DOCX-PROP**: callgrind of 101 queries with no match: 38.4 M instructions against the leaf loop's 34.9 M. The walker reads the condition's one property through `Inspect::prop`, which DOCX leaves at the default that builds the run's whole `format` (a boxed `Props::Run`, the invalid-property scan, the `Format` table) where the leaf reads `direct_properties().bold`. Follow-up: a DOCX `Inspect::prop` that reads one key of the parsed run properties (record 0326 §7's cached property set). |
| odt `query_all` ×6, 2.6–2.9 (0315) | **F-0337-ODT-QUERY**: callgrind (1,000 paragraphs ×10): 1,739 instructions a match against the leaf's 896; leaf `Node::text` (839) is in both; the rest is the walker's fast step (~170), six moves of the node (~180 of `memcpy`), `Children::next` (~64), `Address::with` (~44) and the path (~44). Follow-up: a lending match item (a new public shape, record 0315 §6) and a compact ODT `Address`. |
| rtf fan-out ×8, 1.10–1.26 (0315) and rtf single sets ×6, 1.12–1.24 (0326) | **F-0337-RTF-SINGLE**: the per-call planning of the RTF write adapter around one leaf edit (0310 §2 G9: `Edit::set_with_policy` 8.6 K of 90.6 K instructions); the compiled-plan route still misses, so the compile is not the cost. Follow-up: stage a lone target without the coalescing plan. |
| xlsx 50,000-row `row[Sales>1000]` bytes/file 1.70–1.97 (0315) | **F-0337-XLSX-HEADER**: header-row plans take the walker's general step (record 0315's `FastPlan` excludes header rows); follow-up: a fast plan for one header-row comparison. |
| xlsb `comment_get` 1.14–1.16 (0326) | **F-0337-XLSB-NOTE**: cache the parsed sheet step of a two-step note path, as record 0315 §3.3 did for XLS cells. |
| keynote batch 1k / 10k (new, §3.3) | **F-0337-KEYNOTE** |

Counts: open misses 101 → 103; with neither a follow-up nor a residual
26 → **0** (audit's count 19 → 0). Of the 103: 49 follow-up, 54 residual.

The DOCX, ODT, XLSX and XLSB items above are read paths this record may
change; each fix needs a public or shared-walker change (lending items,
`FastPlan`), or is record 0326 §7's DOCX property-set design, which was not
started here. The XLS removal (§4) was the one fix this record made.

## 6. Design note: a compact parsed `Path` (record 0326 §7)

### 6.1 Where the 616 bytes are

Measured with `size_of` on this tree (64-bit):

| type | bytes | made of |
|---|---:|---|
| `Path<'a>` | 616 | `absolute` (1, padded to 8) + `Inline<Step, 6>`: `len` 8, six inline steps 576, spill `Vec` 24 |
| `Step<'a>` | 96 | `name: Cow<str>` 24 + `OneOrMany<Predicate>` 72 |
| `Predicate<'a>` = `Expr` = `Condition` | 72 | `Compare { key: Key, op, value: Value }`: `Key` 32 (`Cow<str>` 24 + `attribute`) + `Value` 32 (`Text`/`Regex` over `Cow<str>`) + `op` |
| `Query<'a>` | 176 | |

94 % of a `Path` is six inline steps, and 75 % of a step is room for one
`[@key=value]` comparison, which most steps (`p[3]`, `body`) do not use. A
text-route batch item initialises, moves three times and drops this value
(record 0326 §3.3: 267 instructions of `memcpy`, about 60 ns an item, the
gap between Pages `batch_10000` and `batch_10000_parsed`).

### 6.2 What cannot shrink without a public change

- `Step::predicates()` returns `&[Predicate<'a>]`, so a step must hold its
  predicates as materialised `Predicate` values; it cannot store a smaller
  private form and still lend that slice.
- `Value::Text(Cow<'a, str>)`, `Value::Regex(..)` and
  `Condition::Compare { key, op, value }` are public payloads, so
  `Predicate` cannot shrink in place.
- A borrowed-or-boxed string smaller than `Cow<str>` (24 B) needs pointer
  tagging, i.e. `unsafe`, which the crate forbids.
- Fewer inline steps (`INLINE_STEPS` 6 → 4, private: 616 → 424 B) would
  make `/body/tbl[1]/tr[2]/tc[3]/p[1]/r[1]`, the deepest common OfficeCLI
  form, allocate, which `path_allocations` forbids.

Private-only changes save at most 48 B (folding `Key`'s `@` flag into an
enum leaves `Value` at 32), under 8 %.

### 6.3 Proposed type (glossary-conforming; not implemented)

One text buffer per path, with spans into it, and the step and predicate
types as borrowed views:

```rust
pub struct Path<'a> {
    text: Cow<'a, str>,           // the input, or the unescaped / owned text
    absolute: bool,
    steps: Inline<StepSlot, 6>,   // private
}
struct StepSlot { name: Span, predicate: PredicateSlot }   // 32 B
enum PredicateSlot {                                     // 24 B
    None, Index(Position), Last,
    Exists { key: Span, attribute: bool },
    Compare { key: Span, attribute: bool, op: CompareOp, value: Span, regex: bool },
    Complex(Box<[Predicate<'static>]>), // stacked brackets, and/or/not: allocating today too
}
struct Span { start: u32, len: u32 }   // into `text`; MAX_INPUT_BYTES is 16 KiB
```

Size: 24 + 8 + 6 × 32 + 24 = **248 B** (−60 %). Public surface, in the
glossary's words and the existing names:

- `Path::steps(&self) -> Steps<'_, 'a>`, an `ExactSizeIterator +
  DoubleEndedIterator` with `get(Position)` and `last()`, replacing
  `&[Step<'a>]`;
- `Step<'p>` becomes a `Copy` view (`&Path` + slot) keeping `name()`,
  `index()`, `axis()`, `cell_ref()`, `cell_range()`; `predicates()` returns
  `impl Iterator<Item = Predicate<'p>>`, each value built from its slot on
  demand;
- hand-built paths (`Path::from_steps`, `Step::new(..).with(..)`) move to
  `path::Builder` (`push`, `predicate`, `build() -> Path<'static>`);
- `Predicate`, `Expr`, `Condition`, `Key`, `Value`, `CompareOp`,
  `ParseError`, `Display` and every error offset are unchanged.

Cost:

- **Call sites**: 39 `predicates()` calls and up to 160 `.steps()` calls
  (64 files; some are other types' `steps()`) in `litchi-ops` (`path/*`,
  `query/compile.rs`, the coalescers, the metadata adapters), plus the
  selector tests, move from slice indexing to `get`/iterators.
- **Reads**: `name()` becomes a slice of `text` (an add and a bounds
  check); `predicates()` builds one small enum per call; resolvers read
  each step once.
- **Expected gain**: most of record 0326's 60 ns an item on the text route
  (Pages `batch_10000` 1.47–1.50 against 1.28 parsed): three 368-byte
  copies fewer per item. Not measured.
- **Risk**: both the fast parser (record 0326 §3) and the general parser
  must emit spans; `path_parser_differential` (211,682 inputs) and the
  round-trip tests guard them.

### 6.4 Decision

Not implemented: `steps()` and `predicates()` change their return types
and `Step` changes from an owned value to a view, so the public names
cannot all stay stable, which the brief requires for an implementation.
The change needs glossary rows (`Steps`, `path::Builder`, the `Step`
view), then one migration of `litchi-ops`: follow-up **F-0337-PATH**.

## 7. Gaps and follow-ups

- **F-0337-KEYNOTE** (iWork adapter, `format/iwork/keynote`): the
  quadratic paragraph-set run (§3.3); the only 1k-batch budget miss.
- **F-0337-ODC** (`format/odf_tree/odc.rs`): the unified layer's added
  cost per series label set grows from 0.75 to 3.68 µs between 1k and 10k
  (10,000 sets 1.57–1.59, within budget). Not profiled here; the leaf leg
  indexes the definition's series directly, while each unified item
  resolves `/series[i]` positionally, which is consistent with a walk
  proportional to `i`.
- **F-0337-STRUCT** (`transaction.rs` coalescer): structural `add`/`remove`
  batches are coalesced only for XLS/XLSB cells and RTF pictures (§2).
- **F-0337-DOCX-PROP, -ODT-QUERY, -RTF-SINGLE, -XLSX-HEADER, -XLSB-NOTE**:
  §5.
- **F-0337-PATH**: §6.
- The XLS file owner keeps three source observations per coalesced item
  (§4): the resolver's and the leaf `source_cell`'s two; removing the
  latter needs an unchecked leaf read (a public `litchi-xls` addition).
- The harness's DOC, ODG, ODB and ODF formula cases stop at 1,000 (§2).

## 8. Gates

Tree: this branch with trunk `1d2a8e47c0` (records 0334, 0335) merged at
`10a6e9d5ec`. Build environment per the rules (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`, the memory-capped cargo shim, one cargo command at a
time); every exit status checked (`.coord/perf-closure/gates.sh`, not
committed).

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` (and `cargo fmt --check` in the harness) | 0 |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features xls] --all-targets -- -D warnings` | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib` (929), `--doc`, the 153 test targets in five groups (1,383), includes the query-corpus differential, corpus fidelity, the CRUD checklist goldens and the new XLS stale-run test | 0 each |
| `cargo test -p litchi-ops --all-features --test capabilities --test registry` (not re-blessed) | 0 |
| the same with `--no-default-features --features` every litchi-ops feature but `encryption` (not re-blessed) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,313 bans, 253 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (285 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo clippy` / `cargo check --locked --offline --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml --all-targets` | 0 each |
| `python3 tools/unified_bench_matrix.py --check` | 0 |
| `python3 tools/test_no_3rdparty_references.py` | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so the
delta table and the OfficeCLI corpus were not regenerated. No public name
was added or changed (the XLS change is crate-private; the harness is a
tool); no glossary row changes. The harness adds no dependency beyond
those of the workspace lock (its own `Cargo.lock` is the workspace's,
trimmed).

## 9. Not claimed

- That a 1k structural batch meets the budget in any format but XLS/XLSB.
- That Keynote batches meet the budget.
- Anything about the `Path` proposal's speed (§6.3: not built).
- Process ABBA against a base build for the batch rows: they pair two legs
  of one build; only the XLS change compares two builds (§4).
