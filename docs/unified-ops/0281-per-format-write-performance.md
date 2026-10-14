# 0281 — Per-format write performance misses

Status: implemented on `uo2/perf-writes` (cut from `feat/unified-ops-api` at
80e24f8745); awaiting coordinator review and merge.

Plan: [0248](0248-coordinator-restart-plan.md), owner decision D5 (budget
misses need engineering first, or a recorded explanation). Budget: the
goal's PERFORMANCE BAR — a single-op write within 10% of the hand-written
leaf edit, a batch within 2× of one leaf edit doing the same work.

Inputs: the open misses of records 0251a and 0251f (DOCX), 0253 and 0262
(OOXML metadata), 0259 (XLS), 0269 P1 and P2 (PPTX add, warn batches) and
0273 (find/replace).

## 1. Summary

| Item | Before | After | Status |
|---|---:|---:|---|
| `Unsupported::Warn` 100-cell XLSX batch vs the refusing batch (0269 P2) | 49.5× | 1.00× | fixed |
| PPTX text-box add, 1 / 100 / 1,000 shapes (0269 P1) | 1.39–1.42 / 1.10–1.12 / 1.05–1.07 | 1.00–1.02 / 1.00–1.01 / 1.00 | fixed |
| DOCX cell paragraph `align` (0251a) | 1.250 | 1.085–1.097 | inside |
| DOCX cell `text` (0251a) | 1.225 | 1.073–1.086 | inside |
| DOCX find/replace, 10k paragraphs, 1 match (0273) | 1.118 | 1.069–1.089 | inside |
| XLSB eager `/metadata` set (0253) | 1.061–1.097 | 1.005–1.022 | fixed |
| PPTX source-backed `/metadata` set, repeated (0253) | 1.016 (already fixed by 0262) | 1.018 | inside |
| DOCX single paragraph `text` set (0251f) | 1.187 | 1.19 | **miss, explained (§4.1)** |
| PPTX open + first `/metadata` set (0262) | 1.313 | 1.313 | **miss, explained (§4.2)** |
| XLS reads, get / query (0259) | 1.5–4.3× / 1.6–5.2× | unchanged | **handed to the shared resolver and walker (§4.3)** |

## 2. Changes

### 2.1 Warn batches keep the native runs (`batch/execute.rs`)

Record 0269 disabled set coalescing under `Unsupported::Warn`, because a
coalesced run reported its refusal only when the run closed. Every
coalescing family (DOCX formatting, XLSX/XLSB/XLS/ODS cells, PPTX/PPT
shapes, RTF, ODT, ODP, DOC) admits every property of an item in its plan
step, before staging it. `unsupported_property` is only produced by registry
admission (`registry.rs` `resolve` and `setter`, and the adapters' admission
helpers), never by a native commit. So:

- an atomic Warn batch now calls `Edit::coalesce_set` exactly as a refusing
  batch does;
- when the plan step refuses the item with `unsupported_property` (the
  error carries no item index, so it is the item's own and nothing of it was
  staged), the run is closed and the item runs alone through the existing
  skipping path (`apply_skipping`), which reports its warnings with the
  item's wire index;
- a deferred failure of a closing run (error with an item index) still
  fails the batch, attributed to its own item; any other admission error
  fails the item as before.

The per-item warnings, the atomic subset semantics and best-effort mode are
unchanged. Tests (`tests/unsupported_warn.rs`): a 100-cell warn batch is one
native commit under both policies with identical bytes; a mixed batch
(refused items between runs, an all-refused no-op) equals applying each item
in turn with `apply_with`, warning by warning and byte for byte; a later
invalid value still rolls the batch back with its own index; the same for a
DOCX formatting run.

### 2.2 PPTX text-box add admits once (`format/pptx/structure.rs`)

The text-box add called `Pptx::editable`, which builds and drops a native
edit only for admission, and then built the edit that applies the add. The
native edit construction (notes index, strict graph admission, package
fingerprint) was therefore paid twice. Every other PPTX writer already uses
`editable_owner` and builds its edit once; the add now does too. The add's
error order is unchanged for valid documents: the one native edit is still
built before anything is staged.

### 2.3 DOCX positional row and cell lookup

- `litchi-docx`: `Table::row(index)` and `Row::cell(index)` return one row
  or cell. Without a cached list they read the table (row) only up to the
  addressed child; the rows before it are passed over with
  `NsReader::read_to_end` instead of resolving their content. The scanner
  gains a `skip`/`ControlFlow` variant
  (`scan_word_element_ranges_with_context_until`); the complete scans are
  unchanged. The counting rules are those of `rows()`/`cells()`: rows in
  content controls and custom-XML wrappers count, rows of nested tables and
  foreign look-alikes do not (unit test
  `positional_row_and_cell_match_the_complete_lists`). `Table::cell(r, c)`
  uses them.
- ops DOCX nav: `tr[N]` and `tc[N]` steps use them; `last()` and the
  unselected forms still list the children.

### 2.4 DOCX find scopes (`format/docx/write/find.rs`)

A single target node never names a paragraph twice, so the deduplicating
`HashMap<ElementPath, usize>` is built only for several targets, and one
piece buffer is reused for every paragraph.

### 2.5 XLSB metadata on the workbook

`litchi-xlsb` gains `Workbook::edit_metadata`: the same `Edit` and `Commit`
as `Package::edit_metadata`, staged on the workbook's own graph. The ops
eager XLSB metadata set used to copy the workbook into a package
(`Workbook::to_package`, a full OPC graph clone plus its drop) before the
edit; it now stages on the workbook. Unit test
`workbook_edits_equal_package_edits` (same bytes, metadata and patch size;
no-op keeps every byte).

## 3. Measurements

Release build, `taskset -c 13`, AMD EPYC 9R45, paired leaf vs unified legs
in each harness's ABBA order; "before" at 80e24f8745 (load average about
1.5), "after" at 971ba2c2af (load average 16–21 from other agents). Raw
after outputs and the before summary:
[results/0281/](results/0281/) (all files under 10 KB). Binary SHA-256s:
`results/0281/binaries.sha256`; rebuild with
`cargo build --release -p litchi-ops --features docx,pptx,xlsx,xlsb,xls --example <name>`.

| Harness (example) | Scenario | Leaf | Unified before → after | Ratio before → after |
|---|---|---:|---:|---:|
| `corpus_parity_bench` | XLSX 100-cell batch, Refuse (A) vs Warn (B) | 336.6 µs (A) | 17,863 → 335.7 µs | 49.46 → 0.996–0.998 |
| | PPTX formatted text-box add, 1 shape | 331.4 µs | 477.9 → 339.4 µs | 1.393 → 1.016–1.024 |
| | PPTX plain text-box add, 1 shape | 311.3 µs | 451.7 → 310.8 µs | 1.421 → 0.998–1.002 |
| | same, 100 shapes (formatted / plain) | 2.51 / 2.08 ms | | 1.102 / 1.124 → 0.997–1.000 / 1.004–1.007 |
| | same, 1,000 shapes | 24.5 / 20.0 ms | | 1.049 / 1.069 → 0.990–1.001 / 1.002–1.003 |
| `docx_element_set_bench` | cell paragraph `align` | 159.9 µs | 202.4 → 173.4–175.5 µs | 1.250 → 1.085–1.097 |
| | cell `text` | 158.3 µs | 197.5 → 169.8–170.7 µs | 1.225 → 1.073–1.086 |
| | other cases (hyperlink bold, table add, …) | | | 1.02–1.07 (unchanged) |
| `find_replace_bench` | DOCX 10k paragraphs, 1 match | 15.36 ms | 17.69 → 16.42 ms | 1.118 → 1.069–1.089 |
| | DOCX 10k paragraphs, 1,000 matches | 21.75 ms | | 1.084 → 1.052–1.069 |
| `ooxml_metadata_bench` | XLSB eager `/metadata` set | 31.9 µs | 35.2 → 32.6 µs | 1.094 → 1.020–1.022 |
| `xlsb_metadata_source_bench` | unified-eager / leaf-eager (new, Simple, 8×2000) | | | 1.097 / 1.097 / 1.061 → 1.016–1.021 / 1.011–1.015 / 1.005 |
| `docx_fanout_bench` | `set /body/p[1] text` | 8.45 µs | 10.07 µs | 1.187 → 1.192 (miss, §4.1) |
| `ooxml_metadata_bench` | PPTX open + first set | 374 µs | 491 µs | 1.313 → 1.313 (miss, §4.2) |
| `xls_read_bench`, `xls_write_bench` | see §4.3 | | | unchanged |

Everything else re-measured in these harnesses is inside budget before and
after: PPTX fill/geometry and transition sets 0.996–1.011; XLSX metadata
1.04–1.06; PPTX metadata 1.014–1.018 (eager and source); XLS cell writes
1.00–1.02 (1–10k cells). One single-command scenario stays above 10% by
design: a warn `apply_with` with one refused property costs one isolated
retry (1.33–1.36×, record 0269 §6).

## 4. Misses left, with measured explanations

Instruction counts are from `valgrind --tool=callgrind` on the `profiling`
profile (perf is unavailable on the host: `perf_event_paranoid` = 4).

### 4.1 DOCX single paragraph text set: 1.19×

Per call the unified leg executes about 15.5 K instructions more than the
leaf's 142.6 K (about 1.11× in instructions, 1.19× in time):

- 9.6 K path resolution: the generic resolver (`path::resolve`, 3.5 K) and
  the DOCX body node, which builds a `Document` view only to read the cached
  body index. Of that, `OpcPackage::main_document_part` takes 4.4 K: it
  re-normalizes the `officeDocument` relationship target
  (`PackURI::from_rel_ref`) on every call. The leaf leg gets its position
  from the caller.
- about 6 K wrapping the native commit into the unified `File` and
  `Commit`.

The leaf edit and commit are identical in both legs. Closing the miss needs
the main-part name memoized per OPC graph in `litchi-opc` (also four calls
per leaf commit) or a per-handle body index; neither is in this record's
crates. Follow-up **F1**.

### 4.2 PPTX open + first metadata set: 1.31×

Per call (instructions): the unified source-backed open costs 3.9 M and the
first set 6.0 M (adoption of the source graph into an eager package 4.1 M,
the metadata edit and commit 1.3 M); the leaf's eager open costs 5.8 M and
its edit 1.5 M. The adoption
(`SourceBackedPackage::to_opc_package_preserving_source`) inflates every
member again, and `attach_owned_source_preservation` re-reads every
relationship member (1.4 M) the open's catalog already read; about 11% of
both legs is zero-filling inflate buffers in `soapberry-zip`. The repeated
set on the adopted handle is 1.014–1.018×. Closing it needs either a
source-backed PPTX metadata edit (record 0262 F3: a part overlay in
`litchi_pptx::source::Package`) or reuse of the open catalog's relationship
payloads in `litchi-opc`. Neither is in this record's crates. Follow-up
**F2**.

### 4.3 XLS reads

Per eager `get` (1,000 cells), the unified leg costs about 2.7 K
instructions more than the native lookup plus formatting (1.9 K):

- about 1.6 K in the shared path resolver and vocabulary (`path/resolve.rs`,
  `path/vocabulary.rs`, inlined into the call): step canonicalization and
  cell-reference parsing per call;
- about 0.7 K in the XLS adapter's `Nav::child` (sheet lookup, the leaf
  B-tree lookup, the cell node);
- about 0.4 K in the generic `View::text` around the XLS display formatter.

On positional owners, `Xls::check_source` runs about four times per get or
per scanned cell (`FileSource::version`, an `fstat` each, 16.5 M of 448 M in
the profile): these are the shared resolver's and walker's fences, and the
leaf's own `query_cell` fences once more. The 5× positional scan is the
walker's per-node cost (record 0259). The fences, the step parsing and the
walker belong to record 0280's agent (shared walker); the adapter's own
share is small, so it is left as is. Follow-up **F3** for 0280: one fence
per walk or get, and pre-canonicalized parsed paths.

## 5. Gaps and follow-ups

- **F1** DOCX single set: memoize the main-part name per OPC graph
  (`litchi-opc`), or keep the body index per `Docx` handle.
- **F2** PPTX first metadata set on a source handle: a source-backed
  metadata edit in `litchi-pptx` (0262 F3), or relationship-payload reuse in
  `litchi-opc` adoption.
- **F3** XLS reads: shared resolver/walker fences and step parsing (record
  0280).
- The `Unsupported::Warn` retry of a refused single command (1.33–1.36×) is
  one isolated retry per refused property; pre-admitting against the
  registry could remove it, but no budget covers it.

## 6. Gates

Run at 971ba2c2af plus this record (the record changes no code).

| Gate | Command | Exit |
|---|---|---|
| format | `cargo fmt --all --check` | 0 |
| clippy | `cargo clippy -p litchi-docx -p litchi-xlsb -p litchi-pptx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| clippy per feature | `cargo clippy -p litchi-ops --no-default-features --features <f> --all-targets -- -D warnings` for `docx`, `pptx`, `xlsx`, `xlsb`, `xls`, no features, and every feature except `encryption` | 0 each |
| tests | `cargo test -p litchi-docx --all-features` (2,372 passed) | 0 |
| | `cargo test -p litchi-xlsb --all-features` (1,272 passed) | 0 |
| | `cargo test -p litchi-ops --all-features` (2,105 passed) | 0 |
| | `cargo test -p litchi-ops --no-default-features --features <all but encryption> --test capabilities --test registry` (goldens unchanged, nothing re-blessed) | 0 |
| | `cargo test -p litchi --all-features` (260 passed) | 0 |
| | `cargo test -p litchi-cli --all-features` (30 passed) | 0 |
| rustdoc | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-docx -p litchi-xlsb -p litchi-ops` | 0 |
| crate boundaries | `python3 tools/check_crate_boundaries.py` | 0 |
| non-iWork | `python3 tools/non_iwork_gate.py verify` | 0 |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (26 profiles, 1,201 bans, 39 families) | 0 |
| perf-baseline | `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| Python | `crates/pyo3-litchi/tests/run_tests.sh` (261 tests) | 0 |

No capability or registry change, so the goldens and the OfficeCLI delta
table are unchanged. New public API: `litchi_docx::Table::row`,
`litchi_docx::Row::cell`, `litchi_xlsb::Workbook::edit_metadata` (names as
the existing `rows`/`cells`/`Package::edit_metadata`; no glossary row
changes).

## 7. Commits

- `ae6240d1cf` perf(ops): coalesce atomic batches under `Unsupported::Warn`
- `2dc50c3fb8` perf(pptx): admit a text-box add once
- `d4a76b3485` perf(docx): positional row/cell lookup and leaner find scopes
- `971ba2c2af` perf(xlsb): stage metadata edits on the workbook
- this record and its evidence
