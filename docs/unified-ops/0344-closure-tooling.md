# 0344 — Closure tooling: bench-matrix coverage, the XLS/ODS deferral, milestone follow-ups, owner-memo additions

- Status: implemented on `uo2/closure-tools`, cut from `feat/unified-ops-api`
  at `6994d038de`.
- Date: 2026-10-19
- Input: the second definition-of-done audit (2026-10-18; §1 items 2 and 7,
  §3, §5, §12.1 item 5, §12.2 items 3 and 4, WP8, WP10, WP11, WP13),
  records [0332](0332-officecli-parity-dispositions.md) (dispositions),
  [0337](0337-perf-closure.md) (the matrix rules) and
  [0331](0331-owner-decision-memo.md) (the owner memo).
- Scope: tooling and data only. `tools/unified_bench_matrix.py`,
  `tools/officecli_capability_delta.py`, their unit tests,
  `tools/officecli_dispositions.json`, `results/0289/` (the curated file and
  the generated matrix), the generated capability delta, and the owner memo.
  No Rust source changed.
- `performance_claim`: none. The record measures nothing; it collects what
  records 0307–0342 measured.

## 1. Summary

| item | before | after |
|---|---|---|
| bench-matrix rows | 2,377 | **2,737** (+360 from 18 records) |
| records with evidence under `results/` that contribute no rows | 19 (0272, 0307, 0308, 0314, 0316, 0317, 0318, 0323, 0324, 0325, 0328, 0329, 0330, 0333, 0335, 0336, 0338, 0339, 0342), invisible to `--check` | **0**, and `--check` fails on one; 0272 and 0339 carry a recorded reason (§2.3) |
| open misses | 101 (47 follow-up, 54 residual) | **142** (85 follow-up, 57 residual, 0 neither) |
| misses with neither a follow-up nor a residual | 0 in the matrix; 2 outside it (0327 feature reads, 0338 run set) | **0** |
| XLS/ODS P1 delta rows `planned` | 392 | **0** (`deferred-P1`, §3) |
| all P1 rows `planned` | 426 | **34** (DOCX record 0335 follow-ups) |
| P0 rows `planned` | 26 | 26 (unchanged) |
| milestone-only follow-ups in the registry and capability goldens | 784 (96 texts), the capability goldens unchecked | **784 mapped** to a current id by 21 entries; `--check` covers both golden sets (§4) |
| owner-memo items from the second audit | 0 | **5** (items 25–29, §5) |

## 2. Bench matrix

### 2.1 Records collected

The audit counted about 13 misses outside the matrix in five records. The
matrix in fact read no evidence from 18 records, and 41 of their rows are
open misses. Each record now has a parser over its committed evidence, or
curated rows where the evidence stayed outside the repository:

| record | evidence | rows | open misses |
|---|---|---:|---:|
| 0307 | `docx_mc_bench` runs, before/after | 8 | 0 (2 superseded, §2.2) |
| 0308 | `office_gaps_bench` r2 runs, new-path runs, the ODP scene re-run | 20 | 1 |
| 0314 | iWork after legs (`out/*-after-*`) | 49 | 0 |
| 0316 | `xls_odp_followups_bench` r3 runs and new-path runs | 21 | 0 |
| 0317 | `docx_large_bench`, `pptx_large_bench` after legs | 10 | 0 |
| 0318 | `numbers_unified_bench` after legs | 21 | 0 |
| 0323 | `ppt_elements_probe`, two runs | 14 | 4 |
| 0324 | `doc_parts_bench` trials, `doc_view_regression` | 20 | 8 |
| 0325 | `ooxml_followups_bench`, the read-regression probe | 25 | 3 |
| 0327 | curated from §5 (callgrind; evidence in `.coord/xls-ods-features/bench/`) | 2 | 2 |
| 0328 | `odf_rtf_parts_bench` trials, `odf_rtf_view_regression` summary | 32 | 14 |
| 0329 | iWork after legs | 56 | 1 |
| 0330 | the coalesced XLSX batch, before/after | 1 | 0 |
| 0333 | `pptx_p0_bench`, two runs; the per-item batch before/after | 32 | 0 |
| 0335 | `docx_p0_residue_bench` trials, `docx_view_regression` | 26 | 3 |
| 0336 | `dump_bench` (no budget: dump against a traversal) | 3 | 0 |
| 0338 | `p0_residue_bench`, `json_number_bench` | 13 | 4 |
| 0342 | `v1_closure_bench` | 7 | 1 |

Conventions follow the tool's existing ones. A before/after-build row
divides the median of the after runs by the median of the before runs, as
records 0307, 0308 and 0316 compare them, so one noisy run of a shared host
does not decide it. Series keep earlier records' names where the harness and
scenario are the same: `iwork:*` and `numbers:*` for the iWork benches, and
`office_gaps:doc_text_set` for the DOC set that record 0311 re-measured.
Later measurements therefore supersede earlier ones.

### 2.2 New open misses and their dispositions

Every new open miss has a follow-up or an explained residual. The curated
annotations quote each record's explanation.

| record | misses | disposition |
|---|---|---|
| 0327 | XLS feature listing 1.50×, ODS 2.10× (callgrind, against the record's 1.25 bar) | follow-up **F-0344-SHEET-FEATURES**: measure like for like (one open handle; per-kind reads), then share the sheet's feature read and the ODS content index across the six queries |
| 0338 | XLSX run `set bold` 1.32× | follow-up **F-0344-XLSX-RUN**: stage a lone run-path target without the general cell plan (as F-0337-RTF-SINGLE), with the sheet step cached per handle (as F-0337-XLSB-NOTE) |
| 0338 | XLSX picture `set src` 1.27× | 0338-P1 |
| 0338 | JSON decode, long decimals 1.16×, seeds 1.06× (after/before) | residual: the correctly rounded `float_roundtrip` parse, a correctness fix (record 0338 F1) |
| 0342 | DOCX row set 1.13× | 0342-P1 |
| 0335 | `open_control` 2.78×, cold first writes 1.48× / 1.85× | 0335 F6 |
| 0328 | 4 point reads (3.3×–504×, +0.2–0.7 µs) | 0328 follow-up 7 |
| 0328 | 10 child listings, 1.06–1.19 after/before | 0328 follow-up 8 |
| 0324 | 4 point reads (1.30×–423×) | 0324 follow-up 7 |
| 0324 | 4 format views and queries, 1.05–1.09 after/before (seven new format keys) | 0324 follow-up 7 |
| 0323 | PPT comment and run reads, 1.45× and 3.11× | 0323-F12 |
| 0325 | XLSX note sets 1.49×; rich note read 1.12× | 0325 follow-up 4; for the rich read, also **F-0344-XLSX-NOTE-RUNS** (lend the runs to the node instead of cloning them) |
| 0308 | ODP `Scene::parse` 1.07 after/before (third build), with −0.14 % instructions | follow-up **F-0344-ODP-SCENE**: re-measure on current trunk against `6c8d517eff` (ABBA and callgrind); record 0316's own after/before reads 0.983 |
| 0329 | Keynote `table_cell` 1.058 (budget 1.05 or +20 µs) | residual: record 0329 §7 found both legs of every Keynote scenario 2–5 % slower with unchanged instruction counts, which is code layout and host noise |

Record 0307's 1,000-paragraph DOCX text set (2.48×) and add (2.10×) on the
default owner predate the source-overlay fixes of records 0305 and 0311
(its base `e7fe30fb24` is older than both merges). They share the
`docx_large:*:auto` series, so they show as superseded by record 0311
(0.543 and 0.467).

### 2.3 The coverage rule

`--check` now also fails in two cases:

- **Evidence without rows.** A record from 0249 on has evidence under
  `results/` (a `<record>/` directory with a file, or a `<record>-*`
  entry) and contributes no row, unless the curated file's
  `evidence_without_rows` gives a reason. Two records have one:
  - 0272: a route-identity check, with no timing claim.
  - 0339: a property-visibility probe of the corpus harness, not a
    measurement.

  An exemption without a reason fails, and so does one whose record has
  rows or no evidence.
- **Output size.** A generated file reaches 1 MiB.

The record range now ends at the highest record on the tree, not at a
fixed 0340, so a merged record is listed and checked without editing the
tool. The Markdown's record table has an evidence column.

Size: `bench-matrix.md` was 1,066,464 bytes on the base, already past
1 MiB, and these rows would have made it 1.17 MB. Two changes keep every
output under the limit:

- the per-format row tables moved to `results/0289/bench-matrix/<format>.md`,
  so the main file is 402 KB;
- a superseded miss no longer repeats the latest row's explanation. It
  names the row instead, which saved 168 KB in the TSV.

The TSV is 906 KB. It is at 86 % of the limit and grows with every
record's rows (§7).

## 3. XLS and ODS parity rows deferred

The coordinator's adopted default:

- `xls-parity` and `ods-parity` in `tools/officecli_dispositions.json`
  change from `planned` (WP10, WP11) to `deferred`. By tier, that resolves
  to `deferred-P1` for all 392 rows: XLS 196 and ODS 196, across the
  chart, chart-series, chart-axis, run, table and detected-table families
  and their properties.
- Their refs and reasons name the follow-up records:
  - XLS: `0327/0316` (record 0327 §4's feature-node refusals, record
    0316's XLS structural follow-ups, record 0339 F5);
  - ODS: `0327`.
- Both formats are outside the v1 boundary (`policy.p0Formats` = docx,
  xlsx, pptx), so none of these rows is P0. The P0 rows are unchanged: 26
  planned (DOCX 10, PPTX 16) and 4 out of scope. The delta's
  P0-not-deferred rule still passes.
- The decision is owner-memo item 25.

| tier | planned | deferred-P1 | deferred-P2 | wontfix | out-of-scope |
|---|---:|---:|---:|---:|---:|
| P0 | 26 | 0 | 0 | 0 | 4 |
| P1 | 426 → **34** | 6,023 → **6,415** | 0 | 0 | 55 |
| P2 | 0 | 0 | 659 | 0 | 87 |

The 34 P1 rows still planned are record 0335's DOCX comment-reply and
chart-formatting follow-ups (`docx-picture-comment-follow-ups`). The
OfficeCLI corpus report does not change: no corpus class used these
entries.

## 4. Milestone-only follow-ups

### 4.1 What the audit found

The follow-up texts come from the Rust format specs, through the registry
goldens (`crates/litchi-ops/tests/golden/registry/`) and the capability
goldens (`.../capabilities/`). In them:

- 784 follow-ups (96 distinct texts) name only a milestone, such as
  `M2 (validate)`, `M1-extended (R4x)`, `R4x`, `M6 ODT table grid edits`
  or `M2 XLSX generated style resource lifecycle`;
- the delta's `--check` saw only the registry rows that are delta gap rows,
  and it accepted a milestone text when the row's disposition had a ref;
- the capability goldens were not read at all.

### 4.2 The map

This record maps the texts in the dispositions file and does not edit the
Rust strings. That keeps the change to tooling and data, and avoids a
re-bless of the goldens that the parallel DOCX, PPTX and Numbers lanes
also change.

- A new `followUps` list in `tools/officecli_dispositions.json` has 21
  entries. Each matches follow-up texts by fnmatch pattern and, optionally,
  by format, and gives the id of the record or work package that owns the
  work now, with a note.
- The ids follow the format families' disposition entries:
  - DOCX, XLSX and PPTX breadth: 0332;
  - XLSB: 0322; PPT: 0323; DOC: 0324; ODT, ODP and RTF: 0328;
  - XLS projections: 0327; XLS structural writes: 0316.
- Some texts have a specific owner:

  | text | id | note |
  |---|---|---|
  | `M2 (validate)` | 0336 | stale: `validate` has landed |
  | `M2 (batch command and transfer families)` | 0336 | record 0336 F2 |
  | `M2 (DOCX table grid editing)` | WP4 | `add cell` |
  | `M2 (DOCX table properties)` | 0342 | stale: record 0342 implemented the table and row sets |
  | `M4 follow-up: project computed values in cell views` | 0266 | |
  | `M1-Δ1` (PPTX slide notes) | 0333 | record 0333 F1 |

- The delta tables show the id after the text (`M2 (validate) → 0336`). A
  new "Milestone follow-ups" section of `officecli-capability-delta.md`
  lists each entry with its formats and its count of golden follow-ups.
- `--check` reads every `followUp` of both golden sets. Encryption
  snapshots count for their format. The check fails on a milestone-only
  text that no entry maps, and on an entry that maps nothing. The earlier
  per-row rule stays.
- Unit tests cover:
  - specificity: an entry that names the format wins, then the longer
    pattern;
  - validation: an entry needs an id-shaped ref and a note;
  - the unmapped and stale cases;
  - a clean run on the committed goldens.

## 5. Owner-memo additions

Record 0343, on its own branch, adds items 14–24 to memo 0331. This record
continues the numbering after them in a separate section, so the two merge
without renumbering.

| # | item |
|---|---|
| 25 | XLS/ODS parity: the 392 P1 rows are deferred-P1 (§3) |
| 26 | `litchi::sheet::text` kept as non-Office tooling, against design §16.4 (record 0334 §2) |
| 27 | the capability delta, not `capabilities()`, is the property-level gap list: 3,220 `unlisted` rows, each with a disposition |
| 28 | the 57 residual bench misses accepted (the 54 of record 0337, plus the JSON-number and Keynote residuals of §2.2) |
| 29 | the `Box<dyn>` `litchi-core::sheet` traits recorded as an exception to goal constraint 2 |

## 6. Findings for other lanes

- **Stale capability follow-ups.** The Rust spec strings still read
  `M2's validate` on the issues and forms views (record 0336 landed
  `validate`). The DOCX `table-properties` operation gap says table and
  row property changes are "not integrated", although record 0342
  implemented them. The map records the current id. Correcting the texts
  is the DOCX and cross-format lanes' change, with a capability re-bless.
- **More misses than the audit counted.** About 13 misses were outside the
  matrix by the audit's reading; the actual number was 41. Most are reads
  and child listings of records 0324 and 0328 that already have their
  records' follow-ups.

## 7. Gaps and follow-ups

- **F-0344-SHEET-FEATURES** (XLS/ODS lanes): §2.2.
- **F-0344-XLSX-RUN** (XLSX lane): §2.2.
- **F-0344-XLSX-NOTE-RUNS** (XLSX lane): §2.2.
- **F-0344-ODP-SCENE** (ODP, record 0328 owner): §2.2.
- **F-0344-TSV**: `bench-matrix.tsv` is 906 KB, 86 % of the limit. The next
  ~350 rows need either a per-format TSV split or a shorter explanation
  column. `--check` fails before the file reaches 1 MiB.
- **F-0344-SPEC-TEXTS**: replace the milestone texts in the format specs
  with the mapped ids, and drop the stale ones (§6). Once a text names a
  record, its map entry becomes unused and `--check` asks for its removal.
- The coordinator's gate runs the tool unit tests by name. The new
  `tools.test_unified_bench_matrix` is not in that list yet.

## 8. Gates

Tree: this branch with trunk `926cbde41e` (record 0343) merged. The only
conflict was the owner memo; both sections are kept, items 14–24 and then
25–29. After the merge every report was regenerated. The bench matrix
lists 0343, and the delta and corpus outputs did not change. Build
environment per the rules, with every exit status checked.

| Command | Exit |
|---|---:|
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_unified_bench_matrix` (1,276 tests) | 0 |
| `python3 tools/officecli_capability_delta.py --check` (and the regeneration with `--officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI`: same output) | 0 |
| `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 tools/unified_bench_matrix.py --check` | 0 |
| `python3 tools/check_feature_matrix.py` (31 commands) | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `cargo fmt --all --check` | 0 |

No Rust source, golden or public name changed. The branch's diff against
trunk touches `tools/`, `docs/unified-ops/` and the generated reports
only. Clippy, the cargo test suites, rustdoc, the glossary gate,
`run_tests.sh`, the perf-baseline check and the golden re-bless therefore
have nothing of this record's to verify, and were not run. No glossary
row changes.
