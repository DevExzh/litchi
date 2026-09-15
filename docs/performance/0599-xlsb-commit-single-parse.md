# 0599: an XLSB cell-value commit parses the workbook once, not twice — and a proven no-op parses it not at all

Status: retained, implemented. `performance_claim: none` — the paired medians
and isolation-pair counters below are reported as evidence, not registered as a
claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **XLSB-1** of change
[0587](0587-remaining-opportunity-survey.md) (rank 19), in full. It is the first
XLSB *performance* record in the program: all sixteen prior XLSB records
(0304–0376) are correctness records carrying `performance_claim: none`.

## What was changed

Two files under `crates/litchi-xlsb/src`, plus a fixture generator in the
harness.

`Workbook::apply_cell_values` (`workbook/package.rs:137`) used to do this,
whatever the commit contained:

```rust
let mut candidate = self.package.clone();                      // clone 1
let snapshot = cell_values::workbook::apply_with_external_link_limits(
    &mut candidate, &uri, commit, self.external_link_limits)?;
*self = Self::from_opc_package_with_external_link_limits(      // reparse 2
    candidate, self.external_link_limits)?;
```

and the callee (`cell_values/workbook.rs:60`) used to do this for a changed
commit:

```rust
let mut candidate = package.clone();                           // clone 2
candidate.get_part_mut(worksheet)?.set_blob(updated.clone());  // blob copy
candidate.unsign();
let parsed = crate::Workbook::from_opc_package_with_external_link_limits(
    candidate.clone(), external_link_limits)?;                 // clone 3, reparse 1
…validate…
*package = candidate;                                          // parsed dropped
```

So an **exact no-op** commit — one whose patch reproduces the stored worksheet
bytes, which the callee proves at `updated.as_slice() == part.blob()` and then
returns early on — still paid for one package clone and one complete reparse of
a package it had just proved unchanged. A **real edit** paid for three
`OpcPackage` clones, one copy of the whole new worksheet blob, and **two**
independent full `Workbook` parses of byte-identical candidate bytes, the first
of which (`parsed`) was computed, used for validation, and dropped.

The change splits the callee into a crate-private
`apply_retaining_parse(&OpcPackage, …) -> Result<Applied>` which never
publishes, and a two-variant `Applied`:

* `Applied::Unchanged(Snapshot)` — the patch reproduced the stored bytes. No
  clone is taken and no candidate is parsed. `apply_cell_values` keeps `self`
  and returns the commit's snapshot.
* `Applied::Published { snapshot, workbook }` — the patch changed the
  worksheet. The candidate is cloned **once**, the blob is moved into it (not
  copied), it is unsigned, and it is parsed **once**; that parse is what
  `validate_dependencies` runs against, and it is handed back and installed as
  the new `self` only after every check has passed.

`apply_with_external_link_limits` and `apply` keep their signatures, their
`&mut OpcPackage` publication contract and their error identity: they are now
thin wrappers that call `apply_retaining_parse` and, on `Published`, take the
package back out of the validated workbook through a new crate-private
`Workbook::into_opc_package`. Their two in-crate callers in
`cell_values/root.rs` (`insert_candidate_cell:233`, publishing at `:995`, and
`transfer_cell:1253`, publishing at `:1332`) are untouched and keep working
exactly as before; they lose one of their three clones as a side effect and keep
their own second reparse, which this change does not address — see Limitations.

Net per commit: **one whole-workbook parse removed** on both paths, plus one
package clone on the no-op path and two package clones and one worksheet-blob
copy on the edit path.

Also added, in the harness: `tools/perf-baseline/src/bin/xlsb_synthetic_fixture.rs`,
behind the existing `xlsb-crud` feature. It writes a deterministic XLSB workbook
of a caller-chosen shape through the public `litchi_xlsb::writer` API and proves
it reopens through `Workbook::new` before writing it. This exists because change
0587 named the corpus ceiling — 22,715 bytes, one sheet, 48 cells — as the
blocker for every XLSB scaling question. No production code depends on it.

## Why it is sound

* **The no-op branch is value-identical by construction.** The early return is
  taken only when `commit.patch().apply(part.blob())` reproduced `part.blob()`
  byte for byte. The candidate the old code then reparsed *was* `self.package`
  — the same `Arc`-shared blobs, the same content types, the same
  relationships, not even unsigned (`unsign` runs only on the changed branch).
  `from_opc_package_with_external_link_limits` is a pure function of the package
  and the limits: it takes the package by value, reads it through `&self`, and
  every field it fills — `worksheet_names`, `worksheet_positions`,
  `worksheet_rel_ids`, `active_catalog_position`, `formula_context`,
  `shared_strings`, `styles`, `calc`, `is_1904`, `pivot_cache_definitions`,
  `structured_tables`, `chart_sheets`, `sheet_drawings`, `connections` — is
  derived from those bytes and the caller's `external_link_limits`, which are
  carried across unchanged. Re-deriving that from the same input yields the same
  value, so returning without re-deriving it is the same state. Two tests assert
  exactly this rather than argue it (below).
* **The edit branch reuses the parse it already validated.** `parsed` was built
  by `from_opc_package_with_external_link_limits(candidate.clone(), limits)` and
  the old code then built the published workbook by
  `from_opc_package_with_external_link_limits(candidate, limits)` — the same
  constructor, the same bytes, the same limits. Installing the first instead of
  computing the second is the same value. The package inside it is the candidate
  that the old code published, because the constructor stores the package it is
  given and never mutates it.
* **Atomicity (ADR 0003).** ADR 0003 requires that "public format editors publish
  only after their staged CRUD operation and typed readback succeed". Nothing the
  caller owns is touched until every check has passed: `apply_retaining_parse`
  takes `&OpcPackage`, builds the candidate in a local, and every `?` on the way —
  `get_part`, `require_worksheet`, `patch().apply`, `get_part_mut`, the candidate
  parse, the worksheet-index lookup, `parsed.worksheet(index)`,
  `validate_dependencies` — returns before anything is published. The single
  assignment `*self = *workbook` is the last statement. That is strictly tighter
  than the old shape, which mutated the callee's `*package` argument one
  statement earlier.
* **Error identity.** Every refusal is raised at the same point, from the same
  expression, with the same variant and message. The only reachable error the
  change removes is the one the no-op branch's reparse could have raised — and it
  could not raise one, because the workbook in hand is the result of the same
  constructor over the same bytes.
* **Preservation (ADR 0006).** Output bytes are unchanged. The `unsign()` that
  drops signatures on a real edit is where it was, and is still not run for a
  no-op. A 17-artifact differential (below) proves published bytes, readbacks and
  refusals identical.
* **Contracts untouched.** No public signature changed; the two new items are
  `pub(crate)`. No new `unsafe`, no new dependency, no limit relaxed, no
  malformed-input defence weakened, no allocation added (three are removed), no
  ambient I/O, no threading, no archive type, lock or executor exposed. The
  `Snapshot` returned is `commit.snapshot().clone()` on both branches, as before.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0, eight agents building concurrently. Every measured process was
pinned to **CPU 17** with `taskset`. Both legs are `--release --locked` builds of
`tools/perf-baseline`'s `xlsb_crud` selector — before from the read-only checkout
of `08d968f8e` (sha256 `2694b2b9…`), after from this branch's worktree (sha256
`b1f9d9d2…`). `xlsb_crud`'s own source is identical in both legs; only
`litchi-xlsb` differs.

Four fixtures. Two are real producers from the repository corpus; two are
synthetic, generated by the new harness binary and **not checked in** (they are
reproducible byte for byte from it; see `results/change-0599/fixture/`):

| fixture | bytes | sheets | stored cells, selected sheet | parts |
| --- | ---: | ---: | ---: | ---: |
| `testVarious.xlsb` (POI, real producer) | 22,715 | 1 | 48 | 17 |
| `cond_format.xlsb` (real producer) | 8,253 | 1 | 16 | 8 |
| `synthetic-4x500x8.xlsb` | 63,635 | 4 | 3,750 | 13 |
| `synthetic-4x2000x12.xlsb` | 323,710 | 4 | 22,500 | 13 |

### Instructions per commit (deterministic — **measured**)

Callgrind isolation pairs: the same case profiled at 5 and at 15 samples, the
totals differenced and divided by 10. `--cache-sim=no --branch-sim=no`. The
figure is instructions per *harness sample*, which is one `run_case` plus the
harness's own out-of-band SHA-256 of the saved output; that constant is identical
in both legs, so the **absolute difference** is the change and the percentage is
of the harness sample.

| fixture | case | before Ir/op | after Ir/op | Δ Ir | Δ |
| --- | --- | ---: | ---: | ---: | ---: |
| `testVarious.xlsb` | `noop_transaction_commit_save` | 52,308,801 | 29,325,884 | **−22,982,917** | **−43.94%** |
| `testVarious.xlsb` | `edit_one_existing_scalar_save` | 78,314,813 | 54,162,314 | **−24,152,499** | **−30.84%** |
| `synthetic-4x2000x12.xlsb` | `noop_transaction_commit_save` | 188,809,497 | 175,541,374 | −13,268,123 | −7.03% |
| `synthetic-4x2000x12.xlsb` | `edit_one_existing_scalar_save` | 357,557,295 | 343,509,287 | −14,048,008 | −3.93% |

**The removed work is one workbook parse, and the counters say so exactly.** The
per-symbol difference on `testVarious.xlsb` is *identical to the instruction*
between the no-op case and the edit case — `mce::codec::BoundedOutput::extend_from_slice`
−71,907,399, `realloc` −48,362,376 / −48,373,738, `RawVecInner::finish_grow`
−45,037,760, `quick_xml::attributes::IterState::next` −35,336,132,
`mce::codec::esc` −26,081,604 over the ten extra samples in both
(`results/change-0599/callgrind/symbol-diff-*.txt`). That is the signature of
exactly one `from_opc_package_with_external_link_limits` removed in each case:
the no-op loses its only one, the edit loses the first of its two. The whole
per-op difference between the two cases, 24,152,499 − 22,982,917 = **1,169,582
Ir**, is the edit path's two extra `OpcPackage` clones and its worksheet-blob
copy; the synthetic's equivalent is 779,885 Ir. Package clones are cheap because
`BlobPart` holds `Arc<Vec<u8>>` (`litchi-opc/src/part.rs:167`), so the clone
shares the bytes — the parse, not the copying, is what this change removes.

The reparse costs 22.98 M Ir on `testVarious.xlsb` and 13.27 M on the ten-times
larger synthetic because the cost is in the **features** a workbook declares, not
its cell count: `testVarious.xlsb` carries pivot caches, structured tables, chart
sheets, drawings and connections, whose XML the eager constructor re-parses
through the MCE codec, and the synthetic fixtures carry none.

### Paired timing (**measured**)

Order **A1 B1 B2 A2** (before, after, after, before), then two further `before`
legs **A3 A4** as an A/A control in the same window. 40 samples per leg after 3
warmups, `Instant` around `run_case`, warm page cache, pinned to CPU 17. `before`
is A1+A2 pooled, `after` B1+B2 pooled. All eight cases were run on all four
fixtures; every one is in `results/change-0599/timing/summary.txt`.

| fixture | case | before p50 | after p50 | after vs before | before vs after | p99 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `testVarious.xlsb` | `noop_transaction_commit_save` | 2.926 ms | 1.567 ms | **−46.43%** | +86.68% | −45.53% |
| | `edit_one_existing_scalar_save` | 4.392 ms | 3.013 ms | **−31.40%** | +45.77% | −31.38% |
| | `edit_ceil_one_percent_existing_cells_save` | 4.340 ms | 3.065 ms | **−29.38%** | +41.61% | −30.53% |
| `cond_format.xlsb` | `noop_transaction_commit_save` | 0.095 ms | 0.086 ms | −9.69% | +10.73% | −10.49% |
| | `edit_one_existing_scalar_save` | 0.170 ms | 0.154 ms | −9.69% | +10.73% | −6.76% |
| | `edit_ceil_one_percent_existing_cells_save` | 0.166 ms | 0.152 ms | −8.29% | +9.04% | −6.42% |
| `synthetic-4x500x8.xlsb` | `noop_transaction_commit_save` | 2.002 ms | 1.844 ms | −7.90% | +8.58% | −7.53% |
| | `edit_one_existing_scalar_save` | 4.216 ms | 4.016 ms | −4.76% | +5.00% | −3.76% |
| | `edit_ceil_one_percent_existing_cells_save` | 4.302 ms | 4.088 ms | −4.98% | +5.24% | −6.57% |
| `synthetic-4x2000x12.xlsb` | `noop_transaction_commit_save` | 14.388 ms | 14.157 ms | −1.60% | +1.63% | −1.85% |
| | `edit_one_existing_scalar_save` | 29.751 ms | 29.009 ms | −2.49% | +2.56% | −2.17% |
| | `edit_ceil_one_percent_existing_cells_save` | 32.540 ms | 31.771 ms | −2.36% | +2.42% | −3.02% |

**Nothing regressed above the program's review threshold.** The largest positive
p50 delta anywhere in the 32 case×fixture cells is **+1.93%**
(`cond_format.xlsb`, `full_text` — a path this change does not touch), against a
5% review trigger and this window's 10.04% floor. Every commit case moved down on
every fixture.

**A/A floor, same window.** 64 control pairs (A1/A2 and A3/A4, every case on
every fixture): |p50 delta| median **0.84%**, p90 **3.22%**, **max 10.04%**;
|p99 delta| median 1.75%, p90 10.57%, max 31.59%. The worst five p50 cells are
`synthetic-4x500x8` `edit_ceil…` 10.04%, `synthetic-4x500x8` `full_stored_cell_scan`
7.26% and `selected_worksheet_cell` 7.04%, `synthetic-4x2000x12` `full_text`
6.61%, `synthetic-4x500x8` `noop…` 4.23%. This window's floor is therefore
**worse** than the host's standing p50 ≈ 4%, and the reading has to respect it:

* On `testVarious.xlsb` the three commit deltas (−46.4%, −31.4%, −29.4%) are four
  to five times the worst floor cell anywhere and an order of magnitude above
  that fixture's own. **Reportable.**
* On `cond_format.xlsb` the three commit deltas (−9.7%, −9.7%, −8.3%) sit above
  that fixture's floor (its worst A/A cell is 1.7%) while its five read-only
  cases moved **+0.5% to +1.9%** in the same legs. The direction is commit-only.
  **Reportable.**
* On `synthetic-4x500x8.xlsb` the commit deltas (−7.9%, −4.8%, −5.0%) are inside
  that fixture's own A/A range, one of whose cells is the very case measured.
  **Not separable by timing.**
* On `synthetic-4x2000x12.xlsb` the commit deltas (−1.6% to −2.5%) are the same
  size as the drift its five untouched read-only cases show (−0.4% to −4.9%).
  **Not separable by timing.**

**The read-only cases are the in-run control.** `open_identify`,
`worksheet_catalog`, `selected_worksheet_cell`, `full_stored_cell_scan` and
`full_text` never reach `apply_cell_values`, so their delta is drift plus code
layout: −3.5% to −0.3% on `testVarious.xlsb`. Subtracting that offset from the
no-op delta gives ≈ −43%, against the deterministic **−43.94%** of instructions —
an independent agreement between the two methods that neither figure alone
provides.

## Correctness evidence

### A corpus differential over every XLSB artifact

`results/change-0599/differential/` retains the probe source and both reports.
Two binaries share one source file; one links `litchi-xlsb` at `08d968f8e`, the
other at this branch. Over **17 artifact paths** — all 15 `.xlsb` fixtures under
`test-data/`, of which two are byte-identical copies of `62815.xlsb`, plus the
two synthetic ones — and all **35 worksheets** they contain, each leg prints:

* the worksheet catalog, the worksheet count, and the SHA-256 of the saved
  package;
* per worksheet: the snapshot's source-byte digest, its stored-cell count, and a
  digest of every cell's reference, style index and typed value;
* per worksheet: an **exact no-op** commit — its `patch().is_empty()` flag, the
  published snapshot's source digest, the saved package digest, and the
  workbook's worksheet names, shared-string count and cell-XF count after
  publication;
* per worksheet: one **real scalar edit** — the same fields, plus a readback of
  the edited cell through a fresh `cell_values` call;
* one **refused publication** (an out-of-range style index) — the exact typed
  error string, and whether the saved package is byte-identical to what it was
  before the attempt.

**The two reports are byte-identical**, 208 lines each. Within them: all 35
no-op patches are empty and all 35 no-op publications save bytes identical to the
untouched baseline; the 19 worksheets that carry an editable scalar each produce
a changed save and a readback of the new value, and the other 16 report that they
have none; all 17 refusals give the same
`Unrecognized Cell iStyleRef: 16777215 (cell XF count N)` and `unchanged=true`.

### The harness's own gates

`xlsb_crud` exits non-zero on any failed gate. Across 6 legs × 4 fixtures × 8
cases = **192 observations**, every gate that applies is `true`
(`results/change-0599/timing/harness-gates.txt`): representative-output reopen
(72), semantic readback (192), exact no-op patch identity (24), output identity
across all 40 samples (72), untouched package members (72), malformed-input
refusal (192), tight read-limit refusal (192), tight cell-limit refusal (192),
sparse iteration without rectangular expansion (192). The changed-member list is
**empty for every no-op case** and exactly `['/xl/worksheets/sheet1.bin']` for
every edit case, on all four fixtures and in all six legs. Output SHA-256 is
identical across all six legs for every case.

### Tests added

`crates/litchi-xlsb/src/workbook/tests.rs` — a projection helper plus three
tests. `workbook_projection` destructures `Workbook` **exhaustively**, so a field
added later will not compile until the projection covers it; it renders every
field with `Debug`, sorts `StylesTable::num_fmts` (the one hash-ordered field),
and hashes every package part's content type, bytes and sorted relationships.

* `exact_noop_cell_value_commit_publishes_nothing` — the control commit's patch
  is empty; after publication the workbook's projection is **unchanged**, and it
  equals the projection of a fresh
  `from_opc_package_with_external_link_limits` over its own package. This is the
  oracle for the skipped reparse.
* `cell_value_publication_installs_the_parse_that_validated_it` — after a real
  edit the projection differs from before, and equals the projection of a fresh
  parse of the bytes it published. This is the oracle the brief asks for.
* `refused_cell_value_publication_leaves_the_workbook_unchanged` — an
  out-of-range style index is refused and the projection is identical to the
  pre-attempt one (ADR 0003).

`crates/litchi-xlsb/tests/cell_value_edit.rs` — two public-boundary tests:

* `real_fixture_exact_noop_publication_saves_byte_identical_output` — save,
  publish an exact no-op, save again: **not one byte differs**, and the
  stored-cell list reads back identically.
* `real_fixture_refused_publication_saves_byte_identical_output` — the same
  byte-identity assertion after a refusal.

### Gates

Run in the worktree at `perf/0599-xlsb-commit-single-parse`; tails in
`results/change-0599/gates.txt`.

| gate | result |
| --- | --- |
| `cargo fmt --all --check` (workspace) | pass, no diff |
| `cargo fmt --all --check` (`tools/perf-baseline`) | pass, no diff |
| `cargo clippy -p litchi-xlsb --all-targets` | pass, no warning (workspace lints are `deny`) |
| `cargo clippy --features xlsb-crud --bins` (`tools/perf-baseline`) | pass, no warning |
| `cargo test -p litchi-xlsb` | pass — 17 test binaries, 0 failures; 573 library tests, 7 `cell_value_edit` integration tests |
| `cargo doc -p litchi-xlsb --no-deps` | pass (rustdoc lints are `deny`) |
| 17-artifact XLSB differential | byte-identical between legs |
| `xlsb_crud` gates, 192 observations | every applicable gate `true` |

No pre-existing failure was encountered.

## Validation preserved

Every validation the commit path ran still runs, exactly once, in the same order:
`require_worksheet` on the target part; `commit.patch().apply` against the stored
bytes; the complete `Workbook` reparse of the candidate, which is the
whole-workbook readback ADR 0003 requires; the candidate worksheet-URI lookup and
its `WorksheetNotFound` refusal; `parsed.worksheet(index)`, which decodes the
candidate worksheet through the typed codec and discards it; and
`validate_dependencies`, which proves every committed cell's style index, font,
fill, border, number format, shared-string index and rich-string run fonts
against the candidate's own tables. The bounded-allocation defences in the
worksheet decoder (`entries.try_reserve(1)` per cell), `Cursor::guard`,
`raw::record::Limits`, `cell_values::Limits`, `litchi_opc::ReadLimits` and
`ExternalLinkLimits` are untouched, and the harness re-proves the last three
refuse on every fixture in every leg. What is no longer run is the **second**
parse of the same bytes, whose only output was discarded, and the no-op path's
re-derivation of state the workbook already held. Validation still never mutates,
and an exact no-op is still exact.

## Limitations — what is not claimed

* **No claim is registered.** `performance_claim: none`.
* The saving is in **CPU only**. Bytes read, syscalls, peak RSS and allocation
  counts were not measured. Allocations are removed by construction — a whole
  parse and one package clone on the no-op path, a parse, two package clones and
  one worksheet-blob copy on the edit path — but none was counted, because
  `xlsb_crud` has no allocator-metrics hookup.
* Scoped to this host, this build (`rustc 1.95.0`, `--release --locked`, default
  release profile), a warm in-memory source, a single pinned core, and the three
  `xlsb_crud` commit scenarios on the four fixtures named. **Nothing is claimed
  for the two synthetic fixtures' timing**: on both, the deltas are inside the
  A/A range measured in the same window, and only the instruction counts separate
  the legs there.
* **The saving tracks a workbook's feature surface, not its size.** It is the
  cost of one eager `Workbook` parse, which is dominated by the XML of drawings,
  pivot caches, structured tables, chart sheets and connections. On the
  feature-rich 22.7 KB real producer it is 44% of a no-op commit and save; on a
  feature-free 324 KB synthetic it is 7%. **A large real-producer `.xlsb` is
  still absent from this corpus**, so the figure for one is unknown, bounded
  below by 7% and above by 44%.
* Callgrind counts `rep movsb` per byte and runs SHA-256 in software, so the Ir
  percentages overstate the copying and hashing shares; they are used here to
  count parses and to attribute symbols, and the timing legs carry the latency
  statement.
* **Two sibling sites keep the same duplication and were not changed.**
  `cell_values::root::insert_candidate_cell` (`root.rs:233`) and
  `transfer_cell` (`root.rs:1253`) still clone, call
  `apply_with_external_link_limits`, and then reparse the package a second time
  into `*workbook`; each would lose one further parse by taking
  `Applied::Published` directly, exactly as `apply_cell_values` now does. They
  are reached through `apply_workbook_structure`, for which **no harness selector
  exists**, so the change was left out of this record rather than landed
  unmeasured. `Workbook::apply_sparklines` (`package.rs:247`) and
  `apply_cell_watches` (`package.rs:280`) have a third shape: they publish
  straight into `&mut self.package` and never refresh the derived fields at all.
  Whether that is intended is a correctness question this record does not answer.
* **`Workbook::edit_opc` is already minimal** (one clone, one parse) and was not
  touched.
* The synthetic fixtures are producer-free: no pivot cache, table, chart sheet,
  drawing, connection, external link or VBA project. They answer "does the commit
  path scale with cell count" and nothing about real-producer feature surface.
* Change 0587's XLSB-1 entry predicted "+84% no-op, +178% one edit" as the cost
  over a read-only baseline and named its falsification condition as
  "`validate_dependencies` rather than the reparse dominates the delta". That
  condition did **not** fire: the per-symbol difference is entirely XML and
  allocator work inside the removed parse, and `validate_dependencies` runs the
  same number of times before and after. The survey's own single-leg figures are
  reproduced within this window's drift (its 2.766 ms no-op and 4.189 ms edit
  against 2.926 and 4.392 here, both single legs on a contended host).

## Retained evidence

[`results/change-0599/README.md`](results/change-0599/README.md) — the callgrind
isolation pairs and per-symbol differences with their scripts, all 24 paired
timing reports with the analysis script and the harness-gate extract, the
17-artifact differential with its probe source, the synthetic-fixture provenance,
`gates.txt`, `decision.json` and `log-sections.md`.
