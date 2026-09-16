# 0633: the XLS source-backed numeric commit's second complete target parse is deleted; the open's two framing passes are priced at 1.6% each and frozen

Status: retained. `performance_claim: none` — this record carries a
value-identical removal of one complete `Workbook::new` from
`Transaction::commit_source_backed`, measured in both directions beside its
floor; the residue of a second, much smaller reuse in `Snapshot::from_bytes`;
and one frozen design for the fusion this change was briefed to attempt and did
not, with the measurement that says why. No claim is registered.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is queue item **XLS-9** of change
[0630](0630-queue-refresh-after-the-first-wave.md), the XLS commit's remaining
whole-workbook work, priced by change
[0620](0620-xls-edit-save-attribution.md). 0620 attributed the path, removed one
of the four complete parses an open-edit-save runs, and left two items on the
table with their sizes: **(a)** `Snapshot::from_bytes` frames the Workbook
stream twice, 12.5% and 82.1% of the open, and **(b)** the commit parses the
composed target twice, the second parse costing 137,426,515 Ir at 0620's base
(130,182,770 Ir at this one). This record takes both. It implements (b), which
turns out to be value-identical and worth 139 M Ir. It does **not** implement (a): the measurement below shows the
*framing* in those two passes is 1.60% of the open each, not 12.5% and 82.1%,
and that neither direction of the briefed fusion is value-identical. The
value-identical residue of (a) — one deep copy of the shared-string property
table — is implemented.

## Method

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
cargo 1.95.0, valgrind 3.26.0. Base commit `c7326f680`, branch
`perf/0633-xls-commit-single-framing`; the before leg is the shared read-only
checkout at `/home/zhuhe/code/litchi-worktrees/before-c7326f680`. Every measured
child is pinned to CPU 9 with `taskset` under `setarch -R`,
`RAYON_NUM_THREADS=1`.

Change 0620's probe, corpus differential, capture scripts and analyzer are
reused unchanged except for the CPU pin and the scratch path
(`results/change-0633/probe/`, `capture_*.sh`); this record adds a third scratch
binary, the 0541-style first-error matrix (`probe/matrix-main.rs`), and two
analyzers (`framing_attribution.py`, `make_analysis.py`). Fixtures and
operations are 0620's: `54016.xls` (984,576 B, one worksheet, ~30 k cell
records), `WithCustomViews.xls` (165,888 B, three worksheets),
`FormulaEvalTestData.xls` (macro-bearing, so only the generic paths run) and
`59858.xls`, through `open`, `number-plan`, `number-source-backed`,
`number-generic`, `string-generic` and `noop-generic`. Instruction attribution
is change 0574's isolation method with `--separate-callers=2`.

One fix to the reused analyzer is worth naming, because it changes what the
retained annotations can be read for: 0620's `analyze.py` matched
`^\s*([\d,]+) \(([\d.]+)%\)`, and `callgrind_annotate` right-aligns the
percentage, so every row under 10% was silently dropped. The copy in this packet
matches `\(\s*([\d.]+)%\)`. Nothing 0620 reported depended on the dropped rows;
everything in section 1 below does.

## What one `Snapshot::from_bytes` is actually made of

Before leg, `54016.xls`, one `open` operation, callgrind inclusive Ir
(`results/change-0633/framing-attribution.txt`; `WithCustomViews.xls` beside
it):

| phase | Ir/op | share |
| --- | ---: | ---: |
| whole operation | 160,833,535 | 100.00% |
| `Snapshot::from_bytes` | 155,960,385 | 96.97% |
| ↳ `PackageEditor::open` (CFB capture) | 4,172,688 | 2.59% |
| ↳ the constructor | 151,787,620 | 94.38% |
| **pass 1** — `parse_worksheet`, the offset inventory | **7,864,092** | **4.89%** |
| ↳ **framing: `Records::next`** | **2,579,172** | **1.60%** |
| ↳ `push_entry` | 2,732,787 | 1.70% |
| ↳ `parse_reference` | 444,120 | 0.28% |
| `resolve_shared_strings` | 3,629,090 | 2.26% |
| **pass 2** — `Workbook::new`, the complete eager parse | **130,517,621** | **81.15%** |
| ↳ `OleFile::open_stream` (the second Workbook-stream extraction) | 1,948,501 | 1.21% |
| ↳ `SharedStringTable::parse_from_records` | 11,451,901 | 7.12% |
| ↳ `parse_worksheet_records_with_compatibility` | 113,616,319 | 70.64% |
| &nbsp;&nbsp;↳ **framing: `Records::next`** | **2,579,104** | **1.60%** |
| &nbsp;&nbsp;↳ `add_cell` | 54,395,849 | 33.82% |
| &nbsp;&nbsp;&nbsp;&nbsp;↳ `BTreeMap::insert` | 32,438,095 | 20.17% |
| drop glue of the discarded `Workbook` | 7,524,989 | 4.68% |

**The two framing passes are 2,579,172 and 2,579,104 Ir — 1.60% of the open
each, and they agree to 68 instructions because they frame the same bytes.**
On `WithCustomViews.xls` they are 136,680 and 136,476 (0.92% and 0.91%).
Perfectly fusing them would remove at most **1.60%** of one open. What is
actually expensive is the semantic work each pass does *around* its framing: the
eager parser's twenty per-record collectors and its `BTreeMap<(u32,u32), Cell>`
(`add_cell` alone is 33.82% of the open), and, on the inventory side, the entry
vector and the per-`LabelSst` text copy.

This is the same lesson change 0604 recorded in a different form: a count of
walks is not a cost. Both owners walk every record; only one of them pays for
it.

## What was changed

Scope: `crates/litchi-xls` only — in `cell_values/mod.rs` one constructor split
in two, one private struct added, one helper extracted and one function deleted;
in `workbook/model.rs` one `pub(crate)` accessor; and three tests. No public API, error type, limit or output byte changed; no new
`unsafe`, no new dependency.

### 1. The commit verifies its target on the parse its own reopen already ran

`commit_source_backed_numeric` materialized the composed target, reopened it
through `Snapshot::from_bytes` — which runs a complete `Workbook::new` and then
**drops it** — and then called `verify_source_backed_numeric_target`, which ran
a *second* complete `Workbook::new` over the identical bytes to answer the five
checks the publication boundary owes:

```rust
let target = Snapshot::from_bytes(target_bytes.clone())?;
if target.bytes() != target_bytes.as_slice() { /* refuse */ }
verify_source_backed_numeric_target(&source, &target, &changes)?;   // Workbook::new again
```

`Snapshot::from_package_editor` is now `Snapshot::open_package(package,
verification)`, taking an optional `NumericTargetVerification` that carries the
materialized bytes, the source snapshot and the changes. The plain constructor
passes `None` and is unchanged. The commit passes `Some(…)`, and the five checks
run inside the constructor at the point they ran before, on the `Workbook` the
constructor already holds:

```rust
if target_bytes != self.expected_bytes { /* “…changed during complete reopen” */ }
if self.source.inner.workbook_path != target_workbook_path { /* “…changed the Workbook owner” */ }
require_public_worksheet_coverage(workbook, &self.source.inner.sheets)?;
require_unprotected_workbook(workbook)?;
require_macro_free_workbook(workbook)?;
let _ = carry_fixed_numeric_inventory(self.source, target_workbook_stream, self.changes)?;
verify_public_numeric_readback(workbook, self.source, self.changes)
```

The second `Workbook::new` is gone. Nothing else moved.

### 2. The snapshot retains the shared-string property table instead of copying it

The constructor built its `shared_string_properties` by cloning the reader's
table one entry at a time:

```rust
for index in 0..strings.len() {
    properties.push(workbook.shared_string_properties(index).cloned().map(Box::new));
}
```

`Workbook` already keeps exactly that vector — same element type — behind an
`Arc`. `retained_shared_string_properties` takes the shared handle when the
table covers the shared-string list one entry for one and the length fits `u32`,
and otherwise falls back to the copying loop, which still truncates a longer
table and pads a shorter one exactly as before. This is the residue of part (a):
the one thing the inventory and the eager parse both build that the inventory
can take from the parse without moving a refusal.

## Why it is sound

**The removed parse could only have reproduced the one that remains.**
`Workbook::new` is a pure function of its bytes and `OpenOptions::default()`,
and both calls read the same bytes under the same default `Limits`: the first
parses `package.finish()`, which becomes `target.inner.bytes`; the second parsed
`target.bytes()`. The first check of the verification — the one 0620's code ran
one line earlier — proves those bytes are the materialized target byte for byte
before any check consumes the parse. This is change 0620's own argument for
reusing `SourcePolicyFacts`, applied one level out, and the corpus differential
below tests it rather than assuming it. The parse reaches no clock, no RNG and
no ambient I/O; `OleFile::list_streams` walks the directory tree into a `Vec`
and `pivot_cache_stream_paths` sorts by stream id, so no hash iteration order
reaches a verdict.

**The order of every refusal is preserved, and it is preserved by construction.**
The verification runs *after* `parse_workbook_stream`, after `Workbook::new`,
after `SourcePolicyFacts::from_workbook` (which is the target's own coverage
check against the target's own sheet list) and after `resolve_shared_strings` —
the same four owners, in the same order, that ran before
`verify_source_backed_numeric_target` was reached. Inside the verification the
five checks keep their order, so the case 0620 identified as the blocker is
respected exactly: `require_public_worksheet_coverage` still runs against the
**source's** sheet list *before* `carry_fixed_numeric_inventory` proves the two
lists describe the same bytes. Nothing was reordered, so nothing needed to be
proved equal.

**ADR 0003's publication boundary is untouched.** "Publish only after their
staged CRUD operation and typed readback succeed" and "the complete package is
reopened under the retained limits" both still hold: the target is still
materialized, still reopened through `PackageEditor::open` with
`Targets::default()` and `Limits::default()`, still parsed in full by an
independent `Workbook::new`, still carried through `carry_fixed_numeric_inventory`
and still read back by `verify_public_numeric_readback`. What is gone is the
*duplicate* of that parse, not the parse. Change
[0016](changes/0016-xls-commit-editor-reuse.md)'s standing instruction — "the
final full reader and typed readback remain the publication boundary … target a
different source of whole-workbook work, not remove either retained validation
layer" — is the instruction this change follows; both layers remain, and the
callgrind rows below show the checks costing the same as before (4,420,759 Ir
after against 4,419,288 Ir before) with only the parse removed.

ADR 0005's cache contract is not engaged. The snapshot retains no reader: the
`Workbook` is a local binding inside one constructor call, dropped before the
`Snapshot` is built, and the plain open still drops it at the exact statement it
dropped at before. `Snapshot`, `Commit`, `Patch` and every public type are
byte-for-byte the same shape.

**The property table reuse is value-identical or it does not happen.** The
shared handle is taken only when `properties.len() == strings.len()`, and each
element of the shared vector is by definition equal to the `Clone` of itself
that the loop produced. The `u32::try_from(len)` guard keeps the "SST index
exceeds `u32`" refusal reachable on a table the loop would have refused. The
vector is immutable behind an `Arc` in both the reader and the snapshot, and the
reader is dropped inside the constructor, so no aliasing is created that
outlives the call.

## Measured

### Counts (deterministic, before → after)

Per operation, the probe's counting allocator armed for exactly one measured
iteration (`results/change-0633/counters-*.jsonl`, rendered in
`analysis.txt` §2). Twenty (fixture, operation) rows.

The implemented commit change:

| fixture | metric | before | after | delta |
| --- | --- | ---: | ---: | ---: |
| `54016.xls` `number-source-backed` | allocation calls | 161,751 | 121,641 | **−40,110 (−24.80%)** |
| | allocated bytes | 76,204,234 | 59,090,612 | **−17,113,622 (−22.46%)** |
| | peak live bytes | 26,082,823 | 25,628,658 | −454,165 (−1.74%) |
| `WithCustomViews.xls` `number-source-backed` | allocation calls | 20,945 | 15,163 | **−5,782 (−27.61%)** |
| | allocated bytes | 9,911,311 | 7,700,045 | −2,211,266 (−22.31%) |
| | peak live bytes | 3,176,770 | 2,969,885 | −206,885 (−6.51%) |

The property-table reuse, visible on every `open` and therefore on every
operation:

| fixture, `open` | allocation calls | allocated bytes | peak live bytes |
| --- | ---: | ---: | ---: |
| `54016.xls` | 51,818 → 51,814 | 26,140,189 → 26,076,893 (−0.24%) | identical |
| `WithCustomViews.xls` | 6,839 → 6,781 | 3,285,160 → 3,278,252 (−0.21%) | identical |
| `FormulaEvalTestData.xls` | 28,718 → 28,716 | 4,772,340 → 4,762,308 (−0.21%) | identical |
| `59858.xls` | 9,535 → 9,523 | 8,187,958 → 8,184,698 (−0.04%) | identical |

`peak_live_bytes` is identical on all twenty rows except the two source-backed
commits, where it falls; `published_bytes` and the complete source-backed
diagnostics object (splice count, replacement bytes, changed spans, source and
target Workbook lengths, fingerprints, operation shape, versions) are identical
on every row.

### Instructions and cycles (isolation pairs, before → after)

| scenario | callgrind Ir/op | native cycles/op | native instructions/op |
| --- | --- | --- | --- |
| `54016/number-source-backed` | 887,985,767 → 748,473,025 (**−15.71%**) | 179,110,791 → 128,547,808 (**−28.23%**) | 467,095,195 → 328,369,298 (−29.70%) |
| `cv/number-source-backed` | 114,019,837 → 101,121,849 (**−11.31%**) | 16,660,730 → 12,244,019 (**−26.51%**) | 43,664,526 → 30,876,780 (−29.29%) |
| `54016/open` (control) | −0.25% | −1.36% | −0.22% |
| `54016/number-plan` (control) | −0.09% | −3.45% | −0.15% |
| `54016/number-generic` (control) | −0.12% | −3.39% | −0.19% |
| `54016/string-generic` (control) | −0.41% | −4.03% | −0.25% |
| `54016/noop-generic` (control) | −0.28% | −4.55% | −0.29% |
| `cv/*` (five controls) | −0.33% to +0.08% | +1.36% to +1.90% | −0.09% to +1.15% |
| `formula/*` (four controls) | +0.38% to +0.98% | −2.67% to +3.35% | −3.88% to +1.01% |

The callgrind line that carries the change is exact. On `54016.xls`:

| chain | before | after |
| --- | ---: | ---: |
| `Workbook::new'verify_source_backed_numeric_target'commit_source_backed` | **130,182,770** | **0** |
| ↳ `parse_workbook` under it | 130,018,950 | 0 |
| drop glue of that `Workbook` | 9,024,905 | 0 |
| the checks' own body (`verify_source_backed_numeric_target` → `NumericTargetVerification::run`) | 4,419,288 | 4,420,759 |

and on `WithCustomViews.xls` 12,109,922 → 0 with 739,993 → 0 of drop glue. The
whole-operation delta, −139,512,742 Ir, is those two rows; the checks themselves
cost 1,471 Ir more, which is the price of reading a reference instead of a local.

`formula/*` is the one control that moves the wrong way in instructions, by
+0.38% to +0.98%. The whole of it sits in the constructor's own inlined body —
no named callee moves by more than 20,000 Ir, `Workbook::new` under it moves
−7,195, and the same scenario moves −0.28% in *native* instructions. It is the
inlining of one renamed function differing between two separately compiled
binaries, not work added; the deterministic allocation counters for that fixture
move only downward.

### Paired timing (A1 B1 B2 A2, `commit` phase p50, nanoseconds)

**Three rounds are retained, and the host was never quiet.** Seven other agents
shared the machine for this change's whole measurement window (load average
9-22). Round 1's same-binary A/A floor on `54016.xls` reached +12.3% at p50 and
its `cv` tail is unusable (p95 and p99 disagree with p50 by 120 points); round 2
is better; round 3 is the only one whose floor on both changed scenarios is
under 2%. All three are in the packet (`latency/`, `latency-round2/`,
`latency-round3/`); round 3 is quoted.

Round 3 (`latency-round3/latency-summary.json`), 40 samples on `54016.xls`, 120
on the others:

| scenario | a1 p50 | b1 p50 | forward | reverse | p95 fwd/rev | p99 fwd/rev | A/A floor | B/B floor |
| --- | ---: | ---: | ---: | ---: | --- | --- | ---: | ---: |
| `54016/number-source-backed` | 25,781,630 | 15,856,600 | **−38.50%** | **−38.25%** | −39.48% / −41.33% | −38.49% / −40.70% | **−0.15%** | +0.25% |
| `cv/number-source-backed` | 2,432,018 | 1,583,743 | **−34.88%** | **−36.75%** | −34.39% / −36.86% | −36.75% / −36.61% | +1.92% | −1.01% |

The six independent p50 readings of the changed scenario across the three rounds
are **−38.13%, −43.18%, −41.82%, −44.06%, −38.50%, −38.25%** on `54016.xls` and
−37.50%, −35.95%, −37.83%, −37.33%, −34.88%, −36.75% on `WithCustomViews.xls`.

**The controls say how much of that spread is the host.** In round 3 the `open`
phase — the same code on both legs, run inside every one of these processes —
moves +1.95% to +13.17% forward on `54016.xls` and +0.18% to +2.34% on the other
two fixtures; the three unchanged `54016` commits move +7.46% to +15.29% forward
and +1.79% to +8.16% reverse, with same-binary A/A floors from +0.05% to +10.54%
in the same round. Across the three rounds those same controls range from −4.29%
to +24.89%. Their *instruction* counts move −0.09% to −0.41% and their native
cycles −3.39% to −4.55%, so none of that is work: on this host, an 11 ms XLS
operation has a control band of roughly ±15% at p50 while eight agents share 32
cores. `WithCustomViews.xls` and `FormulaEvalTestData.xls`, whose operations are
1-2 ms, stay within ±5% in every round and are the fixtures whose timing carries
weight.

The changed scenario is two to three times outside even the `54016` band, agrees
in both directions in every round, and agrees with the deterministic counts. The
sub-microsecond `noop-generic` commit is excluded: its floors exceed its deltas.

### The registered selectors still cannot see this, for the reason 0620 gave

All twelve registered cases ran A1 B1 B2 A2 (10 warmups / 100 samples for the
numeric group, 20 / 200 for the others). Whole-case `elapsed_ns` p50:

| case | forward | reverse | A/A | B/B |
| --- | ---: | ---: | ---: | ---: |
| `xls_numeric_source_backed_number_edit_save` | +0.32% | −1.86% | +0.52% | −1.67% |
| `xls_numeric_source_backed_rk_mulrk_edit_save` | −1.94% | −1.84% | −0.29% | −0.18% |
| `xls_numeric_plan_only_number_edit_save` (control) | −0.01% | +0.04% | −0.29% | −0.24% |
| `xls_numeric_plan_only_rk_mulrk_edit_save` (control) | −0.55% | +0.19% | −0.55% | +0.19% |
| `xls_numeric_eager_number_edit_save` (control) | −2.28% | −0.62% | −0.85% | +0.83% |
| `xls_numeric_eager_rk_mulrk_edit_save` (control) | −1.68% | +2.98% | −0.68% | +4.02% |

The two cases that exercise `commit_source_backed` do not move, on a change that
moves the same function by 35-44% on a real workbook. 0620 explained why and
this reproduces it: the registered Number corpus is a 16,995,840-byte CFB whose
Workbook stream is 80,946 bytes (0.48%) and the RK/MulRK corpus is 1,665 of
202,752 (0.82%), while the removed work is proportional to the Workbook stream,
which is 94.2% of `54016.xls` and 90.7% of `WithCustomViews.xls`. The brief for
change 0601's successor stands unchanged and is now supported by a second
measurement.

All twelve cases keep their `output_sha256` across all four rounds, including
the Number `f8f37064…` and RK/MulRK `ddf5d5b8…` digests change 0138 recorded.

**One selector group exceeded the 5% review trigger and was chased, not
explained away.** `xls_owned_source_open`, `…_list_worksheets` and
`…_one_cell` moved +7.06% to +9.07% forward and +4.35% to +8.30% reverse at p50
against same-binary floors under 2.5%. Those three cases drive
`litchi_xls::SourceBackedWorkbook::from_read_at` — the lazy source-backed
reader, on a path this change does not touch. A callgrind isolation pair over
the three of them (`--warmup 2`, `--samples 5` against `--samples 25`,
differenced and divided by the 60 extra operations) prices one measured
operation at **3,978,024 Ir before and 3,978,917 Ir after, +0.02%**. The work is
identical; the wall-clock move is the host and code layout on a 90 µs operation.
Evidence in `owned-source-chase.txt`.

## Correctness evidence

**Corpus differential (the brief's oracle).** 0620's corpus binary, rebuilt
against both legs, drives all five publication paths over every `.xls` in the
repository and prints, per file and operation, either the exact refusal text or
the SHA-256 of the complete published artifact with its diagnostics and patch
lengths. 126 fixtures, 591 rows per run, three runs per leg, **3,546 rows
compared** (`corpus/`, `corpus-summary.txt`):

| oracle | rows | mismatches |
| --- | ---: | ---: |
| exact typed refusal text | 591 | **0** |
| every reported field except the digest (published length, changed cells, touched streams, splice count, replacement bytes, changed spans, target Workbook length, `is_noop`, patch lengths and emptiness, inverse digest) | 591 | **0** |
| the artifact digest | 591 | **0** |

**Zero rows had to be excluded**, where 0620 had to exclude four. Sixteen runs
per leg over the two fixtures 0620 identified now yield exactly one digest each:
change [0625](0625-cfb-writer-deterministic-storage-order.md) serialized explicitly created
storages in a canonical order, and 0620's CFB storage-order finding is resolved
at this base. That is reported here because it makes this change's digest oracle
strictly stronger than its predecessor's.

**0541-style first-error matrix.** A third scratch binary
(`probe/matrix-main.rs`) rewrites one small valid package's Workbook stream with
one defect at one chosen position — in the globals BOF, at `FilePass`, at a
duplicated or truncated `SST`, with every `XF` removed, at a `BoundSheet8`
pointing past the end, at a duplicated position and a duplicated name, at the
worksheet BOF, at a missing worksheet EOF, inside a `Number`, `RK`, `BoolErr`,
`LabelSst` and `Formula` payload, at a stray `String`, at a truncated frame
header and an overrunning payload, and at the container — and prints the exact
first typed refusal `Snapshot::from_bytes` produces. **29 cases, 24 typed
refusals, 5 accepted; 0 rows differ between the legs**
(`matrix/matrix-{before,after}.jsonl`). The matrix is the oracle the frozen
design needs and is reproduced as a unit test.

**Unit tests.** Three added to `cell_values/tests.rs`:
`snapshot_open_first_error_matrix_is_frozen` (15 of the matrix rows with their
exact messages, as a frozen table, so any future fusion attempt fails loudly),
`snapshot_retains_the_open_shared_string_property_table` (three packages, every
index of the retained table compared against the complete open's own answer) and
`source_backed_numeric_target_verification_agrees_with_an_independent_parse` (a
published artifact re-parsed independently and put through all four checks plus
the readback, which is the property that made the second parse redundant).

**Gates.** `cargo fmt --all --check`; `cargo clippy -p litchi-xls --all-targets`;
`cargo test -p litchi-xls` (72 suites, 1,402 tests); `cargo doc -p litchi-xls
--no-deps`; the consumer `cargo test -p litchi --features docx,xlsx,pptx,xls`
(26 suites, 266 tests); and the harness's own `cargo test` in
`tools/perf-baseline` (19 suites, 531 tests). All clean; tails in
`results/change-0633/gates.txt`.

## Validation preserved

Every mandatory validation that ran before this change still runs, on the same
bytes, in the same order. The target is materialized, reopened through
`PackageEditor::open` under `Targets::default()` and `Limits::default()`,
inventoried by `parse_workbook_stream`, parsed in full by `Workbook::new`,
checked by `SourcePolicyFacts::from_workbook` against its own sheet list,
resolved by `resolve_shared_strings`, proved byte-identical to the materialized
plan, checked for a changed Workbook owner, checked by
`require_public_worksheet_coverage` against the *source's* sheet list, by
`require_unprotected_workbook` and `require_macro_free_workbook`, carried
through `carry_fixed_numeric_inventory`, and read back by
`verify_public_numeric_readback`. `require_macro_free_container`'s independent
CFB-directory check, the source lineage check, the overlay fingerprints, the
same-length splice validation, the artifact-length check in
`materialize_numeric_plan`, the exact `Patch` and its inverse, and the exact
no-op fast path are untouched. No limit moved, no `unsafe` was added, no public
type changed, and the plan-only, eager, generic and no-op paths are untouched
and were measured only as controls.

## The frozen design: sharing one framing pass between the inventory and the eager parse

The brief asked for one framing pass to be shared, "either derive the offset
inventory from the parse's own record walk or feed the inventory's frames to the
parse … keeping error identity and order". **Neither direction is admissible,
and the prize is 1.60% of an open.** The design is frozen here with what it
would take to revisit it.

**The prize.** Section 1 measures it: the inventory's `Records::next` is
2,579,172 Ir (1.60% of the open on `54016.xls`, 0.92% on
`WithCustomViews.xls`), the eager parse's is 2,579,104 (1.60%, 0.91%). A perfect
fusion removes one of the two. On this host that is inside the A/A floor at p50
even in a quiet window, so it could only ever be reported as a count.

**Direction 1, deriving the inventory from the eager parse's walk, is not
value-identical, and the matrix shows it by row.** The inventory frames first,
so its refusals *shadow* the eager parser's on any stream both owners can see.
`cell-xf-past-end` is the clean case: a `Number` whose XF index is 4095 refuses
with `Invalid record 0x0203: cell XF index 4095 is outside 21 workbook
resources`, which is `parse_workbook_stream`'s message; the eager parser
validates the same index through `Formatting::validate_cell_xf` with a different
one, and never gets to speak. `number-truncated`, `number-nan`,
`boolerr-bad-flag`, `formula-string-cache-orphan`, `globals-duplicate-sst`,
`globals-no-xf` and `boundsheet-past-end` are seven more. Worse, the eager
parser *swallows* every per-worksheet failure — `package.rs:217` is `Err(_) =>
{}` — so a sheet the inventory refuses precisely is a sheet the eager walk
simply drops, and the refusal the user would see becomes
`require_public_worksheet_coverage`'s generic "worksheet at tab position N was
not published by the complete XLS reader" (matrix rows `worksheet-bof-missing`
and `stray-string-after-formula` are exactly that message arriving for two
different defects). Deriving the inventory from that walk would replace precise
per-record refusals with a per-tab one and would produce no inventory at all for
a sheet the eager parser dropped — which is the case
`cell_values/mod.rs:535-539` says the coverage check exists for.

**Direction 2, feeding the inventory's frames to the parse, costs more than it
saves.** `litchi_biff::Records` is a borrowed, allocation-free iterator; handing
its frames to a second consumer means materializing them. `54016.xls` frames
about 50,000 records; a `Vec<RecordRef>` of them is roughly 2 MB of allocation
and copying to avoid 2.58 M Ir of framing — a net loss on both counts, and a new
unbounded intermediate in a path whose bounded-resource discipline is the point.

**What is redundant between the two passes, and what happened to it.** Three
things, with their prices, all of them larger than the framing:

1. **The property table.** The constructor deep-copied a vector the reader
   already owned behind an `Arc`. **Implemented** above; it is the only one of
   the three that is value-identical as it stands.
2. **The Workbook stream is extracted from the CFB twice** —
   `OleFile::open_stream` inside `parse_workbook` is 1,948,501 Ir, 1.21% of the
   open, over bytes `PackageEditor` already captured into
   `Inner::workbook_stream`. **Frozen**: handing the eager reader bytes the edit
   owner extracted would make the validation owner stop independently reading
   the stream it validates, which is a reading of ADR 0003's "the complete
   package is reopened under the retained limits" that this change is not
   entitled to make. Revisiting it needs an ADR amendment saying what
   "independent" means for the stream bytes, plus a proof that
   `PackageEditor`'s capture and `OleFile::open_stream` agree on every fixture
   including the malformed ones.
3. **The eager `Workbook` is built and discarded.** `add_cell` is 33.82% of the
   open and its drop glue another 4.68%, to answer one bit per sheet
   (`parsed_worksheet_index().is_some()`) and hand over the SST. Replacing that
   owner is change 0620's frozen design with its five admission gates; nothing
   here weakens or strengthens them.

**The per-`LabelSst` text copy is the fourth, and it is not fusable either.**
`resolve_shared_strings` is 3,629,090 Ir (2.26%) and allocates one `String` per
`LabelSst` cell; the eager parse independently clones the same text into its own
cells (`String::clone` under `Cell::from_record_with_formula_context`). Sharing
would mean `Value::Text` holding a shared handle, and `Value` is public. Making
the resolution lazy would move the `LabelSst index N is outside the SST` refusal
— matrix row `labelsst-past-sst`, the one row that is refused after *both*
passes.

## Limitations

- **No claim is registered.** `performance_claim: none`.
- The timing is warm, in-memory, single-CPU, one process per round, on a host
  shared with seven other agents. No cold-cache, RSS, physical-I/O, throughput
  or producer claim is made. The allocation and peak figures are the probe's own
  counting allocator over one measured iteration, not process RSS.
- The host was never quiet: eight agents shared 32 cores throughout. Round 1's
  A/A floor on `54016.xls` is +12.3% at p50 and its `WithCustomViews.xls` tail is
  unusable; round 3, whose floors on the two changed scenarios are −0.15% and
  +1.92%, is the one quoted. All three rounds are retained and all six p50
  readings of the changed scenario lie between −34.9% and −44.1%. The unchanged
  `54016.xls` controls move by up to ±15% at p50 across the three rounds with
  instruction counts flat to within 0.5%, so no `54016.xls` *control* timing in
  this packet should be read as a result in either direction.
- The commit change affects **only** `Transaction::commit_source_backed` on XLS
  `cell_values` — fixed-width `Number`/`RK`/`MulRk` edits that retain a complete
  target artifact and a reversible `Patch`. The property-table change affects
  every `Snapshot::from_bytes`, by 2 to 58 allocation calls and 0.04% to 0.24%
  of allocated bytes on the four measured fixtures; it is a count, not a
  latency result.
- The `noop-generic` path never reaches the changed code (`plan.is_noop()`
  returns the source snapshot) and is reported only as a control.
- The framing attribution is scoped to two fixtures on one host and one build.
  It says what one open is made of on those files; a workbook with a different
  record mix will split differently — change 0620 measured `Formula` decoding at
  20.5% of `FormulaEvalTestData.xls`'s open and 0% of `54016.xls`'s.
- 33 of the 126 `.xls` fixtures never reach a snapshot at all, and of the 93
  that do, 35 publish through both source-backed numeric paths. That population,
  which 0620 censused, is what every number here is scoped to.

## Retained evidence

`docs/performance/results/change-0633/README.md` lists every file: the three
scratch binaries' sources, the six capture scripts, the three analyzers, the raw
callgrind annotations and `perf stat` CSVs for both legs, the per-operation
counter lines, the six corpus runs and the two 16-run repetitions, all three
A1 B1 B2 A2 latency rounds, the registered selector runs and the isolation pair
that chased their one review trigger, the first-error matrix for both legs,
`analysis.txt`, `framing-attribution.txt`, `owned-source-chase.txt`,
`decision.json`, `gates.txt`, `environment.txt`, `binaries.sha256` and
`log-sections.md`.
