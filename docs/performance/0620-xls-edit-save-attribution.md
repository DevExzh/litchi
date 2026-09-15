# 0620: an XLS source-backed numeric save parses the whole workbook three more times, and one of the three was already parsed at open

Status: retained. `performance_claim: none` — this record carries an attribution
of the XLS edit-and-save path (callgrind isolation pairs split by call site,
native cycle counts, allocation and peak-retained-byte counts, a 126-fixture
corpus census), one implemented value-identical reuse measured in both
directions beside its floor, and one frozen design that is **not** implemented.
No claim is registered.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is item **XLS-9** (rank 31) of change
[0587](0587-remaining-opportunity-survey.md), "Edit and save: the candidate
readback is a complete eager open". 0587 sized it **unknown** — "no attribution
of the XLS edit path exists" — and gave it a falsification test. This record
supplies the attribution, resolves the falsification test, implements the one
reuse that is value-identical, and freezes the design for the part that is not.

## Why this record exists

Every XLS `cell_values` publication path ends in an independent complete reopen,
and change [0016](changes/0016-xls-commit-editor-reuse.md) closed with a standing
instruction:

> the final full reader and typed readback remain the publication boundary. The
> next XLS optimization should target a different source of whole-workbook work,
> not remove either retained validation layer.

Changes 0136-0138, 0168 and 0172 then removed target-artifact retention and
fingerprint scans from the plan-only path, and 0138 reported p50s — Number
105.3 ms plan-only against 145.4 ms source-backed — but no record has ever said
*where* the time goes. 0587's survey named the mechanism from the source and
stopped, because the harness's XLS attribution binary supports only `open`,
`list` and `one-cell`.

So the first job here is to count, not to change anything.

## Method

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
cargo 1.95.0, valgrind 3.26.0. Base commit `1e4198321`, branch
`perf/0620-xls-edit-save-attribution`. Every measured child is pinned to CPU 14
with `taskset` under `setarch -R`, `RAYON_NUM_THREADS=1`.

Three measured fixtures, chosen because the `cell_values` editor admits them
(see the census below): `54016.xls` (984,576 B, one worksheet, 16,055 `LabelSst`
+ 2,122 `RK` + 11,431 other cell records — the program's primary XLS fixture),
`WithCustomViews.xls` (165,888 B, three worksheets) and
`FormulaEvalTestData.xls` (178,176 B, four worksheets, 1,416 `Formula` records;
macro-bearing, so the two source-backed paths refuse it and only the generic
path runs).

Six operations per fixture, through a scratch probe retained in the packet
(`results/change-0620/probe/main.rs`): `open` (`Snapshot::from_bytes` alone),
`number-plan` (`commit_source_backed_plan`), `number-source-backed`
(`commit_source_backed`), `number-generic` and `string-generic`
(`Transaction::commit`, the latter re-pointing a `LabelSst` cell at an SST entry
that already exists so no resource change is staged), and `noop-generic` (the
exact-no-op fast path). The probe times open, staging, commit and publication
separately, and arms a counting global allocator for exactly one measured
iteration so its allocation, allocated-byte and peak-live-byte figures are per
operation.

Instruction attribution uses change 0574's isolation method with one addition:
`callgrind --separate-callers=2`. That is what makes it an attribution rather
than a total — a source-backed commit calls `Workbook::new` three times, and
without caller separation `callgrind_annotate` sums all three under one symbol.
Callgrind runs SHA-256 in software because valgrind masks the SHA CPUID bit, so
every `sha2` row below is an upper bound and is paired with a native `perf stat`
count taken by the same isolation method.

## What one XLS open-edit-save is made of

All figures per operation, before leg, callgrind inclusive Ir
(`results/change-0620/analysis.txt` has the full tables and the other two
fixtures).

**One `Snapshot::from_bytes` on `54016.xls` is 167,311,992 Ir**, and it is
almost entirely one thing:

| phase | Ir | share of the open |
| --- | ---: | ---: |
| `PackageEditor::open` (CFB directory, stream capture) | ~4.2 M | 2.5% |
| `parse_workbook_stream` inventory + SST property clones + `resolve_shared_strings` | ~20.8 M | 12.5% |
| **`Workbook::new` — the complete eager parse** | **137,421,818** | **82.1%** |
| ↳ of which `parse_worksheet_records_with_compatibility` | 120,526,281 | 72.0% |

The offset inventory the edit owner actually keeps costs an eighth of what the
independent validation owner costs.

**One whole `number-source-backed` operation is 1,053,779,059 Ir**, of which
`commit_source_backed` is 775,712,141 (73.6%). Inside that commit:

| phase inside `commit_source_backed` | Ir | of commit | of operation |
| --- | ---: | ---: | ---: |
| `Workbook::new` for the three source policy checks (`:5260`) | **137,061,668** | **17.7%** | 13.0% |
| `Snapshot::from_bytes` reopen of the materialized target | 162,364,929 | 20.9% | 15.4% |
| ↳ its own `Workbook::new` | ~137.2 M | 17.7% | 13.0% |
| `verify_source_backed_numeric_target` | 152,487,225 | 19.7% | 14.5% |
| ↳ its own `Workbook::new` (`:5506`) | 137,426,515 | 17.7% | 13.0% |
| planning, splice, `materialize_numeric_plan`, fingerprints | ~324 M | 41.7% | 30.7% |

So **one open-edit-save through `commit_source_backed` parses the complete
workbook four times**: once at open, and three more times inside the commit.
Together those four parses are **548,831,368 Ir — 52.1% of the whole
operation**; the three inside the commit are about 411.4 M, **53.0% of the
commit**. (The chain `Workbook::new'from_package_editor'from_bytes` covers two
of the four — the outer open and the target reopen — at 274,343,185 Ir for the
pair, because two caller levels do not separate them; the `open` scenario prices
a single such call at 137,421,818.)

The other publication paths pay the same complete parse once per commit:

| scenario (`54016.xls`) | operation Ir | commit Ir | commit's `Workbook::new` | of commit |
| --- | ---: | ---: | ---: | ---: |
| `number-plan` | 525,020,551 | 250,629,394 | 137,118,729 | **54.7%** |
| `number-generic` | 345,125,462 | 172,327,731 | 137,201,682 | **79.6%** |
| `string-generic` | 358,341,422 | 180,305,483 | 137,260,317 | **76.1%** |
| `number-source-backed` | 1,053,779,059 | 775,712,141 | ~411,409,550 (3×) | **53.0%** |
| `noop-generic` | 170,855,242 | ~1 k | 0 | — |

0587's falsification test for XLS-9 was "falsified if a callgrind of
`xls_numeric_plan_only_*` shows `Workbook::new` under 15% of the commit". It is
54.7%. The item is not falsified.

Two further terms are worth naming because they bound what any future change can
win. `cfb::overlay::fingerprints` plus the publication-time hash are 19.5% and
19.7% of the plan-only operation *in callgrind*; natively the whole plan-only
operation is 101,580,947 cycles against 310,143,787 instructions, so the
hardware-SHA cost is a small fraction of that — the callgrind share is roughly
five times the native one, exactly the caveat `GOAL_AUDIT.md` carries. And the
`noop-generic` commit is about 1 µs on `54016.xls` against a 10.2 ms open: the
exact-no-op fast path is already free, and nothing here is about it.

### What the edit owner admits

The probe ran the same five publications over every `.xls` in the repository
(126 files, 591 file/operation rows,
`results/change-0620/corpus/corpus-*.jsonl`; census in
`results/change-0620/admission-census.txt`). **33 of the 126 fixtures never
reach a snapshot at all** — `Snapshot::from_bytes` refuses them, and 24 of the
33 with the same message: `worksheet at tab position N was not published by the
complete XLS reader` (the rest are three password-protected files, two corrupt
FATs, and four malformed globals). Among them is
`ConditionalFormattingSamples.xls`, the flagship every other XLS record in this
program measures: the *edit* owner cannot open it, so no XLS edit measurement in
this program has ever been taken on it or can be.

Of the 93 fixtures that do open, 35 publish through both source-backed numeric
paths. The 58 that do not divide into 53 where the probe found no
`Number`/`RK`/`MulRk` cell to edit (a limitation of the probe's target
discovery, not a refusal) and **5 genuine editor refusals: four
`macro-bearing XLS sources are not eligible for source-backed numeric edits` and
one `protected worksheets are not eligible`**. Of the seven fixtures over 100 KB
that open, three admit a source-backed numeric edit (`54016.xls`,
`WithCustomViews.xls`, `45365-2.xls`) and four do not. That is the population any
XLS source-backed edit claim is scoped to.

## What was changed

Scope: `crates/litchi-xls` only; one function body and three tests.

`commit_source_backed_numeric` (`cell_values/mod.rs:5258-5263`) opened a
complete `Workbook` over the source snapshot's own bytes and ran three
requirement checks over it:

```rust
let source_workbook = Workbook::new(Cursor::new(source.bytes()))?;
require_public_worksheet_coverage(&source_workbook, &source.inner.sheets)?;
require_unprotected_workbook(&source_workbook)?;
require_macro_free_workbook(&source_workbook)?;
```

It now reads the verdicts the snapshot already recorded:

```rust
source.inner.source_policy.require()?;
```

This is not a new mechanism. `SourcePolicyFacts` was introduced by commit
`237309eea` precisely for this, and its doc comment on the plan-only path says
so: *"The complete source Workbook validation already ran when this immutable
snapshot was opened. Reuse its private policy facts here; the target is still
checked independently below over the composed positional reader."* That commit
converted `commit_source_backed_numeric_plan` and left
`commit_source_backed_numeric` — the older, artifact-retaining sibling —
untouched. This change finishes it.

## Why it is sound

**The reused facts come from the same bytes under the same limits.**
`SourcePolicyFacts::from_workbook` is called at `Snapshot::from_package_editor`
(`:542`) and `Snapshot::from_fixed_numeric_package_editor` (`:592`) on
`Workbook::new(Cursor::new(source.as_slice()))`, where `source` is the `Vec<u8>`
that is then moved into `Inner::bytes` as an immutable `Arc<[u8]>`.
`commit_source_backed_numeric` parsed `Workbook::new(Cursor::new(source.bytes()))`
— the same `Arc<[u8]>`, the same `Cursor`, the same default `Limits`, the same
`CompatibilityProfile::Strict`. `Workbook::new` is a pure function of those
inputs, so the second parse could only ever reproduce the first.

**Every `Snapshot` carries the facts.** `Inner` is constructed at exactly three
sites: `from_package_editor` (`:562`), `from_fixed_numeric_package_editor`
(`:596`) and `retag_source_version` (`:619`, which copies them). The first two
compute `source_policy` from the same `Workbook` they validate against; the
third is a pure retag. There is no path to a `Snapshot` whose facts were taken
from different bytes.

**Refusal identity is exact for the two reachable verdicts.**
`SourcePolicyFacts::require` returns *character-identical* strings to
`require_unprotected_workbook` ("protected or shared workbooks are not eligible
for source-backed numeric edits", "protected worksheets are not eligible for
source-backed numeric edits") and to `require_macro_free_workbook`
("macro-bearing XLS sources are not eligible for source-backed numeric edits").

**The coverage verdict cannot differ, because it is not reachable.**
`public_worksheet_coverage` is only ever set to `true`:
`SourcePolicyFacts::from_workbook` propagates
`require_public_worksheet_coverage`'s error with `?` before constructing the
struct, so a snapshot whose coverage failed does not exist — that refusal has
already been paid at `Snapshot::from_bytes`, with its own per-tab message. The
`require()` branch for it is a defensive dead branch, exactly as it already is
on the plan-only path.

**Order and reachability are preserved.** The removed `Workbook::new` sat after
the structural/resource guard, the `changes_are_fixed_numeric` guard and
`SemanticPatch::from_transaction`, and before the publisher opens. The reuse
sits in the same slot. The only behaviour that moves is that a *protected* or
*macro-bearing* source now refuses without first paying a complete parse of
itself — the same error, sooner.

**Nothing the readback owns is removed.** The target is still materialized,
reopened through `Snapshot::from_bytes` (a complete independent `Workbook::new`
plus the offset inventory) and validated again by
`verify_source_backed_numeric_target`, which runs its own complete
`Workbook::new` over the target and all four checks including
`verify_public_numeric_readback`. ADR 0003's "publish only after their staged
CRUD operation and typed readback succeed" and ADR 0006's validation contract are
untouched; this removes a redundant validation of the *source*, not the
readback. 0016's instruction — "target a different source of whole-workbook
work, not remove either retained validation layer" — is the instruction this
change follows.

## Measured

### Counts (deterministic, before → after)

Per operation, counting allocator armed for one measured iteration
(`results/change-0620/counters-*.jsonl`):

| fixture | metric | before | after | delta |
| --- | --- | ---: | ---: | ---: |
| `54016.xls` | allocation calls | 201,886 | 161,751 | **−40,135 (−19.88%)** |
| | allocated bytes | 93,192,484 | 76,204,234 | **−16,988,250 (−18.23%)** |
| | peak live bytes | 39,572,027 | 26,082,823 | **−13,489,204 (−34.09%)** |
| `WithCustomViews.xls` | allocation calls | 26,643 | 20,945 | −5,698 (−21.39%) |
| | allocated bytes | 12,109,985 | 9,911,311 | −2,198,674 (−18.16%) |
| | peak live bytes | 4,584,255 | 3,176,770 | −1,407,485 (−30.70%) |

The peak figure is the interesting one and it was not predicted. The removed
`Workbook` was a local binding with drop glue, so it stayed alive across the
splice planning, the target materialization, the target `Snapshot::from_bytes`
and `verify_source_backed_numeric_target` — a complete parsed model of the
source held alongside a complete parsed model of the target. Peak retained bytes
for one `commit_source_backed` fall by a third.

Every other counter is identical on both legs, on all 24 (fixture, operation)
rows the counter capture produced — four fixtures by six operations, the
flagship contributing none because it is refused at open: published bytes,
splice count, replacement bytes, changed spans, source and target Workbook
lengths, the four typed refusals, and every counter of `open`, `number-plan`,
`number-generic`, `string-generic` and `noop-generic`.

### Instructions and cycles (isolation pairs, before → after)

| scenario | callgrind Ir/op | native cycles/op | native instructions/op |
| --- | --- | --- | --- |
| `54016/number-source-backed` | 1,053,779,059 → 906,199,478 (**−14.00%**) | 228,627,848 → 171,248,047 (**−25.10%**) | 625,940,966 → 484,563,358 (−22.59%) |
| `cv/number-source-backed` | 128,816,213 → 115,336,872 (**−10.46%**) | 20,812,677 → 16,521,723 (**−20.62%**) | 57,733,826 → 44,775,228 (−22.45%) |
| `54016/open` (control) | 167,311,992 → 166,819,776 (−0.29%) | 48,376,435 → 48,467,675 (+0.19%) | −0.33% |
| `54016/number-plan` (control) | 525,020,551 → 524,570,766 (−0.09%) | +1.37% | −0.15% |
| `54016/number-generic` (control) | 345,125,462 → 344,657,596 (−0.14%) | +2.41% | −0.15% |
| `54016/string-generic` (control) | 358,341,422 → 357,345,827 (−0.28%) | +3.73% | −0.31% |
| `54016/noop-generic` (control) | 170,855,242 → 170,351,988 (−0.29%) | +0.45% | −0.32% |
| `formula/*` (controls, 4 operations) | −0.18% to −0.06% | +0.88% to +2.34% | −0.18% to +0.91% |

The callgrind line that carries the change is exact: the chain
`Workbook::new'commit_source_backed'main` is 137,061,668 Ir before and **0
after**. Nothing else moves by more than 0.5% in instructions. Native cycle
counts on the unchanged controls move by up to +3.7% because each `perf stat`
isolation pair is a single run; the callgrind column is the deterministic one.

The native cycle saving (−25.1%) exceeds the callgrind instruction saving
(−14.0%) because callgrind's software SHA-256 inflates the fingerprint terms
that the change does *not* touch, deflating the share of the term it does.

### Paired timing (A1 B1 B2 A2, 40 samples on `54016.xls`, 120 elsewhere)

`commit` phase p50, nanoseconds
(`results/change-0620/latency/latency-summary.json`):

| scenario | a1 p50 | b1 p50 | forward | reverse | p95 fwd/rev | p99 fwd/rev | A/A floor | B/B floor |
| --- | ---: | ---: | ---: | ---: | --- | --- | ---: | ---: |
| `54016/number-source-backed` | 35,940,283 | 24,400,164 | **−32.11%** | **−33.66%** | −31.85% / −31.28% | −31.82% / −31.49% | +1.70% | −0.62% |
| `cv/number-source-backed` | 3,336,644 | 2,443,878 | **−26.76%** | **−26.19%** | −26.13% / −26.40% | −25.81% / −25.78% | −0.20% | +0.58% |

Controls in the same window, `commit` p50: `54016/number-plan` +4.05%/+3.65%,
`54016/number-generic` +2.23%/+0.59%, `54016/string-generic` +1.73%/+1.95%,
`cv/number-plan` +0.93%/+0.88%, `cv/number-generic` +1.21%/+1.33%,
`cv/string-generic` +5.27%/+2.11%, `formula/number-generic` +4.12%/+2.20%,
`formula/string-generic` +1.50%/+0.36%. The `open` phase, which is the same code
on both legs and runs inside every one of those processes, moves +0.20% to
+4.65% forward and −3.47% to +4.88% reverse.

The measured same-binary A/A floor in this window is small (±2% at p50 on every
scenario above), but the unchanged controls move by up to +5.3%, so the honest
control band for a two-binary comparison here is about ±5% at p50 — code layout
differs between the two binaries and the `open` phase shows it. The changed
scenario is −26% to −34% in both directions at p50, p95 and p99, five to six
times outside that band.

The sub-microsecond `noop-generic` commit (210-615 ns) is reported in the packet
but is not evidence in either direction: its floors are as large as its deltas.

## Correctness evidence

**Corpus differential.** A second scratch binary
(`results/change-0620/corpus/`, source in `probe/corpus-main.rs`) drives all five
publication paths over every `.xls` in the repository and prints, per file and
operation, either the exact refusal text or the SHA-256 of the complete
published artifact with its diagnostics and patch lengths. 126 fixtures — 93 of
which the editor opens, 33 of which it refuses at `Snapshot::from_bytes` — 591
rows per run (297 typed refusals, 294 publishing), three runs per leg, **3,546
rows compared**. Three oracles, `results/change-0620/corpus_oracle.py`:

| oracle | rows | before-vs-after mismatches |
| --- | ---: | ---: |
| exact typed refusal text | 591 | **0** |
| every reported field except the artifact digest — published length, changed cells, touched streams, splice count, replacement bytes, changed spans, target Workbook length, `is_noop`, patch before/after lengths, patch emptiness, inverse digest | 591 | **0** |
| the artifact digest | 587 | **0** |

Four rows are excluded from the digest oracle because they are not reproducible
*within a single binary* — see the finding below. All four are on the generic
`Transaction::commit` path; every `number-plan` and `number-source-backed` row
agrees across all six runs.

**Unit tests.** Three tests added to `cell_values/tests.rs`, mirroring the ones
`237309eea` added for the plan path:
`source_backed_numeric_publication_consumes_cached_source_policy_facts` (flips
each of the three facts on an otherwise valid snapshot and requires the
publication to refuse — this is what pins the reuse to the facts rather than to
a re-parse), `source_backed_numeric_publication_preserves_protection_refusal_reasons`
(exact refusal strings for a protected workbook and a protected worksheet), and
`source_backed_numeric_publication_and_plan_refuse_sources_identically` (a
differential over `protected_package(false)`, `protected_package(true)`,
`macro_package()` and `empty_macro_storage_package()`, requiring the two
publication entry points to produce the same typed refusal text).

**Registered harness selectors.** `xls_numeric_source_backed_number_edit_save`
and `xls_numeric_source_backed_rk_mulrk_edit_save` are the two registered cases
that exercise this function; `xls_numeric_plan_only_*`, `xls_numeric_eager_*`,
`xls_semantic_*_edit_save` and `xls_owned_source_*` are the controls. All twelve
cases ran A1 B1 B2 A2, 10 warmups / 100 samples (20 / 200 for the non-numeric
groups), and every `output_sha256` is identical across all four rounds — Number
`f8f37064…` and RK/MulRK `ddf5d5b8…`, the same two digests change 0138 recorded.

### The registered selectors cannot see this change, and the reason is the corpus

| case, `commit` phase p50 (ns) | a1 | b1 | forward | reverse | A/A floor | B/B floor |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `xls_numeric_source_backed_number_edit_save` | 62,795,500 | 60,033,820 | −4.40% | −0.88% | −1.65% | +1.97% |
| `xls_numeric_source_backed_rk_mulrk_edit_save` | 661,879 | 644,248 | −2.66% | −2.74% | +0.08% | −0.00% |
| `xls_numeric_plan_only_number_edit_save` (control) | 16,813,416 | 16,784,908 | −0.17% | −0.33% | −0.41% | −0.57% |
| `xls_numeric_eager_number_edit_save` (control) | 30,720,682 | 30,532,632 | −0.61% | +5.93% | −6.01% | +0.17% |

The Number case disagrees with itself between directions and its eager control
drifted 6%, so it resolves nothing. The RK/MulRK case is consistent at −2.7% in
both directions against a floor under 0.1%, which is real but an order of
magnitude below what the same change does on a real fixture.

The reason is visible in the corpus evidence the selectors themselves report:
the Number corpus is a **16,995,840-byte CFB whose Workbook stream is 80,946
bytes — 0.48% of the archive**, and the RK/MulRK corpus is 1,665 of 202,752
bytes (0.82%). The removed work is proportional to the *Workbook stream*, while
the fingerprints and the artifact copy that dominate those cases are
proportional to the *archive*. On real files the ratio is inverted: the Workbook
stream is 94.2% of `54016.xls`, 90.7% of `WithCustomViews.xls`, 88.9% of
`FormulaEvalTestData.xls` and 78.9% of `59858.xls`.

This is change 0587's first finding — "the measured path is not the path real
files take" — reproduced on a fourth area, and it is why the paired timing above
is taken on real fixtures through a retained scratch probe rather than on the
registered corpora. It is also a concrete brief for the successor of change
[0601](0601-perf-harness-real-producer-shape.md): the XLS numeric corpora need a
Workbook stream proportioned like a real workbook before they can price anything
in the BIFF record path.

**Gates.** `cargo fmt --all --check`; `cargo clippy -p litchi-xls --all-targets`;
`cargo test -p litchi-xls`; `cargo doc -p litchi-xls --no-deps`. Tails in
`results/change-0620/gates.txt`.

## Validation preserved

The source's complete `Workbook::new` still runs — at `Snapshot::from_bytes`,
where it always ran, and where its three verdicts are recorded. The candidate's
complete reopen still runs twice, once inside `Snapshot::from_bytes` for the
target and once inside `verify_source_backed_numeric_target`, with
`require_public_worksheet_coverage`, `require_unprotected_workbook`,
`require_macro_free_workbook`, `carry_fixed_numeric_inventory` and
`verify_public_numeric_readback` unchanged. `require_macro_free_container`'s
independent CFB-directory check is unchanged. Source lineage
(`source_version != source.inner.source_version`), the overlay fingerprints, the
same-length splice validation, the artifact-length check in
`materialize_numeric_plan`, the exact `Patch` and its inverse, and the exact
no-op fast path are all untouched. No limit moved, no `unsafe` was added, no
public type changed.

## The frozen design: replacing the readback's owner is not admissible as proposed

0587's XLS-9 proposed that the candidate readback's `Workbook::new` be replaced
by "a source-backed candidate readback (globals plus a scan of the edited
sheets) … if `require_public_worksheet_coverage`, `require_unprotected_workbook`,
`require_macro_free_workbook` are proven equivalent on the source-backed owner".
This section is the frozen answer. **It is not implemented, and on the present
evidence it should not be.**

**The candidate can be presented to the lazy owner.** `ComposedOverlaySource`
implements `ReadAt` (`litchi-cfb/src/overlay.rs:599-627`) and is `Clone`, and
`SourceBackedWorkbook::from_read_at(Arc<dyn ReadAt>)`
(`workbook/source.rs:754`) takes exactly that. `ComposedPositionalReader`
(`cell_values/mod.rs:5075`) exists only because `Workbook<R>` needs
`Read + Seek`; the lazy owner would not need it. `ComposedOverlaySource::version`
is stable, so the `ensure_current` fence would not trip perpetually. That part
is free.

**Three of the four checks are answerable; the fourth is not.**

| check | status on `SourceBackedWorkbook` |
| --- | --- |
| `verify_public_numeric_readback` | **equivalent today.** `SourceBackedWorksheet::cell` builds the value through the same `Cell::from_record_with_formula_context` and the same `Formatting::parse_globals` the eager path uses, with `validate_cell_xf` on every record and the same last-record-wins rule for a duplicated position. |
| `require_macro_free_workbook` | **mechanically addable.** Of its three legs, the BoundSheet8 `dt` kinds are already parsed and retained in `SheetEntry::kind` (`source.rs:384`, `2086-2091`) and need only a `pub(crate)` accessor; the CFB `_VBA_PROJECT_CUR` leg is already checked independently by `require_macro_free_container`; only `ObProj` (0x00D3) would need one more arm in the globals semantic loop, over a record vector that is already materialized. |
| `require_unprotected_workbook` | **addable at the globals level, expensive at the worksheet level.** The seven workbook-scope records (`PROTECT` 0x0012, `WINDOWPROTECT` 0x0019, `PASSWORD` 0x0013, `PROT4REV` 0x01AF, `PROT4REVPASS` 0x01BC, `WRITEPROT` 0x0086, `FILESHARING` 0x005B) are all in globals, which `parse_globals` already frames — seven more match arms, no extra I/O. The worksheet-scope records (`PROTECT`, `OBJPROTECT` 0x0063, `SCENPROTECT` 0x00DD, `PASSWORD`) are inside each worksheet substream, which the lazy owner reads **no bytes of** at open, so answering this leg means a full validated scan of *every* worksheet — which is the cost the proposal was trying to avoid. |
| `require_public_worksheet_coverage` | **not equivalent, and not made equivalent by parsing more records.** |

**Why coverage is structurally unanswerable.** The predicate is
`workbook.sheet(i).parsed_worksheet_index().is_some()`, and that field is set at
exactly one place — `workbook/package.rs:210` — only when
`parse_worksheet_from_position` returned `Ok`. The loop's error arm
(`package.rs:217`, `Err(_) => {}`) silently drops any worksheet the eager parser
could not project, for any of roughly forty distinct reasons: an out-of-range or
reserved-slot XF index anywhere in the sheet, a string-valued `Formula` without
its `String` record, a `ShrFmla` without its anchor, an `Array` anchor not at its
`RefU` top-left or overlapping another, a `DVAL` promising more `DV` records than
arrive, an orphan `PtgExp`, array formulas without `Dimensions`, and nine
per-sheet collector `finish()` failures including the protection collector's.
The fact the check asserts is therefore a property *of the eager parser's
success*, not a property of the bytes.

The lazy owner's corresponding field is assigned unconditionally:

```rust
// crates/litchi-xls/src/workbook/source.rs:2093
let current_worksheet_index = (kind == SheetKind::WorksheetOrDialog).then(|| { … });
```

from the BoundSheet8 `dt` byte alone, with no worksheet substream read. So
`worksheet_index.is_some()` is *always* true for every worksheet tab — it is
exactly the predicate the eager check exists to falsify. A source-backed
readback would silently accept a candidate whose splice damaged an unedited
sheet's cross-record grammar, which is the case the comment at
`cell_values/mod.rs:535-539` says the check exists for.

The two owners also disagree in the other direction, which matters for any
future attempt: the eager parser slices each worksheet to the *end of the whole
Workbook stream* with no per-sheet upper bound and ignores a nested BOF
(`worksheet.rs:453`), so a sheet missing its EOF runs into the next substream and
still returns `Ok`; it never validates the BOF substream type (only
`kind != 0x0809`); and it accepts an encryption password. The lazy owner bounds
each sheet at `[start, end)`, calls `validate_worksheet_bof`, refuses a second
SST, a non-empty globals EOF payload, duplicate or misordered BoundSheet8
offsets, and refuses `FilePass` unconditionally. Their worksheet *index*
numbering also diverges the moment any sheet fails eagerly, because the eager
counter increments only on success; only the BoundSheet8 tab position is a
stable join key.

**Admission gates.** This design may be revisited only when all five hold:

1. A proposed ADR (or an amendment to ADR 0003/0006) states that the candidate
   readback's owner may be the source-backed reader, and says what "typed
   readback" then guarantees. ADR 0003 currently says "the complete package is
   reopened under the retained limits".
2. `require_public_worksheet_coverage`'s meaning is redefined as a property of
   the bytes rather than of the eager parser — for example as an explicit list
   of per-sheet grammar invariants the lazy owner also checks — and a
   differential over the whole `.xls` corpus shows the two owners produce the
   same accept/reject set on every candidate, including deliberately damaged
   ones.
3. The worksheet-scope protection records are collected by the lazy owner and
   shown equivalent on the corpus, including the eager parser's coupling whereby
   a self-inconsistent protection collector drops the whole sheet.
4. A measurement shows the replacement actually wins: the lazy owner would have
   to scan every worksheet (gates 2 and 3), and `0605` measured its whole-sheet
   walk at 9,146 Ir per cell against the eager parser's per-cell cost, so the
   win is not obvious and must be demonstrated before the contract is moved.
5. The two `Workbook::new` calls the commit still makes over the *target* are
   shown not to be reducible to one first — see the next section.

## What was left on the table

**The target is parsed twice inside `commit_source_backed`, not once.**
`Snapshot::from_bytes(target_bytes)` runs a complete `Workbook::new` and records
its `SourcePolicyFacts`; `verify_source_backed_numeric_target` then runs a
second complete `Workbook::new` over the identical bytes (identity is proved one
line earlier by `target.bytes() != target_bytes.as_slice()`). The three
requirement checks in the second parse are *nearly* the reuse this change made,
but not quite: `require_public_worksheet_coverage(&workbook, &source.inner.sheets)`
checks the **source's** sheet list against the target's workbook, whereas
`target.inner.source_policy` was computed against the **target's** sheet list.
They are equal only once `carry_fixed_numeric_inventory` has proved every other
Workbook-stream byte unchanged — and that runs *after*. Reordering it would move
when a refusal happens. The remaining `verify_public_numeric_readback` needs the
parsed `Workbook` itself, so the second parse cannot be removed outright without
the snapshot retaining its reader, which ADR 0005's cache contract makes a
subsystem rather than a small change (0605 reached the same conclusion for the
read path). Left open, unmeasured beyond the 137,426,515 Ir it costs.

**The generic `Transaction::commit` copies the source bytes to re-open a second
`PackageEditor`** (`:2643`) after the snapshot already consumed one. Not
value-identical to remove: the first editor is consumed by `finish()` at open.

**`Snapshot::from_bytes` frames the Workbook stream twice** — once for the
offset inventory (`parse_workbook_stream`) and once inside `Workbook::new`. At
12.5% and 82.1% of the open respectively, fusing them is the largest remaining
XLS edit-path item, and it is a different change with its own design questions.

## Finding: the generic XLS commit's published bytes are not reproducible across processes

Reproduced on the untouched before checkout, so **pre-existing and not caused by
this change**. Three consecutive runs of the same before-leg binary over
`59858.xls` through `Transaction::commit` produce three different artifact
digests (`0995dac5…`, `f790bf18…`, `4cac9882…`); sixteen runs produce sixteen.
Across the 126-fixture corpus exactly 4 of 591 rows are affected, all on the
generic path, on two fixtures: `59858.xls` and
`libreoffice-core/sc/qa/unit/data/xls/pivottable_dates_grouping.xls`. On the
latter the output is drawn from a **3-element** digest set per operation, and the
before and after legs draw from the *same* set — which is what a permutation of
a small number of storages looks like. The differing bytes are 250 of 1,103,360
on `59858.xls`, all inside a run of UTF-16LE directory entry names: the CFB
directory entry *order* changes.

The site is `crates/litchi-cfb/src/writer/core.rs:977`:

```rust
for storage_path in &self.storages {
    directory.add_storage_path(storage_path)?;
}
```

`storages` is a `HashSet<Vec<String>>` (`core.rs:263`), so its iteration order is
randomized per process by `RandomState`, and storage directory entries are
assigned SIDs in that order. The adjacent field carries the comment "Using Vec
instead of `HashMap` to preserve insertion order for directory entries" — the
streams table was fixed and the storages table was not.

ADR 0006 says "Serialization is deterministic unless a `Clock`, actor identity,
or cryptographic RNG is explicitly supplied", so this reads as a contract
deviation rather than a licence. It is out of this change's scope (litchi-cfb,
and a path this change does not touch) and is reported, not fixed. The
source-backed and plan-only paths are unaffected because they splice source
bytes in place rather than re-rendering the CFB directory — which is why all 278
publishing rows on those paths were byte-identical across six runs.

## Limitations

- **No claim is registered.** `performance_claim: none`.
- The timing is warm, in-memory, single-CPU, one process per round. No
  cold-cache, RSS, physical-I/O, throughput or producer claim is made. The
  allocation and peak figures are the probe's own counting allocator over one
  measured iteration, not process RSS; change 0172 showed RSS can move the other
  way on this path.
- The change affects **only** `Transaction::commit_source_backed` on XLS
  `cell_values` — fixed-width `Number`/`RK`/`MulRk` edits that retain a complete
  target artifact and a reversible `Patch`. `commit_source_backed_plan` already
  had the reuse; `Transaction::commit`, the eager path, `sheet_visibility`,
  `comments`, and every other crate are untouched and were measured only as
  controls.
- The attribution is scoped to three fixtures on one host and one build. It says
  what the path is made of on those files; it does not predict the shares on a
  workbook with a different record mix. `FormulaEvalTestData.xls` already shows
  the spread: `Formula` decoding is 20.5% of its open and 0% of `54016.xls`'s.
- The `noop-generic` timings are at timer granularity and are reported only for
  completeness.
- The frozen design section establishes that the proposed readback replacement
  is not equivalent as specified. It does not prove no source-backed readback
  can ever work; it states the five gates one would have to pass.
- The CFB storage-order finding is reported from two fixtures and one site read
  from the source. No fix is proposed and none was measured.

## Retained evidence

`docs/performance/results/change-0620/README.md` lists every file: the probe and
corpus-differential sources, the four capture scripts, the analyzer, the raw
callgrind annotations and `perf stat` CSVs for both legs, the per-operation
counter lines, the six corpus runs, the A1 B1 B2 A2 latency JSON, the registered
selector runs, `decision.json`, `gates.txt`, `environment.txt` and
`log-sections.md`.
