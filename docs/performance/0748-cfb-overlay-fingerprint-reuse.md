# 0748 — Sealed owned CFB overlay plans hash their artifact once: every XLS edit that publishes through the overlay stops re-proving digests of bytes that cannot change

Status: retained, implemented. `performance_claim: none` — this record carries
paired ABBA timings, isolation-pair instruction and cycle counts, deterministic
allocation counts and a two-leg correctness census. They are reported as
evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `ab29ac6291` (the `feat/office-format-completeness` tip: it includes 0745,
not 0746); branch `perf/0748-cfb-overlay-fingerprint-reuse`; commits
`3cdcead75c` (the CFB contract), `604132d280` (sealed ingress for two immutable
sources), `d663c504f4` (harness evidence schema v2) and `0a3476edd2` (one more
test). Evidence: [`results/change-0748/`](results/change-0748/README.md).

## Result

A same-length CFB overlay plan records the SHA-256 of its source artifact and
of the composed target. Until now every later step re-hashed the whole artifact
to prove those digests still held, even when the plan's source was an owned,
immutable in-memory allocation. A plan over **sealed owned bytes** now computes
both digests once, in its planning pass, and its checked composed view, direct
write and atomic save no longer re-hash anything. Generic positional sources
(a caller's `ReadAt`, a file, a composed view handed back in) keep every pass.
Two immutable sources that were opened generically — the object editor's
copy-through render and the XLS sheet-visibility source-backed commit — are now
opened sealed. On this base an XLS generic commit renders twice, so it goes from
**24 whole-artifact digests to 4**.

Paired ABBA on CPU 16, 8 processes per case, median of process p50s, paired
after/before ratio with a bootstrap 95% interval:

| case (corpus, bytes) | before | after | paired ratio | 95% CI |
| --- | ---: | ---: | ---: | --- |
| probe `commit`, Number edit (`54016.xls`, 984,576) | 23.68 ms | 14.26 ms | **0.600** | [0.591, 0.610] |
| probe `commit`, LabelSst edit (`54016.xls`) | 24.47 ms | 14.97 ms | 0.619 | [0.577, 0.646] |
| probe `commit`, Number edit (`xls-large`, 163,840) | 3.446 ms | 1.866 ms | 0.546 | [0.530, 0.559] |
| `xls_semantic_one_edit_save` (`xls-large`) | 3.377 ms | 1.800 ms | **0.533** | [0.527, 0.546] |
| `xls_visibility_eager_edit_save` (2,135,552) | 25.33 ms | 5.47 ms | **0.216** | [0.216, 0.216] |
| `xls_numeric_eager_rk_mulrk_edit_save` (202,752) | 2.379 ms | 0.507 ms | **0.213** | [0.213, 0.214] |
| `xls_visibility_source_backed_edit_save` | 10.15 ms | 2.24 ms | 0.220 | [0.219, 0.221] |
| `xls_comments_source_backed_edit_save` (16,995,840) | 33.49 ms | 20.26 ms | 0.605 | [0.491, 0.607] |
| `xls_numeric_source_backed_number_edit_save` | 78.40 ms | 30.61 ms | **0.391** | [0.389, 0.392] |
| `xls_numeric_plan_only_number_edit_save` | 33.82 ms | 17.96 ms | 0.531 | [0.531, 0.532] |
| probe plan-only publication phase (`54016.xls`) | 0.950 ms | 0.043 ms | 0.045 | [0.045, 0.048] |
| `cfb_file_owned_same_length_overlay_atomic_save` (16,913,408) | 65.81 ms | 49.79 ms | 0.758 | [0.753, 0.758] |
| control: `cfb_file_same_length_overlay_atomic_save` (generic) | 98.17 ms | 97.94 ms | 0.998 | [0.992, 1.002] |
| control: `xls_semantic_open` (`xls-large`) | 1.498 ms | 1.481 ms | 0.982 | [0.962, 1.051] |

The full table (29 cases, plus one control run separately), including every
control, is under "Measurements". The generic-source control executes the same
user-space instructions per sample (9,239.267 M vs 9,239.263 M); so do the DOC,
PPT and DIFAT-gated XLS controls.
Allocation counts fall by exactly the removed fingerprint buffers (eight
984,576-byte buffers per `54016.xls` generic commit, less 16 bytes of slightly
larger source adapters); peak live bytes do not move. Every published byte,
refusal text and recorded fingerprint is identical to the base across a
126-fixture census (1,319 lines).

## Why this record exists

Change [0746](0746-xls-validation-only-parse.md) (on its own branch, not in
this base) left the XLS generic commit on `54016.xls` with one
`render_copy_through`, 53% of the commit, and priced the SHA-256 work inside it:
planning ×2, the composed-source preflight, the write preflight, the emission
hash and the post-emission preflight, each hashing the whole 985 KB artifact
twice (source and target). Change [0745](0745-ppt-lazy-artifact-digests.md)
listed the same site as a follow-up: `render_copy_through` wraps the editor's
immutable original `Arc<Vec<u8>>` in an `OwnedSource` and opens it through the
generic `SharedOleFile::open`, so the CFB overlay treats bytes that cannot
change as bytes that can. It named the XLS visibility source-backed open as the
same pattern.

On this base the generic commit renders twice (0746's handoff is not here). A
flat profile of the probe's commit loop (`number-generic`, `--reuse-source`,
`54016.xls`) puts **46.8%** of all samples in
`sha2::sha256::x86_sha::compress`. The frame-pointer tree gives each of the two
`render_copy_through` calls 22.5% of the commit, 21.5 points of it in its six
passes: planning 7.1 (two), then the composed preflight, the write preflight,
the emission hash and the post-flight at 3.6 each (`profiles/before-*`). Changes
[0172](changes/0172-cfb-owned-numeric-publication.md) and
[0178](changes/0178-cfb-owned-planning-fingerprint.md) had already specialized
`open_owned` plans by dropping the outer fences, but both kept the composed-view
preflight and the emission hash as a narrow first step, without arguing that
either proves anything about sealed bytes.

## What changed

**1. `3cdcead75c` — a sealed plan computes its digests once
(`crates/litchi-cfb/src/overlay.rs`, `shared.rs`).**

- `ValidatedOverlayPlan::composed_source` takes its complete two-digest
  preflight only for a generic positional source.
- `write_to` and `save` pass a private `EmissionFence` to `write_validated`:
  `Sealed` (read, overlay and write every chunk; hash nothing), `Hashed`
  (atomic save of a generic source; its pre-rename preflight follows) or
  `HashedAndRechecked` (direct write of a generic source; its post-emission
  preflight follows). A debug assertion ties `Sealed` to the sealed provenance
  and to nothing else.
- `SharedOleFile::open_owned_vec(Arc<Vec<u8>>, SourceVersion)` is new: the same
  sealed ingress as `open_owned(Arc<[u8]>, ..)` for callers that hold their
  bytes as a shared vector, without copying them. Both constructors wrap a
  private `SealedBytes` enum in the private `OwnedArcSource` adapter; those two
  types are the only way to reach the sealed rules.
- `OverlayOperationShape` states the sealed contract: an owned plan reports
  `composed_source_preflight_{scans,bytes,chunks} = 0` and `*_emission_bytes =
  0` (its emissions still count one scan and every publication chunk). Generic
  values are unchanged.
- The documentation of the XLS owners and of
  `SourceBackedOverlayPublisher::open_owned` that described the old owned
  rechecks is corrected, and the XLS unit test that pinned the old owned shape
  pins the new one.

**2. `604132d280` — seal the two immutable sources that were opened
generically (`crates/litchi-ole-common/src/object/codec.rs`,
`crates/litchi-xls/src/sheet_visibility.rs`).**

- `Package::render_copy_through` opens the editor's immutable original
  `Arc<Vec<u8>>` with `open_owned_vec`. Its version becomes the allocation's
  `(address, length)` identity, the convention of the XLS comments and
  visibility owners, instead of a fresh process counter; it only has to stay
  stable while this call's views are alive, and nothing outside the call
  observes it. The DIFAT and length gates, the composed reopen and the
  read-back of every stream are unchanged.
- The sheet-visibility source-backed commit opens the snapshot's immutable
  `Arc<[u8]>` with `SourceBackedOverlayPublisher::open_owned` under the same
  `(address, length)` version its private `SnapshotSource` adapter reported,
  so the composed target's version is unchanged; the adapter is deleted. The
  comments owner already did this.

**3. `d663c504f4` — version the harness evidence (`tools/perf-baseline`,
`tools/perf_abba_summary.py`).** The XLS numeric harness labels its operation
evidence `xls_numeric.operation_evidence.v2`. The ABBA validator checks v2
owned rows against the sealed contract and still checks v1 rows (the
registered 0268 package among them) against the v1 contract; generic rows are
the same under both, and any other label is refused. The harness test pins the
v2 values.

**4. `0a3476edd2`** extends one test to atomic `save` (below).

## The contract argument

A plan records two digests, `H(source)` and `H(target)`, where the target is
the source with the plan's changed physical spans applied. Every later complete
fingerprint pass recomputes both over the bytes the plan would read now and
compares them with the recorded ones. What that proves is freshness: the bytes
about to be viewed or emitted are the bytes planning reopened through the CFB
parser, precondition-checked and handed to the owner's validator. It is needed
because a generic `ReadAt` may keep a stable version token while its bytes
change (ADR 0005's `SourceChanged` only covers an honest version). The second
planning scan, the view preflight, the write pre- and post-flights, the
emission hash and the save fences each close that window at a different moment.

For sealed owned bytes the window does not exist. The reader holds a strong
clone of an `Arc<[u8]>` or `Arc<Vec<u8>>`; neither `[u8]` nor `Vec<u8>` has
interior mutability, and safe Rust hands out no `&mut` into a shared `Arc`
(`get_mut` needs a unique owner, `make_mut` copies). The spans are immutable
`Arc<[u8]>` replacement ranges owned by the plan, and `apply_spans` is a pure
function. The target is therefore one fixed byte string for the plan's
lifetime, and recomputing `H` over it can only return the recorded value. Each
removed pass re-proved a fact that cannot have changed:

| complete source + target pass | generic source | sealed, base | sealed, after |
| --- | ---: | ---: | ---: |
| planning | 2 | 1 | 1 |
| `composed_source` preflight | 1 | 1 | **0** |
| `write_to`: preflight, emission hash, post-flight | 3 | 1 (emission hash) | **0** |
| `save`: pre-temp, emission hash, pre-rename | 3 | 1 (emission hash) | **0** |

`render_copy_through` did planning, `composed_source` and `write_to` over a
generic source, so its six passes become one; an XLS generic commit on this
base renders twice.

The provenance is typed and cannot be forged. `source_is_owned_immutable` is
crate-private and set only by `open_owned`/`open_owned_vec` through the private
`SealedBytes`/`OwnedArcSource` pair; `ValidatedOverlayPlan` has no public
constructor, so the digests a sealed plan carries are always the ones its own
planning pass computed over the exact allocation it retains. A caller cannot
mark an arbitrary `ReadAt` sealed with a flag, a version token or a wrapper
type.

The emission hash also compared the 64 KiB-chunked output with the 1 MiB-
chunked planning digest, which would expose a chunk-boundary bug in
`apply_spans`. That is a property of the code, not of the source, so it is now
tested instead of being re-proved on every publication: the new provenance test
publishes five overlay sets (five mini and FAT streams at once including a 130
KB stream whose spans cross 64 KiB boundaries, one mini stream, an exact no-op,
an empty list, and a version-4 4,096-byte-sector file) through the generic,
owned-`Arc<[u8]>` and owned-`Vec` provenances and requires the direct write, a
second write, the atomic save and an end-to-end read of the composed view to
equal the generic plan's output and to hash to the recorded target digest. The
census checks the same equalities on 35 real XLS files.

Why the target digest stays at planning rather than moving into the emission
pass: the composed view's version is derived from the source version and the
target digest, and the XLS numeric, DOC and PPT owners record that version as
the committed snapshot's version. The target digest is therefore needed before
the first byte is written; moving it into the pass that writes would need a
different version derivation and would change a recorded value. Nothing else
moves either: planning computes both digests with
the same function over the same bytes, so every value recorded in plans,
publish reports and owner diagnostics is unchanged. The composed CFB reopen,
selected-stream checks, owner validation, `render_copy_through`'s reopen and
read-back of every stream, the version and length check on every positional
read, the sink progress contract and atomic save's flush, fsync, rename and
parent-sync sequence are all unchanged.

## Breaking changes

Under change 0652's first standing trade-off:

- **`OverlayOperationShape` values for sealed plans.** Fields and types are
  unchanged. An owned plan now reports `composed_source_preflight_{scans,bytes,
  chunks} = 0` and `target_materialization_emission_bytes =
  direct_emission_bytes = atomic_save_emission_bytes = 0`. The three consumers
  that pinned the old owned values (one XLS unit test, the harness test and the
  ABBA validator) are updated.
- **Harness operation evidence `xls_numeric.operation_evidence.v2`.** Reports
  written by this revision carry v2. The validator accepts a v1 row only with
  the v1 contract and a v2 row only with the v2 contract.
- **A sealed plan's `composed_source`, `write_to` and `save` no longer return
  `SourceFingerprintChanged` or `TargetFingerprintChanged`.** Those refusals
  could not fire for bytes that cannot change. Every generic plan still returns
  them exactly as before (tests below).

`SharedOleFile::open_owned_vec` is additive. The private `SnapshotSource`
adapter in `sheet_visibility.rs` is deleted.

## Authority and constraints

- **ADR 0003.** Snapshots stay immutable and cheap to share. Patches still
  authorize by exact bytes, not fingerprints, and every recorded fingerprint
  value is unchanged.
- **ADR 0005.** Input stays positional `ReadAt` with a stable version. A
  mutation of any source that can mutate is still refused by the same passes in
  the same order. The sealed specialization rests on byte ownership the CFB
  reader retains, as change 0172 established for `open_owned`.
- **ADR 0006.** Validation never mutates, and no malformed-input defence is
  weakened: the composed reopen, owner validation and read-back all still run.
  Output bytes are identical (census, tests), and determinism is unchanged.
- **Change 0652, trade-off 1:** the shape values and the evidence schema move,
  as stated above. **Trade-off 2:** every freshness proof that guards a source
  which can change is kept. Only proofs over sealed bytes are removed, so no
  conflict between speed and safety had to be resolved. **Trade-off 3:** the
  in-memory owned source is the common path of every XLS edit and of the object
  editor. A mutating or malicious source can only arrive through a generic
  `ReadAt`, which still pays every pass.
- **Scope:** `litchi-cfb`, `litchi-ole-common`, `litchi-xls` and the harness
  tooling. No ODF or iWork crate, no dependency, no `unsafe`, no thread,
  filesystem, clock or network behaviour. `crates/litchi-cfb/src/writer/`
  (record 0749's file) is untouched.

## Proof obligations and tests

- **Fingerprint values unchanged.** The census prints each source-backed
  commit's diagnostics fingerprints and its publish report's fingerprints; the
  before and after outputs are byte-identical.
  `sealed_and_generic_plans_publish_identical_bytes_and_digests` and
  `owned_vec_source_is_sealed_and_publishes_the_generic_bytes` require the
  sealed plans' digests to equal the generic plan's and the independent
  SHA-256 of the source and of the published target.
- **Published bytes unchanged.** The same census covers every generic,
  source-backed and plan publication over 126 fixtures. In `litchi-ole-common`,
  `sealed_copy_through_publishes_the_generic_overlay_bytes` requires the
  editor's copy-through output (the commit render, snapshot `finish`, editor
  `finish`, the commit patch's `after`, and a chained two-stream edit) to equal
  a generic `SharedOleFile::open` plan's `write_to` output byte for byte. That
  is the base's route, and the comparison includes a free-sector byte that a
  re-render would drop.
- **Composed version unchanged.** It is still derived from the source version
  and the target digest. A test compares two sealed plans' views, and the
  visibility owner passes the same `(address, length)` version its deleted
  adapter reported.
- **Refusals for sources that can change still fire.** Existing tests keep
  every generic fence: `version_and_stable_token_byte_changes_are_caught_before_output`,
  `stable_token_mutation_of_an_emitted_chunk_is_caught_before_success`,
  `atomic_path_late_stable_token_mutation_leaves_destination_unchanged`, the
  no-op family, and `no_op_inverse_and_stale_source_contracts_are_exact`, which
  mutates a splice plan's source between plan and `write_to`/`composed_source`.
  Two new tests pin the fences this change sits next to:
  - `generic_composed_view_still_rechecks_both_digests`: a `MutableSource`
    double changes an unselected byte under a stable version token between plan
    and view. The view and `write_to` refuse with `SourceFingerprintChanged`
    before any byte, a version change is refused as `SourceChanged`, and the
    view still reads one complete pass.
  - `generic_atomic_save_still_hashes_its_emission`: the double changes a byte
    of a not-yet-emitted chunk during `save`. The emission hash, not the
    pre-rename preflight, refuses it (proven by the read count), and the
    destination is untouched.
- **Sealed work is what the contract says.**
  `owned_composed_view_and_direct_write_take_no_recheck_scan` and
  `owned_atomic_save_reads_only_its_emission` count reads on a sealed source:
  the view reads nothing, and a direct write or an atomic save reads each 64 KiB
  chunk exactly once. `operation_shape_matches_generic_and_owned_overlay_policy`
  covers owned `Arc<[u8]>`, owned `Vec`, generic and no-op shapes.
- **Sink progress unchanged.** `hostile_sink_progress_is_typed` now runs its
  zero-write, partial, over-reporting and failed-flush sinks against both a
  generic and a sealed plan.
- **Stale and foreign owner inputs.** The XLS owner-level tests are unchanged
  and pass, among them
  `source_backed_numeric_plan_rejects_stale_and_foreign_change_metadata`,
  `stale_sources_are_rejected_before_selected_reads` and
  `exact_patch_rejects_stale_and_durable_inverse_restores_source`.
- **The validator** (`test_xls_numeric_v2_owned_rows_take_the_sealed_contract`)
  accepts sealed v2 rows for source-backed and plan-only selectors. It refuses
  sealed values under v1, a v2 owned row that claims a composed preflight or an
  emission hash, a v2 generic row that drops either, and an unknown label.

## Correctness evidence

The change-0746 corpus census (`probe/corpus.rs`), extended to print each
source-backed plan's fingerprints (diagnostics and publish report), the digest
of a second publication from the same plan, and, for the numeric plan, the
digest of the composed view read end to end and of an atomic `save`. It runs
every `.xls` of 0746's 126-fixture list through five `cell_values` paths, the
comments and visibility owners (open, generic, source-backed) and the public
reader. Before and after outputs are **byte-identical**: 1,319 lines, SHA-256
`433b688e…afe843` both. That covers 40 Number, 51 LabelSst and 40 no-op generic
commits, 35 source-backed and 35 plan publications, 9+9 comments and 62+62
visibility publications, and every refusal text. On the after leg, each of the
35 numeric plans' published bytes, second publication, composed-view read-back
and atomic save hash to its target fingerprint, and every diagnostics
fingerprint equals its publish report's (`correctness/`).

## Measurements

**Method.** The before leg is a detached worktree at `ab29ac6291` with the
harness-only commit cherry-picked (per the briefing; it changes one string
constant and tests, not a timed region). The after leg is `d663c504f4`. Every
binary is built by the identical command with rustc 1.95.0; `binaries.sha256`
lists them. Every process is pinned to CPU 16. Each case runs 8 processes in the
order A B B A A B B A, each with 3–5 warmups and 20–60 samples (per case in
`latency/summary.json`); a corpus identity check confirms all 8 used the same
artifact. The host was shared with other implementers (load average 3–23), so
paired ratios and instruction counts carry the conclusion. The timing probe is
0746's `probe/main.rs`, unchanged (SHA-256 `9bd72f55…da03d6`), rebuilt against
each leg; its `commit` phase is `Transaction::commit`, and on this base that
renders twice.

**Latency.** Median of process p50s and p95s, paired p50 ratio (median of the
four A/B pairs) with a bootstrap 95% interval (10,000 resamples, seed 748):

| case | before p50 (ms) | after p50 (ms) | ratio | 95% CI | before p95 | after p95 |
| --- | ---: | ---: | ---: | --- | ---: | ---: |
| `xls_semantic_one_edit_save` large | 3.377 | 1.800 | 0.533 | [0.527, 0.546] | 3.457 | 1.843 |
| `xls_semantic_one_edit_save` tiny (4,096 B) | 0.0802 | 0.0420 | 0.524 | [0.517, 0.526] | 0.0880 | 0.0467 |
| `xls_visibility_eager_edit_save` | 25.33 | 5.471 | 0.216 | [0.216, 0.216] | 25.40 | 5.508 |
| `xls_visibility_eager_batch_edit_save` | 25.34 | 5.470 | 0.216 | [0.216, 0.217] | 25.43 | 5.500 |
| `xls_numeric_eager_rk_mulrk_edit_save` | 2.379 | 0.507 | 0.213 | [0.213, 0.214] | 2.391 | 0.517 |
| `xls_visibility_source_backed_edit_save` | 10.15 | 2.238 | 0.220 | [0.219, 0.221] | 10.17 | 2.258 |
| `xls_visibility_source_backed_batch_edit_save` | 10.25 | 2.292 | 0.224 | [0.223, 0.224] | 10.32 | 2.323 |
| `xls_comments_source_backed_edit_save` | 33.49 | 20.26 | 0.605 | [0.491, 0.607] | 33.67 | 20.39 |
| `xls_comments_source_backed_batch_edit_save` | 36.76 | 20.90 | 0.568 | [0.502, 0.570] | 37.00 | 21.09 |
| `xls_numeric_source_backed_number_edit_save` | 78.40 | 30.61 | 0.391 | [0.389, 0.392] | 78.77 | 30.93 |
| `xls_numeric_source_backed_rk_mulrk_edit_save` | 0.824 | 0.263 | 0.318 | [0.316, 0.319] | 0.829 | 0.276 |
| `xls_numeric_plan_only_number_edit_save` | 33.82 | 17.96 | 0.531 | [0.531, 0.532] | 33.93 | 18.08 |
| `xls_numeric_plan_only_rk_mulrk_edit_save` | 0.403 | 0.219 | 0.542 | [0.541, 0.544] | 0.414 | 0.228 |
| `cfb_file_owned_same_length_overlay_atomic_save` | 65.81 | 49.79 | 0.758 | [0.753, 0.758] | 69.40 | 52.67 |
| probe `54016.xls` Number `commit` | 23.68 | 14.26 | 0.600 | [0.591, 0.610] | 24.35 | 15.05 |
| probe `54016.xls` LabelSst `commit` | 24.47 | 14.97 | 0.619 | [0.577, 0.646] | 24.85 | 15.70 |
| probe `54016.xls` `commit_source_backed` | 15.78 | 14.05 | 0.891 | [0.869, 0.931] | 16.14 | 14.56 |
| probe `54016.xls` `commit_source_backed_plan` | 11.88 | 11.72 | 0.995 | [0.940, 1.062] | 12.32 | 12.80 |
| probe `54016.xls` plan publication (`write_to`) | 0.950 | 0.043 | 0.045 | [0.045, 0.048] | 0.957 | 0.048 |
| probe `xls-large` Number `commit` | 3.446 | 1.866 | 0.546 | [0.530, 0.559] | 3.524 | 1.926 |
| probe `xls-large` `commit_source_backed` | 2.100 | 1.751 | 0.839 | [0.819, 0.868] | 2.165 | 1.811 |
| control: `cfb_file_same_length_overlay_atomic_save` | 98.17 | 97.94 | 0.998 | [0.992, 1.002] | 100.87 | 109.72 |
| control: `ppt_source_backed_one_shape_text` large | 0.0287 | 0.0280 | 0.978 | [0.962, 0.987] | 0.0308 | 0.0307 |
| control: `xls_comments_eager_edit_save` (DIFAT gate) | 21.42 | 21.00 | 0.980 | [0.972, 0.998] | 22.41 | 21.46 |
| control: `xls_numeric_eager_number_edit_save` (DIFAT gate) | 28.08 | 28.00 | 0.995 | [0.985, 1.015] | 28.84 | 28.48 |
| control: `ole_common_one_edit_save` (length gate) | 20.89 | 20.78 | 0.989 | [0.971, 1.024] | 21.61 | 21.59 |
| control: `ole_common_finish_render` (length gate) ² | 4.499 | 4.473 | 0.987 | [0.982, 1.017] | 4.706 | 4.575 |
| control: `doc_semantic_one_edit_save` large | 0.698 | 0.695 | 0.994 | [0.974, 1.008] | 0.804 | 0.802 |
| control: `xls_semantic_open` large | 1.498 | 1.481 | 0.982 | [0.962, 1.051] | 1.556 | 1.575 |
| control: probe `54016.xls` `open` | 11.75 | 11.88 | 1.010 | [0.979, 1.057] | 12.29 | 12.53 |

² Run separately after the gates, same method (`latency/ole-common-finish-render/`).

Treatments fall into three groups. The generic copy-through commits (the
semantic, visibility-eager and RK/MulRK-eager selectors and the probe's generic
commits) lose 20 of 24 digests; on the 2.1 MB visibility corpus that is
four-fifths of the commit. The visibility source-backed commit, generic before,
goes from 10 digests (two planning passes and three write passes, each hashing
source and target) to 2. The already-owned plans lose the composed preflight and
the emission hash. The wide intervals of the two `xls_comments_source_backed_*`
rows come from one unusually fast after process in each (17.7 and 18.4 ms,
against 20.2–21.0 ms for the other three); the before processes sit at 33.4–36.8
ms. The plan-only `commit` phase does the same work as before (one planning
pass), so its 0.995 is the expected null; its publication phase is 22× faster.

The controls: the generic CFB file source keeps every pass; the DIFAT-gated XLS
commits and the length-changing object edit decline copy-through before any
overlay code; the DOC edit never reaches it; and the PPT source-backed edit is a
generic `ReadAt`. All are within ±2.2% at p50.

**Instructions and cycles.** User-space `instructions:u` and `cycles:u` per
iteration, `(24 − 4) / 20` from `perf stat` pairs, median of 3 repeats per leg
in A B B A order. The probe rows use `--reuse-source` (edit, commit and publish
per iteration). Harness rows include the selector's untimed per-sample work
(for example re-opening the output to verify it), so their ratios are diluted;
the absolute differences are what the change removed:

| case | before instr (M) | after (M) | ratio | before cycles (M) | after (M) | ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| probe `54016.xls` Number `commit` + publish | 213.62 | 162.43 | 0.760 | 108.42 | 67.81 | 0.625 |
| probe `54016.xls` `commit_source_backed` | 179.08 | 163.77 | 0.915 | 76.01 | 63.43 | 0.835 |
| probe `54016.xls` plan + publication | 148.27 | 143.13 | 0.965 | 61.27 | 56.87 | 0.928 |
| probe `xls-large` Number `commit` | 34.10 | 25.54 | 0.749 | 14.82 | 8.09 | 0.546 |
| `xls_semantic_one_edit_save` large | 227.01 | 218.35 | 0.962 | 71.37 | 64.51 | 0.904 |
| `xls_visibility_eager_edit_save` | 180.16 | 69.62 | 0.386 | 135.36 | 46.58 | 0.344 |
| `xls_visibility_source_backed_edit_save` | 97.94 | 53.74 | 0.549 | 73.37 | 36.71 | 0.500 |
| `xls_numeric_eager_rk_mulrk_edit_save` | 17.07 | 6.54 | 0.383 | 12.31 | 3.47 | 0.282 |
| `xls_numeric_source_backed_rk_mulrk_edit_save` ¹ | 8.09 | 4.96 | 0.613 | 5.15 | 2.73 | 0.530 |
| `xls_numeric_plan_only_rk_mulrk_edit_save` ¹ | 4.48 | 3.44 | 0.767 | 2.84 | 2.03 | 0.716 |
| `xls_comments_source_backed_edit_save` | 346.83 | 260.14 | 0.750 | 282.69 | 212.13 | 0.750 |
| `cfb_file_owned_same_length_overlay_atomic_save` | 8,811.5 | 8,638.1 | 0.980 | 3,455.2 | 3,312.0 | 0.959 |
| control: `cfb_file_same_length_overlay_atomic_save` | 9,239.267 | 9,239.263 | 1.000 | 3,774.1 | 3,768.8 | 0.999 |
| control: `xls_comments_eager_edit_save` | 204.39 | 204.35 | 1.000 | 187.16 | 184.08 | 0.984 |
| control: `ppt_source_backed_one_shape_text` large | 6.844 | 6.844 | 1.000 | 1.627 | 1.629 | 1.001 |
| control: `doc_semantic_one_edit_save` large | 41.391 | 41.386 | 1.000 | 10.87 | 10.96 | 1.008 |
| control: probe `54016.xls` `open` | 148.50 | 148.41 | 0.999 | 51.09 | 52.13 | 1.020 |

¹ From a second run with a 200-iteration spread (10 vs 210) and 5 repeats
(`instructions/isolation-long-spread.json`). The first run's 20-iteration
spread gave the plan-only row 0.767 instructions but 1.320 cycles, which is not
resolvable at 0.4 ms per sample; both runs are retained. SHA-NI retires few
instructions per byte hashed, so cycles fall further than instructions.

**Allocations.** The probe's `alloc-count` build counts the first measured
iteration (open, stage, commit and publish); two runs per leg are identical:

| fixture, operation | calls | allocated bytes | peak live bytes |
| --- | --- | --- | --- |
| `54016.xls` Number `commit` | 118,272 → 118,264 | 71,167,056 → 63,290,464 | 23,289,770 (same) |
| `54016.xls` LabelSst `commit` | 127,898 → 127,890 | 74,276,715 → 66,400,123 | 23,269,070 (same) |
| `54016.xls` `commit_source_backed` | 121,659 → 121,658 | 59,091,760 → 58,107,184 | 25,628,658 (same) |
| `54016.xls` plan, `open` | same | same | same |
| `xls-large` Number `commit` | 4,476 → 4,468 | 12,090,171 → 10,779,467 | 3,995,094 (same) |
| `xls-large` `commit_source_backed` | 4,394 → 4,393 | 10,320,166 → 10,156,326 | 4,605,289 (same) |
| `xls-large` plan, `open` | same | same | same |

Each removed `fingerprints` pass allocated one `min(length, 1 MiB)` buffer. The
generic commit loses 8 of them (4 passes × 2 renders: 8 × 984,576 bytes on
`54016.xls`); the sealed adapter is 8 bytes larger than the `OwnedSource` it
replaces, so the net is 16 bytes less than eight buffers.
`commit_source_backed` loses the one composed-preflight buffer.

**Profiles** (`profiles/`, summaries only). After the change, the
frame-pointer tree of the same `54016.xls` commit loop gives each
`render_copy_through` 7.0% of the commit (was 22.5%), 5.5 points of it the one
remaining planning pass. The flat profile's SHA-256 share falls from 46.8% to
12.3%. On this base the commit is now dominated by the eager `Workbook::new`
(63% of the commit), which 0746 replaces.

## Regression flags (every change above +5%)

- **`cfb_file_same_length_overlay_atomic_save` (control), p95.** The median of
  process p95s rises by 8.8% (100.9 → 109.7 ms) while p50 is 0.998; a longer
  round (60 samples) gives p50 1.016 [0.931, 1.041] and p95 +15.7%. An **A/A
  round with the before binary in both arms** gives process p95s from 99 to
  143 ms and a p95 "ratio" of 0.888, and the two builds retire identical
  instructions per sample (9,239.267 M vs 9,239.263 M; cycles 0.999). The
  generic save path's code and I/O sequence are unchanged; this case's tail is
  fsync noise on a disk shared with other builds (`latency/cfb-generic-long/`,
  `latency/cfb-generic-aa/`).
- **Single-pair p50 flags on unchanged work:** `xls_semantic_open` large pair 3
  at 1.051 (mean 1.054; median 0.982), probe `commit_source_backed_plan` pair 3
  at 1.062 (median 0.995), probe `open` pair 4 at 1.057 (mean 1.054; median
  1.010; instructions 0.999). None of these paths executes changed code in its
  timed region.
- **Cycles, first isolation run:** plan-only RK/MulRK at 1.320 (see ¹ above;
  0.716 with a resolvable spread; wall clock 0.542).
- **Cycles:** probe `open` 1.020 with instructions 0.999 (control).

No treatment regressed at p50, mean or p95.

## What is not claimed

- No registered speedup: `performance_claim: none`.
- Warm, in-memory, serial timings on one shared host pinned to CPU 16, with no
  cold-I/O, concurrency, RSS, cross-platform or compiler-sweep result. The
  filesystem pair is `--filesystem-cache warm` only.
- The XLS generic commit on this base still renders twice (0746's handoff is
  not in the base), so its numbers are for this base. They will shrink by a
  different factor once 0746 merges, because 0746 removes the second render
  outright.
- The sealed digests are not free: each sealed plan still hashes its source
  and target once at planning. Only the repeated passes are removed.
- DOC and PPT source-backed snapshots, the generic CFB file source and every
  caller-provided `ReadAt` are unchanged by design.

## What is left

- **The planning pass itself.** A sealed plan still hashes the source and the
  target once. The source and target byte streams agree up to the first
  changed byte, so the target digest could start from a clone of the source
  hasher's state there. That is value-identical, and it applies to generic
  passes too. Over the census's effective edits the first changed byte sits at
  a median of 35% (LabelSst) to 50% (Number) of the artifact, and at 32.3% on
  `54016.xls` (`profiles/first-change-fractions.txt`). This is a separate change
  that needs its own measurement. A lazy source digest is an alternative, since
  `render_copy_through` never reads it, but it would make plan accessors lazy.
- **The second render of the XLS generic commits on this base** is removed by
  0746's handoff, not here. After both changes, each generic commit hashes its
  artifact twice (one source and one target digest) instead of 24 times.
- **DOC's owned opens stay generic.** `body_text::SourceSnapshot::from_bytes`
  and `parse` wrap owned bytes in an `OwnedSource` and open them through the
  generic path; the survey counts about 23 complete scans in
  `publish_operation`. Sealing those opens needs the fence argument of changes
  0644 and 0659 re-made for owned bytes, and no harness selector covers it.
- **The stream-move inverse plan** is built over the forward plan's composed
  view with generic provenance, even when the forward plan is sealed (0172
  reviewed that downgrade). A sealed forward's view is itself immutable, so the
  inverse could be sealed too; nothing measures it.

## Verification

`results/change-0748/gates.txt`, at `0a3476edd2`:

All commands and exit codes are in `gates.txt`, run with
`CARGO_TARGET_DIR=targets/0748` and `TMPDIR` in the scratch directory so no
test staged a file under `/tmp`:

- `cargo fmt --all --check` and `cargo fmt --manifest-path
  tools/perf-baseline/Cargo.toml --check`: exit 0.
- `cargo check --all-targets` for `litchi-cfb`, `litchi-ole-common`,
  `litchi-xls` and every in-scope crate that depends on them (`litchi-doc`,
  `-ppt`, `-vba`, `-sign`, `-crypto`, `-ograph`, `-docx`, `-pptx`, `-xlsx`,
  `-xlsb`, `-opc`, `-drawingml`, `-ooxml-common`, `-spreadsheet-drawing`), and
  `cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt
  --all-targets`: exit 0.
- `cargo clippy -p litchi-cfb -p litchi-ole-common -p litchi-xls --lib --no-deps
  -- -D warnings` and `cargo clippy -p litchi-cfb -p litchi-ole-common
  --all-targets --no-deps -- -D warnings`: exit 0. `cargo clippy -p litchi-xls
  --all-targets --no-deps -- -D warnings` exits 101 on the pre-existing
  `unusual_byte_groupings` error in `tests/xls_query_index_cache.rs:546`, a
  file this change does not touch; the same command on the base exits 101 with
  the same error, and with that one lint allowed it exits 0 here.
- `cargo test`: `litchi-cfb` (375 library tests and the integration targets),
  `litchi-ole-common`, `litchi-xls`, `litchi-doc`, `litchi-ppt`, `litchi-vba`
  `-sign` `-crypto` `-ograph`, and the facade with `doc,docx,ppt,pptx,xls,xlsx,
  xlsb,odt`: exit 0.
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-cfb -p litchi-ole-common -p
  litchi-xls --no-deps`: exit 0.
- `cargo test --manifest-path tools/perf-baseline/Cargo.toml` (535 harness unit
  tests passed and 1 ignored, plus its other targets), `python3 -m unittest
  tools.test_perf_abba_summary`, and the CRUD coverage validator: exit 0.
- `check_crate_boundaries.py`, `non_iwork_gate.py verify` and
  `check_perf_claims.py --mode structural` (which re-validates the registered
  ABBA packages, including 0268's v1 XLS numeric rows): exit 0.

## Cleanup

`results/change-0748/cleanup.json`. The census source, both manifests, every
script, every raw timing report (gzipped), the isolation and allocation JSON,
profile summaries and the after census output are retained in the packet. The
binaries, `perf.data` files, target directories (`targets/0748`,
`targets/0748-before`, `targets/0748-release`, `targets/0748-probe`,
`targets/0748-before-probe`, `targets/0748-probe-fp`,
`targets/0748-before-probe-fp`), the `0748-before-src` worktree and the scratch
directory were deleted after the SHA-256s were recorded. The worktree and
branch are kept.
