# 0760: the snapshot slide-root memo re-applied under ADR 0032 — a one-edit commit on the 100 × 100 deck stops rescanning its 99 untouched slides: commit 21.19 → 1.61 ms (−92.4%), one-edit edit/save 45.88 → 26.77 ms (−41.8%), with every value, refusal, patch and published byte unchanged

Status: retained, implemented in `crates/litchi-pptx` (`dfde1e43bb`).
`performance_claim: none` — the paired medians, instruction, cycle and
allocation counts below are reported as evidence beside an A/A floor measured
in the same session, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded and untouched.

Base `1d1044e3ac`; branch `perf/0760-pptx-slide-root-memo-reapply`. The
measured head is the production commit `dfde1e43bb`.

## Result

The harness's own cases. Both legs are built with the identical command,
`cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`,
from detached worktrees of equal path length (`0760-before-src` at the base,
`0760-branch-src` at the head) into target directories of equal length, and run
from binary paths of equal length; eight processes per arm in four ABBA blocks,
pinned to core 4, paired within each block. The A/A floor is the before binary
against a byte-identical copy, two blocks, in the same session.

| case | deck | before p50 | after p50 | paired p50 change [95% CI] | paired mean change | A/A p50 [95% CI] |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `pptx_semantic_one_edit_save` | large | 45.884 ms | 26.774 ms | **−41.79%** [−42.87%, −40.97%] | −41.93% | −0.18% [−0.82%, +2.93%] |
| `pptx_semantic_one_percent_edit_save` | large | 150.345 ms | 150.203 ms | −0.09% [−0.28%, +0.10%] | −0.15% | +0.37% [+0.02%, +1.50%] |
| `pptx_semantic_noop_edit_save` | large | 24.003 ms | 24.108 ms | +0.19% [−0.66%, +1.31%] | +0.72% | +0.10% [−0.83%, +3.02%] |
| `pptx_semantic_full_text` (control) | large | 27.307 ms | 27.378 ms | +0.62% [−1.89%, +2.18%] | +0.28% | −0.51% [−5.89%, +3.67%] |
| `pptx_semantic_one_edit_save` | medium | 1.522 ms | 1.332 ms | **−12.64%** [−13.32%, −12.06%] | −12.36% | +0.25% [−0.49%, +0.65%] |
| `pptx_semantic_one_percent_edit_save` | medium | 1.517 ms | 1.331 ms | **−12.45%** [−12.75%, −11.89%] | −12.45% | −0.11% [−0.40%, +0.71%] |
| `pptx_semantic_noop_edit_save` | medium | 0.814 ms | 0.817 ms | +0.14% [−0.45%, +0.99%] | +0.26% | +0.16% [−0.59%, +0.70%] |
| `pptx_semantic_full_text` (control) | medium | 0.342 ms | 0.344 ms | +0.75% [−1.44%, +1.21%] | +0.15% | −0.28% [−0.83%, +0.25%] |

The large deck is `pptx-semantic-large` (100 slides × 100 text boxes); medium
is 12 × 8. The large one-percent case edits one shape on each of the 100
slides, so every slide is rewritten and nothing can hit: its unchanged time is
the expected null result. The medium one-percent case makes one edit, like the
one-edit case.

The harness phase case (`pptx_semantic_opened_transaction_phases`, one edit,
median of per-process medians; A/A is the same session's floor):

| phase | large before | large after | change | A/A | medium before | medium after | change | A/A |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `opened_presentation` (capture) | 23.720 ms | 23.785 ms | +0.28% | +1.12% | 0.722 ms | 0.724 ms | +0.21% | −0.11% |
| `Snapshot::edit` | 0.066 ms | 0.062 ms | −5.81% | −10.58% | 0.014 ms | 0.014 ms | +1.45% | +0.98% |
| `set_shape_text` | 0.783 ms | 0.788 ms | +0.68% | −1.36% | 0.086 ms | 0.087 ms | +0.42% | −0.79% |
| `Transaction::commit` | 21.193 ms | 1.607 ms | **−92.42%** | +1.16% | 0.540 ms | 0.346 ms | **−36.00%** | −0.01% |
| `apply_opened_presentation_commit` | 0.324 ms | 0.267 ms | −17.43% | −18.11% | 0.064 ms | 0.065 ms | +1.22% | +1.62% |
| `Package::to_bytes` | 0.467 ms | 0.426 ms | −8.68% | −12.61% | 0.098 ms | 0.097 ms | −0.49% | −0.46% |
| total | 46.317 ms | 26.952 ms | **−41.81%** | +0.73% | 1.525 ms | 1.336 ms | **−12.42%** | +0.16% |

The paired p50 change of the phase totals is −41.85% [−43.38%, −41.04%]
(large) and −12.49% [−13.35%, −11.98%] (medium). The sub-millisecond apply and
publication phases move as much in the A/A floor as in the A/B run, so none of
their movement is attributed to this change. The harness's own output digest
of the phase case is identical in every process of both arms
(`cb34b499…` on the large deck).

Per region, from probes built the same way against the same two trees (user
instructions and cycles counted at two iteration counts and differenced, ABBA,
median of four pairs; the edit/save cycles have their per-iteration package
construction subtracted):

| region | deck | before instructions | after instructions | change | before cycles | after cycles | change |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `Transaction::commit`, one edit | large | 391.62 M | 26.34 M | **−93.27%** | 94.33 M | 6.99 M | **−92.59%** |
| `Transaction::commit`, one edit | medium | 8.62 M | 5.08 M | −41.02% | 2.22 M | 1.43 M | −35.66% |
| `Transaction::commit`, one percent (100 slides rewritten) | large | 687.44 M | 687.50 M | +0.01% | 158.35 M | 160.62 M | +1.43% |
| `Package::opened_presentation` (capture) | large | 384.09 M | 384.13 M | +0.01% | 90.95 M | 93.88 M | +3.22% |
| `Package::opened_presentation` (capture) | medium | 7.26 M | 7.26 M | +0.07% | 1.84 M | 1.91 M | +4.14% |
| no-op edit/save cycle | large | 413.44 M | 413.52 M | +0.02% | 104.13 M | 108.83 M | +4.51% |
| no-op edit/save cycle | medium | 11.08 M | 11.09 M | +0.10% | 3.69 M | 3.79 M | +2.94% |
| one-edit edit/save cycle | large | 822.77 M | 457.56 M | **−44.39%** | 203.15 M | 118.75 M | **−41.55%** |
| one-edit edit/save cycle | medium | 21.73 M | 18.21 M | −16.19% | 6.84 M | 6.14 M | −10.27% |
| one-percent edit/save cycle | large | 2,815.89 M | 2,817.13 M | +0.04% | 676.31 M | 671.82 M | −0.66% |
| `presentation().text()` (control) | large | 507.48 M | 507.35 M | −0.03% | 122.86 M | 124.68 M | +1.48% |
| `presentation().text()` (control) | medium | 5.90 M | 5.91 M | +0.02% | 1.48 M | 1.49 M | +0.99% |

The capture and no-op cycle rows add about 40,000 instructions per capture on
the large deck (about 400 per slide: an owner lookup, an `Arc` clone and a
32-byte entry each, then one sort). Their cycle rows are a placement effect of
the two probe builds, not work: the probes' `setup` region
(`Package::from_vec`, which never reaches the changed code) moves by +3.4% in
cycles with identical instructions (5.60 → 5.79 M cycles, 21.65 M instructions
both), and the harness binaries, laid out differently, show
capture at +0.28% against a +1.12% A/A floor and the no-op case at +0.19%
against +0.10%.

Exact allocations per region (probes with a counting global allocator,
probe-only; two repeats agree exactly):

| region | large before | large after | medium before | medium after |
| --- | ---: | ---: | ---: | ---: |
| capture calls / bytes | 75,071 / 18,755,734 | 75,073 / 18,758,974 | 3,458 / 2,997,909 | 3,460 / 2,998,333 |
| one-edit commit calls / bytes | 78,727 / 5,301,054 | 17,745 / 1,365,579 | 4,162 / 326,564 | 3,460 / 277,961 |
| one-edit edit/save calls / bytes | 166,942 / 26.62 MB | 105,964 / 22.69 MB | 10,410 / 4.22 MB | 9,712 / 4.17 MB |
| no-op edit/save calls / bytes | 83,487 / 19.66 MB | 83,491 / 19.66 MB | 5,450 / 3.21 MB | 5,454 / 3.21 MB |
| one-percent edit/save calls / bytes | 624,921 / 111.50 MB | 624,927 / 111.51 MB | as one edit | as one edit |

A capture gains exactly two allocations and 3,240 bytes on the 100-slide deck —
100 entries of 32 bytes and the 40-byte `Arc` header — and 424 bytes
(12 × 32 + 40) on the 12-slide deck, the resident cost ADR 0032 states. Each
further snapshot that builds or projects a memo (the commit's recapture, the
publication's rebind) adds the same two allocations, which is the +4 and +6
calls of the no-op and one-percent rows.

## What was changed

`dfde1e43bb` re-applies change 0743's withdrawn commit `99ce9c5e34` (commit C
of [0743](0743-pptx-semantic-text-and-edit-path.md)), ported onto the moved
base and changed where ADR 0032 requires:

| file | change |
| --- | --- |
| `notes/package.rs` | `SlideRootRecord`, `SlideRootEntry`, `SlideRootMemo` (`from_records`, `project`, `lookup`); `SlideRootProof::record`; a test-only reservation-refusal hook |
| `notes/mod.rs` | exports |
| `parts/slide.rs` | `finish_from_processed` consults the memo for the exact raw observation MCE just read; a hit is re-derived by `debug_assert_eq!`; test-only hit and scan counters |
| `parts/mod.rs` | test-only exports |
| `presentation/package.rs`, `presentation/model.rs` | `capture_slides_with_mce` passes the memo through |
| `opened/model.rs` | `Snapshot::slide_roots`; `capture_internal` builds the memo after the owned package's digest memo exists; `rebound_to` projects it; `capture_with_revision_and_digests_and_mce` takes the parent memo |
| `opened/transaction.rs` | `Transaction::commit` passes its source snapshot's memo |
| `opened/cross_copy_plan.rs` | the four cross-slide copy captures pass no parent memo, as before |
| `opened/mce_retention_tests.rs` | its commit-shaped capture helper passes the memo, as a commit does |
| `opened/slide_root_memo_tests.rs` | 22 tests (below) |

Differences from `99ce9c5e34`:

1. **Typed reservation refusal (ADR 0032 section 3).**
   `SlideRootMemo::from_records` returns `Result<Self>`; a refused table
   reservation is `Error::Allocation { resource: "opened-presentation
   slide-root memo", .. }`, raised by the capture after every validation has
   passed, so it can only replace a success and never another refusal. The
   original built an intermediate record vector that also became empty when it
   could not be reserved; that vector is gone. The memo is built directly from
   the capture's proofs (an `ExactSizeIterator` of records), which now live
   until then.
2. **Only returned classifications are memoized (ADR 0032 section 1).** The
   original kept `Option<Conformance>`, so a scan's `None` — the notes graph's
   refusal "invalid sld root or namespace" — could in principle be kept. Entries
   now hold `Conformance`; a slide the scan refuses is rescanned every time.
3. **An owner `Arc` must alias its key.** `from_records` and `project` admit an
   entry only when the `Arc` the owner hands over has the entry's address and
   length. The digest memo's owner already guarantees this; the check makes the
   admission independent of that guarantee.
4. **The rebind projection is unchanged:** `project` keeps the empty-memo
   degradation of the part-digest projection beside it, since a rebind is
   infallible (ADR 0032 section 3).
5. **Port.** Change 0751 deleted `capture_with_revision` and added four
   `capture_with_revision_and_digests_and_mce` calls in the cross-slide copy;
   they pass `None`, which is what 0751's merge note anticipated and what the
   copy consulted before. The spec-gap merge (0759) touched none of the
   functions this change edits.

No public function, type, trait, variant, durable format, limit or dependency
is added or removed. `Snapshot` gains a crate-private field; its hand-written
`Debug` output is unchanged. The facade retains no slide-root memo: only
snapshots hold one, as in `99ce9c5e34`.

## Authority

* **ADR 0032** (Accepted 2026-09-24 by the owner's decision 1 of
  [0758](0758-owner-decisions-2026-09-24.md)) admits snapshot memos of small
  derived values of payload bytes under the ADR 0005 memo amendment's
  conditions, and section 4 authorizes this re-application with section 3's
  typed-error rule.
* **ADR 0005**, the 2026-09-16 memo amendment and its 2026-09-24 clarification:
  every condition is proven by a test (next section). The facade adoption and
  fill clauses are vacuous because no facade retains this memo; a test shows
  it.
* **ADR 0003** — `commit()` still validates the complete staged package; the
  only work not repeated is the classification of bytes the source snapshot's
  capture already classified, proven identical by allocation identity. Exact
  no-ops remain exact (the no-op commit returns the source snapshot, memo
  included); revisions, patches and conflicts are unchanged.
* **ADR 0006** — validation never mutates, refusals stay typed and identical,
  and no output byte moves.
* **0652** trade-off 2 (correctness first) decided the three differences above;
  trade-off 3 (benign common path) is the scope: a commit that leaves most
  slides untouched.
* ADRs 0002/0024/0011: one crate, no new dependency, no archive type in any
  signature. ADR 0030 (lazy part decode) is untouched; no `unsafe`, thread,
  clock, I/O or ambient state is added.

## Each ADR 0032 condition and the test that proves it

All in `crates/litchi-pptx/src/opened/slide_root_memo_tests.rs`.

| condition | proven by |
| --- | --- |
| `f` is a function of the payload bytes alone | `a_new_allocation_with_equal_bytes_is_a_miss_not_a_hit` (equal bytes, new allocation: miss, same result); `memo_assisted_recaptures_equal_plain_ones_over_the_pptx_corpus`; the differential below |
| keyed on allocation identity with a strong reference | `an_entry_holds_its_allocation_strongly_so_its_key_cannot_be_recycled` (the entry's own strong count keeps the allocation, and so its address, alive after every other owner is gone; equal bytes elsewhere and sub-slices miss; dropping the memo releases it) |
| entries only for allocations the owning snapshot's package holds | `every_entry_names_an_allocation_its_snapshots_package_holds` (capture, commit, publication, rebind); `from_records_admits_only_owned_aliasing_successful_classifications`; `a_foreign_arc_that_does_not_alias_its_payload_is_never_memoized_or_pinned`; `a_published_snapshot_keeps_classifications_only_for_its_own_allocations` |
| a failure of `f` is never memoized | `from_records_admits_only_owned_aliasing_successful_classifications` (a `None` record is not admitted); `the_memo_never_turns_a_refusal_into_a_success` (CDATA in the first and middle slide, an unbound prefix in the last: assisted, plain and cold captures refuse identically) |
| projected, never inherited, on a rebind | `a_rebind_projects_the_memo_and_never_inherits_a_moved_allocation` (the moved slide's entry is dropped, kept entries retain the rebound package's own `Arc`, the rebind takes no reference to the old allocation, and a later commit misses the moved slide) |
| a miss is an ordinary recomputation; no value, refusal or byte depends on hits | `published_bytes_patches_and_revisions_never_depend_on_memo_contents` (full, empty and partial memos: identical patches, revisions, slides, chained commits and published bytes); `a_refused_projection_is_an_empty_memo_like_the_digest_projection`; `a_no_op_commit_and_an_all_slide_commit_keep_their_exact_meaning` |
| built or projected, never mutated | `a_memo_is_built_or_projected_and_never_mutated` (the source memo's `Arc` and every entry are unchanged after commits, a publication, a rebind, a second commit and a refused commit; clones share one memo; commits and rebinds make new ones) |
| resident cost bounded and charged | `the_memo_costs_one_32_byte_entry_per_captured_slide` (32-byte entries, a 40-byte header, exact reservation, bounded by the capture's slide limit); the allocation counts above |
| reservation refusal is a typed error with no partial snapshot | `a_refused_memo_reservation_is_a_typed_allocation_error_and_leaves_no_partial_snapshot` (capture, commit and publication each return `Error::Allocation` with the memo's resource; the facade offers no digest memo, the source snapshot's memo is untouched and the facade's published bytes do not move; without the refusal every step succeeds and equals a cold capture) |
| every reused value is re-derived in test and debug builds | `every_hit_is_re_derived_in_debug_builds` (scans = misses + hits in debug builds); `a_planted_wrong_classification_is_caught_by_the_re_derivation` (a planted wrong classification panics with the assertion's message at its first hit) |
| no added payload observation | `a_memo_hit_adds_no_payload_observation_over_a_plain_recapture` |
| the reuse is real | `a_one_slide_commit_rescans_only_the_rewritten_slide_and_equals_a_cold_capture` (four of five slides hit); `the_notes_owning_deck_reuses_proofs_and_keeps_its_notes_graph` |
| the facade holds none | `the_facade_retains_no_slide_root_memo` (after publication and dropping every snapshot, each slide payload is held only by the facade's package and its digest memo) |

ADR 0032's verification list is met: its ten focused tests are the ten ported
ones (adapted to hits of type `Conformance`); item 2 is the refusal test above;
item 3, the 24 MCE-retention tests of change 0704 (including the
observation-count test), pass with the memo in place; item 4 is this record's
measurement.

## Evidence that motivated it

Change 0743's probe measured this memo at 21.45 → 2.13 ms for the one-edit
commit on the large deck, and its profiles showed why: every capture classifies
each slide with the complete notes-graph scan (88% of capture before 0743's
retained changes), and `Transaction::commit` recaptured the staged package from
scratch, although the staged package shares every slide payload allocation the
transaction did not rewrite. At this base the phase case confirms it: the
one-edit commit was 21.19 of 46.32 ms on the large deck, the same order as the
capture before it.

What remains in the commit at the head is profiled in
[`profiles/commit-branch-fp.txt`](results/change-0760/profiles/commit-branch-fp.txt)
(frame-pointer probe, 1.575 ms per commit): the recapture is about three
quarters of it — per-slide MCE preprocessing checks 13% of the process, the one
rewritten slide's notes-root scan 13%, the presentation-level notes-graph load
15%, slide-reference resolution 10% — then compaction 9%. Building
`Capabilities::ooxml_baseline()`, a `HashSet` of namespace strings, on every MCE
call accounts for about 7.5%.

## Measurements

Host AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, toolchain 1.95.0
(the worktree's `rust-toolchain.toml`); every measured process `taskset -c 4`;
other agents built and measured concurrently (one-minute load average 11–29
before the processes, logged per process). Large-deck processes take 15 samples
after 3 warmups, medium ones 60 after 10, the phase case 15 after 3 on both
decks. Each paired ratio is after / before within one ABBA block (positions 0/1
and 3/2); the interval is a percentile bootstrap over the paired processes
(10,000 resamples, seed 760); samples within a process are never treated as
independent. No measurement ran while this change's own builds ran.

Binaries (SHA-256): before
`5c25958299cc2c92403590bbfd6ff8cda4128b8115e850df20c951caebf7b5df`, after
`5905ca2837f2c9d77edf36e670ef1c6adfa1451c0525e6a3a6f5ecb9e3907122`; the A/A
copy equals the before binary; the probes are listed in
[`binaries.sha256`](results/change-0760/binaries.sha256).

Every harness process also ran under `perf stat`. Whole-process user
instructions (which include corpus construction and verification) fell 3.46%
(large group), 3.67% (medium) and 14.49% (phases), each within 0.02
percentage points across the eight pairs; the A/A pairs moved +0.00%,
−0.00% and −0.00%.

### Regression flags

Every paired ratio above 1.05 in p50, p95 or mean is listed in
[`timing/tables.md`](results/change-0760/timing/tables.md). None is on a p50.
The A/B run has seven; the A/A floor has none.

* Round 3, positions 3/2 (large group): `pptx_semantic_full_text` p95 +320.06%
  and mean +68.13%, `pptx_semantic_noop_edit_save` p95 +158.40% and mean
  +10.30%, `pptx_semantic_one_percent_edit_save` p95 +5.12%. The after process
  at position 2 was time-shared: its single worst samples were 129.06 ms (full
  text) and 65.01 ms (no-op) against p50s of 27.88 and 24.47 ms, its wall time
  12.39 s against 10.3–10.9 s for the other after processes, while its user
  cycles (46.24 G) were in line with them (44.97–45.79 G). The load average rose
  from 17 to 28 during that block.
* `pptx_semantic_noop_edit_save` large p95 +6.81% (round 0, positions 3/2) and
  +5.92% (round 1, positions 0/1): the after p95s (25.40, 25.47 ms) lie inside
  the before arm's own p95 range (23.78–28.22 ms), and the p50 moves +0.19%
  [−0.66%, +1.31%] against the A/A's +0.10% [−0.83%, +3.02%]. The no-op path
  does build one memo per capture and one projection per publication (two
  allocations, about 40,000 instructions each on the large deck); nothing of
  that size is resolvable in wall time here.

## Correctness evidence

* **Focused tests.** The 22 tests above; the whole `litchi-pptx` suite passes
  (1,189 tests), and so does the facade's (382).
* **Differential.** A probe built against the base and the head ran the same
  public flows over the repository's 78 PPTX fixtures and 67,712 mutated
  packages (18 structured mutations of the first, middle and last slide —
  CDATA, unbound prefix, DTD, processing instruction, comment, Strict namespace,
  wrong root, duplicate attribute, MCE alternate content, undeclared and
  declared ignorable namespaces, BOM, whitespace, text change, empty, truncated,
  swapped and shared slide payloads — plus 850 seeded random byte mutations per
  fixture). Each package runs open, capture, a no-op commit and publication,
  one-shape edits of the first, middle and last slide (each committed, chained
  into a second commit from the committed snapshot, published, then edited,
  committed and published again from the published snapshot), an all-slide
  edit, a notes edit, a shape removal, a slide move and an MCE-retention-off
  capture with an edit. The 340,198 outcome lines — every refusal's text, and
  every revision, slide identity, patch and published-bytes digest — are
  byte-identical between base and head (37,732 commits and publications, 1,206
  identical commit refusals and 53,862 identical capture refusals among them);
  see [`differential/README.md`](results/change-0760/differential/README.md).
* **Outputs.** The published archive, durable patch, revision and full text of
  the tiny, medium and large decks under the no-op, one-edit and one-percent
  edits — 30 artifacts — are byte-identical between base and head.

## What is not claimed

* That any real producer deck improves by these amounts. The corpus is the
  harness's generated text-box deck, compact and marker-free; the effect scales
  with the number of slides a commit leaves untouched, and is nil when every
  slide is rewritten (the large one-percent case).
* Anything for `Package::opened_presentation` itself, including a capture right
  after a publication: the facade keeps no slide-root memo, so every fresh
  capture still scans every slide. Only a commit's recapture hits.
* That the capture, no-op or full-text cases moved: their wall-clock changes are
  within the A/A floor, and their cycle changes in the probes are matched by an
  untouched region of the same binaries.
* A cross-slide copy, removal-plan, patch or history result: those captures pass
  no parent memo and are unchanged.
* An RSS, cold-cache, concurrency, facade (`litchi::Presentation`), source-backed
  or range-source result. Allocation counts are requested calls and bytes on the
  probe's system allocator, not RSS.
* That the commit is at a floor (see the profile above).

## Follow-ups (not done here)

* A facade that retains the published snapshot's slide-root memo, as ADR 0032
  section 2 and ADR 0005's clarified amendment allow, would let
  `opened_presentation()` after a publication reuse every untouched slide's
  classification: 23.7 ms of the 26.95 ms large-deck total is that capture.
  The harness's cases start every iteration from a fresh package and would not
  show it.
* `Capabilities::ooxml_baseline()` is rebuilt on every MCE call.
* Under `--all-features`, three golden cross-copy tests fail identically at this
  base and at the head (see Verification).

## Verification

Gates run in the worktree at `dfde1e43bb` with its own target directory, every
Cargo command with `--locked --offline` and `TMPDIR` under the change's scratch
directory; commands, exit codes and output tails are in
[`gates.txt`](results/change-0760/gates.txt):

| gate | exit |
| --- | --- |
| `cargo fmt --all --check` | 0 |
| `cargo check -p litchi-pptx --all-targets` | 0 |
| `cargo check -p litchi-pptx --all-targets --all-features` | 0 |
| `cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets` (the facade that depends on this crate) | 0 |
| `cargo clippy -p litchi-pptx --lib --no-deps -- -D warnings` | 0 |
| `cargo clippy -p litchi-pptx --all-targets --no-deps -- -D warnings` | 101 at the head **and 101 at the base** `1d1044e3ac`: the same three `clippy::err_expect` lints in `opened/tests.rs`, which this change does not touch; not fixed here |
| `cargo test -p litchi-pptx` | 0 — 1,189 passed, 0 failed, 3 ignored |
| `cargo test -p litchi-pptx --all-features` (with `--no-fail-fast`) | 101 — 1,200 passed, 3 failed, 3 ignored; the three failures (`digest_reuse_tests::every_patch_revision_output_and_refusal_matches_the_base`, `media_transfer_tests::captures_that_do_not_fit_record_the_recompressing_route`, `media_transfer_tests::legacy_lpcp0003_patches_are_refused_by_name`) are golden values of compressed cross-copy output that fail with byte-identical assertion values at the base; the default-feature run passes all 56 cross-copy tests. Not in the briefing's known-failure list: reported, not fixed |
| `cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` | 0 — 382 passed, 0 failed, 7 ignored |
| `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-pptx --no-deps` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 1 — the known failure at this base (litchi-xldm registration), per the wave briefing |
| `python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural` | 0 (this record registers no claim) |

The harness is unchanged, so its own suite and the coverage validator were not
required; its semantic PPTX cases nevertheless ran every measured iteration
through the harness's reopen verification, and every measured process exited 0.
The root `Cargo.lock` copied from the main checkout predates the spec-gap merge
and failed `--locked`; `cargo metadata --offline` resolved it minimally (the new
`litchi-xldm` member and dependency edges between workspace packages, no
external version change) before the gates ran. It is gitignored and not
committed.

## Cleanup

Recorded in [`cleanup.json`](results/change-0760/cleanup.json). Binary digests
are recorded in the packet; the record and packet are committed first, and then
the target directories `targets/0760`, `targets/0760-before` and
`targets/0760-branch`, the scratch contents (binaries, raw outputs of the
differential and the dumps, logs, `TMPDIR`) and the two detached source
worktrees `0760-before-src` and `0760-branch-src` (`git worktree remove
--force`) are removed. The worktree and branch are kept.

## Retained evidence

[`results/change-0760/README.md`](results/change-0760/README.md).
