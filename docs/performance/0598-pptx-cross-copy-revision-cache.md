# 0598: the cross-package slide copy stops re-serializing and re-hashing packages it has already proven

Status: retained. `performance_claim: none` — this record carries exact
callgrind call counts and instruction differentials from two isolation pairs,
and paired native timings reported beside the host's A/A floor. No claim is
registered.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **PPTX-2** of change
[0587](0587-remaining-opportunity-survey.md) (rank 5), building on change
[0590](0590-pptx-opened-transaction-revision-reuse.md), which introduced
`capture_with_revision`. Change [0454](changes/0454-pptx-unnamed-lossless-copy.md)
introduced the physical revision this record reuses; its "dedup only with proven
equivalence" rule is untouched.

## The mechanism 0587 found

A cross-package slide copy proves two things about each of three packages: a
*semantic* revision (`package_fingerprint`, a SHA-256 over the complete OPC
graph) and a *physical* revision (`physical_package_fingerprint`, a SHA-256 over
the exact serialized archive, taken by streaming `OpcPackage::to_stream` into a
hashing sink). 0587 measured the whole child of `pptx_cross_copy_media_rich` at
166 G `Ir` with SHA-256 at 79.5% of it, `package_fingerprint` at 46.5% in 65
calls, `physical_package_fingerprint` at 25.9% in 37 full serializations and
`bounded_package_bytes` at 13.1% in 8, while the closure inventory the copy
actually performs was under 0.5%. Those figures were taken at
`2fc5fc65`; change 0590 has landed since, and this record re-measures the whole
path on its own base (`f22f93935`), where the same profile is 158,965,107,851
`Ir` with 60 semantic hashes, 37 archive serializations and 8 bounded candidate
archives.

Most of those passes re-derive a value the same call already holds:

1. `plan_cross_slide_copy` hashes the source and the destination archives
   (`cross_copy_plan.rs:890`, `:892` at the base), builds the candidate,
   captures it — which hashes its graph — and then hashes that same graph again
   for `target_revision` (`:912`) and serializes it a second time for
   `target_physical_revision` (`:913`).
2. `apply_plan` hashes the live source and destination graphs and archives
   (`:457`, `:463`, `:469`, `:477`), captures both — hashing both graphs a
   second time — and then replans, which hashes both archives a third time.
3. After `validate_candidate` has proven the candidate, `apply_plan` serializes
   and hashes it once more (`:523`) to compare against a value it proved a few
   lines earlier.

`apply_patch` repeats the same shape.

## What was changed

Everything below removes a recomputation whose result is already held, in the
same call, for the same bytes. No revision value, proof format, refusal,
published byte or limit changes.

**`crates/litchi-pptx/src/opened/model.rs`** — `Snapshot` gains one private
field, `physical_revision: Arc<OnceLock<(usize, [u8; 32])>>`: the
serialized-archive revision of its immutable package, memoized together with
the archive bound it was taken under. A snapshot owns its `OpcPackage` behind an
`Arc` and never mutates it after capture, so the archive it publishes — and
therefore that archive's digest — is a pure function of the snapshot.
`Snapshot::rebound_to` starts an **empty** cache rather than inheriting the
source snapshot's: `packages_equal` proves the `package_fingerprint` inputs are
identical and says nothing about ZIP ordering, compression or retained source
bytes, so a physical revision must not travel across a rebind.

**`crates/litchi-pptx/src/opened/cross_copy_plan.rs`**

- `seal_physical_revision(archive_digest, length)` is factored out of
  `physical_package_fingerprint`. It is now the only place the
  `litchi-pptx-cross-physical-v2` domain string and the `u64` length prefix are
  combined, so every route to a physical revision produces the same value by
  construction.
- `snapshot_physical_revision(snapshot, limits)` replaces the two direct calls
  in `prepare_cross_slide_copy_for_slides`. On a cache hit under the same
  archive bound it still runs `reject_unknown_non_part_members` — the refusal
  `physical_package_fingerprint` performs first — and then returns the memoized
  digest. A different bound is a miss and recomputes.
- `remember_physical_revision(snapshot, limits, revision)` seeds the cache at
  application with the value the caller just proved on the package `capture`
  cloned.
- `BoundedVecWriter` now digests the bytes it accepts, and
  `bounded_package_bytes` returns `(Vec<u8>, [u8; 32])`. The bounded `Vec` is
  *not* removed: `build_candidate` reopens it with
  `OpcPackage::from_vec_reusing_payloads`, so the archive is the candidate, not
  a by-product. What is removed is the second serialization that used to hash
  it. `build_candidate` returns the candidate, the revision its own capture
  produced, and — only when the reopened package reports
  `is_unmodified_owned_source()` — the sealed archive revision. Owned ingress
  retains the archive it was opened from and `PackageWriter::write_counted`
  republishes those bytes verbatim, so the digest taken while the archive was
  built is exactly what a serializing hash sink reads back. A reopen that did
  not retain that authorization returns `None` and is recomputed.
- `candidate_physical_revision(candidate, limits, known)` consumes that
  `Option` at the position the recomputation occupied, so the unknown-member
  refusal still runs in the same order.
- `apply_plan` and `apply_patch` keep the four revisions they compute for their
  staleness proofs. The two semantic revisions go to `capture_with_revision`
  (change 0590) instead of being recomputed inside `capture`; the two physical
  revisions seed the resulting snapshots, so the replan inside the same call
  reuses them. Every comparison against the plan or the patch happens exactly
  where it did, on values computed from the **live** packages; the caches only
  ever cover the immutable clones that `capture` produced.
- `validate_application_candidate` reports whether it rebuilt the candidate from
  the destination. `published_archive_revision` recomputes for a rebuilt
  candidate and, for the candidate that reached publication unchanged, returns
  the revision the call already proved — the value the final comparison would
  have recomputed from the same bytes.

Three `debug_assert!`s re-derive the reused values in test and debug builds:
in `snapshot_physical_revision`, `remember_physical_revision` and
`published_archive_revision`, plus the one `candidate_physical_revision` adds.

## Why it is sound

**ADR 0003 and 0006: no revision value and no proof format changed.** The
semantic revision is still `package_fingerprint` and the physical revision is
still SHA-256 over `litchi-pptx-cross-physical-v2`, the `u64` little-endian
archive length and the archive digest. The durable `LPCP0002` encoding carries
six 32-byte revisions and is bit-identical; a `CrossSlideCopyPatch` serialized
before this change still applies after it, and one serialized after it still
applies to the base. No published byte changes.

**ADR 0005: the cache is semantically invisible.** ADR 0005 states that "cache
behavior is semantically invisible". The memo is keyed on the archive bound it
was taken under, so the `Error::Limit` a smaller bound raises is unaffected; it
lives on the immutable snapshot, never on a live package; it is never inherited
across `rebound_to`; and a miss is an ordinary recomputation. The bounded
candidate archive is still bounded by `max_patch_bytes` in exactly the same
writer — ADR 0005's bounded-memory rule is why `bounded_package_bytes` survives
rather than being replaced by a hashing sink: its `Vec` is the archive
`from_vec_reusing_payloads` reopens, and hashing it in place costs nothing extra
while removing a whole second pass.

**Every refusal is where it was.** The staleness proofs in `apply_plan` and
`apply_patch` are computed from the live `&OpcPackage` arguments, before any
capture, exactly as before. `reject_unknown_non_part_members` still runs on
every cached read and on the candidate. The `fresh` versus `plan` comparison,
`validate_before`, `validate_after`, `validate_candidate`'s revision check and
the final published-archive comparison all run in the same order on the same
values. The one comparison that becomes tautological — the final
published-archive revision for a candidate that was not rebuilt — is a
comparison of a value against itself: `prepare` computed it from that exact
package and the call already refused unless it equalled the plan's.

**0454's rule is untouched.** Nothing about "dedup only with proven
equivalence" changes: the copy still copies the source closure byte-for-byte,
still refuses to reuse a destination member on byte equality alone, and still
proves the layout inheritance graph before reusing a layout.

**No contract moved.** No new `unsafe`, no weakened limit or defence, no hidden
global pool, no ambient I/O, no public leakage of archive types, locks or
executors, and no public API change: `Snapshot`'s new field is private. The
optimization-order step is 1 (eliminate unnecessary work) for the reused
revisions and 2 (unnecessary I/O, serialization and hashing) for the removed
serializations.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0; every measured process pinned to CPU 21 while other agents of
the same wave built and measured on the other cores. Both legs built `--release
--locked` from `tools/perf-baseline`: before from the shared read-only checkout
of `f22f93935`, after from this branch.

Two selectors, both `Snapshot::plan_cross_slide_copy` followed by
`Package::apply_cross_slide_copy_plan`: `pptx_cross_copy_plain` (a 30 KB
destination, one copied part) and `pptx_cross_copy_media_rich` (a 16.8 MB
destination and nine copied parts totalling 16,784,397 bytes of incompressible
media). The `_lifecycle` variants time the same work plus setup and reopen.

### Deterministic counts (callgrind isolation pair, `--warmup 0`)

`--samples 1` and `--samples 3`, differenced over the two extra samples. Both
selectors produce identical counts, so one table serves for both. A "lifecycle"
is one plan plus one apply.

| per lifecycle | before | after |
| --- | ---: | ---: |
| `package_fingerprint` calls (semantic revision) | 12 | **8** |
| `physical_package_fingerprint` calls (serialized-archive revision) | 9 | **4** |
| `snapshot_physical_revision` calls | 0 | 4 (2 hits, 2 misses) |
| `bounded_package_bytes` calls | 2 | 2 (unchanged) |
| `capture_internal` calls | 8 | 8 (unchanged) |
| `PackageWriter::write_to_stream` calls | 12 | **7** |
| `prepare_cross_slide_copy_for_slides` calls | 2 | 2 (unchanged) |

The five removed serializations are, per lifecycle: the source and destination
archives at the replan inside `apply_plan` (two, now cache hits), the candidate
archive at each of the two plans (two, now hashed while the archive is built),
and the published candidate's archive at the end of `apply_plan` (one, now the
value the call already proved). The four removed semantic hashes are the two
`capture` re-hashes inside `apply_plan` and the two `target_revision`
recomputations of a candidate its own capture had just hashed.

### Instructions

| per lifecycle | `pptx_cross_copy_plain` | `pptx_cross_copy_media_rich` |
| --- | ---: | ---: |
| instructions | 267,299,098 → **242,266,452** (**−9.37%**) | 35,610,254,229 → **26,847,795,508** (**−24.61%**) |
| `package_fingerprint` inclusive | 59,182,098 → 39,466,336 (−33.31%) | 15,787,570,855 → 10,525,038,674 (−33.33%) |
| `physical_package_fingerprint` inclusive | 14,724,276 → 6,500,081 (−55.85%) | 10,501,036,338 → 3,501,942,793 (−66.65%) |
| `bounded_package_bytes` inclusive | 12,846,889 → 16,154,652 (+25.75%) | 6,181,109,826 → 9,680,023,702 (+56.61%) |
| `PackageWriter::write_to_stream` inclusive | 27,432,146 → 22,563,384 (−17.75%) | 16,709,918,068 → 13,209,783,239 (−20.95%) |
| software `sha2::sha256::compress256` | 75,154,750 → 50,665,038 (−32.59%) | 28,037,044,392 → 19,274,528,456 (−31.25%) |
| `zlib_rs::deflate::deflate` | 2,494,129 → 2,494,166 (**+37 Ir**) | 5,955,591,941 → 5,955,602,461 (**+10,520 Ir in 5.96 G**) |

Whole child at `--samples 1`: `pptx_cross_copy_plain` 4,496,073,687 →
4,389,546,486 (−2.37%); `pptx_cross_copy_media_rich` 158,965,107,851 →
123,862,232,137 (**−22.08%**). Most of the whole child is the harness's own
corpus construction and its ten untimed refusal gates.

**`bounded_package_bytes` grows, and that is the change working.** Its
`BoundedVecWriter` now digests the bytes it accepts, so the candidate's archive
hash is *inside* it rather than in a second `to_stream`. The arithmetic closes
exactly on the media-rich corpus, at the 52.1 `Ir`/byte software SHA-256 rate
this build shows: the 3.50 G it gains is the two candidate hashes (2 × 33.6 MB)
that moved in; `physical_package_fingerprint` loses 7.00 G, of which 3.50 G
moved and 3.50 G — the two cached source/destination reads at the replan
(33.6 MB) and the published candidate's own recomputation (33.6 MB) — is gone.
Semantic hashing loses 5.26 G, which is about 101 MB at the same rate: the two
captures' re-hash of the source and destination graphs and the two candidate
`target_revision` recomputations of a graph its own capture had just hashed. 5.26 + 3.50 = 8.76 G, and the per-lifecycle
total falls by 8,762,458,721 `Ir`. Nothing is unaccounted for.

**Deflate is the identity cross-check.** The candidate serialization deflates
the copied media; that work is untouched, and the two legs differ by 10,520
instructions in 5.96 G per lifecycle (0.0002%). The published archive digest is
identical in all four timing legs of all four selectors
(`3e9ae2805a5d…` for the plain corpora, `6a3536fc7d90…` for the media-rich
ones).

### Paired native timing

Four legs in run order **A1 B1 B2 A2** (before, after, after, before), 30
measured samples after 3 warmups each, `taskset -c 21`. A/A is A2 against A1 and
B/B is B2 against B1, both inside the same window.

| selector | before p50 | after p50 | p50 Δ | p95 Δ | p99 Δ | A/A p50 | B/B p50 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_cross_copy_plain` | 8.249 ms | 7.849 ms | **−4.84%** (before is +5.09% of after) | −4.82% | −4.77% | +2.29% | −1.19% |
| `pptx_cross_copy_media_rich` | 739.602 ms | 651.978 ms | **−11.85%** (before is +13.44% of after) | −11.69% | −11.52% | **−4.99%** | +2.96% |
| `pptx_cross_copy_plain_lifecycle` | 9.547 ms | 9.426 ms | −1.26% | −0.32% | −0.30% | +1.97% | −4.68% |
| `pptx_cross_copy_media_rich_lifecycle` | 754.305 ms | 678.257 ms | **−10.08%** (before is +11.21% of after) | −9.10% | −9.08% | −4.27% | −4.89% |

**The floor is large in this window and is not hidden.** The media-rich A/A
floor is −4.99% at p50 — the two before legs, minutes apart in the same window
while other agents built, differ by almost 5%, just under the briefing's 5%
review threshold — and
the plain lifecycle's B/B floor is −4.68%. Only two of the four wall-clock
results clear their own floor by a comfortable margin
(`pptx_cross_copy_media_rich` at −11.85% against −4.99%, and its lifecycle
variant at −10.08% against −4.27%/−4.89%), and `pptx_cross_copy_plain` at
−4.84% is inside the noise of its own selector. **The counts and the
instruction differentials, which are exact, are what this record rests on.**

The selectors' own phase clocks localize the effect, because plan and apply are
timed separately:

| phase | `pptx_cross_copy_media_rich` before p50 | after p50 | Δ |
| --- | ---: | ---: | ---: |
| plan | 329.537 ms | 307.957 ms | −6.55% |
| **apply** (reported as `commit`) | **403.226 ms** | **336.783 ms** | **−16.48%** |
| publication (`to_stream` of the result) | 6.247 ms | 6.437 ms | +3.04% |
| reopen | 21.576 ms | 21.883 ms | +1.42% |

Apply loses three of its seven complete-graph hashes and four of its seven
archive serializations — half of its fourteen passes — and moves −16.5%; plan
loses one of its two hashes and one of its four serializations and moves −6.6%;
publication and reopen contain no changed code and move +3.0% and +1.4%, inside
their legs' spread. (Plan four plus apply seven serializations is eleven; the
twelfth in the count table is the harness's own publication `to_stream`, and
three of the twelve semantic hashes are the harness's own
`opened_presentation()` calls on the source, the destination and the reopened
output.) The same shape holds on
`pptx_cross_copy_media_rich_lifecycle` (apply −15.32%, plan −4.80%) and on the
plain corpus (apply −5.62%, plan −3.64%, reopen +0.13%).

**Native `perf stat`**, whole child, the same A1 B1 B2 A2 order over all four
selectors in one process at `--warmup 1 --samples 10`:

| event | before mean | after mean | Δ | A/A | B/B |
| --- | ---: | ---: | ---: | ---: | ---: |
| cycles | 114,125,032,514 | 104,551,234,104 | **−8.39%** | −0.36% | +1.90% |
| instructions | 240,486,326,936 | 228,814,083,925 | −4.85% | −0.57% | +1.23% |

This is the SHA-NI counterweight the briefing asks for: callgrind prices
SHA-256 in software and puts it at 78.6% of the media-rich whole child, where
natively the same work is worth −8.39% of whole-child cycles against a −0.36%
A/A floor. The whole child includes corpus construction and ten untimed refusal
gates per selector, so it dilutes rather than isolates.

### What `build_candidate` does with the copied media

0587 asked whether `build_candidate` loses physical provenance and re-deflates
the media. **It does not lose provenance**: the candidate serialization goes
through `soapberry_zip::preserve::PreservationIndex`, which copies the
destination's own unchanged members from its retained archive. What it deflates
is the nine *copied* parts — 16,784,397 bytes that are new members of the
destination archive and have no preserved image there. **Measured** at 5.96 G
`Ir` per lifecycle over two candidate builds, or about 178 `Ir` per copied byte,
exactly the rate for 16.78 MB rather than the 33.6 MB the whole candidate
carries. It is therefore not a provenance bug; it is one deflate of the copied
closure per candidate build, and the copy builds the candidate twice per
lifecycle (once to plan, once to re-prove at apply). Removing the second
deflation means retaining the planned candidate archive in the plan, which
changes `CrossSlideCopyPlan`'s memory profile under ADR 0005 and is left to a
design record.

## Correctness evidence

Six tests were added in a new
`crates/litchi-pptx/src/opened/cross_copy_plan/revision_cache_tests.rs`, and one
in the existing `bounded_writer_tests.rs`.

| test | what it pins |
| --- | --- |
| `snapshot_physical_revision_caches_the_freshly_computed_value` | The cache starts empty, the first read equals a direct `physical_package_fingerprint`, the second read returns the same value, a fresh recomputation still agrees, and a snapshot clone shares the same package and the same cache. |
| `snapshot_physical_revision_is_keyed_by_the_archive_bound` | A different `max_patch_bytes` is a miss and recomputes the same digest — the value does not depend on the bound, only the refusal does — and a bound of one byte still raises `Error::Limit { resource: "cross-slide serialized archive bytes", .. }` with a warm cache, after which the ordinary bound still returns the cached value. |
| `rebound_snapshot_starts_an_empty_physical_revision_cache` | Two packages with the same graph and different retained archives (one carries a ZIP end-of-central-directory comment) have equal semantic revisions and satisfy `packages_equal`, but different physical revisions. `Snapshot::rebound_to` onto the second starts an empty cache, and the rebound snapshot computes the second package's own revision, not the first's. This is the test that fails if the cache were inherited. |
| `planned_and_published_revisions_equal_fresh_serializations` | A plan's `source_physical_revision` and `destination_physical_revision` equal fresh fingerprints of the live packages, and after `apply_cross_slide_copy_plan` the plan's `target_physical_revision` and `target_revision` equal a fresh physical fingerprint and a fresh `package_fingerprint` of the published package. This is the test that fails if the archive hashed while it was built were not what the reopened candidate republishes. |
| `warm_snapshot_caches_do_not_hide_a_stale_source_or_destination` | Planning twice from the same two snapshots warms both caches and produces identical revisions; a destination whose slide XML then drifts is refused with the exact `apply_cross_slide_copy_plan` / "the complete destination package graph changed after cross-slide planning" pair and is left byte-identical; a source with the same graph but a different retained archive reaches the physical proof and is refused with "the serialized source package changed after cross-slide planning"; and the unchanged pairing still applies and still publishes the planned archive revision. |
| `candidate_physical_revision_without_a_known_value_recomputes` | The fallback path equals the recomputation, the known-value path returns it, and a seeded snapshot returns the value it would have computed for itself. |
| `bounded_writer_digests_exactly_the_accepted_bytes` | The writer's digest is SHA-256 of exactly the bytes it accepted; a write the archive bound refuses never reaches the digest. |

Four `debug_assert!`s re-derive every reused value in debug and test builds —
in `snapshot_physical_revision` (on each cache hit), `remember_physical_revision`
(on each seed), `candidate_physical_revision` and `published_archive_revision` —
so every one of the 867 `litchi-pptx` tests that reaches a cross-package copy
re-serializes and re-hashes the package and compares. The pre-existing suites
`source_backed_cross_copy.rs` and `source_backed_cross_copy_adversarial.rs`
exercise the stale-source, stale-destination, foreign-source, forged-patch,
collision-remap, limit and reversibility routes through the same code: 58 and 7
tests, all passing.

**Gates** (run in the worktree; tails in
[`results/change-0598/gates.txt`](results/change-0598/gates.txt)):

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-pptx --all-targets` | clean (workspace lints are deny) |
| `cargo test -p litchi-pptx` | **867 passed, 0 failed** (lib 563, integration and doc tests the rest) |
| `cargo doc -p litchi-pptx --no-deps` | clean (rustdoc lints deny) |
| `tools/check_perf_claims.py --mode strict` | `OK: 10 performance claims validated (strict)` — this record adds no entry |
| `tools/check_report_claim_classification.py` | `OK: 167 REPORT rows classified` |

**Differential check on published bytes.** All four measured selectors verify
their published archive against an expected digest and report it; the reported
`output_sha256` is one single value per corpus across all four timing legs
(`3e9ae2805a5d…`, `6a3536fc7d90…`), and the callgrind profiles show
`zlib_rs::deflate::deflate` within 10,520 instructions of 5.96 G per lifecycle.
The durable `CrossSlideCopyPatch` round trip is one of the ten untimed corpus
gates the harness refuses to run without.

## Validation preserved

No validation was removed, weakened, reordered or made conditional.
`snapshot_physical_revision` and `candidate_physical_revision` still run
`reject_unknown_non_part_members` on every call, in the position
`physical_package_fingerprint` ran it. `capture_with_revision` (change 0590)
runs every structural, limit and ADR 0013 notes-topology check unconditionally
and skips only the hash of bytes the caller already hashed. The four staleness
proofs at the top of `apply_plan` and `apply_patch` are unchanged and are taken
from the live packages. `validate_candidate` still captures the candidate
itself and still compares its revision against the plan's `target_revision`.
The bounded candidate archive is still refused above `max_patch_bytes` by the
same writer, and a physical fingerprint taken under a smaller bound still
raises `Error::Limit` rather than reading a memo.

## Limitations

- **Not claimed:** any speedup. `performance_claim: none`, no claim-registry
  entry, and the paired timings are reported beside the A/A and B/B floors
  rather than as a result. The floors in this window are large — the media-rich
  A/A floor is −4.99% at p50 — so `pptx_cross_copy_plain` (−4.84%) and
  `pptx_cross_copy_plain_lifecycle` (−1.26%) are inside their own noise and are
  reported, not relied on.
- The counts and instruction differentials are exact for these two selectors on
  their fixed synthetic corpora, on this host and these two builds. They are not
  a statement about real PowerPoint decks: both corpora are generated by
  `litchi-pptx-cross-slide-copy-evidence-v1` and the media-rich one is
  deliberately incompressible.
- Callgrind runs SHA-256 in software because valgrind masks the SHA CPUID bit,
  so every instruction share attributed to hashing here is roughly five times
  its native cycle share. The `perf stat` row is the counterweight and it is
  whole-child, so it dilutes the timed effect rather than isolating it.
- The external-package copy binary (`tools/perf-baseline/src/bin/pptx_external_cross_copy.rs`)
  was **not** run. Its pinned LibreOffice QA fixture is not in the tree: change
  0454 fetched it over the network to a temporary path and the artifact remains
  outside tracked source. No external-fixture result is offered.
- **What remains, measured.** Four complete-package semantic hashes and four
  archive serializations still run per lifecycle. The largest remaining
  duplication is that the whole candidate is built, serialized and captured
  **twice** per lifecycle — once to plan and once to re-prove at apply — which
  costs the two candidate deflations of the copied closure (5.96 G `Ir` per
  lifecycle, untouched here) and two of the four remaining captures. Retaining
  the planned candidate archive in `CrossSlideCopyPlan` would remove the second,
  but it changes the plan's memory profile under ADR 0005 and needs a frozen
  design record.
- `validate_candidate` captures the candidate a second time after
  `build_candidate` already captured it; reusing that capture means threading a
  proven snapshot into `validate_candidate`, whose revision check is the proof
  that binds the candidate to the plan. That is a contract question, not a
  value-identical reuse, and it is left open rather than taken here.
- 0587's item PPTX-1(c) — memoizing per-part digests and redefining the
  complete-package revision as a hash over them — would cut the four remaining
  semantic hashes, and change 0590 already recorded why it needs a frozen design
  record and a magic bump for `LPRM0001` and `LPCP0002` first. Nothing in this
  record moves that.
- No allocation, RSS, cold-cache, range-source, concurrency or cross-platform
  measurement was taken, and no source-backed (`SourceBackedPresentationEditor`)
  cross-copy selector exists to measure.

## Retained evidence

[`results/change-0598/README.md`](results/change-0598/README.md) — both legs'
callgrind annotations, call counts and isolation pairs, every paired timing and
`perf stat` report, the gates, the decision record and the log paragraphs.
