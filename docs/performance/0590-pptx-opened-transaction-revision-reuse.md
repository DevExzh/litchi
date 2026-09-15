# 0590: one complete-package hash per opened PPTX commit, and the publication stops re-deriving the snapshot it was handed

Status: retained. `performance_claim: none` — this record carries exact
callgrind call counts and instruction differentials from an isolation pair, and
paired native timings reported beside the host's A/A floor. No claim is
registered.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements parts (a) and (b) of item **PPTX-1** of change
[0587](0587-remaining-opportunity-survey.md) (rank 3). Part (c) is written up
below as a design note and is **not** implemented; part (d) was examined and
rejected on inspection.

## The mechanism 0587 found

`litchi-pptx`'s opened-presentation lifecycle hashed the complete package —
every part blob, media included — four times, and ran the slide-name and notes
scans three times. 0587 measured `package_fingerprint` at 34.9% of the profiled
instructions of `pptx_eager_batch_edit_save`; this record reproduces that number
exactly (13,840,168,680 of 39,612,371,219 `Ir`, 34.94%).

The four sites, in order:

1. `capture_with_provenance` (`crates/litchi-pptx/src/opened/model.rs:308`) at
   `opened_presentation_transaction()`.
2. `Transaction::commit` (`opened/transaction.rs:1198`) hashes the staged
   package to decide whether to call `unsign()`.
3. `Transaction::commit` then calls `capture` on that same staged package,
   which hashes it again (`transaction.rs:1213` at the base commit).
4. `Package::apply_opened_presentation_commit` (`package/model.rs:413`)
   discarded `commit.snapshot` entirely, kept only the patch and called
   `opened::apply`, which captured — and hashed — the candidate a fourth time.

Site 3 hashes bytes that site 2 had just hashed, unless `unsign()` changed them
in between. Site 4 re-derives a snapshot the commit had already derived and
validated on content the patch is about to reproduce.

## What was changed

**`crates/litchi-pptx/src/opened/transaction.rs`** — `Transaction::commit` keeps
the revision it computed for the `unsign()` decision and hands it to the
recapture:

```rust
let mut revision = package_fingerprint(&working)?;
if revision != self.source.revision {
    let signed = working.is_signed();
    working.unsign();
    if signed {
        revision = package_fingerprint(&working)?;
    }
}
```

`OpcPackage::unsign` (`crates/litchi-opc/src/package.rs:942`) calls
`strip_signature_graph`, which removes signature-infrastructure parts, the root
relationships that are or reach signatures, and the same relationships inside
every part. `is_signed` (`package.rs:780`) is true whenever any of those exist,
and is also true for the one case stripping does not touch — an opaque non-part
member under the reserved signature directory. It is therefore a conservative
predicate for "unsign may rewrite a fingerprint input": when it is false,
nothing the fingerprint feeds can change, and when it is true the package is
re-hashed exactly as before. `unsign`'s other effects — revoking exact-source
authorization and clearing the signature edit policy — are package state the
fingerprint has never covered and are unchanged by this record.

**`crates/litchi-pptx/src/opened/model.rs`** — `capture_with_provenance` is
split into `capture_internal` plus a new `capture_with_revision`, which takes a
revision the caller computed from identical content and skips only the repeated
hash. Every validation the ordinary capture runs — the slide-reference
resolution, the one-to-one identity checks, the relationship-type checks, the
per-slide name parse, `notes::load_snapshot` and the name-index build — still
runs, in the same order, before the revision is taken. A `debug_assert!`
re-hashes and compares in test and debug builds.

The same file gains `packages_equal`, a direct comparison of exactly the inputs
`package_fingerprint` feeds: part count, and per part the part name, content
type, payload and relationships; the package-root relationships; and the opaque
non-part members. Payload comparison short-circuits on `Arc` pointer identity,
which the opened path preserves for every part an edit did not rewrite. It
proves the two packages carry the same revision *by content*, without hashing
and without relying on collision resistance. A false result proves nothing, so
every caller falls back to the ordinary capture.

`Snapshot::rebound_to` rebinds a snapshot's derived state — presentation root,
slide identities, name index, revision, limits, provenance — onto a
content-identical package, taking its own `Arc<OpcPackage>` clone exactly as
`capture` does.

**`crates/litchi-pptx/src/opened/patch.rs`** — `apply_with_revision` gains an
optional already-captured snapshot, consulted only at the point where it would
otherwise call `capture` on the finished candidate. `apply_committed` is the
entry point that supplies one; `apply` and `apply_exact_revision` pass `None`
and are byte-for-byte the routes they were.

**`crates/litchi-pptx/src/package/model.rs`** —
`apply_opened_presentation_commit` stops discarding `commit.snapshot` and passes
it to the shared publication helper. `apply_opened_presentation_patch` is now a
one-line call into that same helper with `None`; every policy refusal keeps the
`apply_opened_presentation_patch` operation identity both routes reported
before.

## Why it is sound

**ADR 0003's revision binding is untouched.** The revision is still
`package_fingerprint` of the complete package: the same domain string, the same
sorted part order, the same per-part name, content type, payload and
relationship feed, the same root relationships and non-part members. No value
changed, so the durable encodings that carry revisions — `LPRM0001`
(`opened/remove_plan.rs:26`) and `LPCP0002` (`opened/cross_copy_plan.rs:26`) —
are bit-identical to what the base produces, and a patch serialized before this
change still applies after it. ADR 0003 says a `commit()` returns "a named
`Commit<T>` containing the new snapshot, a reversible patch, and diagnostics";
part (b) makes the PPTX facade *use* that snapshot instead of throwing it away.

**The signature policy is preserved.** An `unsign` that changes bytes is
re-hashed. The new test
`commit_rehashes_a_package_whose_unsign_strips_signature_bytes` builds a package
with a signature-origin part and its root relationship, commits an edit through
it, and asserts the committed revision equals a fresh fingerprint of the
stripped package and differs from the source revision; after publication the
origin part is gone, the package is no longer signed, and the published
revision equals a fresh fingerprint of the facade's own package.

**ADR 0013's notes topology still gates every mutation.**
`notes::load_snapshot` still runs inside every capture, and the capture that is
reused is the one `Transaction::commit` already ran on the staged package before
the patch existed. Part (b) never skips a capture of content that was not
captured; it declines to capture the *same* content twice.

**Every refusal is where it was.** The candidate is built, `validate_before`
runs before it, `validate_after` runs after it, and the assignment to the
package happens in the same place. The reuse is consulted after the candidate
exists and only replaces a `capture` call whose inputs are proven identical. If
the destination package moved on outside the patch write set — the case
`apply_opened_presentation_commit` has always tolerated, because it applies a
patch rather than an exact-revision plan — `packages_equal` is false and the
candidate is captured from scratch. The new test
`commit_publication_captures_a_package_that_drifted_outside_the_write_set`
exercises exactly that and asserts the published snapshot describes the drifted
package, not the committed one.

**No contract moved.** No output byte changes, no limit is relocated, no
validation is skipped or reordered, no `unsafe` is added, no lock, executor or
archive type is exposed, and no I/O is performed. The optimization-order step is
1: eliminating unnecessary work.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0; every measured process pinned to CPU 10 while seven other
agents built and measured on the other cores. Both legs built `--release
--locked` from `tools/perf-baseline`: before from the shared read-only checkout
of `08d968f8e`, after from this branch.

### Deterministic call counts (callgrind, `--warmup 0 --samples 1`)

Both legs, `callgrind`, pinned to CPU 10. `capture` is
`capture_with_provenance` before and `capture_internal` after; the two are the
same function body. Counts are whole-child: the harness builds and verifies its
corpus in the same process, so a selector's child runs more lifecycles than it
times.

| selector | `package_fingerprint` calls | `capture` calls | `packages_equal` / reused | whole-child `Ir` |
| --- | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 16 → **10** | 13 → **10** | 3 / 3 | 39,612,371,219 → **33,611,197,006** (−15.15%) |
| `pptx_eager_multi_slide_batch_edit_save` | 16 → **10** | 13 → **10** | 3 / 3 | 39,684,224,947 → **33,684,642,197** (−15.12%) |
| `pptx_slide_remove_boundary_save` | 139 → **126** | 97 → **91** | 6 / 6 | 4,918,071,028 → **4,829,566,830** (−1.80%) |
| `pptx_slide_move_boundary_save` | 120 → **102** | 88 → **80** | 8 / 8 | 4,741,071,794 → **4,619,030,888** (−2.57%) |

`packages_equal` does not exist in the before leg. In the after leg every one of
its calls was followed by a `Snapshot::rebound_to` call, so no comparison failed
and no fallback capture ran during these profiles.

Inclusive symbol shares, `pptx_eager_batch_edit_save`, whole child:

| symbol | before | after |
| --- | ---: | ---: |
| `sha2::sha256::compress256` | 20,506,872,843 (51.77%) | 14,990,136,045 (44.60%) — **−26.90%** |
| `opened::model::package_fingerprint` | 13,840,168,680 (34.94%) | 8,322,478,025 (24.76%) — **−39.87%** |
| the capture entry point | 13,192,681,206 (33.30%) | 7,186,298,675 (21.38%) — **−45.53%** |
| `litchi_perf_baseline::sha256_hex` (the harness's own output check) | 3,543,866,958 | 3,543,866,958 — **identical** |
| `zlib_rs::deflate::deflate` | 9,395,198,020 | 9,395,191,418 — **−6,602 Ir (0.00007%)** |

The last two rows are the cross-check: the harness's own verification hashing is
bit-identical and the deflate work is stable to 6,602 instructions in 39.6 G, so
what was removed is hashing and capturing, not anything the selector produces.

### Instructions (callgrind isolation pair)

`pptx_eager_batch_edit_save` profiled at `--samples 1` and `--samples 3`,
differenced over the two extra samples:

| per lifecycle | before | after |
| --- | ---: | ---: |
| instructions | 12,679,371,772 | **10,680,374,116** (−1,998,997,656, **−15.77%**) |
| `package_fingerprint` calls | 6 | **4** (−33.3%) |
| `capture` calls | 5 | **4** (−20.0%) |
| `packages_equal` calls | 0 | 1 |

The same pair on the two boundary selectors, which time one lifecycle per
sample:

| per lifecycle | `pptx_slide_move_boundary_save` | `pptx_slide_remove_boundary_save` |
| --- | --- | --- |
| instructions | 61,284,176 → **47,252,677** (**−22.90%**) | 97,318,009 → **96,518,458** (−0.82%) |
| `package_fingerprint` calls | 5 → **3** | 9 → **9** (unchanged) |
| `capture` calls | 4 → **3** | 6 → **6** (unchanged) |
| `packages_equal` calls | 0 → 1 | 0 → **0** |

`pptx_slide_remove_boundary_save` is the control this batch did not plan for:
its per-sample lifecycle reaches neither changed function, so its counts are
identical on both legs and its instruction difference is under 1%. Its six
`packages_equal` calls are constant between `--samples 1` and `--samples 3`,
which places them in the selector's untimed verification gates. Read its
wall-clock row below against that.

**Why six and not 0587's four.** 0587 counted the canonical lifecycle: capture,
`unsign()` decision, commit recapture, publication. `pptx_eager_batch_edit_save`
adds two more publications per lifecycle, replaying the patch's inverse and then
the inverse's inverse to prove reversibility, and each of those receives a bare
`Patch` with no commit behind it. Part (a) removes the commit recapture's hash;
part (b) removes the hash and the scans of the publication that does have a
commit behind it. The two bare-patch replays keep theirs, which is why four
remain rather than two. `pptx_slide_move_boundary_save`, which publishes once,
goes from five hashes to three.

### Paired native timing

Four legs per selector in run order **A1 B1 B2 A2** (before, after, after,
before), 30 measured samples after 3 warmups each, `taskset -c 10`, seven other
agents active on the other cores. The A/A floor is A1 against A2 and the B/B
floor is B1 against B2, both taken inside the same window.

| selector | before p50 | after p50 | p50 Δ | p95 Δ | p99 Δ | A/A p50 floor | B/B p50 floor |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 345.818 ms | 319.714 ms | **−7.55%** (before is +8.16% of after) | −8.20% | −9.49% | +0.16% | −0.29% |
| `pptx_eager_multi_slide_batch_edit_save` | 347.596 ms | 322.447 ms | **−7.24%** (before is +7.80% of after) | −7.26% | −7.43% | −0.46% | −0.40% |
| `pptx_slide_move_boundary_save` | 0.837 ms | 0.515 ms | **−38.55%** (before is +62.72% of after) | −38.48% | −38.41% | +0.25% | +1.11% |
| `pptx_slide_remove_boundary_save` | 2.004 ms | 2.046 ms | **+2.13%** (before is −2.08% of after) | +3.30% | +0.04% | +0.45% | +2.65% |

**The one selector that got worse.** `pptx_slide_remove_boundary_save` is
+2.13% at p50 — below the 5% review trigger, above its A/A floor and below its
B/B floor. Its timed region contains none of the changed code: the selector
plans with `Snapshot::plan_slide_removal` and publishes with
`Package::apply_slide_removal_plan`, which routes to `opened::apply_removal_patch`
and never touches `Transaction::commit` or `apply_opened_presentation_commit`.
Its call counts moved only because the selector's *untimed* verification gates
run opened transactions. The selector's own phase clocks say the same thing:

| `pptx_slide_remove_boundary_save` phase | before p50 | after p50 | Δ |
| --- | ---: | ---: | ---: |
| plan | 603.6 µs | 618.6 µs | +2.48% |
| commit | 1,249.4 µs | 1,275.4 µs | +2.08% |
| publication | 153.9 µs | 155.7 µs | +1.14% |
| reopen | 690.2 µs | 694.1 µs | +0.57% |

Every phase drifts up by a similar amount, including publication, which
contains no changed code, no capture and no hash at all. A second, independent
four-leg run at 100 samples after 5 warmups reproduced the direction and the
ambiguity: +2.69% at p50 against a B/B floor of −2.67% and an A/A p95 floor of
+4.60%. The deterministic counts settle it: this selector's timed lifecycle
issues exactly 9 hashes and 6 captures on both legs and reuses nothing, and its
per-lifecycle instruction count moves −0.82%. The regression is real in this
window, below the 5% review trigger, and not attributable to the change; it is
reported rather than folded into a mean.

**The one selector that isolates the changed region.**
`pptx_slide_move_boundary_save` times the same four phases, and its commit phase
is exactly `Transaction::commit()` followed by
`Package::apply_opened_presentation_commit`:

| `pptx_slide_move_boundary_save` phase | before p50 | after p50 | Δ |
| --- | ---: | ---: | ---: |
| plan | 9.1 µs | 8.9 µs | −1.79% |
| **commit** | **714.4 µs** | **391.2 µs** | **−45.24%** |
| publication | 112.3 µs | 111.9 µs | −0.37% |
| reopen | 710.9 µs | 712.1 µs | +0.17% |

This is the cleanest attribution in the batch: the phase that contains the
change moves 45%, the three phases that do not move by under 2%, and the
selector's per-lifecycle instruction count falls 22.90% with two of its five
hashes and one of its four captures gone. The wall-clock fall (−38.55%) is
larger than the instruction fall (−22.90%), which is the direction
`GOAL_AUDIT.md`'s standing note predicts for this kind of work: a capture is a
pointer chase through the package graph, priced high in cycles and low in
instructions, next to a lifecycle whose remaining bulk is deflate and copying.

**Native `perf stat`**, same four legs, whole child:

| selector | cycles before | cycles after | cycles Δ | A/A cycles | instructions Δ |
| --- | ---: | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 74,732,115,238 | 70,892,966,932 | **−5.14%** | +0.39% | −4.18% |
| `pptx_eager_multi_slide_batch_edit_save` | 75,688,676,368 | 71,581,657,626 | **−5.43%** | +1.07% | −4.31% |
| `pptx_slide_move_boundary_save` | 1,780,125,550 | 1,730,468,213 | −2.79% | −1.64% | −4.69% |
| `pptx_slide_remove_boundary_save` | 1,971,181,751 | 1,996,882,156 | +1.30% | +0.20% | +0.23% |

These are whole-child counts over 30 samples plus 3 warmups plus corpus
construction plus every untimed verification gate, so they dilute the timed
effect rather than isolating it; the harness's own p50 above is the timed
figure. They are reported because the briefing asks for native cycles beside
every callgrind hashing share, and because they independently confirm the
direction and the order of magnitude on the two eager selectors.

**Published bytes are identical.** All four selectors verify their output
against an expected digest and report it; the reported `output_sha256` is the
same value in all four legs of all four selectors
(`8371618225b8…`, `23d6d7b8dd43…`, `a23546ab442d…`, `57734a7dfe4d…`).

### The eager-corpus caveat, on every timing above

The eager PPTX harness corpora are built through
`SourceBackedPackage::from_read_at(...).into_opc_package()` and
`Package::from_opc_package`, which sets `physical_source_provenance = false`
(`crates/litchi-pptx/src/package/codec.rs:262`). Every eager save in these
selectors therefore re-deflates all media instead of copying unchanged members,
which 0587 measured at roughly 42% of native cycles on
`pptx_eager_batch_edit_save`. That save dominates the timed region, which is
why the wall-clock improvement on the two eager selectors (−7.5% and −7.2% at
p50) is about half the instruction improvement (−15.8% per lifecycle): the
denominator is inflated by recompression a production `open` or `from_vec`
would not perform. A production package would show the same absolute saving
against a smaller total — a larger relative gain — but that is **modelled**, not
measured, and the `open`/`from_vec` corpus variant 0587 asks for has to exist
before anyone measures it. This caveat applies to every wall-clock number above
and is not repeated per row.

## Correctness evidence

Six tests were added to `crates/litchi-pptx/src/opened/tests.rs`.

| test | what it pins |
| --- | --- |
| `commit_rehashes_a_package_whose_unsign_strips_signature_bytes` | A commit whose `unsign()` changes bytes gets a fresh revision. The fixture attaches a `/_xmlsignatures/origin.sigs` part with the OPC digital-signature-origin content type and its root relationship; the committed revision equals a fresh fingerprint of the *stripped* package, differs from the source revision, and after publication the origin part is gone, `is_signed()` is false and the published revision equals a fresh fingerprint of the facade's own package. |
| `commit_revision_binds_the_staged_package_without_a_signature_graph` | The ordinary path: the reused revision is exactly `package_fingerprint` of the snapshot's own package, and differs from the source. |
| `published_commit_snapshot_equals_a_fresh_capture` | The snapshot handed to the caller after apply equals the one a fresh capture produces — revision, presentation root, slide identities, limits, provenance flag, complete package content, and name-index resolution for every slide — on a lifecycle that edits shape text, moves a slide and rewrites notes, and again on an exact no-op commit, which reaches the same reuse through the empty-patch path. |
| `commit_publication_captures_a_package_that_drifted_outside_the_write_set` | A part the patch never names is added between commit and publication; the published snapshot's revision differs from the committed one, equals a fresh fingerprint of the real package, contains the drifted part and matches a fresh capture field for field. This is the test that fails if `packages_equal` were ever incomplete. |
| `stale_write_set_still_refuses_a_committed_publication` | A stale-revision apply still conflicts: a slide inside the write set is rewritten, the commit is refused, and the package's own fingerprint is unchanged by the refusal. |
| `packages_equal_tracks_every_fingerprint_input` | Seven single-input mutations — payload, content type, added part, removed part, part relationship, root relationship, external relationship — each make `packages_equal` false in both directions *and* change `package_fingerprint`; and a part rewritten to its own bytes in a fresh allocation is still accepted, with an equal fingerprint. |

Both `capture_with_revision` and `Snapshot::rebound_to` carry a `debug_assert!`
that re-derives their precondition, so every one of the 860 tests that reaches
either of them re-hashes the package and compares, or re-runs the full package
comparison.

**Gates** (run in the worktree; tails in
[`results/change-0590/gates.txt`](results/change-0590/gates.txt)):

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-pptx --all-targets` | clean (workspace lints are deny) |
| `cargo test -p litchi-pptx` | **860 passed, 0 failed** (lib 556, integration and doc tests the rest) |
| `cargo doc -p litchi-pptx --no-deps` | clean (rustdoc lints deny) |
| `tools/check_perf_claims.py --mode strict` | `OK: 10 performance claims validated` — this record adds no entry |
| `tools/check_report_claim_classification.py` | `OK: 167 REPORT rows classified` |

**Differential check on published bytes.** All four measured selectors verify
their published archive against an expected digest and report it. The reported
`output_sha256` is identical in all four legs of all four selectors, and the
callgrind profiles show `zlib_rs::deflate::deflate` within 6,602 instructions of
39.6 G and `litchi_perf_baseline::sha256_hex` bit-identical — the published
bytes and the harness's own verification are unchanged.

## Validation preserved

No validation was removed, weakened, reordered or made conditional.
`capture_with_revision` skips only the recomputation of a hash whose value the
caller holds; every structural check, limit check and notes-topology check in
`capture_internal` runs unconditionally on the package being captured.
`capture_candidate` skips a whole capture only when `packages_equal` has proven
the candidate byte-identical to a package that was already captured
successfully in this same lifecycle — so every check that capture performs has
already been performed on this exact content, with the same limits and the same
provenance flag, both of which are compared before the reuse.

## Limitations

- **Not claimed:** any speedup. `performance_claim: none`, no claim-registry
  entry, and the paired timings are reported beside the A/A floor rather than as
  a result.
- The counts and instruction differentials are exact for the four measured
  selectors on their fixed synthetic corpora, on this host and this build. They
  are not a statement about real PowerPoint files: the eager corpora carry no
  physical provenance (above) and the boundary selectors are generated decks.
- Callgrind runs SHA-256 in software because valgrind masks the SHA CPUID bit,
  so every instruction share attributed to hashing here is roughly five times
  its native cycle share. The native timings are the counterweight and they are
  reported whether or not they clear the floor.
- The cross-package copy path (`pptx_cross_copy_*`, item PPTX-2) is untouched.
  Its plan and apply build their own snapshots and do not run through
  `apply_opened_presentation_commit`.
- `Presentation::slide(i)` still reparses `presentation.xml` per call (0587
  item PPTX-3), the eager save still regenerates every slide (SAVE-3), and the
  remaining full-package hashes stay: four per lifecycle of the eager selectors,
  three of `pptx_slide_move_boundary_save`.
- No allocation, RSS, cold-cache, range-source, concurrency or cross-platform
  measurement was taken.

## Item (c), not implemented: per-part digest memoization

0587's part (c) proposes memoizing per-part digests keyed by blob `Arc`
identity and length inside `Snapshot`, so that a recapture hashes only the parts
whose `Arc` changed, with the complete-package revision becoming a hash over the
sorted per-part digests and the `litchi-pptx-opened-v1` domain string bumped.

**It is a durable wire-format change and needs a frozen design record first.**
The value `package_fingerprint` returns is not private: it is serialized into
`SlideRemovalPatch` (`LPRM0001`, two 32-byte revisions in the header,
`opened/remove_plan.rs:107`) and `CrossSlideCopyPatch` (`LPCP0002`, six
32-byte revisions, `opened/cross_copy_plan.rs:26`), and publication compares a
stored revision against a freshly computed one. Redefining the hash makes every
previously serialized patch of both formats fail its revision check. A design
record has to state the magic bump, whether old patches are rejected with a
typed error or migrated, and what the new proof binds, before any code.

**Predicted saving, modelled** from the counts above. After (a) and (b) a
`pptx_eager_batch_edit_save` lifecycle still hashes the complete package four
times, at about 0.83 G Ir each under callgrind (8,322,478,025 Ir over 10 calls).
(c) would leave the first capture hashing everything once and reduce the other
three to the parts an edit actually rewrote — one slide part of a 200-slide
deck, and none of its 16 MiB of media — so about **2.5 G Ir per lifecycle, or
23% of the 10.68 G that remains**. Natively those three passes are worth roughly
a fifth of that share, because callgrind prices SHA-256 in software. The
`pptx_slide_move_boundary_save` commit phase, which this record already moved
from 714 µs to 391 µs, is where it would show first.

## Item (d), examined and rejected

0587's part (d) proposes building `SlideNameIndex` lazily and caching the notes
index per (presentation blob, slide blob) identity. Neither half is the small,
obviously safe change the brief allows.

- **Lazy name index.** `SlideNameIndex::build`
  (`opened/model.rs:167`) is a `HashMap` over already-parsed slide names with no
  XML work in it; the per-slide name *parse* that dominates happens earlier, in
  `slides()`, and fills a `Snapshot` field every caller receives. Deferring the
  map would also defer its two allocation failures out of capture, moving where
  a typed `Error::Allocation` is raised. Cost removed: near zero. Contract
  moved: yes. Rejected.
- **Notes index cache.** `notes::load_snapshot` is a free function over a
  borrowed `&OpcPackage`; there is no owner to hold a memo, so a cache means
  either ambient process state or a new field threaded through every capture
  call site. ADR 0013 makes the notes-topology check a precondition of every
  mutation touching notes, so the memo's invalidation rule is a correctness
  contract, not an implementation detail. That is a design record, not a small
  change. Deferred with its 0587 size (about 108 M Ir per capture on a
  200-slide deck, **retained**).

## Retained evidence

[`results/change-0590/README.md`](results/change-0590/README.md) — both legs'
callgrind annotations and call counts, the isolation pair, every paired timing
report, the gates, the decision record and the log paragraphs.
