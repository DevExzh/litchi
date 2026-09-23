# 0751 — the owned PPTX cross-copy stops re-hashing proven bytes: media-rich lifecycle p50 180.8 → 102.3 ms, commit 88.7 → 25.7 ms

Status: retained, implemented in `litchi-opc` and `litchi-pptx`.
`performance_claim: none` — no claim-registry entry; the paired medians,
counter ratios and allocator counts below are reported as evidence, not
registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `6d989cad63` (the tip of `feat/office-format-completeness`, with records
0742, 0744–0747). Production commits `4c7cbc4b8c` (`litchi-opc`) and
`f84522f8af` (`litchi-pptx`) on `perf/0751-pptx-cross-copy-apply-digest-reuse`.
Evidence packet: [`results/change-0751/`](results/change-0751/README.md).

## Result

On change 0742's media-rich pair (8 × 2 MiB incompressible PNGs per deck,
16.8 MB archives), both legs built with the identical command:

| case | before median p50 ms | after median p50 ms | median paired ratio [95% bootstrap] |
|---|---:|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 180.778 | 102.309 | 0.5637 [0.5366, 0.5849] |
| `pptx_cross_copy_media_rich` | 161.920 | 86.380 | 0.5242 [0.4755, 0.5431] |

- **Commit (application)** moves 88.691 → 25.674 ms (paired 0.2890). Its
  five SHA-256 passes of about 33.6 MB become one: the retained archive's
  digest, which change 0646's gate G5 requires to be recomputed at
  application, and which this change keeps.
- **Planning** moves 63.885 → 48.430 ms (0.7541). The candidate capture no
  longer re-hashes the payloads the snapshots already hashed; the first
  hashes of both archives and of the new candidate archive remain.
- **Per loop iteration**, `perf stat` isolation pairs give 0.7239 of the
  cycles and 0.7272 of the instructions.
- **Allocated bytes** fall by 33.58 MB per lifecycle (−12.3%) and peak live
  bytes by 16.4 MB (−5.4%): the application no longer copies the plan's
  retained 33.6 MB archive.

The plain pair moves 0.9768 (lifecycle) and 0.9701; the source-backed
control 0.9991. Every output digest equals the base's. The semantic controls
move 0.9952, 1.0030 and 1.0065 (see *Regression flags*).

Every `LPCP0004` patch byte, recorded revision, published byte and refusal is
the base's: a 98-line transcript generated on the base pins them (*Proof
obligations*).

## What changed

The owned cross-presentation slide copy proved the same bytes several times per
operation. A frame-pointer profile of this base (packet `attribution/`) puts
the commit phase at 79.8% SHA-256, in five passes over about 33.6 MB each, and
21.8% of planning in one more:

| pass | share of commit cycles (base) | bytes it hashed |
|---|---:|---|
| candidate capture (C1) in `build_candidate` | 16.4% | every payload of the reopened candidate |
| published capture (C2) in `validate_candidate` | 16.3% | the same package, again |
| live semantic revisions (`package_fingerprint`) | 16.0% | every payload of the live source and destination |
| live physical revisions (`to_stream` into a hash sink) | 16.0% | the live source's and destination's archives |
| retained archive digest | 15.9% | the plan's retained candidate archive |
| retained archive copy and its page faults | 8.8% | (a 33.6 MB copy, not a hash) |

Each hash pass is now answered from a digest already bound to the bytes it
describes. For the retained archive, the bytes are shared instead of copied,
and their digest is still recomputed. No check is removed, reordered or made
conditional.

### `crates/litchi-opc` (`4c7cbc4b8c`)

- **A digest memo bound to the retained owned source.** `OpcPackage`'s
  `source_archive` becomes a private `owned_source::OwnedSource`: the archive
  `Arc<Vec<u8>>` and an `Arc<OnceLock<[u8; 32]>>`. Its only constructor,
  `OwnedSource::new`, creates an empty memo. The fields are private to that
  nested module, so no struct update or assignment elsewhere in the crate can
  pair one archive with another's digest (0745's `Artifact` construction). The
  memo is filled at most once, from exactly those bytes. Clones of an ingress
  share both the archive and the memo.
- **`OpcPackage::exact_source_sha256(&self) -> Option<[u8; 32]>`** (new,
  additive) reads the memo, computing it on first request. It answers only while
  the exact-source authorization is intact, exactly when `exact_source_shared`
  answers. For such a package `to_stream` writes the retained archive and
  nothing else (`PackageWriter::write_counted`), so this is the digest of the
  bytes the package publishes. Debug and test builds re-derive every read.
- **Shared-archive ingress** (new, additive):
  `OpcPackage::from_shared_vec_reusing_payloads(Arc<Vec<u8>>, ReadLimits,
  &donor)` and `OpcPackage::from_shared_vec_with_limits(Arc<Vec<u8>>,
  ReadLimits)`. These are the existing owned ingresses over an archive the
  caller keeps. The package becomes a further owner of the immutable bytes
  instead of a copy of them. It reads, validates, decodes, donates and refuses
  exactly as the owned ingress does. `from_vec_reusing_payloads` and the
  owned-bytes path now route through them: they wrap the `Vec` in the `Arc`
  before parsing instead of after, which moves no byte and allocates the same.

### `crates/litchi-pptx` (`f84522f8af`)

`opened/cross_copy_plan.rs`:

- **Live revisions consult the facades' memos.** `apply_plan` and `apply_patch`
  take a `LiveDigests` pair, the payload-digest memos the facades retain for
  the source and the destination (change 0655). They recompute both live
  complete-package revisions with `package_fingerprint_with_memo`, and keep the
  memos those passes fill. The application's source and destination snapshots
  are captured with those memos (`capture_with_revision_and_digests_and_mce`)
  instead of the empty ones `capture_with_revision` gave them.
- **C1 consults the snapshots' memos.** The reopened candidate takes every
  payload whose bytes equal the built graph's from that graph
  (`from_vec_reusing_payloads`' donation). The built graph's parts share the
  destination's allocations (untouched parts) and the source's (copied parts).
  So `build_candidate` captures the candidate with the union of the two
  snapshots' memos (`PartDigests::union`, new) as its parent, at planning and
  at application.
- **C2 consults C1's memo.** `build_candidate` returns C1's memo with the
  candidate (`Prepared::digests`). `validate_application_candidate` passes it
  to `patch::validate_candidate`, which still captures the candidate after
  `validate_before`, under the patch's limits, and compares the revision. A
  candidate rebuilt from the destination is captured as before. The inverse
  route's restored clone is captured consulting the live destination's memo,
  and its published capture consults that capture's memo.
- **Live physical revisions are sealed from the bound digest.**
  `physical_package_fingerprint` of an unmodified owned source checks the
  archive bound first, refusing with the error the stream raises. It then seals
  `exact_source_sha256` into the `litchi-pptx-cross-physical-v2` revision.
  Every other package is streamed as before (`streamed_physical_revision`).
  The plan's snapshots compute the digest (planning's first hash) into the memo
  their live packages share, so application reads it. Debug builds re-derive
  every sealed value, and every debug re-derivation of a physical revision
  (`snapshot_physical_revision`, `remember_physical_revision`,
  `candidate_physical_revision`, `published_archive_revision`) now streams, so
  none of them rests on the memo.
- **The retained archive is shared, and its digest still recomputed.** On the
  reuse arm, `build_candidate` reopens the plan's retained `Arc` with a shared
  ingress instead of copying it into a fallibly reserved `Vec`. The digest is
  recomputed at application over those bytes, through the reopened package's
  own new memo (0646 G5; see *Soundness*). That memo then stays bound to the
  published destination's archive.

`opened/patch.rs`: `validate_candidate` takes the optional parent memo.
`opened/model.rs`: `PartDigests::union` (new); a `#[cfg(test)]` log of the
payloads each fingerprint hashed; `capture_with_revision` is removed (its only
caller was the cross copy).

`package/model.rs` and `package/codec.rs`, the facade:

- `Package::part_digests` becomes a `OnceLock<Arc<PartDigests>>`, so that a
  capture, which takes `&self`, can fill an empty slot.
- `opened_presentation_with_limits` captures with the held memo as its parent,
  when one is held. When the slot is empty and every part is built in, it keeps
  the capture's memo (`offer_part_digests`).
- The cross-copy applications pass both facades' memos.
- `apply_slide_removal_plan` now adopts its published snapshot's memo, as every
  other publication already did (see *Soundness*).

No `unsafe`, no new dependency, thread, clock or I/O. The facade still has no
archive dependency, and litchi-pptx still names no archive type.

## Authority

- **The dispatch** set the rule: a digest may be skipped only when its input
  bytes are provably bytes already hashed in the same operation, or are held in
  an immutable allocation whose digest is memoized and bound to it. Every
  freshness, revision, physical-provenance, budget and refusal check stays; a
  memo never survives a mutation of its bytes; `LPCP0004` and every recorded
  revision stay byte-identical.
- **0652's trade-offs.** Trade-off 2 (correctness first) decided the three
  places this change declines to skip a hash: the retained archive's digest
  (0646 G5), the first hash of each archive at planning, and the cold captures
  of packages no memo describes. Trade-off 3 is the scope: the common benign
  path (unmodified owned packages, opened and captured through the facade) is
  the one made cheaper. Trade-off 1 would allow breaking changes; none was
  needed. The new `litchi-opc` API is additive.
- **ADR 0005, the 2026-09-16 memo amendment (0655).** Every memo this change
  consults or retains meets its conditions, proved in *Soundness*.
- **ADR 0005, the 2026-09-16 retention amendment (0656).** The retained archive
  is still one allocation an operation already made. Sharing it with the
  published package is not a second copy.
- **Change 0646's frozen rules**, adopted by 0652 decision 5 through 0656:
  - G5: the archive digest is recomputed at application over the retained
    bytes. Kept.
  - G6: `validate_candidate`'s capture and revision check still run
    unconditionally, and the plan retains no `Snapshot`. Kept.
  - §6: C2 may not be reused. C2 is not reused. See *Soundness*.
  - G13: the copy at application is fallible. The copy no longer exists.
- **ADR 0003:** plans and patches stay exact-source-checked, reversible and
  deterministic. No verdict moves (*Proof obligations*).
- **ADR 0006:** preservation and published bytes are unchanged.
- **ADRs 0010/0011:** the archive stays inside `litchi-opc`. litchi-pptx passes
  an `Arc<Vec<u8>>` and reads a digest.

## Soundness

### Each skipped hash

| pass | how the digest is known | why the bytes are the same |
|---|---|---|
| live semantic revisions | the facade's memo (0655) | an entry retains the payload `Arc` it names, so a hit proves allocation identity and therefore byte identity; a foreign part whose `blob_arc` does not alias its `blob` is never memoized |
| C1 | the two snapshots' memos, joined | the same entry-level proof; the union copies entries with the `Arc`s they retain and lives only for the capture |
| C2 | C1's memo, from the same call | C1's memo names allocations of C1's clone of the candidate, which a clone of built-in parts shares with the candidate itself |
| live physical revisions | `exact_source_sha256`, bound to the retained archive | the memo is created with the archive by one constructor and filled only from it; the archive is immutable; `to_stream` of an exact source writes exactly it |

Every miss is an ordinary hash. The complete-package revision is still
computed from every part's name, content type, payload digest and
relationships, in sorted order, over the root relationships and non-part
members, and is compared with the recorded one. `package_fingerprint_with_memo`
still calls `get_part` on every part, so a deferred decode's refusal surfaces
exactly where it did. The physical bound is checked before the digest is
read, and the error is the stream's.

### Mutations cannot keep a memo alive over changed bytes

- **Payload memos** are keyed on allocation identity with the allocation
  retained. `set_blob` and `set_blob_shared` install a new allocation, so an
  edited part can only miss (0655).
- **The archive digest** is read only while exact-source authorization holds,
  and every mutating entry point of `OpcPackage` revokes it, even a no-op
  (0742). After a revocation, `exact_source_sha256` is `None` and the package
  is streamed. Nothing can mutate an archive in place while a package holds
  it: a shared `Arc` refuses `get_mut`, and `make_mut` copies. No path of
  `litchi-opc` mutates a retained archive.
- **The facade's memo** is emptied by every mutation that publishes no
  snapshot, and replaced by every publication (0655). This change found one
  publication that did neither, `apply_slide_removal_plan`, and makes it adopt.
  Correctness never depended on this, since a stale memo can only miss. It
  bounds retention: without it, the common flow of `opened_presentation`,
  `plan_slide_removal`, `apply_slide_removal_plan` would keep the removed
  slide's payload alive through the capture's memo. The facade-memo test now
  covers this route.

### The facade keeps a capture's memo, under the memo amendment

The amendment lets a facade retain a memo across operations when all of its
conditions hold. It says such a facade adopts the memo of the snapshot each
publication produces. This change also fills an empty slot from a capture of
the current graph. Each condition still holds:

- keyed on allocation identity with a strong reference: unchanged;
- every entry names an allocation the owning package holds: the capture's
  package is a clone of `self.opc`, and a clone of a built-in part shares its
  payload allocation, including a deferred part's decode cell; a facade holding
  a caller-defined part keeps nothing (tested with a part that copies its
  payload when cloned);
- released at every mutation that publishes no snapshot, adopted at every
  publication: unchanged, with the removal-plan gap above closed;
- a miss is an ordinary recomputation; no value, refusal or byte depends on
  it: unchanged;
- rebuilt or projected, never mutated in place: the slot is filled once and
  replaced wholesale;
- resident cost bounded by the part count: the kept memo is the snapshot's own
  `Arc`, so keeping it allocates nothing; it keeps 57 bytes per part alive
  after the snapshot is dropped (0655's figure);
- re-derived by `debug_assert!`: unchanged.

This reads the amendment's publication sentence as a requirement, not as an
exhaustive list of adoption points. The coordinator may prefer to amend the
ADR's wording; this record does not edit the ADR.

### Change 0646's frozen rules

**§6 ("C2 may not be reused"), point by point:**

1. **Order.** C2 still runs after `validate_before`, and C1 still runs before
   it, as at the base. Only C2's payload hashes are answered from C1's memo.
2. **Identity of the published package.** C1's memo can answer only for
   allocations the candidate holds. In the rebuild branch the prepared
   candidate and its memo are dropped, and the rebuilt candidate is captured
   exactly as before.
3. **Limits.** C2 runs under the patch's limits. A memo carries no limit.
4. **Retention.** Nothing is retained across calls: C1's memo, the union and
   the live memos live only for one planning or one application. The plan
   retains no memo and no `Snapshot`. The retained bytes are still reopened and
   read as a graph at application, by C1 and by C2.

**G5.** 0646 kept the digest recomputed so that `target_physical_revision` would
never be "a value that travels with the plan rather than a hash of the bytes
about to be published". The reuse arm now hashes the shared retained bytes
through the reopened package's new memo. That hash is taken at application,
over the bytes about to be published. The plan carries no digest.
`a_substituted_retained_archive_…` still passes in both builds.

### The bound digest memo in `litchi-opc`, under ADR 0005

It holds one 32-byte value per owned ingress, beside the archive the package
already holds. It is created together with the archive, never mutated in place
(filled once), and shared by clones. Every read is re-derived in debug builds.
A miss is a hash. It raises no error and holds no limit. It follows
`Snapshot::physical_revision`'s `Arc<OnceLock>` shape (change 0598) and
`transfer_index`'s per-ingress cell (change 0742). Its resident cost is one
small `Arc` allocation per owned ingress. That allocation is infallible, as
`transfer_index`'s and the archive `Arc`'s are at the same point.

### Shared ingress

The retained archive becomes the published destination's exact source. After
an application, the plan and the destination hold one allocation instead of
two copies. `retained_candidate_bytes` still reports what the plan holds, and
releasing or dropping the plan returns the plan's handle. The destination keeps
its own handle. `Limits::max_retained_candidate_bytes`' documentation said the
application copies the archive. It now says the archive is shared. The
retention test asserts the sharing and the handle counts.

## Proof obligations

**Byte identity against the base.** `digest_reuse_tests.rs` (new) records a
98-line transcript over two deterministic pairs: a media pair, with a stored
photo and a deflated image, and a plain pair. The transcript holds:

- the plan's `LPCP0004` bytes and its inverse's (SHA-256), the six recorded
  revisions and the encoding flag;
- the published archive (SHA-256) and snapshot revision for:
  - a retained plan, a released plan, and the plan into its warm planning
    facade;
  - the durable patch forward, undo, redo and undo again;
  - redo by plan;
  - a destination whose exact-source authorization an edit revoked;
  - a source whose exact-source authorization an edit revoked while keeping
    its bytes: its replanned patch and that patch's application;
- the refusal text and the untouched bytes for:
  - a stale patch;
  - a moved destination and a moved source;
  - a foreign source;
  - a destination with equal content and different archive bytes.

The module compiled unchanged on the base tree, which printed `GOLDEN`. The
after tree's transcript is identical, SHA-256 `cb410eef…` for both (packet
`golden/`). The test runs in debug builds, with every memo reuse re-derived,
and in release builds, where the reuse arm skips its captures.

**The memoized routes are taken.** `memo_path_tests.rs` (new) reads the test
log of how many payloads each memo-consulting fingerprint hashed:

- **Planning:** the candidate capture hashes 2 payloads of 28: the rewritten
  presentation part and the stored photo's copy.
- **An application** runs four such fingerprints, which hash `[0, 0, 2, 0]`
  payloads: the live source, the live destination, C1 and C2. The base hashed
  every payload of four packages.
- **The forward durable patch** hashes `[0, 0, 2, 0]` in the same four
  fingerprints.
- **The inverse durable patch:** the restored clone's capture hashes 1 payload
  and the published capture 0.
- **Without facade memos,** the two live passes hash every payload, and C1 and
  C2 still hash `[2, 0]`.
- **A second `opened_presentation`** hashes nothing.

The stored photo's copy is hashed because the reopen does not take it from the
built graph. The source's deferred decode allocates one byte of spare
capacity, the eager reopen reads a Stored member at exactly its length, and
donation refuses a larger donor. That is behavior of the base.

**Soundness tests** (`memo_path_tests.rs`, `opened/tests.rs`,
`litchi-opc`):

- **The facade memo.** After a capture, a commit, a typed edit, a removal
  patch and a removal plan, the facade's memo names only its own
  allocations.
- **A caller-defined part.** A facade holding a part that copies its payload
  when cloned keeps no capture memo; the snapshot's allocation for that part
  differs.
- **The exact-source revision.** An exact source's sealed physical revision
  equals the streamed one. Over the bound it raises the stream's exact error,
  and at exactly the bound it seals. An edited clone streams.
- **The `litchi-opc` digest.**
  - It equals SHA-256 of the streamed bytes on every owned and shared ingress
    (`from_vec`, `from_reader`, `open`, `from_vec_reusing_payloads`, both
    shared ingresses), and is absent for borrowed and newly authored packages.
  - Clones share one memo; a second ingress of equal bytes has its own.
  - A revoked clone answers `None` while the original still answers.
- **Shared ingress** matches owned ingress: the exact source is the caller's
  allocation, parts are equal, the same payloads are donated, and the same
  refusals are raised.
- **The retention test** asserts that the published package shares the
  plan's archive, and the handle counts after release and after the drops.

**Unchanged behavior.** These pass unchanged:

- 0742's media-transfer suite: stale and foreign sources, a byte-identical
  modified destination, undo then redo by patch and by plan, caller-defined
  parts, budgets and read limits;
- 0656's retention suite, including the release-build substitution refusal;
- 0598's revision-cache staleness tests.

The harness ran its ten untimed corpus gates before every measured process:
stale, foreign and borrowed refusals, durable round trip, source immutability
and semantic output.

## Evidence that motivated it

- [0742](0742-pptx-owned-cross-copy-media-transfer.md) *Where the remaining
  time goes*: commit is 80.0% SHA-256 in five passes; plan is 35.9%
  serialization with its digest, 22.0% physical revisions, 21.8% candidate
  capture.
- This base, profiled again (frame-pointer build of `6d989cad63`, SHA-256
  `454f70fa…`, 12 samples, `perf record -e cycles --call-graph fp`, packet
  `attribution/before-attribution.json`):
  - commit is 33.4% of the lifecycle's cycles, 79.75% of it SHA-256;
  - the five passes measure 16.35%, 16.25%, 16.01%, 15.99% and 15.88% of
    commit;
  - the retained archive's copy adds 7.6% kernel page-fault samples and 1.2%
    `memcpy`.

## Measurement

**Builds.** Both legs used the identical command,
`cargo build --release --locked --offline --manifest-path
tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`, plus
`--features allocator-metrics --bin litchi-perf-baseline-alloc` for the
allocator lane, with `CARGO_BUILD_JOBS=6`. The before leg was built from a
detached worktree of `6d989cad63`, the after leg from `f84522f8af`. Binaries
were staged outside the target directories before any run. The harness is
unchanged.

**Runs.**

- Every process was pinned with `taskset -c 4`. Every native-lane process
  ran under `perf stat -e cycles,instructions`, which counts the whole child;
  the allocator lane ran without it.
- ABBA order: four rounds of before, after, after, before, giving eight
  processes per arm.
- Samples after warmups: 20 after 3 for the media-rich cases; 40 after 3 for
  the plain cases and the source-backed control; 200 after 20 for the
  semantic control, which runs three corpora per process. The allocator lane
  ran 3 after 1.
- No run of the matrix was discarded or repeated. No build of this change
  ran during a measurement. Before the matrix, one 8-sample sanity run of a
  work-in-progress build was made. It is not reported, since that build was
  not the committed code.

**Statistics.** The median of process p50s per arm, with its min–max. The
median of the eight paired ratios, (s0, s1) and (s3, s2) in each round, with a
percentile bootstrap (10,000 draws, seed 751). The same pairing is used for
phases and counters. Every process's p50, p95 and mean are listed in the
packet's `analysis/matrix.md`.

| binary | SHA-256 |
|---|---|
| before, native | `736ab51487e3201ccfac5a7033618f1edb39d1c412db18e6a4c956d1f062f446` |
| after, native | `568207422863e86a9629418ff99f5a8e7db7fc63dc6e2bc5c4bc5b1ff162bb4b` |
| before, allocator | `67a211a3b8f0aa886c3cf08da51196ff085e075e61c00ac4d1e72f910cfa3321` |
| after, allocator | `ba94f26436b2e0647a675552d0914692de18eadce3d47efd355f1aff832cc4da` |
| before, frame pointers (profile only) | `454f70fa0fdd2f27aefd9e05bda19b027a51613ff5cdde47873c719043f7e56f` |
| after, frame pointers (profile only) | `3a7554d02c21141e762fb422411440bb1f1f8bac5e3646437da7591c437f4efe` |

**Host.** AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws, shared with other
agents' work on other cores. Rust 1.95.0 from `rust-toolchain.toml`. The
reports' `git_revision` and `git_worktree_dirty` fields are read from the
source tree when the harness runs, not when it is built; the SHA-256 above is
the binary's identity.

### Native latency

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [95% bootstrap] |
|---|---:|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 180.778 [174.664–200.927] | 102.309 [101.536–115.669] | 0.5637 [0.5366, 0.5849] |
| `pptx_cross_copy_media_rich` | 161.920 [153.704–172.583] | 86.380 [75.401–93.649] | 0.5242 [0.4755, 0.5431] |
| `pptx_cross_copy_plain_lifecycle` | 8.320 [8.268–8.373] | 8.148 [8.065–8.238] | 0.9768 [0.9713, 0.9891] |
| `pptx_cross_copy_plain` | 7.079 [6.960–7.225] | 6.883 [6.778–6.935] | 0.9701 [0.9601, 0.9781] |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` (control) | 17.053 [16.942–17.126] | 17.063 [12.139–17.137] | 0.9991 [0.9783, 1.0052] |
| `pptx_semantic_one_edit_save`, large corpus (control) | 53.827 [53.299–55.543] | 53.889 [53.281–54.239] | 0.9952 [0.9897, 1.0061] |
| `pptx_semantic_one_edit_save`, medium corpus (control) | 1.692 [1.682–1.701] | 1.699 [1.685–1.833] | 1.0030 [0.9980, 1.0079] |
| `pptx_semantic_one_edit_save`, tiny corpus (control) | 0.905 [0.900–0.913] | 0.911 [0.907–0.995] | 1.0065 [1.0039, 1.0147] |

Median process p95 and mean, ms:

| case | p95 before → after | mean before → after |
|---|---:|---:|
| media-rich lifecycle | 181.260 → 103.785 | 180.737 → 103.150 |
| media-rich | 162.391 → 86.770 | 161.946 → 86.417 |
| plain lifecycle | 8.455 → 8.316 | 8.308 → 8.139 |
| plain | 7.253 → 7.035 | 7.080 → 6.877 |
| source-backed control | 17.243 → 17.231 | 17.033 → 17.036 |
| semantic, large | 54.940 → 54.920 | 53.803 → 53.952 |
| semantic, medium | 1.709 → 1.713 | 1.693 → 1.701 |
| semantic, tiny | 0.916 → 0.920 | 0.905 → 0.912 |

Phases, median of process medians, ms (median paired ratio [95%]):

| case | plan | commit | publication |
|---|---:|---:|---:|
| media-rich lifecycle | 63.885 → 48.430 (0.7541 [0.7180, 0.7859]) | 88.691 → 25.674 (0.2890 [0.2688, 0.2896]) | 6.294 → 6.218 (0.9929) |
| media-rich | 66.747 → 54.429 (0.7791 [0.7587, 0.8534]) | 88.649 → 25.659 (0.2894 [0.2889, 0.2903]) | 6.400 → 6.297 (0.9699) |
| plain lifecycle | 3.285 → 3.274 (0.9922 [0.9793, 1.0078]) | 3.787 → 3.623 (0.9548 [0.9442, 0.9605]) | 0.001 → 0.001 |
| plain | 3.293 → 3.255 (0.9895 [0.9724, 0.9962]) | 3.793 → 3.619 (0.9534 [0.9429, 0.9609]) | 0.001 → 0.001 |
| source-backed control | 4.116 → 4.119 (1.0003) | — | 12.291 → 12.301 (0.9983) |

The media-rich publication phase is bimodal (about 1.5 ms or 6.3 ms,
independently of the binary, as 0656 documented). Its paired interval spans
[0.9332, 4.1112] for that reason, and no publication result is claimed.

### Cycles and instructions

Whole child under `perf stat` (median of processes; median paired ratio
[95%]). This includes the untimed corpus construction and gates:

| case | cycles before → after | ratio | instructions before → after | ratio |
|---|---:|---:|---:|---:|
| media-rich lifecycle | 35.951 G → 27.302 G | 0.7645 [0.7261, 0.7969] | 49.573 G → 38.822 G | 0.7835 [0.7513, 0.8208] |
| media-rich | 36.297 G → 27.210 G | 0.7435 [0.7207, 0.7597] | 49.758 G → 38.524 G | 0.7675 [0.7447, 0.7791] |
| plain lifecycle | 3.586 G → 3.592 G | 1.0007 [0.9945, 1.0089] | 9.546 G → 9.490 G | 0.9942 [0.9941, 0.9945] |
| plain | 3.571 G → 3.565 G | 0.9978 [0.9946, 1.0005] | 9.508 G → 9.452 G | 0.9941 [0.9939, 0.9942] |
| source-backed control | 20.428 G → 19.676 G | 0.9626 [0.9597, 0.9704] | 33.673 G → 32.677 G | 0.9701 [0.9688, 0.9749] |
| semantic control (three corpora, one child) | 200.223 G → 200.705 G | 1.0027 [0.9936, 1.0112] | 832.516 G → 832.144 G | 0.9996 [0.9980, 0.9998] |

The source-backed control's timed phases do not move (0.9991). Its whole-child
counts fall because its untimed corpus gates run owned plans and applications.

Per loop iteration, by isolation pairs (`scripts/isolation_pairs.py`). Each
pair is two whole-child counts at `--samples 12` and `--samples 2`, differenced
and divided by 10, in the order before, after, after, before, twice. The
difference cancels setup and the gates. It still includes the runner's
untimed per-iteration checks, which are the same code in both arms except
where a capture now keeps its memo.

| case | cycles per iteration | ratio | instructions per iteration | ratio |
|---|---:|---:|---:|---:|
| media-rich lifecycle | 1.217 G → 0.881 G | 0.7239 | 1.545 G → 1.124 G | 0.7272 |
| plain lifecycle | 46.68 M → 44.73 M | 0.9582 | 156.8 M → 155.6 M | 0.9923 |
| semantic control (three corpora) | 899.5 M → 907.7 M | 1.0091 | 3,760 M → 3,758 M | 0.9994 |

### First-touch faults

`scripts/faults.py` is 0742's, unchanged. It groups every media-rich lifecycle
sample by its count of freshly mapped 33.6 MB regions (8,204 minor faults
each):

| arm | fresh regions | samples | lifecycle ms | plan ms | commit ms |
|---|---:|---:|---:|---:|---:|
| before | 0 | 40 | 175.367 | 62.872 | 88.643 |
| before | 1 | 60 | 180.551 | 63.726 | 88.671 |
| before | 3 | 40 | 194.082 | 70.072 | 95.561 |
| before | 4 | 20 | 200.927 | 76.870 | 95.505 |
| after | 1 | 115 | 102.161 | 48.287 | 25.665 |
| after | 2 | 25 | 109.039 | 55.193 | 25.659 |
| after | 3 | 20 | 115.669 | 61.830 | 25.669 |

- At one fresh region the lifecycle ratio is 102.161 / 180.551 = 0.566.
- The base's commit rises to 95.5 ms with three or more fresh regions.
- After this change, commit is 25.66–25.67 ms in every group: it no longer
  maps a fresh 33.6 MB buffer, because the copy is gone.
- The remaining spread is in planning.

### Allocations (allocator lane, lifecycle region, median of process medians)

| field | media-rich before → after | plain before → after |
|---|---:|---:|
| allocation calls | 56,384 → 56,387 (+3) | 46,618 → 46,622 (+4) |
| deallocation calls | 46,087 → 46,090 | 38,275 → 38,279 |
| reallocation calls | 6,272 → 6,272 | 5,298 → 5,298 |
| allocated bytes | 272,766,692 → 239,185,859 (−33,580,833, −12.3%) | 16,615,187 → 16,594,762 (−20,425) |
| region peak live bytes | 305,071,007 → 288,634,911 (−16,436,096, −5.4%) | 1,313,406 → 1,313,403 |

- **Media-rich.** The retained archive (33,599,745 bytes) is no longer copied.
  About 18.9 KB of new allocation offsets part of that: the memos' projections
  and unions, and one digest cell per owned ingress. Peak live bytes fall by
  less than the archive, because the lifecycle's peak is not always during
  application.
- **Plain.** Allocation calls rise by four: the digest cells of the owned
  ingresses and the capture unions. Allocated bytes fall by 20,425 bytes (the
  plain candidate archive is 31,545 bytes).
- Latency in this lane: media-rich lifecycle 0.5743 [0.5471, 0.5986], plain
  lifecycle 0.9681 [0.9512, 0.9802].

### Output bytes

Every process of both arms published the digests 0742 reported: media-rich
`68be0ce5…`, plain `3e9ae280…`, source-backed `809c6172…`. The semantic
control has no output digest; the harness verifies it semantically after the
clock.

### Regression flags

No case's median paired ratio exceeds 1.05.

- **Semantic control, tiny corpus: +0.65% at p50 (1.0065 [1.0039, 1.0147]).**
  The interval excludes 1.0, so this is reported as a real, small regression
  of that 0.9 ms case. Dedicated runs (`counters/tiny-semantic/`,
  `scripts/tiny_semantic.py`, 2 ABBA rounds) show:
  - the same work: per-iteration instructions 0.9999 of the base, 3,069 fewer
    of 25.44 M;
  - more cycles: per-iteration cycles 1.0145, paired p50 1.0083;
  - whole-child front-end stalls: `de_no_dispatch_per_slot.no_ops_from_frontend`
    1.0306, branch misses 1.0073.

  This change's code on that path is `offer_part_digests`, one `OnceLock` set
  and a built-in check over about 20 parts, together with one more `Arc`
  increment per package clone. Instructions do not rise. The dispatch warned
  that this host shows large code-layout timing effects, and the front-end
  counters fit that. The mechanism is not established.
- **Semantic control, medium corpus: +0.30%**, interval [0.9980, 1.0079].
  Not distinguishable from zero. The large corpus is 0.9952.
- **Plain cases.** They fall 2–3%, and commit 4.5%, from the skipped hashes
  of the 31.5 KB candidate and the two archives of about 30 KB. Per iteration
  they use 0.9923 of the instructions and 0.9582 of the cycles. Their
  whole-child cycles are 1.0007 [0.9945, 1.0089] and 0.9978, because setup
  dominates the child.

## Where the remaining time goes

The frame-pointer after build (SHA-256 `3a7554d0…`, 12 samples,
`attribution/after-attribution.json`), shares of cycles within a phase:

- **Commit** (13.4% of the lifecycle's cycles, from 33.4%):
  - 59.5%: the retained archive's SHA-256, recomputed at application (G5);
  - 25.5%: the shared eager reopen (inflate, CRC, donation compare);
  - 4.7%: `prove_package_dialect`;
  - 1.3% and 1.1%: the two captures.
- **Plan** (28.4%, from 24.5%, of a shorter lifecycle):
  - 61.7%: `build_candidate`, whose serialization digest is 27.9% of planning;
  - 28.0%: the snapshots' physical revisions, planning's first hash of each
    archive, now into the memo that application reads;
  - 5.7%: the media captures.

**Next opportunities (not implemented):**

- **G5.** Retaining the reopened candidate's bound digest with the plan (a
  `litchi-opc` handle of archive plus memo) would remove the last commit hash,
  about 15 ms here. That departs from 0646's G5 wording ("recomputed at
  application"), which 0652 decision 5 adopted, so it needs the owner's
  decision.
- **Physical revisions.** Computing each archive's digest once per ingress
  instead of at the first plan would move 28% of planning out of the plan
  phase; it would not remove it.
- **Stored members.** The candidate reopen's donation refuses a Stored member
  whose source decode has spare capacity (above), so such copies are hashed
  and decoded twice.
- **The inverse route.** The inverse durable patch's forward replan builds
  from the reopen of the restored clone's serialization, which no memo names.

## What is not claimed

- No claim-registry entry.
- Only the generated media-rich and plain corpora and the semantic control
  corpora were measured. No real-producer deck.
- The facade memo helps an application only when the caller captured or
  published through the same `Package` values it applies to, as the harness
  and ordinary code do. An application on freshly opened packages hashes both
  live packages once, as at the base (tested).
- **Not timed:**
  - the durable-patch routes (tested, not timed);
  - the normalization path (modified packages);
  - repeated `opened_presentation` calls, which now hash nothing;
  - copies with Stored images.
- No cold-cache, remote-source, concurrency, RSS or cross-platform result.
- The per-iteration counters include the runner's untimed per-iteration
  checks.
- The attribution shares are two diagnostic captures of frame-pointer builds,
  not ordinary-release timings.
- The tiny semantic control's regression is not explained beyond the counter
  evidence above.

The broader non-iWork goal remains active.

## Verification

Gates on `f84522f8af` (commands, exit codes and counts in
[`gates.txt`](results/change-0751/gates.txt)):

- `cargo fmt --all --check`;
- `cargo check --all-targets` of `litchi-opc`, `litchi-pptx` and their
  in-scope dependents (`litchi-ooxml-common`, `litchi-docx`, `litchi-xlsx`,
  `litchi-xlsb`, `litchi-spreadsheet-drawing`, `litchi-ppt`, `litchi`);
- Clippy `-D warnings`:
  - both libraries;
  - `litchi-opc --all-targets`;
  - `litchi-pptx --all-targets` fails only on the three pre-existing
    `err_expect` lints at `opened/tests.rs:464/538/557` that 0742 recorded on
    its base.
- Tests:
  - `litchi-opc` and `litchi-pptx`: 1,741 passed, 0 failed, 4 ignored;
  - `litchi-pptx --release`, the cross-copy and facade-memo modules: 57
    passed, 1 ignored (the golden printer). This runs the release-only
    branches: the retained reuse that skips its captures, and the
    substitution refusal with its debug assertion compiled out;
  - the six other dependents: 5,241 passed, 0 failed, 55 ignored;
  - the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`: 382 passed,
    0 failed, 7 ignored.
- rustdoc `-D warnings` for both crates;
- the crate-boundary, non-iWork and structural claim gates.

**Tests.** Twelve are new, one of them an ignored printer, and three are
updated:

- `litchi-opc`, 3 new:
  - `exact_source_sha256_is_the_digest_of_the_published_bytes`;
  - `exact_source_sha256_memo_is_bound_to_its_archive`;
  - `shared_ingress_matches_owned_ingress_without_copying`.
- `litchi-opc`, updated:
  `clone_shares_owned_source_but_revocation_is_independent` compares the
  bound archive's bytes.
- `litchi-pptx`, 9 new:
  - `every_patch_revision_output_and_refusal_matches_the_base` and its
    printer, `print_the_golden_transcript`, ignored by default;
  - seven memo-path tests: planning; application with and without facade
    memos; the durable routes; the capture-filled facade memo; a caller-defined
    part; the exact-source seal and its bound.
- `litchi-pptx`, updated:
  - `the_facade_memo_never_outlives_the_graph_it_describes`, for the new slot
    type, extended with a capture and a removal plan;
  - `releasing_and_dropping_return_the_retained_bytes`, asserting the
    sharing.

**Merging with change 0743.** 0743 edits `opened/model.rs`:
`capture_internal` gains a `parent_roots` argument, and `capture_with_revision`
and `capture_with_revision_and_digests_and_mce` gain one each. Merging the two
branches requires two things:

- resolve the modify/delete conflict on `capture_with_revision` by deleting it,
  since this change removed its only caller;
- add 0743's trailing argument at the four new
  `capture_with_revision_and_digests_and_mce` call sites in `cross_copy_plan.rs`.

The other hunks, `PartDigests::union`, the test-only log, the `opened/mod.rs`
export line and the `Limits` documentation, do not overlap 0743's.

## Cleanup

See [`cleanup.json`](results/change-0751/cleanup.json). Binary identities are
recorded above and in the packet's `binaries.sha256`, taken before removal.
Removed, 125.0 GB in all:

- the target directories:
  - `targets/0751`, 109.1 GB: this change's debug, release
    and gate builds;
  - `targets/0751-before`, 3.0 GB;
  - `targets/0751-before-fp` and `targets/0751-fp`,
    1.1 GB each;
- the scratch directory's contents, 441 MB: staged binaries, raw
  profiles, working copies of the reports;
- the detached before worktree `0751-before-src`.

Kept: the worktree and branch. The frame-pointer profiles' `perf.data`
(53 MB) is not retained; `attribution/` summarizes it.
