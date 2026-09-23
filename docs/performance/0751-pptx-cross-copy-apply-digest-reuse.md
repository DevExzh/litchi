# 0751 — the owned PPTX cross-copy stops re-hashing proven bytes: media-rich commit 88.9 → 25.8 ms, lifecycle paired ratio 0.61 (0.56 at equal page faults)

Status: retained, implemented in `litchi-opc` and `litchi-pptx`.
`performance_claim: none` — no claim-registry entry; the paired medians,
counter ratios and allocator counts below are reported as evidence, not
registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `6d989cad63` (the tip of `feat/office-format-completeness`, with records
0742, 0744–0747). Production commits on
`perf/0751-pptx-cross-copy-apply-digest-reuse`:

- `4c7cbc4b8c` (`litchi-opc`) and `f84522f8af` (`litchi-pptx`), the change;
- after the independent review (verdict: merge after fixes):
  - `aa606e0016`, the facade memo re-projected at every adoption;
  - `c0fefbc85b`, the `litchi-opc` nits;
  - `098f6ffd4e`, the extended golden transcript.

Evidence packet: [`results/change-0751/`](results/change-0751/README.md).

## Result

On change 0742's media-rich pair (8 × 2 MiB incompressible PNGs per deck,
16.8 MB archives), both legs built with the identical command. The numbers
are from the matrix of `098f6ffd4e`, measured after the review's fixes:

| case | before median p50 ms | after median p50 ms | median paired ratio [95% bootstrap] |
|---|---:|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 184.195 | 109.760 | 0.6083 [0.5228, 0.6438] |
| `pptx_cross_copy_media_rich` | 159.629 | 80.585 | 0.5137 [0.5021, 0.5266] |

- **Commit (application)** moves 88.898 → 25.846 ms (paired 0.2905). Its
  five SHA-256 passes of about 33.6 MB become one: the retained archive's
  digest, which change 0646's gate G5 requires to be recomputed at
  application, and which this change keeps.
- **Planning** moves 67.027 → 55.305 ms (0.8705), and 64.037 → 48.355 ms
  (0.7648) in the non-lifecycle case. The candidate capture no longer
  re-hashes the payloads the snapshots already hashed; the first hashes of
  both archives and of the new candidate archive remain.
- **Page faults set the lifecycle's spread.** Grouped by freshly mapped
  33.6 MB buffers, the lifecycle at one fresh mapping moves 180.44 →
  101.92 ms (0.565). Commit is 25.7–25.9 ms in every group. This matrix's
  after processes landed more often in two- and three-mapping modes than the
  first matrix's, which is why its median paired ratio (0.61) is higher than
  the first matrix's 0.5637 (*Measurement*).
- **Per loop iteration**, `perf stat` isolation pairs give 0.6500 of the
  cycles and 0.6492 of the instructions.
- **Memory.** Allocated bytes fall by 33.56 MB per lifecycle (−12.3%) and peak
  live bytes by 16.4 MB (−5.4%): the application no longer copies the plan's
  retained 33.6 MB archive.

The plain pair moves 0.9666 (lifecycle) and 0.9586; the source-backed
control 1.0013. Every output digest equals the base's.

The semantic controls move 0.9940 (large), 1.0084 (medium) and 1.0112 (tiny).
The two small ones are regressions below the 5% trigger. The tiny corpus
executes 0.17% more instructions per iteration: the review's re-projection of
the facade memo at each capture and publication (*Regression flags*).

Every `LPCP0004` patch byte, recorded revision, published byte and refusal is
the base's: a 196-line transcript generated on the base pins them (*Proof
obligations*).

The first matrix, of `f84522f8af` before the review, measured the lifecycle at
0.5637 and the tiny control at 1.0065. It is kept as a summary.

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
  `source_archive` becomes a private `retained_archive::RetainedArchive`: the
  archive `Arc<Vec<u8>>` and an `Arc<OnceLock<[u8; 32]>>`. Its only
  constructor, `RetainedArchive::new`, creates an empty memo. The fields are private to that
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
- **`OpcPackage::exact_source_len(&self) -> Option<usize>`** (new, additive,
  `c0fefbc85b`): the retained archive's length, answered exactly when
  `exact_source_sha256` is, without taking a handle or hashing. The private
  memo type is named `retained_archive::RetainedArchive`, to avoid
  `litchi_core::OwnedSource`'s name.
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
  archive bound first, with `exact_source_len`, refusing with the error the
  stream raises. It then seals
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
  when one is held. When the slot is empty, it keeps the capture's memo
  re-projected onto `self.opc`'s own allocations (`offer_part_digests`).
- Every publication adopts its snapshot's memo re-projected the same way
  (`adopt_part_digests`, `aa606e0016`). At the base it adopted the memo as it
  was, which could keep a caller-defined part's clone alive (see
  *Soundness*).
- The cross-copy applications pass both facades' memos.
- `apply_slide_removal_plan` now adopts its published snapshot's memo, as every
  other publication already did.

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
- **Coordinator ruling, 0751 review.** Filling an empty facade memo slot from
  a capture of the current, unmutated graph is permitted under the
  amendment. Its adoption clause says when the facade must refresh the memo
  (publications) and when it must release it (mutations that publish
  nothing); it does not forbid this. Its clause "a memo carried onto another
  package is re-projected onto that package's own allocations rather than
  inherited" must be met literally in code, at every adoption point. It is:
  `PartDigests::project` onto `self.opc` (`aa606e0016`). The ADR text is
  unchanged, and this reading is listed for the owner's confirmation.
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

### The facade keeps only memos re-projected onto its own allocations

A published or captured snapshot's memo names the allocations of the
snapshot's own package, a clone of `self.opc`. A clone of a built-in part
shares its payload allocation, including a deferred part's decode cell, but a
caller-defined `Part` may copy its payload when cloned.

**What the review found.** The first version adopted a publication's memo as
it was. With a part that copies a 64 KiB payload when cloned,
`opened_presentation`, `plan_slide_removal`, `apply_slide_removal_plan` left
an entry naming the snapshot's copy, and the facade kept those 64 KiB alive
after the snapshot was dropped. The base pinned nothing on that route; the
base's cross-copy publications had the same gap.

**The fix** (`aa606e0016`). Every adoption point, `adopt_part_digests` after
each publication and `offer_part_digests` after a capture, keeps only
`PartDigests::project` of the snapshot's memo onto `self.opc`. An entry
survives only for an allocation `self.opc` holds, and it retains
`self.opc`'s own `Arc`, never the snapshot's; any other entry is dropped. A
facade holding a caller-defined part therefore keeps a memo without that
part's entry, where the first version kept none. A projection that cannot
allocate keeps nothing.

Each condition of the amendment:

- keyed on allocation identity with a strong reference: unchanged;
- every entry names an allocation the owning package holds, and a memo carried
  onto another package is re-projected onto that package's own allocations:
  by construction, since each kept entry is `self.opc`'s own `Arc` (tested
  entry by entry, with `Arc` identity, after each route);
- released at every mutation that publishes no snapshot, adopted at every
  publication: unchanged, with the removal-plan gap closed;
- a miss is an ordinary recomputation; no value, refusal or byte depends on
  it: unchanged;
- rebuilt or projected, never mutated in place: the facade's memo is always a
  fresh projection, set once and replaced wholesale;
- resident cost bounded by the part count, charged through fallible
  reservation: the projection reserves its table with `try_reserve`, one slot
  per part (57 bytes, 0655's figure), and keeps no payload `self.opc` does not
  hold;
- re-derived by `debug_assert!`: unchanged.

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
196-line transcript over two deterministic pairs: a media pair, with a stored
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
- added after the review, a destination holding a caller-defined part:
  - planned against it, which records the recompressing route, then applied
    by plan and by durable patch, forward and back;
  - a plan and a patch planned against an ordinary destination with the same
    bytes. For the media pair they transfer, and the caller-defined
    destination refuses them with `CallerDefinedPart` and is left untouched.
    For the plain pair they apply.
- added after the review, size-limit refusals end to end. For five bounds,
  the durable patch is read under the bound and applied, and a copy is
  planned under it and applied when it plans. The bounds are the patch's
  length, and each archive's length and one byte under it. Most of them end in
  `Limit { resource: "cross-slide serialized archive bytes", .. }` from the
  length-first physical revision, at application and at planning.

The module compiled unchanged on the base tree, which printed `GOLDEN`. The
after tree's transcript is identical, SHA-256 `32b7dbf9…` for both (packet
`golden/`). The first 98 lines are the first version's transcript, SHA-256
`cb410eef…`, which the review reproduced on the base. The test runs in debug
builds, with every memo reuse re-derived, and in release builds, where the
reuse arm skips its captures.

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
  patch and a removal plan, every entry of the facade's memo retains the very
  `Arc` its own graph holds.
- **A caller-defined part that copies its 64 KiB payload when cloned**
  (`the_facade_memo_never_outlives_the_graph_it_describes`, after the review).
  It is tested on the removal plan route, and on the cross-copy route by plan,
  by durable patch and by the inverse. After each publication:
  - the snapshot's memo names the snapshot's copy;
  - once the snapshot is dropped, a `Weak` to that copy no longer upgrades,
    so nothing keeps it alive;
  - the facade's memo keeps every entry but that part's.

  The removal case fails under the first version's adoption.
- **A capture's memo is projected** (`a_capture_memo_is_projected_onto_the_facade_allocations`).
  The snapshot's memo holds an entry for the snapshot's copy, which the
  facade does not hold. The facade's memo drops exactly that entry, and every
  kept entry retains the facade's own `Arc`.
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
detached worktree of `6d989cad63`, the after leg from `098f6ffd4e`. Binaries
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
- No run of the matrix was discarded or repeated, and no build of this change
  ran during a measurement. The host's load average was about 13 during this
  matrix and 1–4 during the first.

**Statistics.** The median of process p50s per arm, with its min–max. The
median of the eight paired ratios, (s0, s1) and (s3, s2) in each round, with a
percentile bootstrap (10,000 draws, seed 751). The same pairing is used for
phases and counters. Every process's p50, p95 and mean are listed in the
packet's `analysis/matrix.md`.

| binary | SHA-256 |
|---|---|
| before, native | `60e0b2587585855a0557627a6cc05aa4f1e36446a8e249e78708b7199429ec20` |
| after, native (`098f6ffd4e`) | `2fe858b5de84d6b07b7ce5b59fd1b13cdbd96775a5a2e4644e3c48c8e495e960` |
| before, allocator | `014c4bdf26dd3bef8ae0d2ec5140f8d2f6f70c2638e4f066c03e72a07640313a` |
| after, allocator (`098f6ffd4e`) | `4bc7f216098516d82d921b974d6ce48e8cff4819265e4a9b9d342e67febd6ccd` |

The first matrix's binaries, and the frame-pointer builds used for
attribution, are listed in `superseded-f84522f8af/binaries.sha256` and under
*Evidence that motivated it* and *Where the remaining time goes*. The before
binaries were rebuilt for this matrix, with the same command from the same
tree, and their hashes differ from the first matrix's (`736ab514…`,
`67a211a3…`): release builds here are not bit-reproducible.

**Host.** AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws, shared with other
agents' work on other cores. Rust 1.95.0 from `rust-toolchain.toml`. The
reports' `git_revision` and `git_worktree_dirty` fields are read from the
source tree when the harness runs, not when it is built; the SHA-256 above is
the binary's identity.

### Native latency

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [95% bootstrap] |
|---|---:|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 184.195 [175.207–202.181] | 109.760 [101.720–117.298] | 0.6083 [0.5228, 0.6438] |
| `pptx_cross_copy_media_rich` | 159.629 [154.153–173.494] | 80.585 [79.910–94.182] | 0.5137 [0.5021, 0.5266] |
| `pptx_cross_copy_plain_lifecycle` | 8.454 [8.359–8.818] | 8.166 [8.154–8.175] | 0.9666 [0.9486, 0.9729] |
| `pptx_cross_copy_plain` | 7.121 [7.094–7.173] | 6.823 [6.789–6.904] | 0.9586 [0.9519, 0.9655] |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` (control) | 17.085 [11.972–17.581] | 17.117 [12.518–17.354] | 1.0013 [0.9871, 1.4190] |
| `pptx_semantic_one_edit_save`, large corpus (control) | 54.667 [53.643–55.520] | 54.492 [54.043–54.886] | 0.9940 [0.9863, 1.0135] |
| `pptx_semantic_one_edit_save`, medium corpus (control) | 1.700 [1.685–1.702] | 1.711 [1.707–1.719] | 1.0084 [1.0058, 1.0155] |
| `pptx_semantic_one_edit_save`, tiny corpus (control) | 0.907 [0.903–0.913] | 0.919 [0.914–0.922] | 1.0112 [1.0096, 1.0174] |

Median process p95 and mean, ms:

| case | p95 before → after | mean before → after |
|---|---:|---:|
| media-rich lifecycle | 184.709 → 110.175 | 184.215 → 109.753 |
| media-rich | 160.600 → 81.141 | 159.760 → 80.626 |
| plain lifecycle | 8.637 → 8.269 | 8.453 → 8.177 |
| plain | 7.356 → 6.958 | 7.139 → 6.841 |
| source-backed control | 17.349 → 17.363 | 17.119 → 17.120 |
| semantic, large | 55.676 → 55.247 | 54.656 → 54.507 |
| semantic, medium | 1.715 → 1.724 | 1.700 → 1.712 |
| semantic, tiny | 0.917 → 0.927 | 0.907 → 0.920 |

Phases, median of process medians, ms (median paired ratio [95%]):

| case | plan | commit | publication |
|---|---:|---:|---:|
| media-rich lifecycle | 67.027 → 55.305 (0.8705 [0.6848, 0.9798]) | 88.898 → 25.846 (0.2905 [0.2690, 0.2919]) | 6.463 → 6.371 (0.9855) |
| media-rich | 64.037 → 48.355 (0.7648 [0.7477, 0.7831]) | 89.258 → 25.870 (0.2906 [0.2877, 0.2922]) | 6.413 → 6.380 (1.0075) |
| plain lifecycle | 3.354 → 3.270 (0.9765 [0.9507, 0.9869]) | 3.835 → 3.607 (0.9410 [0.9227, 0.9465]) | 0.001 → 0.001 |
| plain | 3.316 → 3.230 (0.9752 [0.9675, 0.9829]) | 3.808 → 3.592 (0.9426 [0.9364, 0.9481]) | 0.001 → 0.001 |
| source-backed control | 4.190 → 4.174 (1.0019) | — | 12.262 → 12.275 (0.9985) |

The media-rich publication phase is bimodal (about 1.5 ms or 6.4 ms,
independently of the binary, as 0656 documented), and no publication result
is claimed. The lifecycle's planning phase follows page faults (below).

### Cycles and instructions

Whole child under `perf stat` (median of processes; median paired ratio
[95%]). This includes the untimed corpus construction and gates:

| case | cycles before → after | ratio | instructions before → after | ratio |
|---|---:|---:|---:|---:|
| media-rich lifecycle | 36.727 G → 28.035 G | 0.7665 [0.7119, 0.7931] | 50.462 G → 39.253 G | 0.7840 [0.7494, 0.8050] |
| media-rich | 36.524 G → 27.128 G | 0.7462 [0.7335, 0.7709] | 49.814 G → 38.327 G | 0.7715 [0.7589, 0.7913] |
| plain lifecycle | 3.701 G → 3.636 G | 0.9851 [0.9788, 0.9869] | 9.547 G → 9.505 G | 0.9957 [0.9955, 0.9958] |
| plain | 3.660 G → 3.607 G | 0.9826 [0.9714, 0.9984] | 9.510 G → 9.468 G | 0.9955 [0.9954, 0.9957] |
| source-backed control | 20.465 G → 19.733 G | 0.9611 [0.9509, 1.0306] | 33.667 G → 32.598 G | 0.9715 [0.9546, 1.0263] |
| semantic control (three corpora, one child) | 200.848 G → 202.571 G | 1.0066 [1.0024, 1.0118] | 832.495 G → 832.125 G | 0.9995 [0.9995, 0.9996] |

The source-backed control's timed phases do not move (1.0013). Its
whole-child counts fall because its untimed corpus gates run owned plans and
applications.

Per loop iteration, by isolation pairs (`scripts/isolation_pairs.py`). Each
pair is two whole-child counts at `--samples 12` and `--samples 2`, differenced
and divided by 10, in the order before, after, after, before, twice. The
difference cancels setup and the gates. It still includes the runner's
untimed per-iteration checks, which are the same code in both arms except
where a capture now keeps its memo.

| case | cycles per iteration | ratio | instructions per iteration | ratio |
|---|---:|---:|---:|---:|
| media-rich lifecycle | 1.242 G → 0.807 G | 0.6500 | 1.575 G → 1.022 G | 0.6492 |
| plain lifecycle | 47.35 M → 44.80 M | 0.9463 | 156.9 M → 155.7 M | 0.9922 |
| semantic control (three corpora) | 903.8 M → 913.2 M | 1.0104 | 3,759 M → 3,758 M | 0.9997 |

### First-touch faults

`scripts/faults.py` is 0742's, unchanged apart from reading gzip. It groups
every media-rich lifecycle sample by its count of freshly mapped 33.6 MB
regions (8,204 minor faults each):

| arm | fresh regions | samples | lifecycle ms | plan ms | commit ms |
|---|---:|---:|---:|---:|---:|
| before | 0 | 40 | 175.597 | 63.033 | 88.902 |
| before | 1 | 40 | 180.440 | 63.432 | 88.625 |
| before | 2 | 30 | 187.670 | 70.555 | 88.654 |
| before | 3 | 30 | 194.690 | 70.396 | 95.833 |
| before | 4 | 20 | 202.181 | 77.551 | 95.856 |
| after | 1 | 40 | 101.918 | 47.867 | 25.725 |
| after | 2 | 60 | 109.657 | 55.180 | 25.850 |
| after | 3 | 60 | 115.974 | 62.461 | 25.845 |

- At one fresh region the lifecycle ratio is 101.918 / 180.440 = 0.565, the
  first matrix's 0.566.
- The base's commit rises to 95.8 ms with three or more fresh regions.
- After this change, commit is 25.7–25.9 ms in every group: it no longer maps
  a fresh 33.6 MB buffer, because the copy is gone.
- The remaining spread is in planning. In this matrix, six of the eight after
  processes sat in two- or three-region modes, against two of eight in the
  first matrix. That moves the median paired ratio from 0.5637 to 0.6083. It
  is per-process allocator state, as 0742 found, and the mechanism that
  selects a mode is not established.

### Allocations (allocator lane, lifecycle region, median of process medians)

| field | media-rich before → after | plain before → after |
|---|---:|---:|
| allocation calls | 56,384 → 56,394 (+10) | 46,618 → 46,628 (+10) |
| deallocation calls | 46,086 → 46,093 | 38,275 → 38,281 |
| reallocation calls | 6,272 → 6,272 | 5,298 → 5,298 |
| allocated bytes | 272,765,116 → 239,200,195 (−33,564,921, −12.3%) | 16,615,187 → 16,600,474 (−14,713) |
| region peak live bytes | 305,070,945 → 288,642,369 (−16,428,576, −5.4%) | 1,313,408 → 1,317,213 (+3,805, +0.29%) |

- **Media-rich.** The retained archive (33,599,745 bytes) is no longer copied.
  About 35 KB of new allocation offsets part of that. It comes from:
  - the memos' unions and projections, including the facade's re-projections
    after the review;
  - one digest cell per owned ingress.

  Peak live bytes fall by less than the archive, because the lifecycle's peak
  is not always during application.
- **Plain.** Allocation calls rise by ten, and peak live bytes by 3,805 bytes.
  The review's re-projection allocates the facade's memo separately from the
  snapshot's instead of sharing it, so both are alive at the peak. Allocated
  bytes still fall by 14,713 bytes (the plain candidate archive is 31,545
  bytes). The first matrix, before the review, measured +4 calls and an
  unchanged peak.
- Latency in this lane: media-rich lifecycle 0.5722 [0.5456, 0.6161], plain
  lifecycle 0.9647 [0.9527, 0.9726].

### The first matrix (superseded)

Before the review, the same matrix ran on `f84522f8af` (binaries in
`superseded-f84522f8af/binaries.sha256`; every process's row in
`superseded-f84522f8af/analysis/`). The review's fixes change what the facade
keeps, so this matrix was repeated in full. The first matrix's median paired
ratios:

| case | first matrix (`f84522f8af`) | this matrix (`098f6ffd4e`) |
|---|---:|---:|
| media-rich lifecycle | 0.5637 [0.5366, 0.5849] | 0.6083 [0.5228, 0.6438] |
| media-rich | 0.5242 [0.4755, 0.5431] | 0.5137 [0.5021, 0.5266] |
| plain lifecycle | 0.9768 [0.9713, 0.9891] | 0.9666 [0.9486, 0.9729] |
| plain | 0.9701 [0.9601, 0.9781] | 0.9586 [0.9519, 0.9655] |
| source-backed control | 0.9991 [0.9783, 1.0052] | 1.0013 [0.9871, 1.4190] |
| semantic, large | 0.9952 [0.9897, 1.0061] | 0.9940 [0.9863, 1.0135] |
| semantic, medium | 1.0030 [0.9980, 1.0079] | 1.0084 [1.0058, 1.0155] |
| semantic, tiny | 1.0065 [1.0039, 1.0147] | 1.0112 [1.0096, 1.0174] |

In the first matrix, commit moved 88.691 → 25.674 ms and, per iteration,
the media-rich lifecycle ran 0.7239 of the cycles and 0.7272 of the
instructions.

### Output bytes

Every process of both arms published the digests 0742 reported: media-rich
`68be0ce5…`, plain `3e9ae280…`, source-backed `809c6172…`. The semantic
control has no output digest; the harness verifies it semantically after the
clock.

### Regression flags

No case's median paired ratio exceeds 1.05.

- **Semantic control, tiny corpus: +1.12% at p50 (1.0112 [1.0096, 1.0174]).**
  The interval excludes 1.0: a real, small regression of that 0.9 ms case.
  Dedicated runs (`counters/tiny-semantic/`, `scripts/tiny_semantic.py`, 2
  ABBA rounds) show:
  - per-iteration instructions 1.0017: 43,347 more of 25.45 M;
  - per-iteration cycles 1.0165, paired p50 1.0159;
  - whole-child front-end stalls
    (`de_no_dispatch_per_slot.no_ops_from_frontend`) 1.0509, branch misses
    1.0068.

  Two mechanisms add up:
  - **The review's re-projection.** Each iteration captures once
    (`opened_presentation_transaction`) and publishes once
    (`apply_opened_presentation_commit`). Each now projects the snapshot's
    memo onto the facade's allocations: one table allocation plus a lookup and
    an insert per part, about 21,700 instructions each on this deck. This
    path never consults the capture's memo, since its commit rebinds the
    committed snapshot, so here the projection is pure cost; it is the price
    of meeting the re-projection clause at every adoption point.
  - **Code layout.** Before the review, the same case moved +0.65% at p50
    with per-iteration instructions 0.9999, 3,069 fewer, and cycles 1.0145.
    The dispatch warned that this host shows large code-layout timing effects,
    and the front-end counters fit that; the mechanism is not established.
- **Semantic control, medium corpus: +0.84%** (1.0084 [1.0058, 1.0155]),
  by the same two mechanisms. Before the review it was 1.0030
  [0.9980, 1.0079]. The large corpus is 0.9940.
- **Plain cases.** They fall 3.3–4.1%, and commit about 6%, from the skipped
  hashes of the 31.5 KB candidate and the two archives of about 30 KB. Per
  iteration they use 0.9922 of the instructions and 0.9463 of the cycles.
  Their whole-child cycles fall 1.5–1.7%. Their peak live bytes rise by
  3,805 bytes (+0.29%) from the separately allocated facade memo.

## Where the remaining time goes

The frame-pointer after build of `f84522f8af`, before the review (SHA-256
`3a7554d0…`, 12 samples, `attribution/after-attribution.json`), shares of
cycles within a phase. The review's fixes add one memo projection to the
destination's publication and one to each capture, and change nothing else
on this path; the profile was not repeated:

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
- **Cheaper facade adoption.** `PartDigests::project` looks each part up
  twice (`iter_parts`, then `get_part`); iterating `try_iter_parts` would halve
  that on every projection, including 0655's. A capture could also offer only
  a `Weak` hint instead of a projected memo, which retains nothing and costs
  nothing on paths that never consult it. Either would reduce the semantic
  controls' regression above.

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
- The semantic controls' regressions are attributed to the review's
  re-projection by the per-iteration instruction count; the cycle share from
  code layout is not explained beyond the counter evidence above.
- The lifecycle's median paired ratio depends on the per-process page-fault
  mode: 0.5637 in the first matrix and 0.6083 in this one, 0.565 at equal
  faults in both.

The broader non-iWork goal remains active.

## Verification

**Gates after the review, on `098f6ffd4e`**, with a fresh target directory
and `TMPDIR` under the scratch directory (commands, exit codes and counts in
[`gates.txt`](results/change-0751/gates.txt)):

- `cargo fmt --all --check`;
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
    substitution refusal with its debug assertion compiled out.
- rustdoc `-D warnings` for both crates.

**Gates before the review, on `f84522f8af`**
([`superseded-f84522f8af/gates.txt`](results/change-0751/superseded-f84522f8af/gates.txt)),
covered the rest:

- `cargo check --all-targets` of both crates and their in-scope dependents
  (`litchi-ooxml-common`, `litchi-docx`, `litchi-xlsx`, `litchi-xlsb`,
  `litchi-spreadsheet-drawing`, `litchi-ppt`, `litchi`);
- the six other dependents' tests: 5,241 passed, 0 failed, 55 ignored;
- the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`: 382 passed;
- the crate-boundary, non-iWork and structural claim gates.

The review's commits change no public item a dependent uses. `litchi-opc`
gains `exact_source_len`, an additive method. The facade tests and the gate
scripts were run again after the re-measurement (see `gates.txt`).

**Tests.** Twelve are new, one of them an ignored printer, and four are
updated:

- `litchi-opc`, 3 new:
  - `exact_source_sha256_is_the_digest_of_the_published_bytes`, which also
    checks `exact_source_len`;
  - `exact_source_sha256_memo_is_bound_to_its_archive`;
  - `shared_ingress_matches_owned_ingress_without_copying`.
- `litchi-opc`, updated:
  `clone_shares_owned_source_but_revocation_is_independent` compares the
  bound archive's bytes.
- `litchi-pptx`, 9 new:
  - `every_patch_revision_output_and_refusal_matches_the_base` and its
    printer, `print_the_golden_transcript`, ignored by default;
  - seven memo-path tests:
    - planning;
    - application, with and without facade memos;
    - the durable routes;
    - the capture-filled facade memo;
    - `a_capture_memo_is_projected_onto_the_facade_allocations`, which after
      the review replaces the test that expected no memo;
    - the exact-source seal and its bound.
- `litchi-pptx`, updated:
  - `the_facade_memo_never_outlives_the_graph_it_describes`, for the new slot
    type, extended with a capture, a removal plan and, after the review, the
    caller-defined removal and cross-copy routes;
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

None of this change's other `opened/model.rs` hunks overlaps 0743's:
`PartDigests::union`, `PartDigests::project` made `pub(crate)`, the test-only
log and accessor, and the `Limits` documentation. Neither does the
`opened/mod.rs` export line.

## Cleanup

Binary identities are recorded above and in the packet's `binaries.sha256`,
taken before removal. Each round committed its record and packet before
deleting anything.

- **After the review**
  ([`cleanup.json`](results/change-0751/cleanup.json), written by
  `scripts/cleanup.py`) removed:
  - the fresh target directories of the review's gates and of the
    re-measurement, `targets/0751` and `targets/0751-before`;
  - the scratch directory's contents: the staged binaries, raw reports and
    `TMPDIR`;
  - the recreated detached base worktree `0751-before-src`, removed with
    `git worktree remove --force`.
- **Before the review**
  ([`superseded-f84522f8af/cleanup.json`](results/change-0751/superseded-f84522f8af/cleanup.json))
  removed 125.0 GB:
  - the target directories `targets/0751` (109.1 GB), `targets/0751-before`
    (3.0 GB), `targets/0751-before-fp` and `targets/0751-fp` (1.1 GB each);
  - 441 MB of scratch;
  - the first detached base worktree.

Kept: the worktree and branch. The frame-pointer profiles' `perf.data`
(53 MB) was not retained; `attribution/` summarizes it.
