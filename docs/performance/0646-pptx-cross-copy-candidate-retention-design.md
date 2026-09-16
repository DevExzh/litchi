# 0646: the cross-package copy can stop deflating the candidate twice, and the price is a second whole package in the plan

Status: design, retained. `performance_claim: none`. **No production code
changed on this branch.** The implementation this record measures is a scratch
patch kept in the evidence packet
([`results/change-0646/patch/`](results/change-0646/patch/)); it is not merged.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This answers queue item **8** of change
[0630](0630-queue-refresh-after-the-first-wave.md) — "the cross-package copy
still builds and deflates the candidate twice (5.96 G Ir); retaining the planned
archive in the plan, an ADR 0005 memory question" — which change
[0598](0598-pptx-cross-copy-revision-cache.md) left open in its Limitations with
the words "it changes the plan's memory profile under ADR 0005 and needs a frozen
design record". Base `c7326f68065edf6f2198ca3cb39c38c48cf00ed9`; branch
`perf/0646-pptx-cross-copy-candidate-retention-design`.

## The mechanism, and what this record decides

A plan-plus-apply lifecycle builds the candidate package **twice**: once inside
`plan_cross_slide_copy`, to prove that a complete candidate exists and to derive
`target_revision` and `target_physical_revision`, and once inside
`apply_cross_slide_copy_plan`, because application replans from the live packages
and compares the fresh plan against the durable one. Each build serializes the
candidate through `OpcPackage::to_stream`, which raw-copies the destination's
unchanged members out of its retained archive (`PreservationIndex`) and
**deflates the copied closure**. 0598 measured that deflate at 5.96 G `Ir` per
lifecycle over the two builds on the media-rich corpus and left it untouched,
because removing the second build means the plan must retain the first build's
archive.

The plan already holds the archive for the duration of planning and then throws
it away. `build_candidate` serializes into a bounded `Vec`, moves that `Vec`
into `OpcPackage::from_vec_reusing_payloads`, which authorizes it as the
package's exact source and keeps it behind an `Arc`, and
`plan_cross_slide_copy_for_slides` then drops the whole candidate package.
**Retention is not a new allocation; it is a decision not to free one.** That is
what makes it worth designing, and it is also exactly what makes it an ADR 0005
question: the bytes stop being a transient inside one call and become state
owned by a public value the caller holds for as long as it likes.

This record decides three things:

1. **What is admissible.** The plan may retain the candidate **bytes**. It may
   not retain the candidate **snapshot**. §"Which capture may be reused" states
   the difference and proves it is not a matter of taste.
2. **What it costs.** Measured below: the retained archive is 31,545 bytes on
   the plain corpus and **33,599,873 bytes** on the media-rich one — 2.00× the
   destination archive — bounded by `max_patch_bytes`, which is 128 MiB by
   default. At the default limits a single plan's owned bytes go from
   `≤ max_patch_bytes` to `≤ 2 × max_patch_bytes` = 256 MiB, which is the entire
   `Resource::Memory` limit of `Profile::Server`. Nothing charges it, because
   the PPTX opened-presentation path has no `Budget` at all.
3. **What it buys.** Measured below, with a scratch implementation.

## The frozen design

### 1. What the plan holds

`CrossSlideCopyPlan` gains one private field:

```rust
candidate: Option<RetainedCandidate>,

struct RetainedCandidate {
    /// The exact bytes `bounded_package_bytes` produced at planning, shared
    /// with the candidate package that was reopened from them.
    archive: Arc<Vec<u8>>,
    /// SHA-256 over those bytes. Used for plan equality and for the debug
    /// assertion only; application recomputes it (see §5).
    archive_digest: [u8; 32],
    /// The `max_patch_bytes` the bytes were accepted under. A different bound
    /// is not a reuse.
    bound: usize,
}
```

The `Arc` is **the one the candidate reopen already holds**.
`OpcPackage::from_vec_reusing_payloads` moves the serialized `Vec` into the
package, `authorize_owned_source` wraps it in an `Arc` and marks it the exact
source, and `build_candidate` takes a second handle to that same allocation.
Planning therefore allocates and copies **nothing new**; the only change is that
the allocation is not freed when the candidate package is dropped at the end of
planning. (The scratch patch needs one doc-hidden accessor,
`OpcPackage::exact_source_arc() -> Option<Arc<Vec<u8>>>`, to take that handle
across the crate boundary; `exact_source()` already exists and returns `&[u8]`.)

`PartialEq` must not become a byte comparison: `RetainedCandidate` implements it
by hand over `(archive_digest, bound, archive.len())`, so `plan_a == plan_b` stays
the comparison it is today.

**One finding the scratch surfaced, and it is a real constraint.**
`build_candidate` already takes eleven arguments and the workspace's
`clippy::too_many_arguments` lint is deny at twelve. Threading "which archive"
as two arguments — an `Option<&RetainedCandidate>` and a `retain` flag — breaks
that gate. The scratch therefore passes one, a three-valued
`enum CandidateArchive { Build, BuildAndRetain, Reuse(&RetainedCandidate) }`,
which is both what the lint requires and what the design wants: `Build` is
today's behaviour, `BuildAndRetain` is planning under an opt-in policy, and
`Reuse` is application. A real implementation would go further and group the
candidate's identity (slides, position, slide ID, relationship ID, layouts,
parts) into a request struct; that is a refactor this record does not specify.

### 2. What it weighs

| corpus | destination archive | **retained candidate archive** | multiple | plan's copied closure (`planned_bytes`) |
| --- | ---: | ---: | ---: | ---: |
| `pptx_cross_copy_plain` | 30,539 B | **31,545 B** | 1.03× | 1,419 B |
| `pptx_cross_copy_media_rich` | 16,814,664 B | **33,599,873 B** | 2.00× | 16,784,397 B |

The retained archive is the published output byte-for-byte: its digest is the
`output_sha256` every measured leg reports (`3e9ae2805a5d…` plain,
`6a3536fc7d90…` media-rich). In general it is
`destination_archive + deflate(copied closure) + ZIP framing`, so its ceiling is
`limits.max_patch_bytes()` — the bound `BoundedVecWriter` already enforces —
**128 MiB by default**.

No real deck in the tree exercises the large end: the largest `.pptx` under
`test-data/` is 972,788 bytes (`ArtisticEffectSample.pptx`), and the 78 fixtures
there are all under 1 MB. The media-rich corpus, at a 16.8 MB destination and
16.78 MB of deliberately incompressible copied media, is the only large sample
available, and it is synthetic. A real deck with embedded video is routinely
tens to hundreds of megabytes, so the 128 MiB ceiling is reachable in practice,
not only in principle.

### 3. The ADR 0005 memory question, stated exactly

**What the plan owns today.** `parts` is metadata (names, content types, byte
counts). `patch` holds a `Patch` whose `Delta`s carry
`ResourceState { content_type, blob: Arc<Vec<u8>>, relationships }`, and the
copied parts' `after` blobs are `Arc` clones of the **source package's** blobs
(`Part::blob_arc`). So while the source package lives, the plan adds no bytes of
its own; once the source is dropped the plan owns `planned_bytes`, which
`prepare_cross_slide_copy_for_slides` already refuses above
`limits.max_patch_bytes()`. Today's bound on a plan's owned bytes is therefore
`max_patch_bytes`, and in the ordinary case it is *shared* with a live package.

**What retention adds.** Up to `max_patch_bytes` more, **unshared**, from the
moment the plan is returned. At the default limits the worst case goes from
128 MiB to **256 MiB per plan** — which is the whole `Resource::Memory` limit of
`litchi_core::budget::Limits::for_profile(Profile::Server)` (256 MiB), and a
quarter of `Profile::Desktop`'s 1 GiB.

**What charges it: nothing.** ADR 0005 says "Every operation charges a
hierarchical resource budget supplied by an execution context", and
`litchi-core` provides `Budget`, `Resource::Memory` and the three finite
profiles. The PPTX **opened-presentation** path does not use them.
`crate::opened::model::Limits` is a separate per-operation limit type carrying
`max_parts`, `max_patch_bytes`, `max_text_bytes`, `max_history_entries` and
`max_history_bytes`: a bound on the size of one operation's *output*, not on
live memory, and not hierarchical. (`Resource::Memory` does appear in
`litchi-pptx`, but in the *source-backed* presentation path, not here.)
Retention is the first thing in this path that would hold a whole serialized
package for an unbounded, caller-controlled time, so it is the first thing that
makes that gap load-bearing rather than theoretical.

**For how long.** From `plan_cross_slide_copy` returning until the plan is
dropped. `apply_cross_slide_copy_plan` takes `&CrossSlideCopyPlan`, so applying
does **not** release the archive, and a plan may legitimately be applied more
than once — to any destination whose six revisions match, for example N clones
of one package. The design therefore adds an explicit release rather than an
implicit one:

- `CrossSlideCopyPlan::retained_candidate_bytes(&self) -> Option<usize>` so a
  caller can see what it is holding;
- `CrossSlideCopyPlan::release_retained_candidate(&mut self)` so it can stop.

Release must not be implicit on a successful apply: the plan is `&`-borrowed
there, making it implicit would need interior mutability, and it would make the
second and later applications of the same plan silently slower.

**At apply the archive is copied once.** `OpcPackage::from_vec*` takes an owned
`Vec<u8>`, so the retained `Arc` is cloned into an owned buffer and the peak
during application carries both: +33,599,873 bytes transiently on the media-rich
corpus. Removing that copy needs a shared-bytes ingress
(`OpcPackage::from_shared_vec(Arc<Vec<u8>>, …)`) in `litchi-opc`, which is an
ADR 0011 physical-ownership question and is **not** taken here. The measurement
below is therefore the conservative one: it pays the copy.

### 4. The rule that decides between rebuilding and retaining

**The crate already has this shape, and it is the shape to copy.**
`opened::Limits` carries `max_history_bytes` (256 MiB by default) and
`max_history_entries` (64), and `History::push`
(`crates/litchi-pptx/src/opened/patch.rs:694`) enforces them on live retained
undo patches: one patch alone above the byte bound is refused with
`Error::Limit { resource: "opened-presentation history bytes", limit }`, and
otherwise the oldest entries are evicted until the new one fits. So this crate
already declares a **live-memory retention ceiling** in `Limits`, already
enforces it, and already has a typed refusal for "does not fit". The retention
ceiling belongs there too, as a sibling member `max_retained_candidate_bytes`,
intersected by `intersect_limits` like every other member, rather than as a new
policy type invented for this one call. A plan cannot evict — there is exactly
one candidate — so the analogue of eviction is "do not retain", and the analogue
of `History::push`'s refusal is the `Require` route below.

That also puts the 256 MiB figure of §3 in proportion: an opened presentation
with a full history is already permitted to hold 256 MiB of retained patches at
the default limits. The difference is not the size. It is that the history bound
is **declared, intersected and enforced**, and an unconditional retention would
be none of those.

Retention is **opt-in**, with the default unchanged:

| policy | behaviour | typed refusal |
| --- | --- | --- |
| `Rebuild` (**default**) | never retain; today's memory profile and today's application cost, exactly | none |
| `RetainUpTo(bytes)` | retain iff `archive.len() <= min(bytes, limits.max_patch_bytes())`, where `bytes` is `limits.max_retained_candidate_bytes()`; otherwise return the plan unretained, observable through `retained_candidate_bytes()` | none |
| `Require(bytes)` | retain, or refuse the plan | `Error::Limit { resource: "cross-slide retained candidate archive bytes", limit: bytes }` |

**Why the default must be `Rebuild`.** ADR 0005 permits invisible caches — "Cache
behavior is semantically invisible" — and change 0598 used exactly that licence
for a 32-byte memo on an immutable snapshot. A 128 MiB hold on a public value the
caller owns is not that. It is a retention policy, and ADR 0005 puts retention
policies under the budget, not under the cache clause. Turning it on silently
would change the memory profile of every existing caller of
`plan_cross_slide_copy` with no way to observe or refuse it.

**Why `Require` exists and what it refuses.** A caller that needs the
application cost bounded — a server applying many plans under a latency
objective — would rather learn at planning time that it cannot have the
guarantee than discover at apply time that it is paying a second deflate. Its
refusal is `Error::Limit`, the same shape every other cross-slide ceiling uses,
with a new resource name. Note the gap this exposes: `litchi_pptx::Error::Limit`
carries `resource` and `limit` only, so the refusal cannot report the observed
`archive.len()`, which ADR 0005 requires of a limit error ("Limit errors identify
the resource, observed value, limit, and object path"). That is a pre-existing
shape problem in this crate, not one retention introduces, but a `Require` route
would be the first limit whose *whole purpose* is to tell the caller how big the
thing was.

**No spill, ever.** ADR 0005: "Scratch storage is an explicit capability. Litchi
never spills decrypted or sensitive content to plaintext temporary files
automatically. Supported scratch providers include memory, encrypted temporary
storage, and caller-defined stores; absence yields a typed resource error." A
candidate archive is a complete PPTX package — document content. This design
therefore never writes it anywhere: above the ceiling it rebuilds, in memory, as
today.

A caller-supplied scratch provider **could** hold the archive, and ADR 0005
permits that precisely because the provider is an explicit capability. The design
does not take that route, for a reason worth recording: ADR 0005's rule that the
*absence* of a provider "yields a typed resource error" would be wrong here.
Litchi is never obliged to hold the candidate; rebuilding it is always correct
and is what the code does today. So a scratch-backed retention must fall back to
`Rebuild` on absence, not refuse — and a design in which the capability's absence
is not an error is a different design from the one ADR 0005 describes. It also
has to clear a cost bar this one does not: an encrypted temporary store must
decrypt and read back 33.6 MB to save a 16.78 MB deflate, which is not obviously
a win, and a caller-defined store's cost is unknowable to litchi. **Scratch-backed
retention is therefore left unspecified and unmeasured.**

### 5. How application proves the retained archive is the one the plan validated

Retention changes **one line** of `build_candidate`. The in-memory candidate
graph is still built from scratch at application — the destination package is
cloned, the copied parts are added with `Arc`-shared blobs, `presentation.xml`
is rewritten — and the whole replan around it still runs every refusal it runs
today. Only `bounded_package_bytes` is skipped, and the retained bytes take its
place:

```rust
let (serialized, archive_digest) = match archive {
    CandidateArchive::Reuse(held)
        if held.bound == archive_limit && held.archive.len() <= archive_limit =>
    {
        debug_assert!(/* a fresh bounded_package_bytes reproduces both */);
        let mut serialized = Vec::new();
        serialized
            .try_reserve_exact(held.archive.len())
            .map_err(|source| Error::Allocation {
                resource: "cross-slide candidate archive",
                source,
            })?;
        serialized.extend_from_slice(&held.archive);
        let mut digest = Sha256::new();
        digest.update(&serialized);
        (serialized, digest.finalize().into())
    },
    _ => bounded_package_bytes(&candidate, archive_limit)?,
};
```

**The copy is fallible.** `BoundedVecWriter::into_bytes` reserves its compacted
buffer with `try_reserve_exact` and returns `Error::Allocation` on failure; an
infallible `Vec::clone` in its place would abort the process where the route it
replaces returns a typed error. GOAL's rule — never trade a typed refusal for a
partial result — applies to the allocation path as much as to the validation
path, and the scratch was corrected before it was measured.

**The digest is recomputed, not carried.** This is the design's central
concession and it is deliberate. Storing `archive_digest` beside the bytes and
sealing it directly would make `target_physical_revision` a value that travels
with the plan rather than a hash of the bytes about to be published; a plan whose
two halves disagreed — through memory corruption or a caller reaching a private
field — would publish an archive whose stated revision is not its own. Recomputing
costs one SHA-256 pass over bytes that `bounded_package_bytes` hashes anyway, so
**the hash is a wash and what retention removes is exactly the serialization and
its deflate**. The measured `sha2::sha256::compress256` line below confirms it:
−0.01% on the plain corpus.

The proof chain at application, in the order `apply_plan` runs it, with what each
link still checks:

| # | check | still a real check? |
| --- | --- | --- |
| 1 | `package_fingerprint(source) == plan.source_revision` | yes — pins the source graph |
| 2 | `physical_package_fingerprint(source) == plan.source_physical_revision` | yes — pins the source archive bytes |
| 3-4 | the same two for the destination | yes |
| 5 | the whole replan: provenance, signature, unknown members, macro, dialect, MCE, protection, collisions, `MAX_SLIDES`, slide IDs, part counts, master/layout graph, content types, slide surface, registered layouts, layout inheritance, owned closure, cycles, collision remap, resulting parts, planned bytes, preflight | yes — **unchanged**, this is not skipped |
| 6 | the retained bytes are re-hashed and sealed into `fresh.target_physical_revision`, compared with `plan.target_physical_revision` (an independent field of the plan) | **yes** |
| 7 | the retained bytes are reopened and captured; `fresh.target_revision` compared with `plan.target_revision` | **yes** |
| 8 | `validate_candidate` captures the reopened candidate again, runs `validate_before` on the live destination and `validate_after` on the candidate, and compares `snapshot.revision()` with `plan.target_revision` | **yes** |
| 9 | `published_archive_revision` returns the value proved at (6) and it is compared with `plan.target_physical_revision` | tautological — as it already is at the base, after change 0598 |

What is **no longer checked** is one thing only: that a fresh `to_stream` of the
in-memory candidate graph would reproduce the retained bytes. Its proof is that
(1)–(4) pin both input packages bit-exactly — graph *and* serialized archive —
`to_stream` is a deterministic function of the package, and the destination's
unchanged members are raw-copied by `PreservationIndex` out of the very archive
(4) pinned. A `debug_assert` in `build_candidate` re-derives it on every reuse in
debug and test builds, which is every one of the 871 `litchi-pptx` tests that
reaches a cross-package copy.

**Both layers are pinned by a test.** Substituting the retained archive for a
different, structurally valid PPTX archive trips the `debug_assert` in a debug
build (`cross-slide reused a retained candidate archive a fresh serialization does
not reproduce`) and, in a **release** build where that assertion is compiled out,
is refused by the recomputed proof —
`Err(Invalid("cross-slide candidate did not publish the reserved slide identity"))`
in this fixture, because `build_candidate`'s own identity check on the reopened
candidate fires before the revision comparison — leaving the destination
byte-identical. A substitution that survived that check would still be caught at
(6) and (7).

### 6. Which capture may be reused, and which may not

Each candidate build performs two captures of the same `OpcPackage`:

- **C1**, in `build_candidate`:
  `capture(&reopened, destination.limits, destination.physical_source_provenance)`,
  whose revision becomes `target_revision`.
- **C2**, in `patch::validate_candidate`:
  `capture(candidate, patch.limits, physical_source_provenance)`, whose
  `snapshot.revision() != result_revision` check is what `apply_plan` binds to
  `plan.target_revision`, and whose snapshot is returned to the caller.

**C1 is the reusable one, and change 0598 already reuses it**: `target_revision`
is `candidate_revision` rather than a second `package_fingerprint`, because a
capture's revision is `package_fingerprint` of the package it captured. That
reuse is value-identical.

**C2 may not be reused. Precisely, what C2 proves that C1 does not:**

1. **Order.** C2 runs *after* `validate_before(destination, patch)`; C1 runs
   before any staleness check against the live destination. Substituting C1 moves
   the `stale()` refusal relative to the structural refusals `capture` raises.
2. **Identity of the published package.** `validate_application_candidate` has a
   branch — `reopened && !destination.is_unmodified_owned_source()` — in which the
   prepared candidate is **dropped** and a different candidate is built from
   `destination.clone()` plus `apply_exact_revision`. C1's snapshot belongs to the
   discarded package.
3. **Limits.** C2 is taken under `patch.limits`, the *intersection* of the two
   snapshots' limits; C1 under the destination snapshot's own limits. At planning
   those differ whenever the source has tighter limits, so C1 can accept a
   candidate whose slide count C2 refuses with
   `Error::Limit { resource: "opened-presentation slides", .. }`. (Inside one
   `apply_plan` call they coincide, because application captures both live
   packages under `plan.patch.patch.limits()`. The divergence is a planning-time
   property — and it is the reason a planning-time snapshot must never be carried
   into application.)
4. **Retention changes what the two captures are captures *of*, and that is the
   reason this record cares.** Without retention, C1 and C2 at application are two
   reads of bytes serialized microseconds earlier in the same call, so skipping one
   of them loses a duplicate. With retention they are the only two places that read
   bytes produced in an *earlier* call and held since by a caller-owned value.
   Retention makes a second reuse look free — the plan already holds the archive, so
   why not hold its `Snapshot` too and skip both captures? — and that reuse would
   leave nothing at application that reads the retained bytes as a graph at all: the
   archive would be reopened, and the plan's own precomputed revision would be
   compared with itself.

So the rule this record freezes is: **a plan may retain the candidate bytes; it
may not retain the candidate snapshot.** 0598's open question — "reusing
`validate_candidate`'s second capture means threading a proven snapshot into
`validate_candidate`, whose revision check is the proof that binds the candidate
to the plan" — is answered **no** for the cross-package copy: points 1-3 refuse it
on order, published identity and limits even without retention, and point 4 says
why retention must not be allowed to revive it.

### 7. Admission gates

This design may be implemented when, and only when, every one of these holds.

| # | gate | why |
| --- | --- | --- |
| G1 | The default planning route retains nothing, proved by a test asserting `retained_candidate_bytes() == None` for `plan_cross_slide_copy`, and retention is reachable only through an explicit policy argument. | §4: a 128 MiB hold on a caller-owned public value is not an invisible cache. |
| G2 | The retained bytes are charged. Either the opened-presentation path gains a `litchi_core::budget::Budget` and charges `Resource::Memory`, **or** the ceiling is `opened::Limits::max_retained_candidate_bytes`, a sibling of the enforced `max_history_bytes`, intersected by `intersect_limits`, with `retained_candidate_bytes()` exposing what is held. | ADR 0005: every operation charges a hierarchical budget. Today this path charges none; `max_history_bytes` is the crate's only enforced live-memory ceiling and is the precedent. |
| G3 | `Require` refuses with `Error::Limit { resource: "cross-slide retained candidate archive bytes", limit }`, and no route can reach a partially retained or partially published state. | ADR 0005 typed limit errors; GOAL rule "never trade a typed refusal for a partial result". |
| G4 | No spill: a source ratchet proving no filesystem, `tempfile` or scratch-provider path is reachable from the retention code. | ADR 0005: "Litchi never spills … automatically." |
| G5 | The archive digest is **recomputed** at application over the retained bytes; the stored digest is used only for equality and the debug assertion. | §5: the physical revision must remain a hash of the bytes being published. |
| G6 | `validate_candidate`'s capture and revision check still run unconditionally on the reopened retained archive, and the plan does **not** retain a `Snapshot`. | §6, point 4. |
| G7 | Every adversarial route produces the identical `Result` with retention on and off — stale source, stale destination, foreign source, forged patch, collision remap, limit, reversibility, borrowed-provenance refusal — and the published `output_sha256` is identical on all four selectors. | The verdict set may not move. |
| G8 | Plan equality is not a byte comparison of the retained archive. | A `derive(PartialEq)` over `Arc<Vec<u8>>` would make `plan_a == plan_b` O(archive). |
| G8b | The archive decision reaches `build_candidate` as **one** argument. | It already takes eleven and `clippy::too_many_arguments` is deny at twelve. |
| G9 | `CrossSlideCopyPatch` is untouched: `LPCP0002`, its six 32-byte revisions and `to_bytes` are bit-identical, and `apply_patch` retains nothing. | The durable representation cannot carry an archive, and must not change. |
| G10 | `release_retained_candidate(&mut self)` exists and a retention probe checkpoint shows live bytes returning to the unretained baseline after it. | §3: applying does not release. |
| G11 | **Native cycles for the application phase fall** on the media-rich corpus by more than that leg's own A/A floor. A count of removed deflate is not a saving until the archive copy it trades into is measured natively (change 0612's rule). | The copy at application is real work. |
| G12 | Peak live bytes during planning do not rise, and the plan-held bytes after planning rise by exactly the archive length — proving retention shares the reopen's allocation rather than making a second one. | §1: retention must be a decision not to free, not a new copy. |
| G13 | The archive copy at application is **fallible**, returning `Error::Allocation { resource: "cross-slide candidate archive", .. }` rather than aborting. | The `BoundedVecWriter` it replaces reserves with `try_reserve_exact`; a typed refusal may not be traded for an abort. |

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0; every measured process pinned to CPU 21 while seven other
agents of the same wave built and measured on the other cores. Both legs built
`--release --locked` from `tools/perf-baseline`: before from the shared
read-only checkout of `c7326f680`, after from this branch **with the scratch
patch applied**. Both binaries were staged outside their target directories
before any measurement (change 0627). Two selectors,
`pptx_cross_copy_plain` (a 30,539-byte destination, one copied part) and
`pptx_cross_copy_media_rich` (a 16,814,664-byte destination and nine copied
parts totalling 16,784,397 bytes of incompressible media), plus their
`_lifecycle` variants which time the same work with setup and reopen included.
A "lifecycle" below is one plan plus one apply.

### Deterministic counts (callgrind isolation pairs, `--warmup 0`)

`--samples 1` and `--samples 3`, differenced over the two extra samples. Call
counts per lifecycle:

| callee, per lifecycle | plain before | plain after | media-rich before | media-rich after |
| --- | ---: | ---: | ---: | ---: |
| `package_fingerprint` (semantic revision) | 8 | 8 | 8 | 8 |
| `physical_package_fingerprint` | 4 | 4 | 4 | 4 |
| `snapshot_physical_revision` | 4 | 4 | 4 | 4 |
| `capture_internal` | 8 | 8 | 8 | 8 |
| `prepare_cross_slide_copy_for_slides` | 2 | 2 | 2 | 2 |
| `build_candidate` | 2 | 2 | 2 | 2 |
| `OpcPackage::from_vec_reusing_payloads` | 2 | 2 | 2 | 2 |
| **`bounded_package_bytes`** | 2 | **1** | 2 | **1** |
| **`PackageWriter::write_to_stream`** | 7 | **6** | 7 | **6** |
| **`PreservationIndex::write_to`** | 2 | **1** | 2 | **1** |
| **`zlib_rs::deflate::deflate`** | 40 | **20** | 1,112 | **556** |
| `sha2::sha256::compress256` | 1,709 | 1,566 | 5,462 | 4,226 |

**Exactly one serialization is removed and nothing else moves.** Every hash,
every capture, every reopen and the whole replan run the same number of times.
The one call-count row that falls without being the point is
`compress256`, and it falls *without its work falling* (below): the removed
`BoundedVecWriter` hashed the archive in ZIP-writer-sized chunks, while the
retained path hashes it in one `update`, so the same bytes reach the same
compression function in fewer, larger batches.

### Instructions per lifecycle

| inclusive `Ir`, per lifecycle | `pptx_cross_copy_plain` | `pptx_cross_copy_media_rich` |
| --- | ---: | ---: |
| **whole lifecycle** | 239,585,652 → **234,548,824** (**−2.10%**) | 26,800,703,086 → **23,736,065,870** (**−11.43%**) |
| `bounded_package_bytes` | 14,177,765 → 7,096,450 (**−49.95%**) | 9,634,402,864 → 4,817,182,700 (**−50.00%**) |
| `PreservationIndex::write_to` | 10,988,511 → 5,496,190 (−49.98%) | 9,624,120,794 → 4,812,044,592 (**−50.00%**) |
| **`zlib_rs::deflate::deflate`** | 2,492,079 → 1,246,063 (**−50.00%**) | 5,955,603,242 → **2,977,783,594** (**−50.00%**) |
| `PackageWriter::write_to_stream` | 20,591,086 → 13,545,970 (−34.21%) | 13,164,161,022 → 8,349,845,603 (−36.57%) |
| `build_candidate` | 45,153,116 → 39,913,666 (−11.60%) | 13,474,961,554 → 10,409,963,212 (−22.75%) |
| `prepare_cross_slide_copy_for_slides` | 133,917,304 → 128,752,209 (−3.86%) | 15,315,609,365 → 12,250,625,815 (−20.01%) |
| `package_fingerprint` | 39,455,467 → 39,467,050 (+0.03%) | 10,525,045,950 → 10,525,048,555 (+0.00%) |
| `physical_package_fingerprint` | 6,502,294 → 6,500,121 (−0.03%) | 3,501,940,703 → 3,501,940,056 (−0.00%) |
| `capture_internal` | 58,659,408 → 58,920,598 (+0.45%) | 8,804,515,388 → 8,804,657,832 (+0.00%) |
| **`sha2::sha256::compress256`** | 50,665,038 → 50,658,460 (**−0.01%**) | 19,274,528,456 → **19,274,471,600 (−0.00%)** |
| `__memcpy_avx_unaligned_erms` | — | 1,074,024,902 → **936,377,586 (−12.82%)** |

Whole child at `--samples 1`: plain 4,323,171,238 → 4,313,031,318 (−0.23%);
media-rich 123,572,072,250 → 117,447,329,387 (−4.96%). Most of the whole child
is the harness's corpus construction and its ten untimed refusal gates.

**The arithmetic closes on the media-rich corpus, and it says exactly what was
traded for what.** One `bounded_package_bytes` is worth 4,817,220,164 `Ir`. Of
that, the deflate of the copied closure is 2,977,819,648 and the SHA-256 of the
33.6 MB archive is about 1.75 G at the 52.1 `Ir`/byte software rate this build
shows; the rest is ZIP framing and the raw member copy. The retained path gives
back the hash — hence `compress256` unchanged to five significant figures — and
adds one fallibly reserved 33,599,873-byte copy. Net: **−3,064,637,216 `Ir` per
lifecycle**, of which 2,977,819,648 is the deflate and the remaining 86,817,568
is the framing: **the deflate plus the framing, and nothing else.**

**Copying falls rather than rises.** The archive copy retention introduces is
smaller than the copying the removed serialization performed:
`BoundedVecWriter` copied every accepted chunk into its buffer and then
compacted the result into an exactly-reserved second buffer, so
`__memcpy_avx_unaligned_erms` drops 137,647,316 `Ir` per lifecycle (−12.82%)
even though a whole archive copy was added. That term is the one callgrind
prices worst — change 0604 measured a 35× overstatement of `rep movsb`-class
copies — so it is reported as a direction, not a size, and the `perf stat` rows
below are the counterweight.

### Memory: what the plan holds, exactly

`litchi-perf-baseline-alloc retention --api owned`, 5 samples after 1 warmup per
leg, medians; the allocator counters are deterministic. The probe checkpoints an
owned cross-copy lifecycle at nine ownership boundaries.

| checkpoint, live bytes | plain before | plain after | Δ | media-rich before | media-rich after | Δ |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| baseline | 161,388 | 161,386 | −2 | 68,909,261 | 68,909,259 | −2 |
| inputs and sink ready | 352,051 | 352,049 | −2 | 169,804,828 | 169,804,826 | −2 |
| documents opened | 689,104 | 689,102 | −2 | 203,743,965 | 203,743,963 | −2 |
| **planned** | 696,934 | 728,517 | **+31,583** | 203,757,827 | 237,357,738 | **+33,599,911** |
| published | 782,847 | 814,430 | +31,583 | 237,438,832 | 271,038,743 | +33,599,911 |
| result dropped | 748,924 | 780,507 | +31,583 | 237,392,085 | 270,991,996 | +33,599,911 |
| **plan dropped** | 741,094 | 741,092 | **−2** | 237,378,223 | 237,378,221 | **−2** |
| document handles dropped | 290,014 | 290,012 | −2 | 136,174,543 | 136,174,541 | −2 |
| sink dropped | 161,388 | 161,386 | −2 | 68,909,261 | 68,909,259 | −2 |

**The delta closes to the byte, and it is the archive plus an `Arc` header.**
The two binaries differ by a constant −2 bytes of static allocation, visible at
every checkpoint where the plan holds nothing. At `planned`:

- plain: 31,545 (archive) + 40 (`ArcInner<Vec<u8>>`: two counters plus the `Vec`
  triple) − 2 = **31,583**;
- media-rich: 33,599,873 + 40 − 2 = **33,599,911**.

Nothing else is retained. **Retention is a decision not to free, not a second
copy** — G12 is satisfied — and **dropping the plan releases it exactly**: at
`plan dropped` the two legs are back to the constant −2.

| aggregate over the lifecycle region | plain before | plain after | Δ | media-rich before | media-rich after | Δ |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| **region peak live bytes** | 1,339,600 | 1,266,634 | **−5.45%** | 271,884,201 | 305,008,696 | **+12.18%** |
| allocation calls | 48,096 | 47,041 | −2.19% | 58,250 | 56,770 | −2.54% |
| reallocation calls | 5,659 | 5,298 | −6.38% | 6,765 | 6,271 | −7.30% |
| allocated bytes | 15,581,584 | 13,175,380 | **−15.44%** | 467,072,982 | 373,765,695 | **−19.98%** |

**The peak moves in opposite directions on the two corpora, and that is the
design's real shape.** On the plain corpus the peak *falls* 5.45%: the removed
second `BoundedVecWriter`, with its geometric growth and its compaction copy, is
worth more than the 31,545 bytes retained. On the media-rich corpus the peak
rises by 33,124,495 bytes — one archive, to within 1.4% — and the rise is not
new work at planning: planning's own transient peak is unchanged in both legs,
because both legs build the archive there. What rises is everything downstream
of `planned`, because the allocation that used to be freed at the end of
planning now sits underneath the application's own peak.

Total bytes passing through the allocator fall by **19.98%** on the media-rich
lifecycle (93,307,287 bytes), which is the removed serialization's churn: one
33.6 MB archive built by doubling and then compacted.

### Paired native timing

Four legs in run order **A1 B1 B2 A2** (before, after, after, before), all four
selectors in one process per leg, 30 measured samples after 3 warmups,
`taskset -c 21`. A/A is A2 against A1 and B/B is B2 against B1, both inside the
same window. Run-window load average 11.3 falling to 5.3; the floors below are
correspondingly tight.

| selector | before p50 | after p50 | p50 Δ | p95 Δ | p99 Δ | A/A p50 | B/B p50 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_cross_copy_plain` | 7.932 ms | 7.695 ms | **−2.99%** | −3.91% | −4.72% | −1.34% | −0.46% |
| `pptx_cross_copy_media_rich` | 659.101 ms | 412.477 ms | **−37.42%** (before is +59.79% of after) | −37.42% | −37.29% | **−0.15%** | +0.22% |
| `pptx_cross_copy_plain_lifecycle` | 9.259 ms | 8.939 ms | **−3.46%** | −4.08% | −3.57% | +0.31% | −0.88% |
| `pptx_cross_copy_media_rich_lifecycle` | 669.283 ms | 430.942 ms | **−35.61%** (before is +55.31% of after) | −36.24% | −36.30% | −3.44% | −0.02% |

The published digest is one value per corpus across all four legs
(`3e9ae2805a5d…`, `6a3536fc7d90…`).

**The effect is in `apply`, and it is far larger than the instruction counts
suggest.** The selectors clock plan and apply separately. Per leg, on
`pptx_cross_copy_media_rich`:

| phase p50 | A1 (before) | A2 (before) | B1 (after) | B2 (after) | Δ |
| --- | ---: | ---: | ---: | ---: | ---: |
| **apply** (reported as `commit`) | **339.427 ms** | **337.876 ms** | **94.427 ms** | **94.540 ms** | **−72.10%** |
| plan | 313.467 ms | 313.936 ms | 311.471 ms | 312.475 ms | −0.64% |
| publication | 6.513 ms | 6.497 ms | 6.301 ms | 6.149 ms | −3.88% |
| reopen | 21.244 ms | 21.141 ms | 21.168 ms | 21.237 ms | −0.02% |

Apply falls **−72.10%** with a before-pair spread of 0.46% and an after-pair
spread of 0.12%. The lifecycle variant agrees: **337.258 → 94.531 ms, −71.97%**.
**Planning is flat** (−0.64%), which is what identical code on identical inputs
should do: retention changes nothing about how a plan is built, only what happens
to the archive afterwards. Reopen and publication are flat too.

Instructions per lifecycle fell 11.43% and the wall clock of the phase that
changed fell 72.10%, and those are consistent rather than contradictory.
Callgrind prices SHA-256 in software, so it attributes 19.27 G of the 26.80 G
lifecycle — **72%** — to hashing that SHA-NI makes about five times cheaper
natively. What retention removes is one deflate of 16.78 MB of deliberately
incompressible media, which callgrind prices at 2.98 G `Ir` (11% of the
lifecycle) and the CPU prices at about 244 ms. The instruction ranking and the
clock are measuring different things, exactly as the program's own rule says, and
here the clock is the one that matters.

**One phase is bimodal, in both directions, and it is not the change.**
`publication_ns` on `pptx_cross_copy_media_rich_lifecycle` reads, per leg,
A1 1.536, B1 1.529, B2 1.527, **A2 6.545** ms — a *before* leg four times slower
than its own pair. In an earlier window, measured on a superseded revision of the
scratch, the same phase read A1 1.632, **B1 6.324**, B2 1.571, A2 1.594 ms, with
an *after* leg as the outlier. The phase therefore sits at either ≈1.53 ms or
≈6.5 ms independently of which binary runs it; the non-lifecycle selector, whose
publication is measured in the same processes, reads 6.51 / 6.30 / 6.15 / 6.50 ms
across all four legs. A2's excursion also drags its `plan` down to 295.9 ms
(−5% against its own pair), which is why that selector's A/A floor is −3.44%
where the non-lifecycle selector's is −0.15%. Publication executes no changed
code, and **the pooled ±4× on this phase is reported as leg variance, not as a
result in either direction.**

**Native `perf stat`**, whole child, the same A1 B1 B2 A2 order over all four
selectors in one process at `--warmup 1 --samples 10`:

| event | before mean | after mean | Δ | A/A | B/B |
| --- | ---: | ---: | ---: | ---: | ---: |
| **cycles** | 103,501,330,199 | 79,587,933,858 | **−23.10%** | −0.77% | −0.18% |
| instructions | 226,998,397,412 | 159,109,153,677 | −29.91% | −0.27% | +0.03% |

This is admission gate **G11**, satisfied: the removed deflate is worth 23.10% of
whole-child native cycles against a 0.77% floor, and the archive copy it trades
into does not eat the saving. It is also the counterweight to the callgrind
numbers: native instructions fall 29.91% where callgrind's per-lifecycle total
falls 11.43%, because SHA-NI executes the unchanged hashing in a small fraction
of the instructions valgrind's software implementation needs, which raises
deflate's share of the native mix in exactly the way the wall clock shows.

## Correctness evidence

**The whole `litchi-pptx` suite runs on the retained path.** The scratch
implementation retains unconditionally under a 64 MiB ceiling, so every test in
the crate that plans and then applies a cross-package copy exercises the reuse,
and every one of them re-serializes the candidate under the `debug_assert`.

| leg | command | result |
| --- | --- | --- |
| before (untouched checkout of `c7326f680`) | `cargo test -p litchi-pptx` | **871 passed, 0 failed, 2 ignored** across 77 suites |
| after (scratch patch applied) | `cargo test -p litchi-pptx` | **874 passed, 0 failed, 2 ignored** across 77 suites |
| after | `cargo test -p litchi-opc` | 707 passed, 0 failed, 1 ignored across 26 suites |

874 − 871 = 3, the three binding-proof tests the patch adds. (`litchi-opc` is
listed because the patch adds one doc-hidden accessor there; it adds no test and
changes no behaviour, so the suite is a no-regression check, not evidence.) **Every test that
passes on the base passes on the retained path, with the same assertions.** The
pre-existing adversarial suites are the substance of that: `source_backed_cross_copy.rs`,
`source_backed_cross_copy_adversarial.rs` and change 0598's
`revision_cache_tests.rs` assert exact `Error` kinds, operations and reason
strings for stale sources, stale destinations, foreign sources, forged patches,
collision remaps, limits and reversibility, and none of them moved.

The three tests the patch adds
(`crates/litchi-pptx/src/opened/cross_copy_plan/retention_scratch_tests.rs`, in
the packet):

| test | what it pins |
| --- | --- |
| `a_plan_retains_the_archive_its_application_publishes` | The plan's retained bytes seal — through `seal_physical_revision` — to the plan's own `target_physical_revision`, `retained_candidate_bytes()` reports their length, and the archive the application publishes is byte-identical to them. |
| `a_substituted_retained_archive_trips_the_debug_assertion` (debug) / `…_is_refused_and_publishes_nothing` (release) | Replacing the retained archive with a different, structurally valid PPTX archive trips the `debug_assert` in a debug build and, in a release build where that assertion is compiled out, is refused — here by `Err(Invalid("cross-slide candidate did not publish the reserved slide identity"))` — leaving the destination byte-identical. This is the test that fails if the retained archive were ever trusted without a recomputed proof. |
| `retention_changes_no_apply_verdict` | The same application with and without the retained archive publishes identical bytes and identical revisions on the accepting route, and produces the identical `Result` on the drifted-destination and foreign-source routes. The unretained plan is the same plan with `candidate` cleared, so the two runs differ in exactly one bit of state. |

**Byte identity across every measured leg.** All four selectors verify their
published archive against an expected digest and report it; across the four
timing legs and the eight callgrind profiles the reported `output_sha256` is one
value per corpus — `3e9ae2805a5dfd8e57ceff4a3662a2772f4197850252f7a23df8e6525d58ee22`
for the plain corpora and
`6a3536fc7d9055eefad6eb115294f0157b7f8acd576a4898b4e402e48c3e7b4f` for the
media-rich ones — identical between the legs. The harness refuses to run any of
them without its ten untimed corpus gates (semantic output, package topology,
dependency closure, source immutability, collision remap, durable patch round
trip, borrowed-provenance refusal, stale-source refusal, stale-destination
refusal, foreign-source refusal), so each measured run is also a verdict check.

## Why this is designed and not built

Three of the fourteen gates are not satisfiable inside this change, and one of
them is not satisfiable inside this crate.

**G2 is the blocker.** ADR 0005 requires every operation to charge a
hierarchical budget, and the opened-presentation path charges none. Retention
is the first construct in this path whose *size* is the point: a 32-byte memo
(0598) needs no budget, a 128 MiB archive does. Adding a `Budget` to
`opened::` is not a cross-copy change — it touches `capture`, `Limits`, every
`Snapshot` and every opened transaction — and
[proposed ADR 0031](../adr/0031-execution-context-budgets.md), which would
define where such a budget comes from, is **not accepted** and may not be cited
as authority. Until the budget exists, the honest alternative is G2's second
branch: `opened::Limits::max_retained_candidate_bytes`, the sibling of the
already-enforced `max_history_bytes`. That is implementable today and needs no
new concept — but `Limits::new` is a public constructor taking exactly five
arguments, so adding a sixth member is a **breaking** public API change, and
`CrossSlideCopyPlan` gains two public methods besides. A breaking public API
change is a decision the performance program does not get to make on its own.

**G1 and G3 are the same decision seen twice.** The saving only exists for a
caller who opts in, so the shape of the opt-in *is* the change. The scratch
implementation measured above deliberately does the opposite — it retains
unconditionally under a 64 MiB ceiling — because the question this record had to
answer first is whether the saving is worth asking for at all. It is; the answer
to the second question, what the API should look like, belongs to whoever owns
`litchi-pptx`'s public surface.

**The other eleven gates are met, or are met by construction in the scratch.**
G11 in particular is met with room: native whole-child cycles fall 23.10%
against a 0.77% A/A floor, so the archive copy retention trades into does not eat
the deflate it removes. G5, G6, G8, G8b, G9, G12 and G13 are properties of the
patch and are pinned by the measurements and tests above; G4, G7 and G10 are
satisfied by the scratch's shape but would need the ratchet, the differential and
the accessor spelled out in a real implementation.

What this record therefore leaves is a design that can be implemented without
re-deriving any of it, fourteen gates that say when, and a patch that has already
been run against the whole `litchi-pptx` suite.

## ADR compliance

**ADR 0005 (I/O, memory, measured performance).** The memory question is stated
in §3 and gated by G2. The cache clause — "Cache behavior is semantically
invisible" — is deliberately **not** used as the licence: §4 explains why a
128 MiB hold on a caller-owned public value is a retention policy rather than a
cache, and why the default must stay `Rebuild`. The scratch-storage clause —
"Litchi never spills … automatically … absence yields a typed resource error" —
is honoured by never spilling (G4), and §4 records why a scratch-backed variant
would need a different fallback rule from the one ADR 0005 states, because
rebuilding is always available. The limit-error clause ("resource, observed
value, limit, and object path") is met in kind but not in full by the existing
`litchi_pptx::Error::Limit`, which carries no observed value; §4 names that gap.

**ADR 0003 (snapshots, edits, patches).** No revision value, proof format or
durable encoding changes. `CrossSlideCopyPatch`, its `LPCP0002` magic and its six
32-byte revisions are untouched, and `apply_patch` retains nothing (G9). The plan
remains an immutable value; the new field is private and does not change what
`plan == plan` means (G8).

**ADR 0011 (OOXML physical package ownership) and the 2026-08-21 OPC
exact-source amendment.** The retained archive is the exact source the candidate
reopen already authorized; nothing about exact-source authorization, preservation
provenance or the "planning evidence only" rule changes. The one place this
design brushes ADR 0011 is the copy at application (§3): eliminating it would
need a shared-bytes ingress in `litchi-opc`, which is an ownership decision, and
this design does not take it.

**ADR 0006 / change 0454.** The "dedup only with proven equivalence" rule is
untouched: the copy still copies the source closure byte-for-byte, still refuses
to reuse a destination member on byte equality alone, and still proves the layout
inheritance graph before reusing a layout. The physical revision 0454 introduced
is still SHA-256 over `litchi-pptx-cross-physical-v2`, the `u64` archive length
and the archive digest, and §5 keeps it a hash of the bytes about to be published.

**`docs/GOAL.md`.** Optimization-order step 2 (unnecessary I/O, decompression,
recompression). No new `unsafe`, no weakened limit or malformed-input defence, no
hidden global pool, no ambient I/O, no public leakage of archive types, locks or
executors. Proposed ADRs 0030 and 0031 are not cited as authority; ADR 0031 is
named in §"Why this is designed and not built" only as the thing that is *not*
available.

## Validation preserved

Nothing is removed, weakened, reordered or made conditional. The replan at
application runs every refusal it runs today (§5, row 5); both captures still
run; `validate_before`, `validate_after` and `validate_candidate`'s revision
check still run in the same order on the same values; the four staleness proofs
are still taken from the live packages before any capture;
`reject_unknown_non_part_members` still runs on every physical read. The archive
bound is still `max_patch_bytes`, enforced at planning by `BoundedVecWriter` and
re-checked at application before a retained archive is accepted
(`held.bound == archive_limit && held.archive.len() <= archive_limit`), so a
tighter bound at application is a miss and rebuilds under the tighter writer.
The allocation refusal is preserved too: the archive copy reserves with
`try_reserve_exact` and returns the same
`Error::Allocation { resource: "cross-slide candidate archive", .. }` the writer
it replaces returns, rather than aborting on an infallible `Vec::clone`.

## Limitations — what is not claimed

- **No speedup is claimed and none is registered.** `performance_claim: none`,
  no claim-registry entry. The wall-clock figures are reported beside the A/A
  and B/B floors measured in the same window, and where a delta is inside its
  own floor the record says so rather than averaging it away.
- **Nothing here is a statement about the shipped library.** No production code
  changed on this branch. Every number describes one scratch implementation of
  one design, on two synthetic corpora, on this host, in these two builds. The
  scratch retains unconditionally under a 64 MiB ceiling; the policy §4
  specifies (`Rebuild` / `RetainUpTo` / `Require`) is neither implemented nor
  measured, and the default it specifies would make the saving reachable only by
  a caller who asks for it.
- **The counts and allocator figures are exact; the interpretation of the
  instruction totals is not.** Instructions rank work, not latency. Callgrind
  runs SHA-256 in software because valgrind masks the SHA CPUID bit, so the
  19.27 G `Ir` of `compress256` on the media-rich lifecycle is roughly five times
  its native cycle share; and callgrind prices `rep movsb`-class copies per byte,
  which change 0604 measured at a 35× overstatement, so the memcpy row is a
  direction and not a size. The `perf stat` rows are the counterweight and they
  are whole-child, so they dilute rather than isolate.
- **No real deck exercises the large end.** The largest `.pptx` under
  `test-data/` is 972,788 bytes and all 78 are under 1 MB. The media-rich corpus
  is generated by `litchi-pptx-cross-slide-copy-evidence-v1` and is deliberately
  incompressible, which maximizes both the deflate that retention removes and the
  archive that retention holds. A deck whose copied closure compresses well would
  save less deflate and retain a smaller archive; one with more copied media would
  do the opposite. **No shape of real deck was measured.**
- **The proof this design replaces with an argument is named, not eliminated.**
  A fresh `to_stream` of the in-memory candidate is no longer compared against the
  retained bytes at application. The argument that it must agree rests on the four
  staleness proofs plus determinism of `to_stream`, and it is re-derived by a
  `debug_assert` in debug and test builds only. A release build trusts the
  argument, backed by the recomputed graph and archive revisions of §5 rows 6-8.
- **The saving is release-only, by construction.** The `debug_assert` that
  re-derives the retained archive re-runs the serialization it removes, so a debug
  or test build does strictly more work than the base, not less. That is the right
  trade for an argument-backed reuse — it is how the 874-test suite becomes the
  proof — but it means no debug-build measurement of this design is meaningful.
- **`apply_patch` is unchanged and unmeasured.** Retention is specified for
  `CrossSlideCopyPlan` only; the durable `CrossSlideCopyPatch` carries no archive
  by construction, so a patch-driven application still builds and deflates the
  candidate exactly as today. The inverse-patch route, which additionally clones
  the destination and replans, is untouched.
- **The `Arc` accessor the scratch adds is scratch.**
  `OpcPackage::exact_source_arc` is `#[doc(hidden)]` and exists to take a shared
  handle across the crate boundary. A real implementation needs that seam to be
  designed, and the zero-copy application variant needs more than that — a
  shared-bytes ingress — which is an ADR 0011 question this record does not answer.
- **Not measured at all:** the source-backed cross-copy plan
  (`SourceBackedCrossSlideCopyPlan`; no selector exists), the external-package
  copy binary (its pinned LibreOffice QA fixture is not in the tree — change 0454
  fetched it over the network), process RSS, cold-cache, range-source,
  concurrency and cross-platform behaviour, and the plan-only lifetime of a plan
  that is never applied, which is the case retention costs most and saves nothing.

## Retained evidence

[`results/change-0646/README.md`](results/change-0646/README.md) — the scratch
patch, both legs' callgrind profiles and annotations, the isolation-pair count
and instruction tables, the allocator retention probe for both corpora, every
paired timing and `perf stat` report, both legs' `litchi-pptx` suite output, the
gates, the decision record and the log paragraphs.
