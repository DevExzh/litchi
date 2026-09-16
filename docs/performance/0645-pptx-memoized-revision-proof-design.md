# 0645: the memoized per-part revision proof, frozen — it is worth 26% of an opened PPTX lifecycle, it costs a durable format bump, and it makes one selector slower

Status: retained, design only. `performance_claim: none` — the numbers below are
callgrind isolation-pair counts and paired native timings reported beside the
host's A/A floor, as evidence rather than as a registered claim. **No file under
`crates/` is modified by this change.** The scratch implementation that produced
the measurement lives only in the measurement worktree and is retained as a
patch in the evidence packet.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is item **7** of change [0630](0630-queue-refresh-after-the-first-wave.md)'s
queue — part **(c)** of **PPTX-1**, the one part of change
[0590](0590-pptx-opened-transaction-revision-reuse.md) that its author declined
to implement because it redefines a value that is serialized into two durable
patch formats.

## What this record carries

Nothing in production. Three things:

1. the **frozen design** — the two-tier proof, the memo's key and its retention
   rule, the `litchi-pptx-opened-v2` domain strings, the `LPRM0002` / `LPCP0003`
   magic bump, the typed refusal for a patch serialized under the old format,
   and the migration policy (there is none: re-plan);
2. the **sizing measurement**, from a scratch implementation built in the
   measurement worktree and never committed as production code — hashes and
   bytes per lifecycle, instructions, native cycles and paired timing with the
   floor, on change 0590's four retained selectors;
3. the **admission gates** a later batch must clear before any of this lands.

## The proof, today

`package_fingerprint` (`crates/litchi-pptx/src/opened/model.rs:458`) is a single
SHA-256 over one flat, length-prefixed encoding of the complete OPC graph:

```
H( "litchi-pptx-opened-v1"
 ‖ root relationships, sorted by rId
 ‖ "non-part-members" ‖ u32 count ‖ (name, reason)*
 ‖ for each part in sorted part-name order:
       name ‖ content type ‖ payload ‖ relationships sorted by rId )
```

Every variable-length field is written as a `u64` little-endian length followed
by the bytes (`feed`, `model.rs:576`), and relationship runs carry a `u32`
count, so the encoding is self-delimiting. Call this encoding `enc₁` and the
tuple it encodes `ℐ(P)` — the *fingerprint input* of package `P`.

The cost is that `ℐ(P)` contains every payload byte, so **every recapture hashes
the complete package**, media included. On the harness's 229-part, 17,568,429-byte
eager deck one pass is **919,619,676 `Ir`** under callgrind, and change 0590 left
four of them in a `pptx_eager_batch_edit_save` lifecycle.

## The design

### The two-tier proof

Replace `enc₁` with `enc₂`, which keeps the header exactly and replaces the flat
per-part field run with a fixed-width sequence of per-part digests:

```
payload digest   Dᵇ(b)     = H( "litchi-pptx-opened-payload-v2" ‖ len ‖ b )
part digest      Dᵖ(p)     = H( "litchi-pptx-opened-part-v2"
                              ‖ name ‖ content type ‖ Dᵇ(payload)
                              ‖ relationships sorted by rId )
revision         R₂(P)     = H( "litchi-pptx-opened-v2"
                              ‖ root relationships, sorted by rId
                              ‖ "non-part-members" ‖ u32 count ‖ (name, reason)*
                              ‖ "parts" ‖ u32 part count ‖ Dᵖ(p₁) ‖ … ‖ Dᵖ(p_k) )
```

`Dᵇ` — and only `Dᵇ` — is memoized.

**Why the payload tier is separate.** The obvious design memoizes the whole
per-part digest keyed on blob identity. That is a correctness trap: a part's
content type and relationships are not behind the blob `Arc`, so a part whose
relationships changed while its payload `Arc` did not would return a stale
digest. Splitting the tiers keeps the memo's key and the memoized value over
exactly the same bytes. The content type, the part name and the relationships
are re-fed on every pass; on the eager deck that is 93,763 bytes against
17,668,953 — **0.53%** of the feed (measured, `counts/memo-per-fingerprint.txt`).

### The memo

```rust
struct PartDigests {
    entries: HashMap<(usize /* payload address */, usize /* length */),
                     (Arc<Vec<u8>>, [u8; 32])>,
}
```

stored on `Snapshot` as `Arc<PartDigests>` beside the `physical_revision` cache
change [0598](0598-pptx-cross-copy-revision-cache.md) already put there.

**The retained `Arc` is load-bearing, not an optimization.** An entry asserts
"the allocation at address `a` of length `ℓ` has payload digest `d`". Without a
strong reference the allocation can be freed and a *different* payload can be
allocated at the same address — the ABA hazard — and the memo would answer with
a digest of bytes that no longer exist. Holding the `Arc` makes the address
unrecyclable for the entry's lifetime, so a hit proves allocation identity and
therefore byte identity. The one degenerate case, `ℓ = 0`, is safe in the other
direction: two distinct empty `Vec`s may share a dangling address, and they have
the same (empty) payload, so a "false" hit returns the right digest.

**The memo pins nothing the snapshot does not already own.** Invariant: *every
entry of a snapshot's memo names an allocation that snapshot's own
`Arc<OpcPackage>` holds.* It is established at capture — the memo is filled while
walking that package's own parts — and preserved by `PartDigests::project`,
which `Snapshot::rebound_to` uses to carry the memo onto a content-equal
package: `project` re-keys onto the new package's own allocations and **drops**
every entry whose allocation the new package does not hold. (0598's
`physical_revision` cache is *not* inherited across a rebind, for a different
reason — `packages_equal` says nothing about ZIP ordering. The part-digest memo
can be projected because it makes a claim about payload bytes only.)

**`Part` is a public trait, so `blob_arc()` cannot be trusted.**
`litchi-opc` already knows this: `from_vec_reusing_payloads`
(`crates/litchi-opc/src/package.rs:435`) refuses a donor payload unless
`std::ptr::eq(blob.as_slice(), visible) || blob.as_slice() == visible`, with the
comment "a custom part cannot donate storage inconsistent with its blob", and
`payload_reuse_tests.rs` carries a `MismatchedBlobArcPart` that lies. The memo
applies the same alias test before keying anything: a part whose `blob_arc()`
does not alias its `blob()` is **never memoized** and is hashed exactly as v1
hashes it. A hostile or merely inconsistent part can therefore cost performance
and can never change a value.

**ADR 0005.** "Cache behaviour is semantically invisible." A miss is an ordinary
hash; the memo holds no limit, raises no error the uncached path would not
raise, never travels to a live mutable package, and is rebuilt rather than
mutated. The reused digest is re-derived by a `debug_assert!` in test and debug
builds.

**Resident cost, measured.** `entries.capacity() × (16 + 40 + 1)` bytes:

| parts | entries | memo bytes | per entry |
| ---: | ---: | ---: | ---: |
| 29 | 29 | 3,192 | 110.1 |
| 85 | 85 | 6,384 | 75.1 |
| 221 | 221 | 12,768 | 57.8 |

The per-entry figure varies only because `HashMap` capacity grows in steps (the
three rows have capacity 56, 112 and 224 for 29, 85 and 221 entries); the
per-slot cost is a flat 57 bytes. A 229-part deck costs about **13 KiB** per
snapshot on top of the 17 MB it already holds — 0.08%. Snapshot clones share the
memo through the `Arc`.

### The format bump

`package_fingerprint`'s return value is not private. It is serialized into two
durable formats:

| format | magic | revisions in the header | where |
| --- | --- | --- | --- |
| `SlideRemovalPatch` | `LPRM0001` | 2 × 32 B (source, target) | `opened/remove_plan.rs:26` |
| `CrossSlideCopyPatch` | `LPCP0002` | 6 × 32 B, of which **3 are semantic** | `opened/cross_copy_plan.rs:26` |

Redefining the hash changes both. The bump is **`LPRM0002`** and **`LPCP0003`**.
`LPCP`'s other three revisions are *physical* — SHA-256 over the serialized
archive under `litchi-pptx-cross-physical-v2` (0598) — and this design does not
touch them; the magic still bumps because the semantic triple beside them
changes.

**What an old patch does under the new build: a typed refusal, at parse.**

```rust
/// Durable patch families and the revision proof they carry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DurablePatchFormat {
    SlideRemovalV1,   // LPRM0001, litchi-pptx-opened-v1
    SlideRemovalV2,   // LPRM0002, litchi-pptx-opened-v2
    CrossSlideCopyV2, // LPCP0002, litchi-pptx-opened-v1
    CrossSlideCopyV3, // LPCP0003, litchi-pptx-opened-v2
}

/// A durable patch carries a superseded complete-package revision proof.
#[error("PresentationML durable patch carries the superseded revision proof of \
         format {found}; this build reads {expected}. Re-plan the edit against \
         the source package.")]
DurablePatchRevisionFormat {
    found: DurablePatchFormat,
    expected: DurablePatchFormat,
},
```

Three properties are deliberate:

- **It fires in `from_bytes` / `from_bytes_with_limits`, not at apply.** The
  caller must never hold a `SlideRemovalPatch` whose public
  `source_revision() -> [u8; 32]` is a value from a foreign algebra. Refusing at
  parse means no mutation is attempted and no comparison is ever made across
  versions.
- **It names the format and says what to do.** A silent mismatch — the patch
  parsing and then failing its source check with `Error::UnsafeEdit {
  reason: "the complete package graph differs from the slide-removal patch
  source" }` — would tell the caller the *package* moved when in fact the
  *proof format* moved. That is exactly the failure the brief forbids, and it is
  what happens if the magic is not bumped.
- **Unrecognized bytes keep their present refusal.** Only a *recognized but
  superseded* magic (`LPRM0001`, `LPCP0002`) produces the new variant; anything
  else still returns `Error::Invalid("… has an unsupported version")`, as it
  does today. The error identity therefore moves for exactly the set of inputs
  that used to be accepted, and for no others. `DurablePatchFormat` deliberately
  has no `Unrecognized` variant carrying input bytes: nothing attacker-supplied
  reaches the message.

`litchi_pptx::Error` is `#[non_exhaustive]` (`crates/litchi-pptx/src/error.rs:178`),
so the variant is additive for downstream matching.

**Migration policy: none, and none is possible.** No dual read, no rewriter.

- A rewriter would have to compute the v2 revisions of the source and target
  packages, which requires **holding those packages**. A holder of the source
  package can already produce a v2 patch by re-running
  `Snapshot::plan_slide_removal` / `plan_cross_slide_copy`. A migration tool
  would be re-planning under another name.
- A dual read would put two revision algebras in one process behind one
  untyped `[u8; 32]`. `Snapshot::revision()`, `SlideRemovalPatch::
  source_revision()`, `SlideRemovalPlan::source_revision()`,
  `SlideCopyPlan::source_revision()` and the six `CrossSlideCopy*` accessors are
  all public and all return a bare `[u8; 32]`; a caller comparing a v1 patch's
  source revision against a v2 snapshot's revision would get a false negative
  with no type error. ADR 0003 requires patches to be *exact-source-checked* with
  deterministic conflicts; a proof in a superseded algebra cannot be
  source-checked at all.
- A dual read would also keep the whole-payload hash alive on the path the
  design exists to make cheap, and would have to run it on the live package,
  where no memo applies.

## Exactly what changes

### Revision **values** change — everywhere

Every producer and consumer of `package_fingerprint`. The complete list at
`c7326f680`:

| site | role |
| --- | --- |
| `opened/model.rs:458` `package_fingerprint` | the definition |
| `opened/model.rs:444` `capture_internal` | every `Snapshot::revision()` |
| `opened/transaction.rs:57` `Transaction::is_changed` | dirty check |
| `opened/transaction.rs:1200,1211` `Transaction::commit` | committed revision, and the `unsign()` re-hash |
| `opened/copy_plan.rs:158` `SlideCopyPlan::source_revision` | same-package copy authorization |
| `package/model.rs:228` `apply_slide_copy_plan` | its staleness check |
| `opened/remove_plan.rs:358` | `SlideRemovalPlan` / `SlideRemovalPatch` target revision |
| `opened/remove_plan.rs:381` `apply_patch` | its source check |
| `opened/cross_copy_plan.rs:458,465,548,555` | cross-copy source and destination semantic revisions |
| `opened/patch.rs` `apply_exact_revision`, `validate_candidate` | post-condition checks |
| durable `LPRM` header | 2 of its 3 fixed fields |
| durable `LPCP` header | 3 of its 6 revisions (the 3 physical ones are unchanged) |

**Not changed:** `physical_package_fingerprint` and its
`litchi-pptx-cross-physical-v2` domain (0598), every published byte, every
archive layout, `packages_equal`, and the `Patch` payload encoding itself.

### Revision **verdicts** change — nowhere

Write `ℐ(P)` for the fingerprint input tuple, `R₁ = H∘enc₁`, `R₂ = H∘enc₂`.

1. **Same domain.** `R₂` is a total function of exactly `ℐ(P)` and nothing else:
   it reads no `Arc` identity, no part insertion order, no ZIP layout, no
   compression and no physical archive. The memo changes no input — it only
   declines to recompute `Dᵇ` for an allocation whose digest is already known,
   and a miss recomputes it. So `ℐ(P) = ℐ(Q) ⟹ R₂(P) = R₂(Q)`: **equal packages
   still compare equal**, unconditionally and without relying on any hash
   property.

2. **`enc₂` is injective.** The header is self-delimiting (every variable field
   length-prefixed, every count fixed-width), it is followed by a `u32` part
   count and then exactly that many fixed 32-byte blocks; `encᵖ` and `encᵇ` are
   self-delimiting for the same reason. So distinct digest sequences come from
   distinct encodings and vice versa.

3. **The part order is canonical.** Both versions sort by part name, and OPC part
   names are unique within a package (`try_add_part` routes every candidate
   through `validate_new_part_name`, which refuses `Duplicate`, `Equivalent`
   and `Derived` conflicts), so the sort is a total order with no ties. Two packages
   whose parts hold the same payloads under *different* names therefore differ,
   which the gate exercises with a payload-transposition case.

4. **Different packages still compare different — with one quantified cost.**
   `R₁(P) = R₁(Q)` with `ℐ(P) ≠ ℐ(Q)` needs a SHA-256 collision; so does
   `R₂(P) = R₂(Q)`, but `R₂` offers **2k + 1** hash instances to collide instead
   of one, where `k` is the part count: the outer digest, `k` part digests and
   `k` payload digests. By the union bound that costs at most
   `log₂(2k+1)` bits of the 128-bit generic collision margin — **about 9 bits at
   `k = 256`**, the default `Limits::max_parts`. The three tiers carry distinct
   domain strings, so a payload digest can never be reinterpreted as a part
   digest or as a revision. This is the honest price of tiering and it is stated
   rather than elided.

5. **Every verdict in the crate is of the form `R(A) == R(B)` for two packages in
   the same process and the same build.** By (1)-(4) that predicate is the same
   function of `(A, B)` under `R₁` and `R₂`, so: a valid application still
   applies; a stale package still conflicts with the same `Error::UnsafeEdit`
   operation and reason; an exact no-op still compares equal, so
   `Transaction::is_changed` is false and `commit()` still returns the empty
   patch and the source snapshot; `apply_exact_revision`'s post-condition still
   passes exactly when the candidate is the planned one; and ADR 0003's rule that
   a patch authorizes exactly one source state is untouched. The *only*
   comparison that is not preserved is one between a v1 value and a v2 value,
   which arises solely for a persisted patch — and the magic bump refuses that
   input instead of answering it.

**The empirical gate.** A pairwise oracle test over an 11-package corpus — the
base, a byte-identical re-allocation, and nine single-input mutations (payload
grows, payload shrinks, content type, added part, removed part, part
relationship, root relationship, external relationship, payload transposition) —
asserts for all 121 ordered pairs that `packages_equal(A,B) == (R₁(A)==R₁(B)) ==
(R₂(A)==R₂(B))`. It passes with 0 disagreements
(`gates.txt`, `revision_v2_preserves_every_pairwise_verdict_of_v1`).

**The whole suite passes under the redefined hash.** `cargo test -p litchi-pptx`
is **874 passed, 0 failed** with the scratch implementation in place, including
every durable-patch round-trip, every staleness refusal and every adversarial
cross-copy test. No test in the crate pins a golden revision value; the durable
formats are only ever round-tripped inside a process. That is evidence for the
scope of the bump: the *only* thing it protects is a patch that outlives the
process that planned it.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0; every measured process pinned to **CPU 20** while seven other
agents of the same wave built and measured on the other cores. Both legs built
`--release --locked` from `tools/perf-baseline`: the before leg from the shared
read-only checkout of `c7326f680`, the after legs from the scratch
implementation in the measurement worktree. Binaries were staged outside every
Cargo target directory before measurement (change 0627's rule).

Two scratch legs are measured, because they answer different questions:

- **A** — the memo lives on `Snapshot` only. The commit's fingerprint gets a
  parent memo from the source snapshot; the two bare-patch replays hold no
  snapshot and stay cold.
- **A+B** — A, plus the facade `Package` retains the memo of the snapshot it
  last published and offers it as the parent of the next publication. This is
  what makes a *durable patch application* cheap.

### Bytes and hits per fingerprint (probe build, exact)

A `pptx_eager_batch_edit_save` lifecycle, 229 parts, 17,568,429 payload bytes.
`fed_bytes` counts every byte handed to a SHA-256 instance inside the fingerprint.

| fingerprint | parts hit | parts hashed | payload bytes hashed | bytes fed |
| --- | ---: | ---: | ---: | ---: |
| `opened_presentation()` — the cold capture | 0 | 229 | 17,568,429 | 17,668,953 |
| `Transaction::commit` | **228** | **1** | **3,499** | **93,763** |
| `apply_opened_presentation_commit` | *(reuses 0590's `packages_equal`; no fingerprint)* | | | |
| replay of the inverse patch *(A+B only)* | **228** | **1** | **3,491** | **93,755** |
| replay of the inverse's inverse *(A+B only)* | **228** | **1** | **3,499** | **93,763** |

A memoized fingerprint of this deck feeds **0.53%** of the bytes a cold one
feeds. The one miss is the edited slide part.

### Instructions per lifecycle (callgrind isolation pair, `--warmup 0`)

`--samples 1` and `--samples 3`, differenced over the two extra samples.

| selector | before | **A** | **A+B** |
| --- | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 10,458,591,613 | 9,556,437,414 (**−8.63%**) | **7,723,437,449 (−26.15%)** |
| `pptx_eager_multi_slide_batch_edit_save` | 10,494,452,610 | 9,596,966,728 (**−8.55%**) | **7,763,085,693 (−26.03%)** |
| `pptx_slide_move_boundary_save` | 45,796,410 | 42,490,714 (**−7.22%**) | 42,628,717 (−6.92%) |
| `pptx_slide_remove_boundary_save` | 95,825,741 | 98,139,554 (**+2.41%**) | 98,170,602 (**+2.45%**) |

The complete-package fingerprint subtree, per lifecycle:

| selector | fingerprints | before | **A** | **A+B** |
| --- | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 4 | 3,678,478,705 | 2,775,430,951 (−24.55%) | **944,205,414 (−74.33%)** |
| `pptx_eager_multi_slide_batch_edit_save` | 4 | 3,678,533,244 | 2,776,751,270 (−24.51%) | **948,232,832 (−74.22%)** |
| `pptx_slide_move_boundary_save` | 3 | 15,087,733 | 11,440,687 (−24.17%) | 11,428,431 (−24.25%) |
| `pptx_slide_remove_boundary_save` | 9 | 44,728,467 | 47,480,233 (+6.15%) | 47,473,915 (+6.14%) |

**Change 0590's model is confirmed.** It predicted "about 2.5 G `Ir` per
lifecycle, or 23% of the 10.68 G that remains". The measurement is
**2,735,154,164 `Ir`, 26.15%** of the 10.46 G that remains at this base.

**What one fingerprint costs, and what memoization leaves.** Splitting the
callgrind call graph by call site on the eager selector gives, per call:

| | before | A |
| --- | ---: | ---: |
| cold fingerprint (`capture_internal` → fingerprint, 7 calls) | 794,800,002 | 797,900,434 (**+0.39%**) |
| memoized fingerprint, derived | — | **≈ 7.3 M** (0.79% of cold) |

The +0.39% is the price of the extra tiers on a fingerprint that cannot use the
memo, and it is the mechanism behind the one selector that regresses.

### The one selector that gets worse

`pptx_slide_remove_boundary_save` is **+2.41%** (A) and **+2.45%** (A+B) per
lifecycle, and it is attributable — unlike 0590's regression on the same
selector, which was not. Its lifecycle issues **9 complete-package fingerprints
and reuses nothing** (`packages_equal` calls: 0 on every leg), so it pays the
per-part tier overhead on all nine and collects no hit. The measurement isolates
the overhead exactly:

> `(47,480,233 − 44,728,467) / 9 = ` **`+305,752 Ir` per fingerprint** on this
> 25-member, 32,396-byte deck.

The same overhead on the 229-part eager deck is `+3,100,432 Ir` per cold
fingerprint — **13,539 `Ir` per part**, which is what two extra SHA-256
initialize/finalize pairs and one hash-map insert cost under callgrind's
software SHA-256.

**The break-even rule this yields.** A memoized part saves the payload's hashing
and pays the tier; a missed part pays the tier for nothing. At this host's
callgrind price of ≈ 49 `Ir` per hashed byte, a part must carry roughly
**280 bytes** of payload for a hit to pay for its own tier, and every part of
every fingerprint that gets no hit is a flat loss of about 13,500 `Ir`. This is
why the media-rich deck gains 26% and the 32 KB boundary deck loses 2.4%.

### Native `perf stat` and paired timing

Four legs per selector in run order **A1 B1 B2 A2** (before, A+B, A+B, before),
30 measured samples after 3 warmups each, `taskset -c 20`, with seven other
agents active on the other cores. The A/A floor is A1 against A2 and the B/B
floor is B1 against B2, both taken inside the same window.

| selector | before p50 | A+B p50 | p50 Δ | mean Δ | A/A p50 floor | B/B p50 floor |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 320.127 ms | 296.628 ms | **−7.34%** (before is +7.92% of after) | −6.07% | +0.52% | +0.54% |
| `pptx_eager_multi_slide_batch_edit_save` | 323.216 ms | 299.623 ms | **−7.30%** (before is +7.87% of after) | −18.97% | −0.55% | −0.48% |
| `pptx_slide_move_boundary_save` | 0.529 ms | 0.450 ms | −14.97% | −31.97% | −2.78% | **−10.17%** |
| `pptx_slide_remove_boundary_save` | 2.083 ms | 2.039 ms | −2.11% | −3.73% | −2.14% | −0.37% |

Only the two eager rows have a floor worth reporting against. The two boundary
selectors run in 0.5-2 ms and this window was noisy for them: the move
selector's own B/B p50 floor is **−10.17%** and its A/A p95 and p99 floors are
−22.01% and −23.84%, which is wider than any effect; the remove selector's
−2.11% at p50 sits exactly on its own −2.14% A/A floor. Their tails are not
interpretable and their per-phase clocks move in phases that contain no changed
code (the move selector's `plan` phase −25.68% and its `publication` phase
−21.93%, neither of which this design touches), which is the signature of the
noise rather than of the change. **For both boundary selectors the counts above
are the result and the wall clock is not.**

**Native `perf stat`**, same four legs, whole child:

| selector | cycles before | cycles A+B | cycles Δ | A/A cycles | perf `instructions` Δ |
| --- | ---: | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 70,235,133,228 | 66,769,076,529 | **−4.93%** | −1.11% | −2.31% |
| `pptx_eager_multi_slide_batch_edit_save` | 70,300,503,970 | 67,080,530,872 | **−4.58%** | −0.81% | −2.45% |
| `pptx_slide_move_boundary_save` | 1,642,225,611 | 1,609,335,224 | −2.00% | −3.86% | −0.30% |
| `pptx_slide_remove_boundary_save` | 1,888,507,408 | 1,896,637,654 | **+0.43%** | −0.85% | **+0.32%** |

Two readings matter here.

**The native saving is much smaller than the callgrind saving, and that is
expected.** Callgrind prices SHA-256 in software because valgrind masks the SHA
CPUID bit, so it charges roughly 49 `Ir` per hashed byte where the hardware
retires `sha256rnds2`; it also counts `rep movsb` per byte (0604 measured a 35×
overstatement). The eager selectors fall **26.15%** in callgrind instructions
per lifecycle and **2.31-2.45%** in native instructions, while native *cycles*
fall **4.6-4.9%** against a floor under 1.2%. Cycles are the number to believe,
and the callgrind counts rank the work rather than predicting the latency —
exactly as `GOAL_AUDIT.md`'s standing note says.

**The instruction regression on `pptx_slide_remove_boundary_save` is real and
does not clear its own floor natively.** It is +2.45% per lifecycle in
callgrind, +0.32% in native whole-child instructions and +0.43% in cycles
against an A/A floor of −0.85%. It is reported rather than folded into a mean,
and it is the reason admission gate 6 exists.

## Correctness evidence

The scratch implementation was gated in the measurement worktree before any
number was taken:

| gate | result |
| --- | --- |
| `cargo test -p litchi-pptx` (scratch A+B in place) | **874 passed, 0 failed** |
| `revision_v2_preserves_every_pairwise_verdict_of_v1` | 121 ordered pairs, 0 disagreements between `packages_equal`, `R₁` and `R₂` |
| `memoized_recapture_equals_a_cold_recapture` | the memoized commit revision equals a fresh fingerprint; every memo key names an allocation the snapshot's own package holds |
| `part_digest_memo_resident_bytes` | the table above |
| `debug_assert!` in `capture_internal` | every test that captures re-derives the revision the cold way and compares — this is why the run above proves value identity and not only "it compiles" |
| `cargo test -p litchi --features docx,xlsx,pptx,xls` | **266 passed, 0 failed** — the facade consumer of the changed crate |
| `cargo clippy -p litchi-pptx --all-targets` | clean (workspace lints are deny) |
| `cargo doc -p litchi-pptx --no-deps` | clean (rustdoc lints deny) |
| `cargo fmt --all --check` | clean |

The harness's own verification is the differential check on output bytes:
`litchi_perf_baseline::sha256_hex` is bit-identical between legs
(3,543,866,958 `Ir` on both), `zlib_rs::deflate::deflate` moves by 30,570 `Ir`
in 9.39 G (**0.0003%**), and every selector reports the same `output_sha256` on
every leg. What the scratch removes is hashing, not anything the selector
produces.

## Validation preserved

No validation is removed, weakened, reordered or made conditional by the design.
`capture_internal` runs the presentation-root resolution, the slide-reference
one-to-one checks, the relationship-type checks, the per-slide name parse,
ADR 0013's `notes::load_snapshot` topology check and the name-index build
unconditionally and in the same order, before the revision is taken. The memo
sits strictly inside the hash, below every check. The limits are untouched:
`Limits::max_parts` still bounds the slide count at the same place, and the memo
allocates through `try_reserve`, so an allocator refusal is still
`Error::Allocation { resource: "opened-presentation part digests" }` rather than
an abort.

## Admission gates

This design may land only when **all** of the following hold. They are written
so that a later batch can fail them.

1. **The format bump is complete.** `LPRM0002` and `LPCP0003` are written; a
   `LPRM0001` or `LPCP0002` input returns `Error::DurablePatchRevisionFormat`
   from `from_bytes` and `from_bytes_with_limits` **before** any package is read
   or mutated; forward and inverse both refuse; an unrecognized magic still
   returns the existing `Error::Invalid`; and a v2 patch round-trips
   `to_bytes`/`from_bytes` to an equal value.
2. **Verdict preservation, corpus-wide.** The pairwise oracle above runs over
   every `.pptx` under `test-data/` as well as the synthetic mutation corpus,
   comparing `packages_equal`, `R₁` and `R₂` on every ordered pair of
   (package, mutated package), with **zero** disagreements. The v1
   implementation is retained under `#[cfg(test)]` for exactly this gate.
3. **The alias gate.** A `Part` implementation whose `blob_arc()` disagrees with
   its `blob()` — the `MismatchedBlobArcPart` shape `litchi-opc` already carries
   — is proven never to be memoized, and the revision it produces is proven
   equal to the one the non-memoizing path produces.
4. **The ABA gate.** A test drops a package's payload, allocates a different
   payload, and proves the memo does not answer for the recycled address. (With
   the retained `Arc` this is a proof that the address cannot be recycled.)
5. **The retention gate.** After a capture, a rebind and a publication, every
   memo entry names an allocation the owning snapshot's package holds; the
   process holds no payload allocation solely because a memo names it. Asserted
   directly, not inferred.
6. **No regression above the 5% review trigger, and the attributable one
   disclosed.** `pptx_slide_remove_boundary_save` is **+2.4%** here. That is
   below the trigger but it is *caused by the change*, so it must be re-measured
   and reported at landing, and the landing record must say why a deck of small
   parts pays for a design that helps a deck of large ones.
7. **The facade carry (part B) lands with the memo, or the design is not worth
   it.** A alone is **−8.6%** of an eager lifecycle; A+B is **−26.2%**. The two
   bare-patch replays are not a harness artifact: they are
   `Package::apply_opened_presentation_patch`, the public durable-patch route.
   If B is judged too much ambient state, this design should be reconsidered
   rather than landed at a third of its value.
8. **Native confirmation.** The instruction win must be confirmed by `perf stat`
   cycles and by paired timing above the window's floor, because callgrind runs
   SHA-256 in software and overstates every hashing share by roughly 5×.
9. **ADR 0003 conflict determinism re-proved at landing**, with the existing
   stale-source, forged-patch and drift tests re-run unchanged.

## Limitations

- **Not claimed:** any speedup. `performance_claim: none`, no claim-registry
  entry, and the paired timings are reported beside the floor rather than as a
  result.
- No production code changed. The measurement is of a scratch implementation
  that passed the crate's own gates — `cargo fmt --all --check`, `cargo clippy
  -p litchi-pptx --all-targets`, `cargo doc -p litchi-pptx --no-deps`,
  `cargo test -p litchi-pptx` (874) and `cargo test -p litchi --features
  docx,xlsx,pptx,xls` (266) — but was not reviewed, does not bump either
  durable magic, does not carry the typed refusal, and was not exercised
  against the corpus gates its own admission list requires. It is a sizing
  instrument, not a candidate.
- The counts are exact for four selectors on their fixed synthetic corpora, on
  this host and these builds. The eager corpora carry no physical source
  provenance and re-deflate all media on save (0590's caveat, unchanged), so
  their wall-clock denominator is inflated by recompression a production `open`
  would not perform.
- Callgrind runs SHA-256 in software because valgrind masks the SHA CPUID bit,
  so every instruction share attributed to hashing here is roughly five times
  its native cycle share. The `perf stat` cycles are the counterweight.
- The `≈ 7.3 M Ir` cost of a memoized fingerprint is **derived** from the
  measured per-call-site split and the measured per-fingerprint overhead, not
  measured in isolation. The bytes-fed figures beside it are measured exactly.
- The 9-bit collision-margin figure is a union bound at `k = 256`, not an
  analysis of SHA-256.
- No allocation-count, RSS, cold-cache, range-source, concurrency, real-producer
  or cross-platform measurement was taken. The memo's resident cost is computed
  from `HashMap::capacity`, not from an allocator counter.
- The cross-package copy path's *physical* revisions and 0598's
  `physical_revision` cache are untouched and unmeasured here.

## Retained evidence

[`results/change-0645/README.md`](results/change-0645/README.md) — both legs'
callgrind annotations, call counts and per-call-site splits, the isolation
pairs, the per-fingerprint memo probe, every paired timing report, the gates,
the scratch patches (A, A+B and the probe instrumentation), the decision record
and the log paragraphs.
