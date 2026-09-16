# 0655: the memoized per-part revision proof lands — an eager PPTX lifecycle loses 27.65% of its instructions, the two durable patch formats bump, and a patch under the old magic is refused by name

Status: retained, implemented in `crates/litchi-pptx`.
`performance_claim: none` — no claim-registry entry is created by this wave; the
instruction counts, hash accounting and paired medians below are reported as
evidence beside the host's A/A floor, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements the design change
[0645](0645-pptx-memoized-revision-proof-design.md) froze, under decision 4 of
change [0652](0652-owner-decisions-for-the-third-wave.md). It is part **(c)** of
**PPTX-1**, the one part of change
[0590](0590-pptx-opened-transaction-revision-reuse.md) its author declined to
implement because `package_fingerprint` is serialized into two durable patch
formats and redefining it is a wire-format change.

## What was changed

Four things, in `crates/litchi-pptx` only.

1. **The complete-package revision is now a two-tier proof.**
   `opened::model::package_fingerprint` keeps its header exactly — the domain
   string, the root relationships sorted by `rId`, and the opaque non-part
   members — and replaces the flat per-part field run with a `u32` part count
   followed by one fixed 32-byte digest per part, in sorted part-name order:

   ```text
   payload digest   Dᵇ(b) = H( "litchi-pptx-opened-payload-v2" ‖ len ‖ b )
   part digest      Dᵖ(p) = H( "litchi-pptx-opened-part-v2"
                             ‖ name ‖ content type ‖ Dᵇ(payload)
                             ‖ relationships sorted by rId )
   revision         R₂(P) = H( "litchi-pptx-opened-v2" ‖ header
                             ‖ "parts" ‖ u32 k ‖ Dᵖ(p₁) ‖ … ‖ Dᵖ(p_k) )
   ```

2. **`Dᵇ`, and only `Dᵇ`, is memoized**, in a `PartDigests` table keyed by
   `(payload address, payload length)` that retains the payload `Arc`. The table
   lives on `Snapshot` beside the `physical_revision` cache change
   [0598](0598-pptx-cross-copy-revision-cache.md) put there. `Dᵖ` is recomputed
   on every pass, because a part's content type and relationships are not behind
   its payload `Arc`.

3. **Both durable formats bump**: `LPRM0001 → LPRM0002` and
   `LPCP0002 → LPCP0003`. A patch whose first eight bytes are a recognized but
   superseded magic is refused in `from_bytes` and `from_bytes_with_limits` with
   the new typed `Error::DurablePatchRevisionFormat`, before any header field is
   read and with no package in hand. An unrecognized magic keeps today's
   `Error::Invalid`. There is no migration and no dual read.

4. **The facade carries the memo.** `litchi_pptx::Package` retains the memo of
   the snapshot each opened-presentation publication produces and offers it as
   the parent of the next one, which is what makes a bare durable-patch
   application cheap. Every mutation that produces no opened-presentation
   snapshot to adopt releases it instead, so the facade never retains a payload
   allocation its own graph has replaced.

## Authority

Change 0652, decision 4, quotes the owner: **"PPTX memoized revision: invalidate
the old patches and make the path faster."** 0652 records that this authorizes
"0645's design in full: the per-part memo on `Snapshot`, the `LPRM0001 →
LPRM0002` and `LPCP0002 → LPCP0003` bumps, the typed refusal for a patch
serialized under the old format, and the facade-carried memo", and requires this
record to prove "0645's nine gates; a patch under the old magic refused by name
before any package is read; equal packages still compare equal and different
ones different, over the corpus; the `pptx_slide_remove_boundary_save`
regression re-measured and disclosed".

0652's standing trade-off 1 ("breaking changes are totally acceptable")
authorizes the format bump and the new public items; trade-off 2 ("correctness
and safety is the primary consideration") is why the memo is projected onto its
own package at every construction and released at every mutation that cannot
adopt an exact one, rather than carried on the faster but looser rule 0645
priced.

## Breaking changes

| item | change |
| --- | --- |
| `litchi_pptx::DurablePatchFormat` | **new** public `#[non_exhaustive]` enum with `SlideRemovalV1`, `SlideRemovalV2`, `CrossSlideCopyV2`, `CrossSlideCopyV3`, a `magic()` accessor and a `Display` that writes the magic |
| `litchi_pptx::Error::DurablePatchRevisionFormat { found, expected }` | **new** variant. `Error` is already `#[non_exhaustive]`, so matching stays additive |
| `SlideRemovalPatch::to_bytes` | writes `LPRM0002`. Bytes written by any earlier build are no longer accepted |
| `SlideRemovalPatch::from_bytes`, `from_bytes_with_limits` | an `LPRM0001` input within the durable byte limit now returns `Error::DurablePatchRevisionFormat` instead of parsing (or, for a malformed old patch, instead of its previous refusal) |
| `CrossSlideCopyPatch::to_bytes` | writes `LPCP0003` |
| `CrossSlideCopyPatch::from_bytes`, `from_bytes_with_limits` | same refusal for `LPCP0002` |
| every public `[u8; 32]` revision accessor | `Snapshot::revision`, `SlideRemovalPatch::{source,target}_revision`, `SlideRemovalPlan::source_revision`, `SlideCopyPlan::source_revision` and the three *semantic* `CrossSlideCopy*` revisions return values in the new algebra. Their types and their meanings are unchanged; only the values move. The three *physical* cross-copy revisions (`litchi-pptx-cross-physical-v2`, change 0598) are untouched |

No other public item changes. No archive layout, no published byte, no
`packages_equal` behaviour and no `Patch` payload encoding moves.

## Why it is sound

### Verdicts do not move

Write `ℐ(P)` for the fingerprint input tuple — the root relationships, the
non-part members, and per part the name, content type, payload and
relationships — and `R₁ = H∘enc₁`, `R₂ = H∘enc₂`.

1. **Same domain.** `R₂` is a total function of exactly `ℐ(P)`: it reads no
   `Arc` identity, no part insertion order, no ZIP layout, no compression and no
   physical archive. The memo changes no input — it declines to recompute `Dᵇ`
   for an allocation whose digest is already known, and a miss recomputes it. So
   `ℐ(P) = ℐ(Q) ⟹ R₂(P) = R₂(Q)`: **equal packages still compare equal**,
   unconditionally and without relying on any hash property.
2. **`enc₂` is injective.** The header is self-delimiting (every variable field
   length-prefixed, every count fixed-width), followed by a `u32` part count and
   exactly that many fixed 32-byte blocks; `encᵖ` and `encᵇ` are self-delimiting
   for the same reason.
3. **The part order is canonical.** Both versions sort by part name, and OPC
   part names are unique within a package, so the sort is a total order with no
   ties. Two packages holding the same payloads under different names therefore
   differ, which the oracle exercises with a payload-transposition case.
4. **Different packages still compare different, at a quantified price.**
   `R₁(P) = R₁(Q)` with `ℐ(P) ≠ ℐ(Q)` needs a SHA-256 collision; so does
   `R₂(P) = R₂(Q)`, but `R₂` offers `2k + 1` hash instances to collide instead
   of one, where `k` is the part count. By the union bound that costs at most
   `log₂(2k+1)` bits of the 128-bit generic collision margin — **about 9 bits at
   `k = 256`**. The three tiers carry distinct domain strings, so a payload
   digest can never be reinterpreted as a part digest or as a revision.
5. **Every verdict in the crate is `R(A) == R(B)` for two packages in the same
   process and the same build.** By (1)–(4) that predicate is the same function
   of `(A, B)` under `R₁` and `R₂`: a valid application still applies; a stale
   package still conflicts with the same `Error::UnsafeEdit` operation and
   reason; an exact no-op still compares equal, so `Transaction::is_changed` is
   false and `commit()` still returns the empty patch and the source snapshot;
   `apply_exact_revision`'s post-condition still passes exactly when the
   candidate is the planned one; and ADR 0003's rule that a patch authorizes
   exactly one source state is untouched. The only comparison not preserved is
   one between a v1 value and a v2 value, which arises solely for a persisted
   patch — and the magic bump refuses that input instead of answering it.

### The memo cannot answer for the wrong bytes

An entry asserts *the allocation at address `a` of length `ℓ` has payload digest
`d`*. Three properties make that safe.

- **Allocation identity, retained.** The entry holds the payload `Arc`, so the
  allocation cannot be freed while the entry lives, so the address cannot be
  recycled by a different payload. This is the ABA argument, and the retained
  `Arc` is load-bearing rather than an optimization. The degenerate case `ℓ = 0`
  is safe in the other direction: two distinct empty `Vec`s may share a dangling
  address, and they have the same payload.
- **The alias test.** `Part` is a public trait and a foreign implementation may
  return an `Arc` whose contents differ from the bytes `blob()` returns —
  `litchi-opc` guards the same way before reusing a donor payload. A part whose
  `blob_arc()` does not alias its `blob()` is never memoized and is hashed
  exactly as the unmemoized path hashes it. A hostile or merely inconsistent
  part can cost performance and can never change a value.
- **Re-derivation under test.** Every reused digest is recomputed and compared
  by a `debug_assert!`, so the crate's own suite proves value identity rather
  than only that the code compiles.

### The memo retains nothing its owner does not own

*Invariant: every entry of a snapshot's memo names an allocation that snapshot's
own package holds.* It is established structurally: `capture_internal` clones
the package into the snapshot's `Arc` **first** and then takes the revision over
that clone, so a computed memo names the snapshot's own allocations by
construction; a memo the caller supplies is projected onto the clone, which
re-keys onto its allocations and drops every entry it does not hold.
`Snapshot::rebound_to` projects the same way. A projection that cannot allocate
falls back to an empty memo rather than failing the capture, because a memo is
optional by definition.

The facade obeys the matching rule: it adopts the exact memo of the snapshot
each publication produces, and releases it at every mutation that produces no
such snapshot (a typed edit, a presentation materialization, a change-tracking
publication). A released memo costs one cold hash and can never answer wrongly.

### ADR reading

- **ADR 0003** (snapshots, edits and patches) is unchanged in substance: patches
  remain exact-source-checked, reversible and deterministic in conflict. A proof
  in a superseded algebra cannot be source-checked at all, which is why a
  superseded patch is refused rather than compared.
- **ADR 0005** is amended by this record, as change 0652 directs, with a dated
  note that states the conditions under which a snapshot or a publishing facade
  may retain an identity-keyed digest memo. The amendment does not weaken the
  eviction rules for semantic payload caches and permits no ambient process-wide
  state.
- **ADR 0001 / ADR 0005 leakage rules** are untouched: `PartDigests` is
  `pub(crate)`, no archive type, lock or executor is exposed, and the only new
  public items are the format enum and the error variant.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0. Every measured process pinned to **CPU 10** while seven other
agents of the same wave built and measured on the other cores. Both legs built
`--release --locked` from `tools/perf-baseline`: the before leg from the shared
read-only checkout of `70d7768cc`, the after leg from this branch. Binaries were
staged outside every Cargo target directory before measurement (change 0627's
rule) and their sha256s are in the packet.

### Hashes and bytes per lifecycle (probe build, exact)

A reverted instrumentation patch (`instrumentation/memo-accounting.patch`,
retained, never present in a timed binary) counts, per fingerprint, the parts
answered from the memo, the payload bytes hashed, and every byte handed to a
SHA-256 instance inside the fingerprint. Fingerprint counts per lifecycle are
isolation pairs: the probe is run at `--samples 1` and `--samples 3` and the
difference is divided by two.

**`pptx_eager_batch_edit_save`** — 229 parts, 17,568,429 payload bytes, four
complete-package fingerprints per lifecycle:

| fingerprint | parts hit | parts hashed | payload bytes hashed | bytes fed |
| --- | ---: | ---: | ---: | ---: |
| `opened_presentation()` — the cold capture | 0 | 229 | 17,568,429 | 17,668,953 |
| `Transaction::commit` | **228** | **1** | **3,499** | **93,763** |
| replay of the inverse patch | **228** | **1** | **3,491** | **93,755** |
| replay of the inverse's inverse | **228** | **1** | **3,499** | **93,763** |

A memoized fingerprint of this deck feeds **0.53%** of the bytes a cold one
feeds. The one miss is the edited slide part. The two replays are
`Package::apply_opened_presentation_patch`, the public durable-patch route: they
hold no snapshot, and it is the facade carry that answers them.

**`pptx_real_file_ordinary_save_edit`** — change [0638](0638-facade-and-ordinary-save-selectors.md)'s real-deck ordinary-save edit phase, on
`test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx`, the
108 KB, 103-member deck change
[0649](0649-pptx-opened-transaction-real-deck-edit.md) attributed. Of its 103
ZIP members, 58 are OPC parts carrying 768,905 payload bytes; the rest are the
content-type item and the relationship items, which the fingerprint feeds
through its header rather than as parts. Two fingerprints per lifecycle, and
the phase re-opens the file for each sample, so the capture is always cold and
only the commit can hit:

| fingerprint | parts hit | parts hashed | payload bytes hashed | bytes fed |
| --- | ---: | ---: | ---: | ---: |
| `opened_presentation_transaction()` | 0 | 58 | 768,905 | 792,152 |
| `Transaction::commit` | **57** | **1** | **16,936** | **37,618** |

The memoized fingerprint feeds **4.75%** of the cold feed. The single miss is
the edited slide, which is 2.2% of this deck's payload bytes rather than the
0.02% the eager corpus's edited slide is.

**`pptx_slide_remove_boundary_save`** — 25 members, 87,003 payload bytes. Its
lifecycle issues **nine** complete-package fingerprints and **every one of them
is cold**: three over the 25-member package and six over the 24-member one,
0 hits, 0 memoized bytes. It pays the per-part tier nine times and collects
nothing. That is the measured mechanism of this selector's regression, and it is
why the design helps a deck of large parts and hurts a deck of small ones.

The other two selectors, for completeness:
`pptx_eager_multi_slide_batch_edit_save` issues four fingerprints, one cold and
three with **221 of 229** parts answered (its edit touches eight slides, so it
misses eight); `pptx_slide_move_boundary_save` issues three, one with 24 of 25
parts answered and two cold. The complete per-lifecycle table is
`counts/memo-per-lifecycle.txt`.

### Resident cost of the memo

`part_digest_memo_resident_bytes` reports `entries.capacity() × (16 + 40 + 1)`
and asserts a bound of 228 bytes per memoized payload:

| parts | entries | memo bytes | per entry |
| ---: | ---: | ---: | ---: |
| 29 | 29 | 3,192 | 110.1 |
| 85 | 85 | 6,384 | 75.1 |
| 221 | 221 | 12,768 | 57.8 |

The per-entry figure varies only because `HashMap` capacity grows in steps; the
per-slot cost is a flat 57 bytes. A 229-part deck costs about **13 KiB** per
snapshot on top of the 17 MB of payload it already holds — 0.08%. Snapshot
clones share the memo through the `Arc`, and the retention invariant is what
keeps the payload bytes themselves at zero extra cost.

### Instructions per lifecycle (callgrind isolation pair, `--warmup 0`)

Each selector is profiled at `--samples 1` and `--samples 3` and the totals are
differenced over the two extra samples.

| selector | before | after | Δ |
| --- | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 9,886,996,492 | **7,153,579,244** | **−27.65%** |
| `pptx_eager_multi_slide_batch_edit_save` | 9,924,133,972 | **7,191,997,153** | **−27.53%** |
| `pptx_slide_move_boundary_save` | 46,267,799 | 42,627,390 | −7.87% |
| `pptx_slide_remove_boundary_save` | 95,811,797 | 98,493,465 | **+2.80%** |
| `pptx_real_file_ordinary_save_edit` | 2,756,984,423 | 2,628,177,821 | −4.67% *(not interpretable, below)* |
| `pptx_real_file_ordinary_save_lifecycle` | 2,697,833,744 | 2,696,576,466 | −0.05% *(not interpretable, below)* |

The complete-package fingerprint subtree, per lifecycle:

| selector | fingerprints | before | after | Δ |
| --- | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 4 | 3,678,482,713 | **944,179,031** | **−74.33%** |
| `pptx_eager_multi_slide_batch_edit_save` | 4 | 3,678,512,559 | **948,226,506** | **−74.22%** |
| `pptx_slide_move_boundary_save` | 3 | 15,095,724 | 11,428,679 | −24.29% |
| `pptx_slide_remove_boundary_save` | 9 | 44,729,621 | 47,487,507 | **+6.17%** |
| `pptx_real_file_ordinary_save_edit` | 2 | 82,269,878 | **44,253,161** | **−46.21%** |
| `pptx_real_file_ordinary_save_lifecycle` | 2 | 82,303,403 | **44,276,493** | **−46.20%** |

**The two real-file rows' whole-lifecycle totals are not the result; their
fingerprint subtrees are.** Subtracting the fingerprint subtree from each total
leaves the part of the lifecycle this change does not touch, and that remainder
moves by **−3.40%** on `pptx_real_file_ordinary_save_edit`
(2,674,714,545 → 2,583,924,660) and by **+1.41%** on
`pptx_real_file_ordinary_save_lifecycle` (2,615,530,341 → 2,652,299,973) between
the same two legs. Both selectors are dominated by
`process_markup_compatibility` — change 0649 showed it is 93.9% of this deck's
edit — and that subtree is unchanged code that nevertheless differs by up to
17.8 M `Ir` per lifecycle between the legs. So the isolation-pair method has
about **±3%** of resolution on this selector family, which is wider than the
1.4% the fingerprint saving is worth, and the whole-lifecycle deltas above
(−4.67% and −0.05%) bracket that saving rather than measuring it. The
fingerprint subtree, which is exactly the code that changed, falls **−46.21%**
and **−46.20%** — the same number twice, independently — and that is what this
selector pair establishes.

**Change 0590's model is confirmed, and change 0645's sizing reproduces at this
base.** 0590 predicted "about 2.5 G `Ir` per lifecycle" for the eager selector;
the measurement is **2,733,417,248 `Ir`, 27.65%** of the 9.89 G that remains at
`70d7768cc`. 0645 sized the same design at −26.15% of the 10.46 G that remained
at `c7326f680`, with the fingerprint subtree at −74.33% — the subtree figure is
identical to three decimal places, and the whole-lifecycle percentage differs
only because the denominator moved with the wave-2 changes.

### What one fingerprint costs, cold and memoized

Splitting the fingerprint subtree by **call site** and differencing the same
isolation pair separates a cold fingerprint from a memoized one on the same leg
(`counts/callsite-pairs.txt`). On `pptx_eager_batch_edit_save` the four
fingerprints of a lifecycle are one cold capture, two memoized replays (both
through `capture_internal`) and one memoized commit:

| | before | after |
| --- | ---: | ---: |
| `Transaction::commit` → fingerprint (1 per lifecycle) | 919,620,607 | **7,149,543** |
| `capture_internal` → fingerprint (3 per lifecycle) | 2,758,862,105 | 937,029,488 |
| …of which the cold capture, by subtraction | 919,620,701 | **922,730,402** |

So on this deck a **memoized fingerprint costs 7,149,543 `Ir`, 0.77% of a cold
one** — change 0645 derived "≈ 7.3 M" from a leg that could not measure it, and
this measures it — and the **cold fingerprint costs 3,109,701 `Ir` more than it
did**, **+0.34%**, or **13,579 `Ir` per part** for the extra initialize and
finalize pairs and the map insert. (The subtraction assumes the two replays cost
what the commit costs; the probe justifies that, since all three answer 228 of
229 parts and feed 93,755 or 93,763 bytes.)

The two boundary selectors show the same two numbers on a 25-part deck, where
they are visible without any subtraction because the commit is the only
memoized fingerprint:

| `pptx_slide_move_boundary_save` | before | after |
| --- | ---: | ---: |
| `Transaction::commit` → fingerprint (1) | 5,024,159 | **735,551** (−85.36%) |
| `capture_internal` → fingerprint (2, both cold) | 5,035,771 each | 5,345,360 each (**+6.15%**) |

`+309,589 Ir` per cold fingerprint over 25 parts is **12,384 `Ir` per part**, the
same flat per-part tier the 229-part deck pays.

On change 0638's real deck the same split is visible directly, because its two
fingerprints per lifecycle are one cold capture and one commit:

| `pptx_real_file_ordinary_save_edit` | before | after |
| --- | ---: | ---: |
| `Transaction::commit` → fingerprint (1) | 41,139,251 | **2,455,718** (−94.03%) |
| `capture_internal` → fingerprint (1, cold) | 41,130,626 | 41,797,443 (**+1.62%**) |

`+666,817 Ir` per cold fingerprint over 58 parts is **11,497 `Ir` per part**: the
same flat tier again, on a third deck shape. The real deck's commit keeps more
of its cost than the eager deck's because its one missed part is 2.2% of the
package's payload bytes rather than 0.02%. The `pptx_real_file_ordinary_save_lifecycle`
selector gives the same pair to within 0.1%: 41,140,676 → 2,457,313 for the
commit and 41,162,727 → 41,819,179 for the cold capture. Unlike the
whole-lifecycle totals of these two selectors, the per-call-site figures are
stable, because they measure one call each rather than a difference of two
multi-gigabyte totals.

### The one selector that gets worse

`pptx_slide_remove_boundary_save` is **+2.80%** per lifecycle, and it is
attributable rather than mysterious. Its lifecycle issues nine complete-package
fingerprints and, as the probe shows, **collects no hit on any of them**. Nine
cold fingerprints × the measured `+306,723 Ir` per cold fingerprint of this
25-part deck is `+2,760,507 Ir`; the measured increase is `+2,681,668 Ir`. The
tier accounts for the whole regression.

Natively it is **+2.57% at p50 and +3.01% at p95** of wall clock, against A/A
and B/B floors of −0.28% and −0.29%, so unlike change 0645's measurement this
one clears its own floor: the selector really is about 51 µs slower per
lifecycle. Its `perf stat` cycles move −0.40% against a +0.61% floor and its
native instructions +0.54% against +0.07%.

This is below the 5% review trigger but it is caused by this change, so it is
reported rather than folded into a mean. The break-even rule behind it: a
memoized part saves its payload's hashing and pays the tier; a missed part pays
the tier for nothing. At this host's callgrind price of about 49 `Ir` per hashed
byte, a part needs roughly **280 bytes** of payload for a hit to pay for its own
tier, and every part of every fingerprint that gets no hit is a flat loss of
about 12,400 `Ir`. A 17 MB media-rich deck gains 27%; a 32 KB boundary deck of
small parts loses 2.8%.

Two smaller costs are inside these numbers and are not separated out. The
projection that establishes the retention invariant walks the parts of every
capture that is handed a caller's memo, and the facade's release calls reset an
`Arc` on mutation paths that publish no snapshot. Both are part-count work with
no hashing; they are the price of the invariant and they are paid on the same
selectors measured here.

### Native `perf stat` and paired timing

Four legs per selector in run order **A1 B1 B2 A2** (before, after, after,
before), 30 measured samples after 3 warmups each, `taskset -c 10`, with seven
other agents active on the other cores. The A/A floor is A1 against A2 and the
B/B floor is B1 against B2, both taken inside the same window. This window was
much quieter than the one change 0645 measured in: every A/A p50 floor except
the real-file lifecycle's is inside ±0.6%.

| selector | before p50 | after p50 | p50 Δ | mean Δ | A/A p50 floor | B/B p50 floor |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 316.779 ms | 297.376 ms | **−6.12%** | −5.97% | −0.51% | +0.13% |
| `pptx_eager_multi_slide_batch_edit_save` | 319.544 ms | 299.431 ms | **−6.29%** | −6.19% | −0.44% | −0.47% |
| `pptx_slide_move_boundary_save` | 0.459 ms | 0.428 ms | **−6.62%** | −6.37% | −0.56% | +0.04% |
| `pptx_slide_remove_boundary_save` | 1.970 ms | 2.021 ms | **+2.57%** | +2.69% | −0.28% | −0.29% |
| `pptx_real_file_ordinary_save_edit` | 127.896 ms | 123.588 ms | **−3.37%** | −3.69% | −0.14% | −0.07% |
| `pptx_real_file_ordinary_save_lifecycle` | 129.751 ms | 126.176 ms | −2.76% | −12.14% | **−2.65%** | −0.48% |

**Native `perf stat`**, same four legs, whole child:

| selector | cycles before | cycles after | cycles Δ | A/A cycles | perf `instructions` Δ |
| --- | ---: | ---: | ---: | ---: | ---: |
| `pptx_eager_batch_edit_save` | 65,775,648,862 | 62,556,033,195 | **−4.89%** | +0.13% | −2.92% |
| `pptx_eager_multi_slide_batch_edit_save` | 65,684,703,620 | 62,782,323,739 | **−4.42%** | +0.00% | −2.14% |
| `pptx_slide_move_boundary_save` | 1,712,119,489 | 1,685,996,879 | −1.53% | **+3.86%** | +0.24% |
| `pptx_slide_remove_boundary_save` | 1,983,673,052 | 1,975,694,387 | −0.40% | +0.61% | **+0.54%** |
| `pptx_real_file_ordinary_save_edit` | 31,662,590,100 | 31,178,424,886 | −1.53% | **+1.38%** | −0.08% |
| `pptx_real_file_ordinary_save_lifecycle` | 32,217,092,274 | 31,242,713,546 | **−3.02%** | +0.29% | −0.09% |

Four readings matter.

**The eager saving is confirmed natively and is a quarter of its callgrind
size.** Cycles fall **4.89%** and **4.42%** against A/A floors of +0.13% and
+0.00%, while callgrind instructions per lifecycle fall 27.65% and 27.53% and
native instructions fall 2.92% and 2.14%. Callgrind prices SHA-256 in software
because valgrind masks the SHA CPUID bit, charging roughly 49 `Ir` per hashed
byte where the hardware retires `sha256rnds2`; it also counts `rep movsb` per
byte (0604 measured a 35× overstatement). The cycles are the number to believe,
and the callgrind counts rank the work rather than predicting the latency.
Change 0645's native figures were −4.93% and −4.58%: reproduced.

**The real deck gains 3.37% of wall clock on an edit change 0649 measured at
133.61 ms and attributed 93.9% to MCE rewriting.** That row's floor is ±0.14%,
so the effect is fifteen times the floor even though the fingerprint is a small
minority of the phase. `perf stat` around the whole child reports only −1.53%
against a +1.38% floor, because the child also runs the corpus setup that the
per-sample p50 excludes; the two are measuring different denominators, and the
paired p50 is the per-operation figure.

**The regression is real in wall clock this time, and it is reported as such.**
`pptx_slide_remove_boundary_save` is **+2.57% at p50 and +3.01% at p95** against
A/A and B/B floors of −0.28% and −0.29%. Change 0645 could not separate this
from its noisy window; this window can, and the answer is that the selector is
genuinely slower by about 51 µs per lifecycle. It stays below the 5% review
trigger, and its mechanism is the measured per-part tier on nine fingerprints
that collect no hit.

**The real-file lifecycle row is not interpretable and is reported in full
anyway.** Its A/A p50 floor is −2.65% against a −2.76% effect, its before leg
carries a filesystem outlier that puts its mean at 143.9 ms against a 129.8 ms
p50, and its A/A p95 floor is −50.99%. Its `perf stat` cycles (−3.02% against a
+0.29% floor) and its fingerprint subtree (−46.20%) are the usable measurements
for that selector; its wall-clock tail is not.

## Correctness evidence

### The gates change 0645 set

| # | gate | result |
| ---: | --- | --- |
| 1 | the format bump is complete | `a_superseded_slide_removal_patch_is_refused_by_name`, `a_superseded_cross_slide_copy_patch_is_refused_by_name`: a v2/v3 patch round-trips `to_bytes`/`from_bytes` to an equal value; the forward **and** the inverse under the superseded magic are refused by `from_bytes` and `from_bytes_with_limits` with `Error::DurablePatchRevisionFormat { found, expected }`, whose message names both magics; an unrecognized magic (`LPRM9999`, `LPCP9999`) still returns `Error::Invalid` |
| 2 | verdict preservation, corpus-wide | `revision_v2_preserves_every_pairwise_verdict_over_the_pptx_corpus`: **78** `.pptx` fixtures under `test-data/`, expanded to **856** packages by the base, a byte-identical re-allocation and nine single-input mutations each, **9,396 ordered pairs**, **0 disagreements** between `packages_equal`, `R₁` and `R₂`. `revision_v2_preserves_every_pairwise_verdict_of_v1` runs 0645's synthetic 11-package corpus: **121 ordered pairs, 0 disagreements**. The v1 implementation is retained under `#[cfg(test)]` for exactly this gate |
| 3 | the alias gate | `a_mismatched_blob_arc_part_is_never_memoized`: a `Part` whose `blob_arc()` returns a different payload from its `blob()` is absent from the memo (`len == part_count − 1`), the revision it produces equals the revision of an honest part with the same visible bytes under both `R₁` and `R₂`, and a parent memo taken over that package answers nothing for it |
| 4 | the ABA gate | `the_memo_retains_every_allocation_it_keys`: the package that owned a 4 KiB payload is dropped, 64 different 4 KiB payloads are allocated, none lands on the retained key's address, and none answers from the memo; the original entry still describes its own bytes |
| 5 | the retention gate | `memoized_recapture_equals_a_cold_recapture` asserts directly, after a capture, a commit, a publication and a durable-patch replay, that every memo key names an allocation the owning snapshot's package holds and that the memo is never the sole owner of a payload (`Arc::strong_count ≥ 2`). `the_facade_memo_never_outlives_the_graph_it_describes` asserts the same of the facade's memo after a publication, after a typed edit that releases it, and after a slide-removal patch that adopts an exact one |
| 6 | no regression above the review trigger, the attributable one disclosed | see **Measured**; `pptx_slide_remove_boundary_save` is re-measured and reported in full below |
| 7 | the facade carry lands with the memo | it does: `Package::apply_opened_presentation_patch` is the route the two bare-patch replays take, and the measurement below separates the two |
| 8 | native confirmation | `perf stat` cycles and paired timing with the window's floor, below |
| 9 | ADR 0003 conflict determinism re-proved | the crate's existing stale-source, forged-patch and drift tests run unchanged and pass. The forged-patch test constructs its header literal from `DurablePatchFormat::SlideRemovalV2.magic()` instead of the `b"LPRM0001"` literal it used before; nothing else about it moved |

### Suite

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-pptx --all-targets` | clean (workspace lints are deny) |
| `cargo doc -p litchi-pptx --no-deps` | clean (rustdoc lints deny) |
| `cargo test -p litchi-pptx` | **885 passed, 0 failed** |
| `cargo test -p litchi --features docx,xlsx,pptx,xls` | **265 passed, 0 failed** |
| `cargo test` in `tools/perf-baseline` | **540 passed, 0 failed, 1 ignored** across 19 test binaries |
| `python3 tools/non_iwork_gate.py verify` | exit 0 |

Every capture in every one of those tests re-derives its revision the cold way
through the `debug_assert!` in `capture_internal`, and every reused payload
digest is recomputed and compared, so the suite is evidence of value identity
and not only of compilation.

## Validation preserved

No validation is removed, weakened, reordered or made conditional.
`capture_internal` still runs the presentation-root resolution, the
slide-reference one-to-one checks, the relationship-type checks, the per-slide
name parse, ADR 0013's `notes::load_snapshot` topology check and the name-index
build unconditionally and in the same order, before the revision is taken. The
memo sits strictly inside the hash, below every check.

The limits are untouched: `Limits::max_parts` still bounds the slide count at
the same place, and every memo allocation goes through `try_reserve`, so an
allocator refusal is `Error::Allocation { resource: "opened-presentation part
digests" }` rather than an abort.

The durable-patch byte limit still runs **first**, before the superseded-magic
test, so an over-long input is still refused with `Error::Limit` and the
malformed-input defence keeps its position. Within the limit, the refusal moves
earlier than any header parse, which is the direction change 0652's trade-off 3
allows: less work on the path that will refuse, and the refusal is typed and
precedes any partial result.

## Limitations

- **Not claimed:** any speedup. `performance_claim: none`, no claim-registry
  entry; the paired timings are reported beside the window's floor rather than
  as a result, and the instruction counts rank work rather than predicting
  latency.
- Callgrind runs SHA-256 in software because valgrind masks the SHA CPUID bit,
  so every instruction share attributed to hashing here is roughly five times
  its native cycle share (0649 measured 6.4×). The `perf stat` cycles are the
  counterweight and they are the number to believe.
- The counts are exact for six selectors on their corpora, on this host and
  these builds, and the cold-fingerprint figure for the eager deck is a
  subtraction rather than a direct reading, because a cold capture and a
  memoized replay both enter through `capture_internal`. The eager corpora carry no physical source provenance and
  re-deflate all media on save (change 0590's caveat, unchanged), so their
  wall-clock denominator is inflated by recompression a production `open` would
  not perform.
- The memo's resident cost is computed from `HashMap::capacity`, not from an
  allocator counter. No RSS, cold-cache, range-source, concurrency or
  cross-platform measurement was taken, and the harness's allocation binary has
  no metrics for the ordinary-save family (0649).
- The 9-bit collision-margin figure is a union bound at `k = 256`, not an
  analysis of SHA-256.
- The corpus oracle compares every ordered pair *within* each fixture's mutation
  corpus, not every pair across fixtures; distinctness across unrelated packages
  rests on the injectivity argument rather than on an exhaustive 856 × 856
  sweep.
- **A memo is a performance asset and this change does not make it a universal
  one.** The facade releases it at every mutation that publishes no
  opened-presentation snapshot, so a caller that interleaves typed edits with
  opened publications pays a cold hash after each typed edit. That is the safe
  direction and it is chosen deliberately; a generation-stamped memo that
  survived such a mutation is possible and is not attempted here.
- The cross-package copy path's *physical* revisions and change 0598's
  `physical_revision` cache are untouched and unmeasured. The per-part memo is
  reachable from the cross-copy path only through the snapshots it captures;
  change 0656 owns that path and may consume `Snapshot::part_digests` directly
  if its candidate retention needs it.

## Retained evidence

[`results/change-0655/README.md`](results/change-0655/README.md) — the callgrind
annotations, call counts and isolation pairs for both legs and every selector;
the memo accounting probe's per-fingerprint output and the reverted
instrumentation patch that produced it; the paired timing and `perf stat`
reports in run order A1 B1 B2 A2 with the A/A and B/B floors; the gate tails;
the binaries' sha256s; the decision record; the cleanup record; and the four log
paragraphs the coordinator merges.
