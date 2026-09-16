# 0656: the cross-package copy retains its planned candidate archive under a declared budget, and stops deflating it twice

Status: retained, implemented in `litchi-pptx` (with one new accessor in
`litchi-opc`). `performance_claim: none` — no claim-registry entry; the paired
medians and counts below are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements the design change
[0646](0646-pptx-cross-copy-candidate-retention-design.md) froze, under the
authority change [0652](0652-owner-decisions-for-the-third-wave.md) decision 5.
Base `70d7768cc6dada420ede063f72c88dc99ad30383`; branch
`perf/0656-pptx-cross-copy-candidate-budget`.

## What was changed

A plan-plus-apply cross-package slide copy used to build the candidate package
**twice**: once inside `plan_cross_slide_copy`, to prove a complete candidate
exists and to derive `target_revision` and `target_physical_revision`, and once
inside `apply_cross_slide_copy_plan`, because application replans from the live
packages and compares the fresh plan against the durable one. Each build
serializes the candidate through `OpcPackage::to_stream`, which raw-copies the
destination's unchanged members out of its retained archive and **deflates the
copied closure**.

`CrossSlideCopyPlan` now keeps the archive the first build produced, and the
second build reuses those bytes instead of serializing and deflating the
candidate again. Everything else about the second build is unchanged: the
in-memory candidate graph is still constructed from scratch, and the whole
replan around it still runs every refusal it runs today.

Retention is bounded by a new member of the crate's existing finite limit
policy, `opened::Limits::max_retained_candidate_bytes`, intersected between the
two snapshots like every other member. A candidate above the budget is **not
retained**, and its application rebuilds the archive exactly as it does today;
exceeding the budget is never a refusal.

Concretely:

- `crates/litchi-pptx/src/opened/model.rs`: `Limits` gains
  `max_retained_candidate_bytes` (default **64 MiB**), its accessor, and the
  sixth `Limits::new` argument with the same nonzero rule as the other five.
- `crates/litchi-pptx/src/opened/cross_copy_plan.rs`: `CrossSlideCopyPlan`
  gains a private `RetainedCandidateSlot`; `build_candidate` takes the
  operation's `Limits` in place of a bare archive bound plus one
  `CandidateArchive` argument (`Build` / `BuildAndRetain` / `Reuse`) that says
  what this build does about its archive; `plan_cross_slide_copy` asks for
  retention, `apply_plan` reuses what the plan holds, and both `apply_patch`
  routes build as before. `intersect_limits` intersects the new member.
- `crates/litchi-pptx/src/opened/patch.rs`: its `intersect_limits` likewise.
- `crates/litchi-opc/src/package.rs`: `OpcPackage::exact_source_shared`, a
  shared handle to the exact owned source archive, so the plan can take a
  second owner of the allocation the candidate reopen already holds instead of
  copying it.

**No new `unsafe`, no new dependency, no weakened limit or malformed-input
defence, no new refusal, and no published byte changes on any route.**

## Authority

Change 0652 decision 5, in the owner's words:

> "Cross-package copy: add a budget for that."

which 0652 reads as authorizing

> 0646's design with a retained-candidate budget as a public `opened::Limits`
> field (a breaking constructor change is acceptable); a candidate over the
> budget falls back to today's rebuild rather than refusing, because rebuilding
> is always available

and which requires this record to prove

> 0646's fourteen gates; the default budget stated with its reason; the fallback
> exercised by a test; the retained bytes charged and released as 0646 measured.

0652's standing trade-off 1 ("breaking changes are totally acceptable")
authorizes the sixth `Limits::new` argument. Trade-off 2 (correctness first)
decided two points below: the archive copy at application stays fallible rather
than becoming an infallible `Vec::clone`, and the physical revision stays a hash
recomputed over the bytes about to be published rather than a value carried
beside them.

**What decision 5 changes about 0646.** 0646's gate G1 demanded that the default
planning route retain nothing and that retention be reachable only through an
explicit policy argument. Its reason was stated exactly: "a 128 MiB hold on a
caller-owned public value is not an invisible cache", and 0646 had no way to
declare or bound that hold. Decision 5 supplies the bound, and with it the
reason G1 existed. A ceiling that is declared in `Limits`, intersected between
the two snapshots, enforced, observable through
`CrossSlideCopyPlan::retained_candidate_bytes`, releasable through
`release_retained_candidate`, and lowerable by any caller to a value no
candidate can meet is not an invisible cache whatever its default; the opt-in
*shape* G1 asked for is the budget. This record therefore sets a default that
retains the ordinary case, gives the four grounds for it below, and states its
measured cost rather than netting it out. 0646's three-valued policy
(`Rebuild` / `RetainUpTo` / `Require`) is **not** implemented: `RetainUpTo` is
what the budget does, and decision 5 removes `Require` by saying the
over-budget case falls back rather than refusing.

## Breaking changes

| item | before | after |
| --- | --- | --- |
| `litchi_pptx::opened::Limits::new` | five arguments | **six**: `max_retained_candidate_bytes` last, rejected when zero like the other five |
| `litchi_pptx::opened::Limits` | five members | six; `Limits::DEFAULT` gains `max_retained_candidate_bytes: 64 * 1024 * 1024`, so `Limits::default()` changes value |
| `litchi_pptx::opened::Limits` `Debug` | five fields | six |
| `litchi_pptx::opened::CrossSlideCopyPlan` `Debug` | seventeen fields | eighteen: a `candidate` field reporting `retained_bytes` and `bound`, never the archive |

Additive, not breaking:

- `litchi_pptx::opened::Limits::max_retained_candidate_bytes(self) -> usize`
- `litchi_pptx::opened::CrossSlideCopyPlan::retained_candidate_bytes(&self) -> Option<usize>`
- `litchi_pptx::opened::CrossSlideCopyPlan::release_retained_candidate(&mut self)`
- `litchi_opc::OpcPackage::exact_source_shared(&self) -> Option<Arc<Vec<u8>>>`
  — a wider surface than the `pub(crate) exact_source()` it sits beside: it
  hands the complete original archive of any owned package to any downstream
  crate. It leaks no new information (for an exact-source-authorized package
  `to_bytes` already short-circuits to exactly those bytes) and the bytes are
  immutable through the handle, because `source_archive` is only ever assigned
  wholesale and is never reached by `Arc::get_mut` or `Arc::make_mut`. It exists
  because the two crates are separate and `litchi-pptx` is its only consumer;
  a narrower seam would be a `litchi-opc` ownership design, which ADR 0011
  governs and this change does not open.

Unchanged: every error type, variant, `resource` string and message; `PartialEq`
and `Eq` on `CrossSlideCopyPlan` (see below); `CrossSlideCopyPatch` in every
respect, including its magic, its six 32-byte revisions and its encoded bytes.
(That magic is `LPCP0002` on this base. Change 0655 bumps it to `LPCP0003` in
the same wave, for reasons that have nothing to do with retention; the test
that pins this checks the `LPCP` family prefix and byte equality between the
retaining and the released plan, so it survives that bump.)

One call site outside the crates had to move with the constructor:
`tools/perf-baseline/src/pptx_slide_boundaries.rs` passes a sixth argument. It
is a boundary probe that does no cross-package copy, so the value it passes
(`1`, retention off) changes nothing it measures; it is named here because the
two measured legs' harness sources therefore differ by that one argument.

### The default, and why 64 MiB

The default `max_retained_candidate_bytes` is **64 MiB**. Four reasons, in the
order they bind:

1. **It retains the ordinary case, which is the point of a budget rather than
   an opt-in flag.** The retained archive is `destination_archive + deflate(copied closure) + ZIP
   framing`. Measured here: 31,545 B on the plain corpus and 33,599,873 B on the
   media-rich one. 64 MiB retains both, the larger at 52% of the ceiling. Every
   `.pptx` under `test-data/` is smaller than 972,788 bytes, so the default
   retains all 78 of them by a factor of at least 69.
2. **It leaves the default worst case inside a server profile.** A plan's owned
   bytes are bounded by `max_patch_bytes` (128 MiB by default) for its copied
   closure; retention adds at most `max_retained_candidate_bytes` more,
   unshared. At 64 MiB the default worst case per plan is 192 MiB rather than
   the 256 MiB an unbounded ceiling would allow — which is the entire
   `Resource::Memory` limit of `litchi_core::budget::Profile::Server`. The
   crate's own precedent sits beside it: `max_history_bytes` already permits
   256 MiB of retained undo patches at the default limits.
3. **It is the configuration 0646 measured.** 0646's scratch retained
   unconditionally under a 64 MiB ceiling, and every number in that record was
   taken under it. Choosing the same value makes this record's measurements
   directly comparable with the design's rather than a new point.
4. **A caller can lower it to 1 and turn retention off.** `Limits::new` rejects
   zero for every member, and no ZIP archive is one byte, so
   `max_retained_candidate_bytes = 1` is the off switch. A caller can also raise
   it; the archive bound `max_patch_bytes` still caps what can be produced at
   all.

The value is **not** swept and is not derived from a cost model. It is a
declared ceiling chosen to retain the measured cases and to keep the default
plan inside a server memory profile.

## Why it is sound

### What the plan holds, and what that costs

`RetainedCandidate` holds two things: `archive: Arc<Vec<u8>>` and
`bound: usize`. The `Arc` is **the one the candidate reopen already holds** —
`build_candidate` serializes into a bounded `Vec`, moves it into
`OpcPackage::from_vec_reusing_payloads` (or `from_vec`), which authorizes it as
the package's exact source behind an `Arc`, and `exact_source_shared` takes a
second owner of that same allocation. Planning therefore allocates and copies
**nothing new**; the only change is that the allocation is not freed when the
candidate package is dropped at the end of planning. Retention is a decision not
to free.

`bound` is the archive bound the bytes were produced under. Reuse requires
`held.bound == archive_limit && held.archive.len() <= archive_limit`, so a
build under a tighter bound is a miss and rebuilds under the tighter writer.

**No digest is carried.** 0646's design stored `archive_digest` beside the bytes
and recomputed it at application anyway; this implementation does not store it
at all, which makes 0646's gate G5 ("the digest is recomputed, not carried") true
by construction rather than by discipline. The recomputation costs one SHA-256
pass over bytes `bounded_package_bytes` hashes anyway, so the hash is a wash and
what reuse removes is exactly the serialization and its deflate.

### Plan equality and `Debug`

`RetainedCandidateSlot` implements `PartialEq` as "always equal", so
`plan_a == plan_b` is exactly the comparison it has always been: the retained
archive is not part of a plan's value. Two plans that agree on every proof field
publish the same bytes whether or not one of them happens to be holding a copy
of them, and `release_retained_candidate` is therefore value-preserving. This is
0646's gate G8 satisfied in its strongest form — plan equality can never become
a byte comparison of two archives, because it does not look at them — and it is
what 0646's own sentence "the new field is private and does not change what
`plan == plan` means" asks for. (0646's scratch compared a stored digest, which
made a released plan differ from its unreleased self.)

`RetainedCandidateSlot` also implements `Debug` by hand, reporting
`retained_bytes` and `bound`. A derived `Debug` would print 33.6 MB of archive
on the media-rich corpus.

### How application proves the retained archive is the one the plan validated

Retention changes one expression in `build_candidate`. The proof chain at
application, in the order `apply_plan` runs it:

| # | check | still a real check? |
| --- | --- | --- |
| 1 | `package_fingerprint(source) == plan.source_revision` | yes — pins the source graph |
| 2 | `physical_package_fingerprint(source) == plan.source_physical_revision` | yes — pins the source archive bytes |
| 3-4 | the same two for the destination | yes |
| 5 | the whole replan: provenance, signature, unknown members, macro, dialect, MCE, protection, collisions, `MAX_SLIDES`, slide IDs, part counts, master/layout graph, content types, slide surface, registered layouts, layout inheritance, owned closure, cycles, collision remap, resulting parts, planned bytes, preflight | yes — **unchanged**, none of it is skipped |
| 6 | the retained bytes are re-hashed and sealed into `fresh.target_physical_revision`, compared with `plan.target_physical_revision` | **narrower on the reuse arm, and the record says so**: it proves the bytes the plan is holding *now* are the bytes planning hashed, so it catches a substituted or corrupted archive; it does not re-prove that a fresh serialization would reproduce them |
| 7 | the retained bytes are reopened and captured; `fresh.target_revision` compared with `plan.target_revision` | yes — the reopen parses the retained bytes into a graph and that graph's revision is compared, so a retained archive that decodes to a different package is caught |
| 8 | `validate_candidate` captures the reopened candidate again, runs `validate_before` on the live destination and `validate_after` on the candidate, and compares `snapshot.revision()` with `plan.target_revision` | yes |

**Be exact about what row 6 stops proving.** Before this change it was the
comparison that established that a from-scratch `to_stream` of the *apply-time*
candidate graph reproduced the plan-time archive, because `fresh` was built by
that serialization. On the reuse arm `fresh.target_physical_revision` is the
seal of a hash taken over the retained bytes themselves, so it can only fail if
those bytes are no longer the ones planning hashed. It is still a real check
against substitution and corruption — that is what
`a_substituted_retained_archive_…` exercises — but it is not the physical
proof it used to be.

What is therefore no longer checked in a release build is one thing: that a
fresh `to_stream` of the in-memory candidate graph would reproduce the retained
bytes. Its proof is that (1)–(4) pin both input packages — the semantic graph
by `package_fingerprint` and the serialized archive by
`physical_package_fingerprint` — and that `to_stream` is a deterministic
function of the package. A `debug_assert` in `build_candidate` re-derives it on
**every** reuse in debug and test builds, which is every one of the 885
`litchi-pptx` tests that reaches a cross-package copy.

**One input to the candidate's serialization is not pinned by (1)–(4), and it
is now covered by a test.** The candidate is published through the
destination's *preservation source* — the archive bytes the destination package
was opened from — not through the destination's own re-serialization, and two
destinations can agree on both fingerprints while carrying different
preservation sources. The reachable case is: plan against a destination whose
exact-source authorization was revoked by an edit, then apply to a clean reopen
of that destination's serialized bytes. Both fingerprints match, but the
plan-time destination republishes its pristine members out of archive `A` and
the apply-time one out of `B = to_stream(A-derived package)`.
`a_dirty_planning_destination_and_a_clean_applying_one_agree` builds exactly
that pairing and asserts the verdict and the published bytes are the same with
and without the retained archive; in a debug build the assertion above runs on
it, so the equality is measured for that fixture rather than argued. The general
statement — that this holds for every package — is *not* proved, and the
Limitations say so.

That assertion is also the positive proof that application takes the reuse
branch at all: `a_substituted_retained_archive_trips_the_debug_assertion`
panics with the assertion's own message, which can only be reached inside the
`Reuse` arm, and its release-build twin is refused rather than published for the
same reason.

### Which capture may be reused, and which may not

0646 §6 froze the rule that a plan may retain the candidate **bytes** and may
not retain the candidate **snapshot**, for four reasons that this
implementation keeps intact: `validate_candidate`'s capture runs *after*
`validate_before` where the planning capture runs before any staleness check;
`validate_application_candidate` has a branch in which the prepared candidate is
dropped and a different one built, so a planning-time snapshot can belong to the
discarded package; the two captures are taken under different limits at planning
time; and retention must not be allowed to make a second reuse look free. This
change retains no `Snapshot` and skips no capture. Change 0598's open question
stays answered **no**.

### Why an over-budget candidate falls back rather than refusing

A typed `Error::Limit` protects a caller from work or memory it did not ask for.
There is nothing to protect it from here: rebuilding the candidate is always
correct, is what the code does today, and is what a caller who never opts into
retention gets. 0646 reached the same conclusion about a scratch-backed variant
("Litchi is never obliged to hold the candidate; rebuilding it is always correct
… so a scratch-backed retention must fall back to `Rebuild` on absence, not
refuse"), and decision 5 makes it the rule for the budget. The budget therefore
changes what a plan **holds** and never what it **proves**, refuses or
publishes. ADR 0005's retained-state wording is amended to say so.

### No spill

A candidate archive is a complete PPTX package — document content. ADR 0005
forbids litchi from spilling such content automatically. Nothing on this path
touches the filesystem, a temporary file or a scratch provider, and a source
ratchet test (`the_retention_path_never_spills`) keeps it that way by refusing
fourteen markers in `cross_copy_plan.rs`.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0; every measured process pinned to **CPU 11** while seven other
agents of the same wave built and measured on the other cores. Both legs built
`--release --locked` from `tools/perf-baseline`: before from the shared
read-only checkout of `70d7768cc`, after from this branch. Every binary was
staged outside its target directory before any measurement (change 0627), and
its sha256 is recorded, because `lto = true` makes release binaries
non-reproducible byte for byte (change 0635). The after leg was built twice —
once for the first three windows and once from the committed tree, after three
edits that change no behaviour (two doc comments, a `derive(Debug)` removed from
a private struct nothing formats, and test code). The counts below are the
first build's and are **re-verified on the committed tree's binary**: identical
call counts, `compress256` inclusive `Ir` identical to the digit, and
per-lifecycle totals within 0.11% (plain) and 0.0001% (media-rich).
Selectors: `pptx_cross_copy_plain` (a 30,539-byte destination, one copied part),
`pptx_cross_copy_media_rich` (a 16,814,664-byte destination, nine copied parts
totalling 16,784,397 bytes of incompressible media), and their `_lifecycle`
variants which time the same work with setup and reopen included. A "lifecycle"
below is one plan plus one apply.

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
The one call-count row that falls without being the point is `compress256`, and
it falls *without its work falling* (below): the removed `BoundedVecWriter`
hashed the archive in ZIP-writer-sized chunks, while the reuse path hashes it in
one `update`, so the same bytes reach the same compression function in fewer,
larger batches. This table is 0646's table to the digit.

### Instructions per lifecycle

| inclusive `Ir`, per lifecycle | `pptx_cross_copy_plain_lifecycle` | `pptx_cross_copy_media_rich_lifecycle` |
| --- | ---: | ---: |
| **whole lifecycle** | 239,779,454 → **233,987,606** (**−2.42%**) | 26,801,253,558 → **23,735,984,702** (**−11.44%**) |
| `bounded_package_bytes` | 14,172,802 → 7,077,395 (**−50.06%**) | 9,634,316,498 → 4,816,910,104 (**−50.00%**) |
| `PreservationIndex::write_to` | 10,987,341 → 5,501,564 (−49.93%) | 9,624,080,075 → 4,811,741,902 (**−50.00%**) |
| **`zlib_rs::deflate::deflate`** | 2,492,187 → 1,246,050 (**−50.00%**) | 5,955,553,990 → **2,977,751,014** (**−50.00%**) |
| `PackageWriter::write_to_stream` | 20,582,065 → 13,527,842 (−34.27%) | 13,164,074,465 → 8,349,574,762 (−36.57%) |
| `build_candidate` | 45,134,285 → 39,525,614 (−12.43%) | 13,474,613,273 → 10,409,950,655 (−22.74%) |
| `prepare_cross_slide_copy_for_slides` | 133,892,085 → 128,274,476 (−4.20%) | 15,315,262,906 → 12,250,542,758 (−20.01%) |
| `package_fingerprint` | 39,471,192 → 39,463,252 (−0.02%) | 10,525,046,995 → 10,525,041,229 (−0.00%) |
| `physical_package_fingerprint` | 6,500,099 → 6,500,444 (+0.01%) | 3,501,940,474 → 3,501,941,859 (+0.00%) |
| `capture_internal` | 58,569,841 → 58,154,634 (−0.71%) | 8,804,820,904 → 8,804,800,449 (−0.00%) |
| **`sha2::sha256::compress256`** | 50,665,038 → 50,658,460 (**−0.01%**) | 19,274,528,456 → **19,274,471,600 (−0.00%)** |

Whole child at `--samples 1`: plain 4,361,066,483 → 4,350,145,029 (−0.25%);
media-rich 123,608,576,899 → 117,482,830,409 (−4.96%). Most of the whole child
is the harness's corpus construction and its ten untimed refusal gates.

**The arithmetic closes on the media-rich corpus.** One `bounded_package_bytes`
is worth 4,817,406,394 `Ir`. The reuse path gives back the SHA-256 — hence
`compress256` unchanged to five significant figures — and adds one fallibly
reserved 33,599,873-byte copy. Net: **−3,065,268,856 `Ir` per lifecycle**, of
which 2,977,802,976 is the deflate and the remaining 87,465,880 is the ZIP
framing and the raw member copy: **the deflate plus the framing, and nothing
else.** Every figure here reproduces 0646's to within 0.01 percentage points on
the media-rich corpus, which is the check that the landed change is the change
that record priced.

### Memory: what the plan holds, exactly

`litchi-perf-baseline-alloc retention --api owned`, 5 samples after 1 warmup per
leg, medians; the allocator counters are deterministic. The probe checkpoints an
owned cross-copy lifecycle at nine ownership boundaries.

| checkpoint, live bytes | plain before | plain after | Δ | media-rich before | media-rich after | Δ |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| baseline | 161,426 | 161,446 | +20 | 68,909,299 | 68,909,319 | +20 |
| inputs and sink ready | 352,089 | 352,109 | +20 | 169,804,866 | 169,804,886 | +20 |
| documents opened | 689,142 | 689,162 | +20 | 203,744,003 | 203,744,023 | +20 |
| **planned** | 696,972 | 728,577 | **+31,605** | 203,757,865 | 237,357,798 | **+33,599,933** |
| published | 782,885 | 814,490 | +31,605 | 237,438,870 | 271,038,803 | +33,599,933 |
| result dropped | 748,962 | 780,567 | +31,605 | 237,392,123 | 270,992,056 | +33,599,933 |
| **plan dropped** | 741,132 | 741,152 | **+20** | 237,378,261 | 237,378,281 | **+20** |
| document handles dropped | 290,052 | 290,072 | +20 | 136,174,581 | 136,174,601 | +20 |
| sink dropped | 161,426 | 161,446 | +20 | 68,909,299 | 68,909,319 | +20 |

**The delta closes to the byte, and it is the archive plus an `Arc` header.** The
two binaries differ by a constant +20 bytes of static allocation, visible at
every checkpoint where the plan holds nothing. At `planned`:

- plain: 31,545 (archive) + 40 (`ArcInner<Vec<u8>>`: two counters plus the `Vec`
  triple) + 20 = **31,605**;
- media-rich: 33,599,873 + 40 + 20 = **33,599,933**.

Nothing else is retained, and **dropping the plan releases it exactly**: at
`plan dropped` the two legs are back to the constant +20. 0646's gate G12 is
satisfied on the landed change — retention is a decision not to free, not a
second copy — and so is 0652's requirement that "the retained bytes [be]
charged and released as 0646 measured", to the byte. (The first after build,
the one the first three timing windows used, gives the same table with a
constant of −4 in place of +20 and the same two `planned` deltas net of it;
both probes are retained.)

| aggregate over the lifecycle region | plain before | plain after | Δ | media-rich before | media-rich after | Δ |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| **region peak live bytes** | 1,339,638 | 1,266,694 | **−5.45%** | 271,884,239 | 305,008,756 | **+12.18%** |
| allocation calls | 48,096 | 47,041 | −2.19% | 58,250 | 56,770 | −2.54% |
| reallocation calls | 5,659 | 5,298 | −6.38% | 6,765 | 6,271 | −7.30% |
| allocated bytes | 15,581,584 | 13,175,380 | **−15.44%** | 467,072,982 | 373,765,695 | **−19.98%** |

**The peak moves in opposite directions on the two corpora.** On the plain corpus
it *falls* 5.45%: the removed second `BoundedVecWriter`, with its geometric
growth and its compaction copy, is worth more than the 31,545 bytes retained. On
the media-rich corpus it rises by 33,124,517 bytes — one archive, to within 1.4%
— because the allocation that used to be freed at the end of planning now sits
underneath the application's own peak. Total bytes through the allocator fall
19.98% on the media-rich lifecycle (93,307,287 bytes). **This is the change's
real cost and it is not hidden in a mean: a caller that plans a media-rich
cross-copy and then applies it now peaks about 12% higher, and the default
budget is the control it lowers if that is the wrong trade for it.**

### Paired native timing

Four legs in run order **A1 B1 B2 A2** (before, after, after, before), all four
selectors in one process per leg, 30 measured samples after 3 warmups,
`taskset -c 11`. A/A is A2 against A1 and B/B is B2 against B1, both inside the
same window. Four windows were run because the host carries seven other agents
of this wave; **window 4 is the one quoted**, because it is the window whose
after binary is built from the committed tree and its A/A floor at p50 is under
5% on every selector. Window 2 is quoted beside it: its floors are the tightest
of the four (A/A p50 under 2% everywhere), and its after binary predates three
edits that changed no behaviour — two doc comments, a `derive(Debug)` removed
from a private struct nothing formats, and test code. That claim is checked, not
asserted: re-running the callgrind isolation pairs on the committed tree's
binary reproduces the counts exactly (deflate calls 20 and 556 per lifecycle,
`bounded_package_bytes` 1, `compress256` inclusive `Ir` identical to the digit)
and the per-lifecycle totals to within 0.11% on the plain corpus and 0.0001% on
the media-rich one. Window 1 (load average 26 rising to 59 during the run) has
A/A p50 floors of +145%, +3.4%, −57% and −6.0%: **its floor exceeds 5%, so its
selector totals say nothing**, and it is retained for its `perf stat` legs and
its phase table. Window 3 is a media-rich-only 40-sample window used to settle
two phases.

Window 4 (load average 19.5 falling to 18.2), the committed tree's binary:

| selector | before p50 | after p50 | p50 Δ | p95 Δ | p99 Δ | A/A p50 | B/B p50 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_cross_copy_plain` | 8.048 ms | 7.676 ms | **−4.62%** | −4.92% | −9.27% | −1.35% | +0.96% |
| `pptx_cross_copy_media_rich` | 663.198 ms | 417.352 ms | **−37.07%** (before is +58.91% of after) | +31.31% | +32.03% | −2.36% | +30.20% |
| `pptx_cross_copy_plain_lifecycle` | 9.248 ms | 9.102 ms | **−1.58%** | +87.51% | +88.70% | −3.12% | +4.99% |
| `pptx_cross_copy_media_rich_lifecycle` | 673.748 ms | 413.485 ms | **−38.63%** (before is +62.94% of after) | −31.77% | −19.78% | +4.71% | +2.36% |

Window 2 (load average 17.8 falling to 14.4), the tightest floors:

| selector | before p50 | after p50 | p50 Δ | p95 Δ | p99 Δ | A/A p50 | B/B p50 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_cross_copy_plain` | 7.937 ms | 7.706 ms | **−2.91%** | −3.11% | −3.23% | +0.65% | −0.32% |
| `pptx_cross_copy_media_rich` | 628.577 ms | 408.349 ms | **−35.04%** | −39.14% | −44.03% | **+0.62%** | −3.87% |
| `pptx_cross_copy_plain_lifecycle` | 9.160 ms | 8.978 ms | **−2.00%** | −3.83% | −3.19% | −1.68% | −0.11% |
| `pptx_cross_copy_media_rich_lifecycle` | 648.941 ms | 432.309 ms | **−33.38%** | −33.01% | −39.78% | **+0.01%** | −0.67% |

**The tails in window 4 are worse, and they are leg noise, not the change.**
`pptx_cross_copy_media_rich` p95 and p99 rise 31–32% and
`pptx_cross_copy_plain_lifecycle`'s rise 88%, entirely because the B2 leg took
excursions: its phase table below reads a 422 ms planning median and a 32 ms
reopen against 316 ms and 21 ms in its own pair, neither of which is code this
change touches. Window 2's tails on the same selectors fall 3–44%, and its B/B
floor is under 4% everywhere. **Both are reported; neither is averaged away.**
The published digest is one value per corpus across every leg of every window
(`3e9ae2805a5d…` plain, `6a3536fc7d90…` media-rich).

**The effect is in `apply`.** The selectors clock plan and apply separately.
Per-leg medians on `pptx_cross_copy_media_rich_lifecycle`, window 4:

| phase p50 | A1 (before) | B1 (after) | B2 (after) | A2 (before) | Δ |
| --- | ---: | ---: | ---: | ---: | ---: |
| **apply** (reported as `commit`) | **328.209 ms** | **89.394 ms** | **90.554 ms** | **341.400 ms** | **−73.41%** |
| plan | 298.080 ms | 297.910 ms | 301.410 ms | 314.366 ms | −3.47% |
| publication | 1.524 ms | 1.518 ms | 6.524 ms | 1.538 ms | +318% |
| reopen | 21.312 ms | 21.247 ms | 21.550 ms | 21.287 ms | +0.60% |

Apply falls **−73.41%** with an after-pair spread of 1.3%. Across the four
windows the same phase reads −70.55%, −70.83%, −72.05%, −72.28% and −73.41% on
the media-rich selectors; 0646 predicted −72.10%.

**Planning is flat, and the windows disagree about its sign, which is the
point.** The four windows read the media-rich lifecycle's planning phase at
+5.19% (w2), +1.54% (w3) and **−3.47%** (w4), and the per-leg medians overlap
in every one of them: in window 4 the two after legs (297.910, 301.410 ms)
straddle one before leg (298.080) and sit well below the other (314.366). The
phase's own leg-to-leg floor is about ±4%. Planning does strictly less work than
before — it no longer frees the archive at the end of the call — so a rise is
not expected, and **no planning result is claimed in either direction**. It is
reported rather than averaged away because one window read it above 5%.

**One phase is bimodal, in both directions, and it is not the change.**
`publication_ns` on the media-rich selectors sits at either ≈1.5 ms or ≈6.4 ms,
and which mode a leg lands in does not follow the binary. Per leg:

| window, selector | A1 (before) | B1 (after) | B2 (after) | A2 (before) |
| --- | ---: | ---: | ---: | ---: |
| 1, `…media_rich` | 6.957 ms | 6.663 ms | 6.368 ms | 7.799 ms |
| 2, `…media_rich` | 1.500 ms | 6.501 ms | 6.232 ms | 1.547 ms |
| **3, `…media_rich`** | **6.362 ms** | **1.664 ms** | **6.888 ms** | **1.622 ms** |
| 4, `…media_rich` | 6.624 ms | 6.357 ms | 7.158 ms | 6.602 ms |
| **4, `…media_rich_lifecycle`** | **1.524 ms** | **1.518 ms** | **6.524 ms** | **1.538 ms** |
| 0646's window, `…media_rich` | 6.513 ms | 6.301 ms | 6.149 ms | 6.497 ms |

Windows 3 and 4 are decisive: window 3 puts one before leg and one after leg in
the high mode and one of each in the low mode inside a single run, and window 4
puts one of an after *pair* in each mode. Publication executes no changed code —
retention changes how the candidate archive is produced, not how the result is
published — and **the pooled ±4× on this phase is reported as leg variance, not
as a result in either direction.** This is the same phase change 0646 chased and
reached the same verdict about; this record adds the within-window witnesses
0646 did not have.

**Native `perf stat`**, whole child, the same A1 B1 B2 A2 order over all four
selectors in one process at `--warmup 1 --samples 10`:

| window | event | before mean | after mean | Δ | A/A | B/B |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| **4** | **cycles** | 104,543,849,314 | 78,786,657,086 | **−24.64%** | +1.71% | −1.86% |
| 4 | instructions | 228,484,701,788 | 157,845,215,208 | −30.92% | +1.16% | −1.34% |
| 2 | **cycles** | 106,120,965,699 | 77,460,759,244 | **−27.01%** | +1.46% | −3.12% |
| 2 | instructions | 229,944,053,935 | 155,664,295,884 | −32.30% | +0.29% | −1.92% |
| 1 | **cycles** | 107,882,282,144 | 78,665,225,380 | **−27.08%** | +0.82% | −3.52% |
| 1 | instructions | 229,780,334,316 | 156,582,850,834 | −31.86% | +1.20% | −2.09% |

This is 0646's admission gate **G11**, satisfied on the landed change: the
removed deflate is worth 24.6–27.1% of whole-child native cycles against A/A
floors of 0.8–1.7% in three independent windows, so the archive copy it trades
into does not eat the saving. It is also the counterweight to the callgrind
numbers: native instructions fall 30.92% where callgrind's per-lifecycle total
falls 11.44%, because SHA-NI executes the unchanged hashing in a small fraction
of the instructions valgrind's software implementation needs, which raises
deflate's share of the native mix in exactly the way the wall clock shows.

## Correctness evidence

**The whole `litchi-pptx` suite runs on the reuse path.** The default budget
retains, so every test in the crate that plans and then applies a cross-package
copy exercises the reuse, and every one of them re-serializes the candidate
under the `debug_assert`.

| leg | command | result |
| --- | --- | --- |
| before (detached worktree of `70d7768cc`) | `cargo test -p litchi-pptx` | **875 passed, 0 failed, 2 ignored** across 78 suites |
| after | `cargo test -p litchi-pptx` | **885 passed, 0 failed, 2 ignored** across 78 suites |
| after | `cargo test -p litchi-pptx --release --lib retention` | 12 passed, including the release-only substitution refusal |
| after | `cargo test -p litchi-opc` | 712 passed, 0 failed, 1 ignored across 27 suites |
| after | `cargo test -p litchi --features docx,xlsx,pptx,xls` | 265 passed, 0 failed, 7 ignored across 29 suites |
| after | `cargo test` in `tools/perf-baseline` | 540 passed, 0 failed, 1 ignored across 19 suites |

885 − 875 = 10, the ten tests this change adds. **Every test that passes on the
base passes on the reuse path, with the same assertions.** The pre-existing
adversarial suites are the substance of that: `source_backed_cross_copy.rs`,
`source_backed_cross_copy_adversarial.rs` and change 0598's
`revision_cache_tests.rs` assert exact `Error` kinds, operations and reason
strings for stale sources, stale destinations, foreign sources, forged patches,
collision remaps, limits and reversibility, and none of them moved.

The ten tests, in
`crates/litchi-pptx/src/opened/cross_copy_plan/retention_tests.rs`:

| test | what it pins |
| --- | --- |
| `a_plan_retains_the_archive_its_application_publishes` | The default budget retains; `retained_candidate_bytes()` reports the archive's length; the retained bytes seal — through `seal_physical_revision` — to the plan's own `target_physical_revision`; and the archive the application publishes is byte-identical to them. |
| `a_candidate_above_the_budget_falls_back_to_the_rebuild` | **The budget's two branches.** At one byte below the candidate's length the plan retains nothing and still applies, publishing the same bytes and the same revisions as the retaining plan; at exactly the candidate's length it retains. Exceeding the budget produces no error of any kind. |
| `the_tighter_retained_candidate_budget_binds` | The budget is intersected: a tight ceiling on either snapshot suppresses retention, and two wide ones admit it. |
| `releasing_and_dropping_return_the_retained_bytes` | `release_retained_candidate` returns the handle (`Arc::strong_count` falls), leaves the plan equal to what it was, and leaves it applicable through the rebuild; applying does **not** release; dropping the plan returns the bytes. |
| `debug_does_not_print_the_retained_archive` | `Debug` reports `retained_bytes` and never the archive, checked structurally (no ZIP local-file signature in the rendering) rather than by a size comparison a fixture could invert. |
| `a_substituted_retained_archive_trips_the_debug_assertion` (debug) / `…_is_refused_and_publishes_nothing` (release) | Replacing the retained archive with a different, structurally valid PPTX archive trips the `debug_assert` in a debug build and, in a release build where that assertion is compiled out, is refused by `Err(Invalid("cross-slide candidate did not publish the reserved slide identity"))`, leaving the destination byte-identical. This is the test that fails if a retained archive were ever trusted without a recomputed proof, and it is also the positive proof that application takes the reuse branch. |
| `retention_changes_no_apply_verdict` | The same application with and without the retained archive publishes identical bytes and identical revisions on the accepting route, and produces the identical `Result` on four rejecting routes, each asserted to be a refusal and not only to agree: a drifted destination, a foreign source, a drifted source, and borrowed graph-only ingress. |
| `the_durable_patch_is_untouched_by_retention` | `CrossSlideCopyPatch::to_bytes` is identical with and without retention and still carries the `LPCP` family magic; a patch-driven application publishes what the plan-driven one publishes. |
| `a_dirty_planning_destination_and_a_clean_applying_one_agree` | The one pairing in which the planning and applying destinations carry different preservation sources while passing all four staleness proofs: plan against a destination whose exact-source authorization an edit revoked, apply to a clean reopen of its own bytes. Verdict and published bytes are the same with and without the retained archive, and in a debug build the re-serializing assertion runs on it. |
| `the_retention_path_never_spills` | Source ratchet: `cross_copy_plan.rs` contains none of fourteen filesystem, temporary-file, scratch-provider or mmap markers. |

**Byte identity across every measured leg.** All four selectors verify their
published archive against an expected digest and report it; across every leg
that ran it — sixteen timing legs (windows 1, 2 and 4 over all four selectors,
window 3 over the two media-rich ones), twelve `perf stat` legs and twelve
callgrind profiles — the reported `output_sha256` is one value per corpus —
`3e9ae2805a5dfd8e57ceff4a3662a2772f4197850252f7a23df8e6525d58ee22` for the plain
corpora and
`6a3536fc7d9055eefad6eb115294f0157b7f8acd576a4898b4e402e48c3e7b4f` for the
media-rich ones — identical between the legs. The harness refuses to run any of
them without its ten untimed corpus gates (semantic output, package topology,
dependency closure, source immutability, collision remap, durable patch round
trip, borrowed-provenance refusal, stale-source refusal, stale-destination
refusal, foreign-source refusal), so each measured run is also a verdict check.

### 0646's fourteen gates, on the landed change

| # | gate | verdict |
| --- | --- | --- |
| G1 | the default route retains nothing, reachable only through an explicit policy argument | **re-read under 0652 decision 5, then met in its amended form.** The budget replaces the opt-in: the hold is declared in `Limits`, intersected, enforced, observable and lowerable. `a_plan_retains_the_archive_its_application_publishes` pins that the default retains and `the_tighter_retained_candidate_budget_binds` pins that a caller's ceiling is obeyed. The record says plainly what the default costs (the peak table above). |
| G2 | the retained bytes are charged — a `Budget`, **or** `opened::Limits::max_retained_candidate_bytes`, a sibling of the enforced `max_history_bytes`, intersected by `intersect_limits`, with `retained_candidate_bytes()` exposing what is held | **met** by G2's second branch, which is the branch 0652 decision 5 chose. Both `intersect_limits` functions carry the new member; `Limits::new` rejects zero for it; `retained_candidate_bytes()` exposes the hold. |
| G3 | `Require` refuses with a typed `Error::Limit`, and no route reaches a partially retained or partially published state | **re-read: `Require` is removed by decision 5**, which makes the over-budget case a fallback. No refusal is introduced. The no-partial half is met: retention is decided at one point, after the archive exists, and is all or nothing; `a_candidate_above_the_budget_falls_back_to_the_rebuild` shows the fallback publishes the same bytes. |
| G4 | no spill: a source ratchet proving no filesystem, `tempfile` or scratch-provider path is reachable | **met** by `the_retention_path_never_spills`. |
| G5 | the archive digest is **recomputed** at application over the retained bytes; a stored digest is used only for equality and the debug assertion | **met, and by construction**: no digest is stored at all. |
| G6 | `validate_candidate`'s capture and revision check still run unconditionally on the reopened retained archive, and the plan does not retain a `Snapshot` | **met**: `capture_internal` runs 8 times per lifecycle in both legs, and `RetainedCandidate` holds bytes and a bound. |
| G7 | every adversarial route produces the identical `Result` with retention on and off, and `output_sha256` is identical on all four selectors | **met**: `retention_changes_no_apply_verdict` over four rejecting routes plus the accepting one, the whole pre-existing adversarial suite green, and one digest per corpus across every leg. |
| G8 | plan equality is not a byte comparison of the retained archive | **met in its strongest form**: equality does not look at the slot, so a released plan equals its unreleased self (`releasing_and_dropping_return_the_retained_bytes`). |
| G8b | the archive decision reaches `build_candidate` as one argument | **met**: one `CandidateArchive` argument; `build_candidate` has twelve parameters, and `clippy::too_many_arguments` (threshold 12) passes. |
| G9 | `CrossSlideCopyPatch` is untouched: its magic, its six revisions and `to_bytes` bit-identical, and `apply_patch` retains nothing | **met**: `the_durable_patch_is_untouched_by_retention` compares the encoded patch with and without retention byte for byte, and both `apply_patch` call sites pass `CandidateArchive::Build`. |
| G10 | `release_retained_candidate(&mut self)` exists and live bytes return to the unretained baseline after it | **met**: the method exists, `Arc::strong_count` falls on release, and the allocator probe's `plan dropped` checkpoint returns to the constant −4. |
| G11 | native cycles for the application phase fall on the media-rich corpus by more than that leg's own A/A floor | **met**: whole-child cycles −24.64% against a +1.71% A/A floor (window 4, the committed tree's binary), −27.01% against +1.46% (window 2) and −27.08% against +0.82% (window 1); the apply phase itself falls 70.55–73.41% across four windows. |
| G12 | peak live bytes during planning do not rise, and the plan-held bytes after planning rise by exactly the archive length | **met**: the `planned` delta is the archive plus a 40-byte `Arc` header, and nothing else. |
| G13 | the archive copy at application is fallible, returning `Error::Allocation { resource: "cross-slide candidate archive", .. }` rather than aborting | **met**: `try_reserve_exact` with the same typed error the `BoundedVecWriter` it replaces returns. |

### Gates run

Listed with their tails in
[`results/change-0656/gates.txt`](results/change-0656/gates.txt):
`cargo fmt --all --check`; `cargo clippy -p litchi-pptx -p litchi-opc
--all-targets`; `cargo doc -p litchi-pptx -p litchi-opc --no-deps`; the six test
suites above; `python3 tools/non_iwork_gate.py verify`. `cargo check --workspace
--all-targets` fails on one `litchi-iwa` example (`ChartData::new` returning the
wrong error type), reproduced on the untouched base checkout and therefore
recorded as pre-existing; it is an iWork crate, which this programme excludes and
this change does not touch.

## Validation preserved

Nothing is removed, weakened, reordered or made conditional. The replan at
application runs every refusal it runs today; both captures still run;
`validate_before`, `validate_after` and `validate_candidate`'s revision check
still run in the same order on the same values; the four staleness proofs are
still taken from the live packages before any capture;
`reject_unknown_non_part_members` still runs on every physical read. The archive
bound is still `max_patch_bytes`, enforced at planning by `BoundedVecWriter` and
re-checked at application before a retained archive is accepted, so a tighter
bound at application is a miss and rebuilds under the tighter writer. The
allocation refusal is preserved: the archive copy reserves with
`try_reserve_exact` and returns the same
`Error::Allocation { resource: "cross-slide candidate archive", .. }` the writer
it replaces returns, rather than aborting on an infallible `Vec::clone`.

The new budget adds no refusal at all. `Limits::new` rejects a zero
`max_retained_candidate_bytes` exactly as it rejects the other five zeros, which
is the only new way to be told no, and it is a constructor rejection rather than
an operation refusal.

## Limitations — what is not claimed

- **No speedup is claimed and none is registered.** `performance_claim: none`,
  no claim-registry entry. The wall-clock figures are reported beside the A/A
  and B/B floors measured in the same window, and the window whose floor
  exceeded 5% is retained and named rather than dropped.
- **The default budget costs memory and the record says so.** A caller that
  plans a media-rich cross-copy and then applies it peaks about 12% higher than
  before. On the plain corpus the peak falls 5.45%. The budget is the control:
  `max_retained_candidate_bytes = 1` restores the old memory profile exactly, at
  the old cost.
- **The default is not swept.** 64 MiB is argued, not optimized: it retains the
  two measured corpora and every fixture in the tree, keeps the default plan's
  worst case inside a server memory profile, and is the value 0646 measured. No
  other value was measured.
- **Two corpora, both synthetic.** Both are generated by
  `litchi-pptx-cross-slide-copy-evidence-v1`, and the media-rich one is
  deliberately incompressible, which maximizes both the deflate the reuse
  removes and the archive the plan holds. The largest real `.pptx` under
  `test-data/` is 972,788 bytes and all 78 are under 1 MB, so **no real deck
  exercises the large end of the budget, and no shape of real deck was
  measured.** A deck whose copied closure compresses well would save less and
  retain less; one with more copied media would do the opposite.
- **The counts and allocator figures are exact; the interpretation of the
  instruction totals is not.** Instructions rank work, not latency. Callgrind
  runs SHA-256 in software because valgrind masks the SHA CPUID bit (change
  0649), so the 19.27 G `Ir` of `compress256` on the media-rich lifecycle is
  roughly five times its native cycle share; and callgrind prices
  `rep movsb`-class copies per byte, which change 0604 measured at a 35×
  overstatement. The `perf stat` rows are the counterweight and they are
  whole-child, so they dilute rather than isolate.
- **One input to the candidate's serialization is not pinned by the staleness
  proofs, and only one fixture covers it.** The candidate is published through
  the destination's preservation source — the archive it was opened from — and
  two destinations can agree on both fingerprints while carrying different
  preservation sources (plan against an edited owned destination, apply to a
  clean reopen of its own bytes). One test builds that pairing and finds the
  verdict and the published bytes identical, with the re-serializing assertion
  running on it; **that it holds for every package is argued, not proved.**
- **The proof this change replaces with an argument is named, not eliminated.**
  A fresh `to_stream` of the in-memory candidate is no longer compared against
  the retained bytes at application. The argument that they must agree rests on
  the four staleness proofs plus determinism of `to_stream`, and it is
  re-derived by a `debug_assert` in debug and test builds only. A release build
  trusts the argument, backed by the recomputed graph and archive revisions.
- **The saving is release-only, by construction.** The `debug_assert` re-runs
  the serialization it removes, so a debug or test build does strictly more work
  than the base. That is the right trade for an argument-backed reuse — it is
  how the 885-test suite becomes the proof — but it means no debug-build
  measurement of this change is meaningful. It also means the release half of
  the substitution proof runs only where the suite is run in release; the debug
  half runs everywhere.
- **The archive is copied once at application.** `OpcPackage::from_vec*` takes
  an owned `Vec<u8>`, so the retained handle is copied into an owned buffer and
  the peak during application carries both. Removing that copy needs a
  shared-bytes ingress (`OpcPackage::from_shared_vec`) in `litchi-opc`, which is
  an ADR 0011 physical-ownership question and is **not** taken here. Every
  measurement above pays the copy.
- **`apply_patch` is unchanged and unmeasured.** Retention is implemented for
  `CrossSlideCopyPlan` only; the durable `CrossSlideCopyPatch` carries no
  archive by construction, so a patch-driven application still builds and
  deflates the candidate exactly as today. The inverse-patch route, which
  additionally clones the destination and replans, is untouched.
- **The planning phase is reported flat and is not proved flat.** One window
  read planning +5.19% on one selector; a 40-sample window says the phase's own
  leg-to-leg floor is about ±4% and the pairs overlap. No planning result is
  claimed in either direction.
- **`publication_ns` is bimodal and unexplained.** It sits at ≈1.5 ms or
  ≈6.4 ms on the media-rich selectors independently of the binary; this record
  adds a within-window witness (one leg of each binary in each mode) and no
  explanation.
- **Not measured at all:** the source-backed cross-copy plan
  (`SourceBackedCrossSlideCopyPlan`; no selector exists), the external-package
  copy binary, process RSS, cold-cache, range-source, concurrency and
  cross-platform behaviour, applying one plan to more than one destination
  (which is the case the retained archive pays for most), and the plan-only
  lifetime of a plan that is never applied, which is the case retention costs
  most and saves nothing.

## Retained evidence

[`results/change-0656/README.md`](results/change-0656/README.md) — both legs'
callgrind annotations, per-callee call counts and run logs, the isolation-pair
count and instruction tables, the allocator retention probe for both corpora,
all three timing windows and both `perf stat` windows with their per-leg phase
medians and floors, both legs' `litchi-pptx` suite output and the four consumer
suites, the measured binaries' sha256, the gates, the decision record, the
cleanup record and the log paragraphs. The eight raw callgrind profiles are not
retained (4.4 MB); the annotations and call counts derived from them are, and
`scripts/` reproduces them.
