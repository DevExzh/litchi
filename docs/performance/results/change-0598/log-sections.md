# Log paragraphs for change 0598

The coordinator merges these into the four logs; this batch did not edit them.
Each is written in the style of the newest section of its target file.

## For `docs/performance/HOTSPOTS.md`

## 0598 — PPTX cross-package copy revision reuse

Item PPTX-2 of the 0587 queue is implemented. A cross-package slide copy
serialized a complete package twelve times and hashed a complete package graph
twelve times per plan-plus-apply lifecycle; it now serializes seven times and
hashes eight. Three reuses, each of a value the same call already holds for the
same bytes: an immutable `Snapshot` memoizes its serialized-archive revision
together with the archive bound it was taken under, so the replan inside
`apply_plan` stops re-hashing the source and destination archives;
`BoundedVecWriter` digests the candidate archive while `build_candidate` writes
it, so the second serialization that used to hash it is gone (the bounded `Vec`
stays — it is the archive `from_vec_reusing_payloads` reopens); and application
hands the semantic revisions it already proved to `capture_with_revision`
(change 0590) and returns the candidate's proven archive revision instead of
recomputing it for the final comparison. **Measured** on the isolation pair
(`--samples 3` minus `--samples 1`, callgrind, CPU 21): per lifecycle of
`pptx_cross_copy_media_rich` 35,610,254,229 → 26,847,795,508 `Ir`, **−24.61%**,
with `physical_package_fingerprint` −66.65%, `package_fingerprint` −33.33% and
software SHA-256 −31.25%; `pptx_cross_copy_plain` −9.37%; whole child −22.08%
and −2.37%. Natively, `perf stat` whole child is −8.39% cycles against a −0.36%
A/A floor. The revision values, the durable `LPCP0002` encoding, every refusal
and every published byte are unchanged — `zlib_rs::deflate::deflate` differs by
10,520 instructions in 5.96 G per lifecycle and the published digest is one
value per corpus across all four timing legs; 867 `litchi-pptx` tests pass,
seven of them new. Wall clock: `pptx_cross_copy_media_rich` −11.85% at p50
against a −4.99% A/A floor, its apply phase −16.48%; the two plain selectors are
inside their own floors and are reported, not relied on. What remains: the whole
candidate is still built, deflated and captured twice per lifecycle, which is
5.96 G `Ir` of untouched deflate, and removing the second build means retaining
the planned archive in the plan — an ADR 0005 memory-profile change needing a
frozen design record. `performance_claim: none`.
[Change and limitations](0598-pptx-cross-copy-revision-cache.md);
[evidence](results/change-0598/README.md). OLE2/OOXML remain active; ODF is
deferred until completion and iWork excluded.

## For `docs/performance/GOAL_AUDIT.md`

## 0598 — PPTX cross-package copy revision reuse

`docs/GOAL.md`'s first two optimization steps — eliminate unnecessary work, then
unnecessary I/O, serialization and hashing — applied to the cross-package slide
copy that 0587 ranked fifth and measured at 85% of a media-rich whole child.
Five of nine complete-archive serializations and four of twelve complete-graph
hashes per lifecycle are removed by reusing values the same call already proved
on the same bytes. ADR 0005's "cache behavior is semantically invisible" is the
licence for the one new piece of state: a private `OnceLock` on the immutable
`Snapshot`, keyed on the archive bound so no `Error::Limit` moves, never
inherited across `Snapshot::rebound_to` (where `packages_equal` proves graph
equality and says nothing about the retained archive), and always a plain
recomputation on a miss. ADR 0005's bounded-memory rule is also why
`bounded_package_bytes` keeps its `Vec` rather than becoming a hashing sink: the
`Vec` is the candidate archive that `from_vec_reusing_payloads` reopens, so the
saving is taken by digesting it in place. ADR 0003 and 0006 are untouched
because no revision value or proof format changes and the durable `LPCP0002`
encoding stays bit-identical; the four staleness proofs still run on the **live**
packages before any capture, so a stale or foreign source or destination is
refused exactly where it was. **Measured** −24.61% `Ir` per media-rich lifecycle
with exact before/after call counts (twelve semantic hashes and nine archive
serializations before; eight and four after), and −8.39% native whole-child
cycles against a −0.36% A/A floor. What stays open: the candidate is still built,
deflated and captured twice per lifecycle (5.96 G `Ir` of deflate on the copied
closure), and retaining the planned archive to remove the second build changes
`CrossSlideCopyPlan`'s memory profile under ADR 0005; `validate_candidate`'s
second capture is the proof that binds the candidate to the plan, so reusing it
is a contract question, not a value-identical reuse; the four remaining semantic
hashes need 0587's PPTX-1(c), which redefines the durable revision. No speedup,
RSS, allocation, cold-cache or real-producer claim follows, and the wall-clock
floors in this window (A/A −4.99% at p50 on the media-rich selector) are stated
beside every timing.

## For `docs/performance/REPORT.md`

## Change 0598: PPTX cross-package copy revision reuse

Change 0598 removes five of the nine complete-archive serializations and four of
the twelve complete-graph hashes in a `litchi-pptx` cross-package slide copy.
`Snapshot` (`opened/model.rs`) gains one private field, an
`Arc<OnceLock<(usize, [u8; 32])>>` memoizing the serialized-archive revision of
its immutable package together with the archive bound it was taken under;
`Snapshot::rebound_to` starts an empty one. In `opened/cross_copy_plan.rs`,
`snapshot_physical_revision` reads that memo — still running
`reject_unknown_non_part_members` on every call, so no refusal moves —
`BoundedVecWriter` digests the candidate archive as `build_candidate` writes it
and `seal_physical_revision` binds that digest and its length exactly as the
streaming sink does, `candidate_physical_revision` consumes the result at the
position the recomputation occupied, and `apply_plan` and `apply_patch` hand the
semantic revisions they proved to `capture_with_revision` (change 0590) and seed
the resulting snapshots with the archive revisions they proved, so the replan
inside the same call reuses both. `published_archive_revision` returns the
already-proven revision for a candidate that reached publication unchanged and
recomputes for one `validate_application_candidate` rebuilt. Validation passed
`litchi-pptx` library `563/563` plus integration and doc tests, 867 in total,
with seven new tests covering the cache's freshness, its archive-bound key and
its `Error::Limit`, the empty cache after a rebind onto a graph-equal package
with a different retained archive, the equality of planned and published
revisions with fresh serializations, stale and foreign refusals with both caches
warm, the no-known-value fallback, and the bounded writer's digest. **Measured**,
callgrind isolation pair on CPU 21: −24.61% `Ir` per `pptx_cross_copy_media_rich`
lifecycle and −9.37% per `pptx_cross_copy_plain` lifecycle, with twelve semantic
hashes and nine archive serializations becoming eight and four; natively −8.39%
whole-child cycles against a −0.36% A/A floor. The paired wall-clock timings are
reported beside this host's A/A and B/B floors and are not a speedup claim: two
of the four selectors sit inside their own floors. See
[Change 0598](0598-pptx-cross-copy-revision-cache.md); `performance_claim: none`.

## For `docs/performance/ADR_COMPLIANCE.md`

## Change 0598 compliance update

Change 0598 keeps ADR 0003's and ADR 0006's revision binding exactly. The
semantic revision is still `package_fingerprint` over the complete OPC graph and
the physical revision is still SHA-256 over the `litchi-pptx-cross-physical-v2`
domain string, the `u64` little-endian archive length and the archive digest —
now produced by a single `seal_physical_revision`, so the streaming sink and the
archive hashed while it was built cannot drift. No revision value changes, so
the durable `CrossSlideCopyPatch` (`LPCP0002`) encoding with its six 32-byte
revisions remains bit-identical and patches serialized before the change still
apply after it; no published byte changes, confirmed by an identical
`output_sha256` per corpus across all four measured legs and by deflate work
stable to 10,520 instructions in 5.96 G per lifecycle. ADR 0005's statement that
"cache behavior is semantically invisible" covers the one new piece of state: the
memo is private, lives only on an immutable `Snapshot` whose `Arc<OpcPackage>` is
never mutated, is keyed on the archive bound so that a smaller bound still raises
its typed `Error::Limit` rather than reading a memo, is never inherited across
`Snapshot::rebound_to`, and degrades to an ordinary recomputation on a miss.
ADR 0005's bounded-memory rule is preserved in the candidate path: the bounded
`Vec` and its `max_patch_bytes` ceiling are unchanged, and the saving comes from
digesting those bytes in place rather than from removing the bound. ADR 0005's
mandatory validation is intact: `reject_unknown_non_part_members` still runs on
every cached read and on the candidate, in the same position; `validate_before`,
`validate_after` and `validate_candidate`'s own revision check are untouched; and
the four staleness proofs in `apply_plan` and `apply_patch` are still computed
from the live `&OpcPackage` arguments before any capture, so the caches never
cover a package a caller can still mutate. ADR 0013's notes-topology check still
runs inside every capture, because `capture_with_revision` skips only the hash.
Change 0454's "dedup only with proven equivalence" is untouched: the source
closure is still copied byte-for-byte, no destination member is reused on byte
equality, and the layout inheritance graph is still proven before a layout is
reused. No `unsafe`, no ambient I/O, no global pool, no weakened limit, no public
API change — `Snapshot`'s new field is private. See
[Change 0598](0598-pptx-cross-copy-revision-cache.md); `performance_claim: none`.
