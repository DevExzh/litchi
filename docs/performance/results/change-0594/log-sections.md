# Log sections for change 0594

Four paragraphs for the coordinator to merge, one per log, in the style of each
log's newest section.

## For `HOTSPOTS.md`

## 0594 — one Deflate decoder per OOXML open, above a measured threshold

Item ZIP-1 of the [0587 queue](0587-remaining-opportunity-survey.md) is
implemented for the structural admission pass and the serial batch wave, gated by
a stated rule: keep the shared session only when it will avoid at least sixteen
decoder constructions. Above the gate an open's allocated bytes fall 58.95% on a
132-member XLSX and 73.02% on a 48-member PPTX, `memset` falls from 17.97% to
2.92% of open-plus-one-part instructions, the total falls 16.35%, and per-open
cycles fall 2.17% on a 445-member PPTX; below it every read takes the
byte-identical pre-change path. The gate exists because the first implementation
regressed `docx_file_source_open` and the regression was attributed rather than
argued: callgrind showed −7.38% Ir per open and `perf stat` −0.81% instructions
with flat cycles, while the selector's own child showed **+8 minor page faults
per fresh-process open at every quantile** — the retained 80,320-byte workspace
raising the open's heap high-water — which at ~430 ns per fault is the entire
1.08% p50 delta. A first gate that scanned member names cost more than it saved
(+4.83% cycles on that corpus) and was replaced by a hook that reuses the count
`source_catalog` already computes. Reach is honest and small: 15 of 179
repository OOXML fixtures carry 16 or more relationship members. The remaining
ZIP-1 half — a pooled session for a *single* cold part read — is still not
implementable without a self-referential borrow; ZIP-4 is the next item on this
path. [Change and limitations](0594-zip-session-reuse-per-open.md);
[evidence](results/change-0594/README.md).

## For `GOAL_AUDIT.md`

## 0594 — one Deflate decoder per OOXML open, above a measured threshold

`docs/GOAL.md` puts unnecessary allocation ahead of layout, algorithms and
parallelism, and requires every claim to be scoped to scenario, corpus, machine,
build and metric. This change removes allocation without touching I/O, limits,
error identity or output bytes — positional requests, requested bytes,
`version()` calls, `syscr` and `rchar` are identical on every measured phase —
and it is scoped by a threshold rather than by assertion, because the measurement
found a regime where the trade loses. The bounded-resource cost is measured, not
argued: one retained 80,320-byte workspace, worth +15 minor page faults per
fresh-process open of a 132-member workbook and −2 on a 445-member PPTX, against
3.29 MB of allocation and zero-fill removed. The audit's standing instruction to
price pointer-chase work in cycles is what caught the first implementation:
instructions fell and cycles did not. `performance_claim: none` and no
claim-registry entry. The P1 row "finish source-backed CRUD adoption across
formats" is unchanged. [Change and limitations](0594-zip-session-reuse-per-open.md);
[evidence](results/change-0594/README.md).

## For `REPORT.md`

## 0594 — one Deflate decoder per OOXML open, above a measured threshold

`IndexedArchive::read` built a fresh `IndexedReadSession` and a fresh
`DeflateDecoder` for every member, so the OPC structural admission pass paid
80,320 bytes per `.rels`. One session now spans that pass through a crate-private
`SessionedArchive` implementing the existing private `ArchiveAccess` trait, and a
new default-no-op trait hook hands it the relationship-member count
`source_catalog` already computes so the decision costs one comparison. Above
sixteen relationship members the open's allocated bytes fall 58.95% and 73.02% on
the two above-threshold fixtures and per-open cycles fall 2.17% on the 445-member
PPTX corpus; the in-process open of the 132-member workbook improves 4.88–5.34%
at p50 against a 0.14% A/A floor. Below the threshold the path is byte-identical
and 164 of the 179 corpus fixtures take it. The first implementation had no
threshold and was held: the selector's own measured child showed +8 minor page
faults per fresh-process open and −1.08% at p50, which the same instrument now
reports as +0.9 faults and +2.45%. Correctness is a member-by-member differential
over 179 packages and 4,215 members that asserts both sides of the threshold are
exercised, a dedicated threshold test and a by-name refusal-recovery test; gates
are clean on both crates with 1,254 tests passing. No speedup, RSS or cold-cache
claim follows; `performance_claim: none`.
[Change and limitations](0594-zip-session-reuse-per-open.md);
[evidence](results/change-0594/README.md).

## For `ADR_COMPLIANCE.md`

## Change 0594 compliance update

Change 0594 threads one `IndexedReadSession` through the OPC structural admission
pass and the serial batch wave, above a stated member threshold. ADR 0005 is
satisfied on its own terms: a session is not a cache — it retains no payload, no
metadata and no verdict, only decoder workspace — and read-order independence is
preserved because every catalog verdict remains a function of member bytes proven
identical member by member across the whole OOXML corpus on both sides of the
threshold, while the threshold itself is a function of the archive's own member
names computed before the first read. No `ReadLimits` or `ArchiveLimits` value,
check or ordering moved; the new hook is called immediately after the
`RelationshipParts` check whose count it reuses, so a package that exceeds that
limit is refused exactly where it was. The single-flight `Loader`/`Waiter`/
`Bypass` arms, the `source.ensure_current()` brackets and the budget reservations
are the ones the sessionless path already used, and a managed package still
refuses to retain decoder workspace across a load (change 0402). ADR 0006 is
untouched: no output byte, no typed refusal and no preservation path changed, and
the validation open keeps its exact `ValidationCatalogPhase` provenance while
remaining non-mutating. No new `unsafe`, no weakened defence, no hidden global
pool, no ambient I/O, and no archive, lock or executor type reaches a public
signature — `SessionedArchive` is `pub(crate)` and the new `SourceBackedPackage`
methods are module-private. See [Change 0594](0594-zip-session-reuse-per-open.md);
`performance_claim: none`.
