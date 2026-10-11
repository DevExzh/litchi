# 0187 — XLSB sheet removal

## Contract

The native primary `Edit::remove_sheet(Selector)` resolves the immutable complete
tab catalog. Missing selectors return `None`; an already selected tab returns
`Some(false)`. Distinct selections extend one sorted set of source positions,
so earlier removals never shift the meaning of later selectors. The operation
is a topology barrier and is validated before staging. Durable tag 10 contains
a checked count and source positions; existing patch tags retain their meaning.

Removal preserves surviving tab identifiers, relationship IDs and physical
parts. It deletes the selected sheet root and its exclusively owned binary
index. Other internal outgoing content relationships require an explicit
disposition and currently receive typed refusal. Original orphan parts are
retained; this operation does not implicitly perform garbage collection.

BookView positions and modeled selected flags retain the surviving active
identity. When the active tab is removed, the fallback is the next visible
source tab, or the last visible tab before it. The final catalog must retain
a visible tab. Scoped names on surviving tabs shift; names scoped to removed
tabs are refused until formula name-index remapping is modeled. Live cell/name formulas that would lose all local targets receive typed
refusal until their error-reference rewrite is modeled. Internal XTI
entry indexes and cardinality stay stable: surviving ranges shrink and a
fully removed target becomes the existing deleted-reference sentinel. External
workbook links remain unchanged. Surviving volatile-reference ordinals shift;
references to removed tabs receive typed refusal.

The shared lexical metadata helper reuses the insertion scanner to remove
catalog title spans and update vector/heading counts, preserving survivors and
trailing named-range metadata. Optional absence remains absent. Unsupported
metadata layouts and markup compatibility remain explicit typed gaps.
Retained XML/VML references to removed names and opaque binary dependency
owners are guarded instead of leaving stale references. These guards include
binary pivot-cache/table owners, worksheet FRT/AC records and formula groups,
and ambiguous or removed-target internal hyperlink locations.

The facade supports direct sheet paths and homogeneous sheet queries for both
eager and retained-source owners. All query targets resolve before one native
edit/commit. Mixed sheet/cell queries fail atomically. Source fences, signature
policy and semantic quotas remain enforced. Structural history now uses the
native checked retained-byte charge, including every selected source position
and the durable envelope. The envelope charge keeps a measured history budget
consistent with native durable-patch admission.

## Remaining work

Disposition-aware related-content removal, scoped-name pruning, broader binary
and textual dependency rewriting, and homogeneous command-batch scaling remain
open. Repeated native removal calls revalidate the staged prefix; this record
does not claim linear multi-removal scaling. Tier 2/3 writes, Python/CLI and
full-goal acceptance remain open.

## Verification

The shared OOXML suite passes 484 tests/doctests. The full native XLSB suite
passes 1,225; after the final formula/history-accounting changes, 41 focused
removal and structural-history tests plus 19 native doctests pass. The combined
DOCX/XLSX/PPTX/XLSB/encryption facade suite passes 1,258 tests/doctests, including
the new eager/source sheet removal and 17-target exact-history-budget checks.

Coverage includes immutable source selectors, repeated/missing selection,
malformed durable tag 10, inverse and merge behavior, complete chart catalogs,
exclusive binary-index deletion, orphan retention, disposition refusal, custom
workbook locations, strict relationships, absent optional app/views, active
fallback and selected flags, XTI ranges, scoped/global names, live formula
references, volatile owners, internal hyperlinks and opaque dependency guards.
Review also retained repeated worksheet insertion support and its runnable
metadata example while sharing the title scanner.

Full and minimal all-target Clippy and affected-crate rustdoc pass with warnings
denied. Both XLSB glossary profiles, scoped formatting, the generated OfficeCLI
parity check and the boundary audit pass. The latter covers 66 packages and
268 internal declarations with 11 existing debt items. No base-only commits
were pending. Repository lockfiles remain unchanged; workspace verification
uses an external copied lockfile with cached cc 1.6.0. All five delegated
worktrees and their branches were removed after integration. This record does
not claim Python, CLI, fuzz or full-goal acceptance.

## Paired measurements

The standalone removal harness compares one native edit/commit with one unified
`File::remove` on 2, 10, 100 and 1,000 existing worksheets. It removes
`floor(N/2)` after preparing native positions and facade paths outside timing.
Source promotion remains inside both source-owner legs. Each matrix uses nine
ABBA rounds on CPU 0 with no concurrent build, test or heavy audit.

The [after](results/0187-xlsb-sheet-remove-after.json) and
[repeat](results/0187-xlsb-sheet-remove-repeat.json) receipts have empty failure
matrices and exact native/facade ZIP hashes in every case. Checks also cover
surviving names/markers, active identity, binary indexes and the absence of the
removed worksheet/index members.

| Owner / existing tabs | After native / facade µs | Repeat native / facade µs |
|---|---:|---:|
| eager / 2 | 850.124 / 877.033 | 865.944 / 885.493 |
| source / 2 | 1047.134 / 1054.035 | 1054.235 / 1064.914 |
| eager / 10 | 2065.489 / 2105.599 | 2064.380 / 2110.759 |
| source / 10 | 2492.182 / 2528.251 | 2483.911 / 2511.601 |
| eager / 100 | 13868.012 / 14107.612 | 13553.890 / 13727.781 |
| source / 100 | 16117.291 / 16264.252 | 16192.362 / 16366.733 |
| eager / 1,000 | 136245.831 / 137005.235 | 130672.068 / 132102.574 |
| source / 1,000 | 154434.872 / 155611.667 | 157037.064 / 157750.536 |

Measured facade overhead is 0.5–3.2% across both matrices.
These measurements qualify only the listed single-removal cases. There is no
pre-existing native removal baseline and no insertion/removal command-batch or
linear-scaling claim. Receipts include the dirty-build source manifest and its
encoding, source digest, toolchain, build command and binary SHA-256; compiled
binaries and build targets are kept outside Git and removed after verification.
