# 0244 — Native-backed RTF run removal

## Contract

Unified `remove` accepts ordinary RTF run paths and scoped queries. One bounded
selector walk collects typed paragraph/run positions, rejects mixed paragraph/run
selections, and normalizes the native source-ordered batch. The adapter delegates
to the existing native atomic run-text transaction with empty replacements. It
retains no shadow document model or second source-rewriting path.

Removing a run preserves its containing paragraph. Removing every run keeps all
paragraphs, including leading/interior empties and terminal empty paragraphs.
The native encoder retains paragraph delimiters, metadata and unselected run
formatting; source-bound patches and inverse history retain exact endpoints.
Native run fragments are the positional targets, including fragments originating
from Unicode controls. Removal does not infer whole words or text spans.

`capabilities()` advertises run removal and its native source-closure restrictions.
Compressed transport, dependent destinations, inline breaks and paragraph
style/list references remain explicit gaps. Opaque body syntax before a removed
run can retain its bytes and anchor; removal that shifts an opaque anchor remains
a typed native refusal until run-text anchor projection is integrated. Paragraph removal keeps its prior
whole-paragraph semantics and refusal classification. Run insertion, move and
swap remain follow-ups.

## Admission and accounting

The shared mutation resolver collects source text lengths while visiting targets,
so removal and text replacement can bound retained before-text payloads before
native staging allocations. Previously run text writes charged only the update
DTO and replacement text. This closes the case where a large source run replaced
by a short value exceeded the intended staging budget while passing admission.
Format-only updates do not clone before-text and therefore retain their existing
charge. Operation, expanded-target and native patch snapshot limits still apply.
Atomic/best-effort batches preserve their sequential addressing semantics; one
query fan-out remains one native publication.

## Verification

Native RTF: **1,513 tests/doctests** pass. Full facade: **1,735** pass;
RTF-only facade: **600** pass. These broad runs include eight new removal tests.
The subsequently added opaque-anchor regression also passes in the final focused
run of all **nine** removal tests. Coverage includes scalar/query parity with
native empty replacement, terminal and interior empty paragraphs, Unicode source
fragments, staging admission, exact inverse/replay, batch rollback, protection,
compressed-source refusals and origin-preserving save without residue.

Native/full and RTF-only all-target Clippy, standalone benchmark Clippy,
warning-denied rustdoc, no-feature checking and both RTF glossary gates pass.
The frozen performance harness passes all **nine** RTF tests, including stable
large-fixture hashes. Tooling passes **1,166** unit tests; crate boundaries pass
for **66 packages / 268 dependency declarations / 11 existing debt items**.
The OfficeCLI capability delta check passes. Parser smoke checks pass for
792 query inputs / 748 plans, 909 paths / 435 accepted, and 776 values /
627 successful coercions. Python bindings are not qualified
by this milestone.

## Paired measurements

Three CPU-0-pinned release trials cover 20 cases each: 1–10,000 groups of three
runs, laid out in separate paragraphs or one dense paragraph, selecting either
the last middle run or every italic run. Two warm-up rounds and nine ABBA rounds
produce 18 samples per route. Owners, selectors and native update DTOs are
prepared outside timing. Staging and commit are timed; serialization, exact byte
comparison, reopened text/format/count checks and source immutability checks are
outside timing. Every check passes in every trial.

Ranges below span the three trial medians; durations are microseconds.

| Shape / groups / route | Native µs | Unified µs | Unified/native |
|---|---:|---:|---:|
| Paragraphs / 1 / last | 4.83–4.95 | 5.48–5.54 | 1.12–1.14× |
| Paragraphs / 1 / all | 4.90–4.95 | 7.29–7.32 | 1.47–1.49× |
| Paragraphs / 1,000 / last | 2,023–2,079 | 2,046–2,090 | 0.984–1.024× |
| Paragraphs / 1,000 / all | 2,046–2,091 | 2,492–2,516 | 1.20–1.23× |
| Paragraphs / 10,000 / last | 39,384–39,932 | 39,562–40,096 | 0.991–1.017× |
| Paragraphs / 10,000 / all | 39,276–39,654 | 45,592–46,084 | 1.16–1.17× |
| Dense / 1,000 / all | 1,732–1,824 | 2,095–2,214 | 1.15–1.27× |
| Dense / 10,000 / last | 26,365–27,077 | 26,842–27,292 | 0.991–1.035× |
| Dense / 10,000 / all | 25,697–26,617 | 30,028–30,759 | 1.13–1.20× |

Large direct-path removals remain within the 10% reference. Query fan-out exceeds
it; tiny documents retain fixed adapter overhead. From 1,000 to 10,000 groups,
unified all-target time grows 18.30–18.42× for separate paragraphs and
13.89–14.34× for dense runs. End-to-end linear scaling is not established.
Generic query resolution, native parse/source-map costs, allocations and repeated
command coalescing still need profiling and qualification. Full performance
acceptance remains open.

Receipts: [trial 1](results/0244-rtf-run-remove/trial-1.json),
[trial 2](results/0244-rtf-run-remove/trial-2.json),
[trial 3](results/0244-rtf-run-remove/trial-3.json). Each contains raw samples,
source commit/tree, compiler/CPU, binary size/SHA-256 and regeneration commands.
Measured implementation: `a6cdf98625`. No binary is committed.

Builds use external copied lockfiles with cached `cc 1.6.0`; repository lockfiles
are unchanged. This turn created no agent worktree; its temporary build artifacts
were removed. The base checkout remains at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; this branch is rebased onto it.

## Remaining work

Run insertion, move/swap, font/color resources, structural-break text, broader
mixed-operation composition, repeated-command coalescing, durable new-op replay,
allocation/scaling qualification, broader RTF elements, remaining formats and
glossary work and Python bindings remain required by the full goal.
