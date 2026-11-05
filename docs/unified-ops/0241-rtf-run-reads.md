# 0241 — Native-backed RTF run reads

## Coverage

Ordinary body paragraphs now expose their native text runs at
`/body/p[N]/r[N]`, with `paragraph` and `run` aliases. Direct lookup, `last()`,
children, generic views and streamed queries use the same borrowed native
`Paragraph::runs()` axis. Run and parent paragraph positions use `Position`.
`Node::text` replaces the paragraph-specific accessor without a compatibility
alias; the scoped `ops-rtf` glossary profile enforces its removal.
Reusable `Plan::compile` now dispatches RTF through the same native-backed
query adapter as ordinary queries.
Run text remains borrowed; the adapter stores no parallel document model.
Child iterators retain one active native axis without boxing; run nodes keep
their parent position without duplicating the native paragraph view.

The native run axis omits explicit inline line-break tokens and paragraph
boundaries. Empty paragraphs therefore have no text-run children. Paragraph
text retains its existing line breaks and empty-paragraph semantics; document
text, annotated, outline and statistics views remain paragraph-based.

Run properties are local native character values. Exact native underline
styles remain distinct in the typed token vocabulary. Properties are readable
and queryable; run mutations are explicit element-level gaps. Paragraph mutation
handlers reject run targets even though a run retains its parent position.

Native character setters currently reject overlapping work on the same span
and use a uniform source-formatting closure. Unified compound run writes need a
native atomic edit path that preserves mixed runs; exposing read support does
not claim that write contract is complete.

Inline native formatting controls now flush preceding text before changing
its character state. This fixes foreground-color and other local character
values being retroactively applied to earlier text. The regression tests cover
all exposed character controls, all underline variants and nested group
restoration. Dense plain paragraphs retain the existing block strategy.

## Verification

The native suite passes 1,497 tests/doctests, the full-feature facade passes
1,705 and the RTF-only facade passes 570. Eight new facade regressions cover
borrowed paths/positions, native break semantics, local character properties,
all 19 underline styles, reusable plans across immutable owners, mutation
refusals, compressed transports and retained corpus reads. Two native tests
cover lexical character transitions and nested group restoration.

Warning-denied all-target Clippy passes for the native/full facade, RTF-only
facade and standalone run benchmark. Warning-denied rustdoc, no-feature
compilation, the scoped glossary JSON audit and scoped formatting/diff checks
pass. Repository tooling passes 1,166 tests; boundaries pass for 66 packages
and 268 internal dependencies with 11 existing debt items. The OfficeCLI delta
is unchanged. All nine RTF-filtered performance-harness tests pass, including
the frozen large-output hashes and native/unified read oracles.

## Paired read measurements

Three CPU-0-pinned trials use 1, 10, 100, 1,000 and 10,000 paragraphs,
three native runs per paragraph, and last-path/all-bold/first-bold routes.
Owners, parsed paths and compiled query plans are prepared outside timing.
Two warm-up and nine measured ABBA rounds yield 18 samples per route.
Repeated reads reduce timer noise; raw batch durations and repetitions are
retained, with medians normalized per read. Every timed checksum agrees with
native, and untimed full text/last-text/source-byte checks pass.

Ranges below are across the three trial medians:

| Paragraphs / route | Native | Unified | Unified/native |
|---|---:|---:|---:|
| 1 / last path | 42.2–42.8 ns | 303.7–312.8 ns | 7.10–7.41× |
| 1 / all bold | 55.9–56.1 ns | 526.5–611.6 ns | 9.39–10.91× |
| 1,000 / last path | 2.42–2.43 µs | 2.70–2.71 µs | 1.11–1.12× |
| 1,000 / all bold | 70.10–71.79 µs | 403.78–469.60 µs | 5.76–6.54× |
| 10,000 / last path | 36.79–37.27 µs | 37.30–37.51 µs | 1.001–1.019× |
| 10,000 / all bold | 820.41–849.40 µs | 4,184.27–4,780.03 µs | 5.03–5.63× |

First-match unified reads stay within 419–523 ns across every size/trial;
there is no document-sized traversal before yielding the first run. Full
unified traversal grows 10.2–10.4× from 1,000 to 10,000 paragraphs, while last
lookup follows the scanned native paragraph axis. The substantial fixed and
per-run generic-query overhead remains a qualification gap. This is evidence
for borrowed streaming behavior, not acceptance of the full performance goal.
Allocation counts, other document shapes and before/after open costs are not
qualified here.

Receipts: [trial 1](results/0241-rtf-runs/trial-1.json),
[trial 2](results/0241-rtf-runs/trial-2.json),
[trial 3](results/0241-rtf-runs/trial-3.json). Each records source commit/tree,
binary name/size/SHA-256, compiler, CPU, timestamp and regeneration commands.
Measured implementation: `13dd098e67`. No binary is committed.

Builds use an external copied lockfile with cached `cc 1.6.0`; repository
lockfiles are unchanged. The owned agent worktree/branch and scratch/build
artifacts were removed after verification. The base checkout remains at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; this branch is rebased onto it.

## Remaining work

Atomic run text/character writes and compound formatting, richer mixed
text/layout writes, formatted paragraph creation, broader elements, repeated
command coalescing, allocation/scaling qualification and earlier durable swap
gaps remain. Native glossary completion, other delayed formats and Python
bindings remain required by the full goal.
