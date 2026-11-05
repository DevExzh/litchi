# 0236 — Unified RTF paragraph insertion

This milestone connects native positioned paragraph insertion to unified `add`,
explicit edits and batches. The native entry uses the ADR's `Position` type and
accepts positions from zero through the current paragraph count, including an
empty document. The facade admits `p`/`paragraph`, an optional text property,
and the shared placement vocabulary, then returns the created canonical path.
An omitted text property creates a real empty paragraph.

The native owner controls source closure, output validation and reversible
history. The facade checks parent/kind/property admission, bounds retained
payload/action bytes and history, and preserves origin and source policies.
Insertion does not build a facade document model. Capabilities expose paragraph
creation and its native source-closure restriction; broader RTF nodes remain
explicit gaps.

## Verification

The complete native RTF suite passes 1,481 tests/doctests; the full-feature
facade passes 1,671 and the RTF-only facade passes 536. Ten new facade tests
cover all placements and canonical created paths, empty documents and paragraphs,
Unicode and control delimiters, encoded-source empty insertion, sequential
mixed edits, exact history, atomic/best-effort batches, typed limits/protection,
and save/origin. Native tests cover payload admission and edit reuse, paragraph
boundaries, durable admission and the shared writer's UTF-16 transport.

Qualification fixed three native details. Literal insertion adds a control-word
delimiter only when needed, preserving numeric and space-prefixed text. Generic
insertion treats an empty added paragraph as structural change even when body
text alone cannot distinguish the result. Unicode encoding now lives in the
shared writer and emits signed UTF-16 units, including surrogate pairs, without
an extra scan in existing edit paths. Its encoded-size bound accounts for both
units. Split/merge retain their existing stricter delimiter admission.

Durable paragraph insertion/removal now dry-runs both directions and refuses
`PatchError::InvalidText` if either cannot reproduce the exact endpoint bytes.
The check also covers serialization of the inverse patch. Some inserted control
delimiters and terminal boundaries remain source-bound; process-local inverse
is exact for all admitted insertions. A boundary-aware durable vocabulary is
still a follow-up.

All-target warning-denied Clippy passes for native/facade all-features, RTF-only
facade and the standalone harness. Warning-denied rustdoc, no-feature compilation,
scoped formatting and diff checks pass. Crate boundaries pass for 66 packages,
36 Python repository-tooling tests pass, and the OfficeCLI delta is unchanged.

## Paired timing

The standalone [harness](../../tools/perf-baseline/unified-rtf-add/README.md)
compares native positioned insertion with unified paragraph addition on warmed
immutable owners. Body paths and new-node arguments are prepared outside timing.
Three trials on one pinned CPU use two warm-up and nine measured ABBA rounds
(18 samples per route/case). Every sample checks exact native/facade bytes,
reopened paragraph order and source immutability outside timing.

| Source paragraphs | Native median µs | Unified median µs | Unified/native |
|---:|---:|---:|---:|
| 0 | 2.530–2.610 | 2.845–2.900 | 1.111–1.125× |
| 1 | 3.145–3.165 | 3.515–3.525 | 1.114–1.118× |
| 10 | 8.285–8.375 | 8.590–8.705 | 1.037–1.048× |
| 100 | 53.355–53.806 | 54.141–54.580 | 1.008–1.023× |
| 1,000 | 475.417–480.242 | 475.688–480.572 | 0.999–1.009× |
| 10,000 | 5076.053–5142.243 | 5084.773–5120.838 | 0.996–1.002× |

The 0/1-paragraph cases remain about 11–13% above native, exceeding the 10%
small-write comparison reference; fixed facade overhead remains a follow-up.
The larger cases stay within that reference in these trials. This is one
insertion into each document, not a repeated-command batch or an allocation
qualification. Shared-host measurements support only the stated cases.

Receipts: [trial 1](results/0236-rtf-add/trial-1.json),
[trial 2](results/0236-rtf-add/trial-2.json),
[trial 3](results/0236-rtf-add/trial-3.json). They include source/tree hashes,
compiler, CPU, binary size/SHA-256, reproduction commands and raw samples.
Measured source: `6da36a3557`. Builds used an external copy of the root lockfile
with cached `cc 1.6.0` and an external target directory; repository lockfiles are
unchanged. Both owned worktrees/branches and owned scratch/build output were
removed after validation. No binary is committed.

## Remaining work

RTF swap, broader paragraph properties and element coverage, repeated-command
coalescing, allocation qualification, native glossary completion, other delayed
formats and Python bindings remain required by the original goal. This record
claims no new Python binding surface and does not close the full goal.
