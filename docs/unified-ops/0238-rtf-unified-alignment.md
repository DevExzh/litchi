# 0238 — Unified RTF paragraph alignment

RTF paragraph nodes expose the typed `align` property for reads, predicates and
writes, with `left`, `center`, `right` and `justify` values. The facade accepts
alignment alone or together with text, resolves selectors against one immutable
owner, and stages both facets in one native edit. Operation limits count each
native property target; facade transaction diagnostics continue to count user
commands. Staging and history retain their finite bounds. Paragraph addition
remains text-only, and the static property table marks alignment unavailable
for `add`.

The native `ParagraphAlignmentUpdate` uses the ADR's `Position` type.
`Edit::set_body_paragraph_alignments` validates ordering, limits and conflicts
before appending any work, resolves before-values in one paragraph traversal,
and leaves failed edits reusable. The scalar setter delegates to it. Existing
alignment conflicts are searched against the ordered updates, avoiding a scan
of every staged text operation for every paragraph in mixed batches.

Pure alignment commits use the native lossless paragraph-property writer,
preserving mixed runs and unrelated paragraph properties. Same-value batches
retain exact source bytes and owner identity. Text-plus-alignment commits retain
the existing stricter native text/character source closure; richer mixed writes
remain a gap. Process-local history restores exact endpoints, and changed
protected sources retain their typed refusal.

Native alignment controls now flush preceding buffered text before changing
alignment. This keeps a pending `\par` boundary in the preceding paragraph's
state instead of assigning the next paragraph's alignment to it. The dense
plain-body preflight/allocation invariant remains covered and passes.

## Verification

The native RTF suite passes 1,490 tests/doctests, the full-feature facade passes
1,688, and the RTF-only facade passes 553. Five native tests cover ordered
batches, admission failures/reuse, mixed text, exact inverse, mixed-run
preservation, same-value batches and direct paragraph-boundary readback. Eight
facade tests cover native-backed alignment predicates, every value and byte
parity, atomic mixed-property fanout, operation/staging/target/history limits,
protected no-ops, invalid values, batch failure attribution, sequential
alignment queries and save/reopen.

Warning-denied all-target Clippy passes for native/facade all-features, RTF-only
facade and the standalone harness. Warning-denied rustdoc, no-feature compilation,
the RTF glossary profile (85 banned names and one transaction family), scoped
formatting and diff checks pass. Crate boundaries pass for 66 packages, all 36
Python repository-tooling tests pass, and the OfficeCLI delta is unchanged.

## Paired timing

The [standalone harness](../../tools/perf-baseline/unified-rtf-alignment/README.md)
compares native alignment batches against unified last-paragraph path updates
and all-paragraph query updates on warmed immutable owners. Native updates and
facade selectors/properties are prepared outside timing; query expansion remains
inside timing. Three trials pinned to one CPU use two warm-up and nine measured
ABBA rounds (18 samples per route/case). Every sample checks exact native/facade
bytes and reopened paragraph text/alignment outside timing.

| Paragraphs | Route | Native median µs | Unified median µs | Unified/native |
|---:|---|---:|---:|---:|
| 1 | Last path | 3.335–3.370 | 3.785–3.825 | 1.135–1.145× |
| 1 | Query all | 3.380–3.425 | 4.685–4.750 | 1.374–1.393× |
| 10 | Last path | 10.890–10.960 | 11.420–11.505 | 1.043–1.050× |
| 10 | Query all | 11.350–11.445 | 13.380–13.450 | 1.170–1.184× |
| 100 | Last path | 112.736–113.490 | 115.235–115.371 | 1.016–1.022× |
| 100 | Query all | 107.225–108.650 | 122.771–126.356 | 1.145–1.170× |
| 1,000 | Last path | 961.074–966.019 | 966.014–971.610 | 1.004–1.006× |
| 1,000 | Query all | 1049.675–1051.415 | 1113.855–1120.475 | 1.059–1.067× |
| 10,000 | Last path | 14860.598–15122.309 | 14900.468–15187.629 | 1.003–1.004× |
| 10,000 | Query all | 16979.607–17231.698 | 17524.665–17724.381 | 1.017–1.041× |

The one-paragraph path case and query fanout through 100 paragraphs exceed the
10% small-write comparison reference. Fixed facade/query overhead remains open.
Larger cases stay within that reference in these trials. Native runtime grows
about 15–16× from 1,000 to 10,000 paragraphs in these samples, so this evidence
does not qualify linear end-to-end scaling. These single-command measurements
do not cover repeated-command coalescing or allocations, and shared-host results
support only the stated cases.

Receipts: [trial 1](results/0238-rtf-alignment/trial-1.json),
[trial 2](results/0238-rtf-alignment/trial-2.json),
[trial 3](results/0238-rtf-alignment/trial-3.json). They retain raw samples,
source/tree hashes, compiler, CPU, binary size/SHA-256 and reproduction commands.
Measured source: `18a0e5c410`. Builds used an external copied root lockfile with
cached `cc 1.6.0`; repository lockfiles remain unchanged. Both owned agent
worktrees/branches and owned scratch/build artifacts were removed after checks.
No binary is committed.

## Remaining work

RTF paragraph layout and character properties, richer mixed text/format writes,
formatted paragraph creation, broader elements, repeated-command coalescing,
allocation/scaling qualification and the earlier no-op/durable swap gaps remain.
Native glossary completion, other delayed formats and Python bindings remain
required by the original goal. This milestone does not complete that goal.
