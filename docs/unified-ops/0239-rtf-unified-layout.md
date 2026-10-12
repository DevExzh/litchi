# 0239 — Unified RTF paragraph layout

## Scope and API

RTF paragraphs expose `spaceBefore`, `spaceAfter`, `indent`, `rightIndent`,
`firstLineIndent`, `keepNext`, `keepLines` and `pageBreakBefore` through native
borrowed reads, property predicates and unified `set`. Spacing is nonnegative;
indentation is signed. Lengths use the shared typed coercion and points display,
with checked conversion to the native signed 32-bit twip range. Boolean fields
use the shared boolean coercion. Capability declarations advertise get/set;
paragraph creation still accepts text only.

The native `ParagraphLayoutUpdate::{new, position}` and
`Edit::patch_paragraph_layout` now use `Position`; consumers are migrated with
no compatibility alias. The batch method preflights existing layout conflicts
through a sorted index, avoiding a scan of every prior alignment operation for
each incoming paragraph. First-conflict diagnostics remain in staging order.

The adapter groups all eight layout fields into one native delta per selected
paragraph. Alignment and layout share one native commit. Target, operation,
staging and history limits are admitted; one layout delta counts as one native
operation regardless of field count, and alignment adds a second operation.
Unified edit diagnostics continue to count caller commands. Immutable owners,
exact no-op identity, inverse history and origin-aware atomic save are retained.
Text plus layout remains a native `ParagraphLayoutTextConflict` refusal; the
adapter does not split that request into separate publications.

Native paragraph controls now flush preceding buffered text before changing
layout or resetting paragraph properties, preserving the previous paragraph's
format at a pending paragraph boundary. The dense plain-body allocation guard
still passes. A regression confirms the existing native compatibility guard
rejects layout after removal, insertion, move or swap without damaging the
already staged structural operation. No new lifecycle restriction is added.

## Verification

The native suite passes 1,492 tests/doctests; the full-feature facade passes
1,697 and the RTF-only facade passes 562. Eight facade tests cover direct native
reads and predicates, exact native publication bytes, alignment/layout fanout,
mixed-run preservation, clearing and exact no-ops, invalid lengths/flags,
native mixed-text refusal, resource limits, sequential predicates, atomic and
best-effort failure attribution, protection and origin-aware atomic save.
Native regressions cover paragraph-boundary layout readback and lifecycle
conflict atomicity; a value-layer test covers signed twip rounding and overflow.

Warning-denied all-target Clippy passes for native/facade all-features, the
RTF-only facade and standalone harness. Warning-denied rustdoc, no-feature
compilation, the RTF glossary profile (85 banned names and one transaction
family), scoped formatting and diff checks pass. Crate boundaries pass for
66 packages, all 36 Python repository-tooling tests pass, and the OfficeCLI
delta is unchanged. Repository lockfiles remain unchanged.

## Paired timing

The [standalone harness](../../tools/perf-baseline/unified-rtf-layout/README.md)
compares native layout batches with unified last-paragraph path updates and
all-paragraph query updates on warmed immutable owners. Each update assigns all
eight fields. Native updates and facade selectors/properties are prepared
outside timing; query expansion remains inside timing. Three trials pinned to
CPU 0 use two warm-up and nine measured ABBA rounds (18 samples per route/case).
Every sample checks exact bytes and reopened text/layout outside timing.

| Paragraphs | Route | Native median µs | Unified median µs | Unified/native |
|---:|---|---:|---:|---:|
| 1 | Last path | 4.165–4.215 | 5.135–5.160 | 1.218–1.236× |
| 1 | Query all | 4.065–4.165 | 5.900–5.945 | 1.425–1.451× |
| 10 | Last path | 11.895–11.980 | 12.945–12.965 | 1.082–1.089× |
| 10 | Query all | 17.610–17.710 | 20.195–20.405 | 1.144–1.152× |
| 100 | Last path | 116.736–118.865 | 119.050–121.866 | 1.020–1.025× |
| 100 | Query all | 180.066–185.881 | 191.481–194.856 | 1.041–1.063× |
| 1,000 | Last path | 999.205–1018.125 | 998.365–1018.405 | 0.998–1.003× |
| 1,000 | Query all | 1886.795–1946.236 | 1978.031–2017.966 | 1.037–1.048× |
| 10,000 | Last path | 16556.684–17331.426 | 16556.327–16938.809 | 0.977–1.004× |
| 10,000 | Query all | 23555.437–23797.641 | 24010.627–24087.632 | 1.012–1.022× |

The one-paragraph path case and query fanout through ten paragraphs exceed the
10% small-write comparison reference. Fixed facade/query overhead remains open.
Other measured cases stay within that reference. Native runtime grows
12.2–17.0× from 1,000 to 10,000 paragraphs across these trials, so this evidence
does not qualify linear end-to-end scaling. These single-command measurements
do not cover repeated-command coalescing or allocations. Shared-host results
support only the stated cases.

Receipts: [trial 1](results/0239-rtf-layout/trial-1.json),
[trial 2](results/0239-rtf-layout/trial-2.json),
[trial 3](results/0239-rtf-layout/trial-3.json). They retain raw samples,
source/tree hashes, compiler, CPU, binary size/SHA-256 and reproduction commands.
Measured source: `884d7f992c`. Builds used an external copied root lockfile with
cached `cc 1.6.0`; repository lockfiles remain unchanged. Both owned agent
worktrees/branches and owned scratch/build artifacts were removed after checks.
No binary is committed.

## Remaining work

Character properties, richer mixed text/layout writes, formatted paragraph
creation, broader elements, repeated-command coalescing, allocation/scaling
qualification and earlier no-op/durable swap gaps remain. Native glossary
completion, other delayed formats and Python bindings remain required by the
original goal. This milestone does not complete that goal.
