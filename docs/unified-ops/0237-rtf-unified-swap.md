# 0237 — Unified RTF paragraph swaps

RTF now exposes `File::swap`, explicit-edit swap and batch swap for ordinary
body paragraphs. Both selectors resolve against the same immutable owner.
The native `Edit::swap_paragraphs(Position, Position)` stages one atomic
lifecycle operation, using the ADR's typed positions. The facade delegates
publication and history to that native operation and admits target, operation,
staging and snapshot limits before adopting the result.

The native projection preserves empty paragraphs, terminal boundaries and
decoded Unicode. Same-position swaps preserve exact bytes, including on
protected or formatted sources; identical plain paragraphs also remain exact
no-ops. Changed swaps retain the existing strict plain-body closure and reject
unsupported rich or protected sources. Capability reporting exposes swap and
its remaining source-closure restriction. Resolved paragraph views supply
staging lengths without a second facade body scan; same-position staging
retains operation descriptors without copying selected text.

Process-local patches restore exact endpoint bytes and retain the existing
identity checks. Changed swap patches and their inverses explicitly refuse
durable serialization with typed `PatchError::InvalidText`: the current wire
vocabulary does not represent this native operation. Extending that vocabulary
with exact replay remains a follow-up.

## Verification

The complete native RTF suite passes 1,485 tests/doctests. Full-feature facade
tests/doctests pass 1,680, and the RTF-only profile passes 545. Four native
swap tests cover lifecycle order, exact inverse, empty/terminal/Unicode
preservation, protected same-position no-ops, durable refusal, admission
failures and edit reuse. Nine facade tests cover every pair in a four-paragraph
body, native-byte parity, empty boundaries, identical and same-position no-ops,
sequential mixed edits, identity-bound history, atomic/best-effort batches,
resource limits, typed native causes and save/reopen.

Warning-denied all-target Clippy passes for native/facade all-features, RTF-only
facade and the standalone harness. Warning-denied rustdoc, no-feature compilation,
the RTF glossary profile (85 banned names and one transaction family), scoped
formatting and diff checks pass. Crate boundaries pass for 66 packages, all 36
Python repository-tooling tests pass, and the OfficeCLI delta is unchanged.

## Paired timing

The standalone [harness](../../tools/perf-baseline/unified-rtf-swap/README.md)
compares swapping the first and last paragraphs through the native and unified
APIs on warmed immutable owners, with selectors and matching limits prepared
outside timing. Three trials pinned to one CPU use two warm-up and nine measured
ABBA rounds (18 samples per route/case). Every sample proves native/facade byte
parity, reopened paragraph order and source immutability outside timing.

| Source paragraphs | Native median µs | Unified median µs | Unified/native |
|---:|---:|---:|---:|
| 1 (no-op) | 0.185–0.185 | 0.590–0.610 | 3.189–3.297× |
| 10 | 6.995–7.080 | 7.405–7.540 | 1.055–1.074× |
| 100 | 45.896–46.430 | 46.865–47.361 | 1.017–1.026× |
| 1,000 | 412.557–415.142 | 413.157–420.477 | 1.001–1.013× |
| 10,000 | 4607.751–4645.926 | 4638.582–4659.062 | 1.003–1.008× |

Changed swaps stay within the 10% small-write comparison reference in these
trials. The one-paragraph no-op exceeds it, with about 0.4 µs of facade overhead;
fixed no-op overhead remains a performance follow-up. These measurements cover
one swap per document, not repeated-command batching or allocation behavior,
and shared-host results support only the stated cases.

Receipts: [trial 1](results/0237-rtf-swap/trial-1.json),
[trial 2](results/0237-rtf-swap/trial-2.json),
[trial 3](results/0237-rtf-swap/trial-3.json). They retain raw samples, source/tree
hashes, compiler, CPU, binary size/SHA-256 and reproduction commands. Measured
source: `4530c5aa8f`. Builds used an external copied root lockfile with cached
`cc 1.6.0`; repository lockfiles remain unchanged. Both owned agent worktrees
and branches and the owned scratch/build artifacts were removed after checks.
No binary is committed.

## Remaining work

RTF rich paragraph properties and element coverage, durable swap replay,
repeated-command coalescing, allocation qualification, native glossary
completion, other delayed formats and Python bindings remain required by the
original goal. This milestone does not complete that goal.
