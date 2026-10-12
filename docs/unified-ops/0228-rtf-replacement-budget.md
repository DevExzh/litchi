# 0228 — RTF semantic replacement admission

Primary `SemanticLimits` now independently bounds aggregate replacement bytes.
Its checked consuming builder accepts zero through `isize::MAX`, with a finite
512 MiB default. Invalid profiles return `InvalidLimit`; runtime refusals return
`Limit { kind: ReplacementBytes, observed, maximum }`. The retained source/output
ceiling remains independently enforced.

Text entry points accept borrowed or owned `Cow<str>` inputs. Admission precedes
transaction-owned copies; owned strings retain their allocation, and borrowed
copies reserve fallibly. Paragraph batches, SmartTag and move-bookmark metadata,
root transfers, split boundaries and picture batches check aggregate payloads
before copying incoming values. Failed admission leaves staged operations and
counters unchanged. Caller-owned batch DTO construction is outside this budget.

Accounting uses UTF-8 bytes for text/metadata, includes inserted paragraph
separators, and charges pictures for preserved hexadecimal transport including
whitespace. Composition now counts SmartTag/move-bookmark content and picture
transport consistently with staging. Commit rechecks aggregate admission.
Composition and durable replay explicitly retain the source-byte ceiling instead
of silently falling back to the ordinary replacement default. This policy does
not bound inverse snapshots or total process memory.

## Verification

The native suite passes 1,456 tests/doctests; the all-features facade passes
1,620. The final tightened metadata boundary passes with all 12 policy tests.
Coverage includes checked profiles, exact and exceeded aggregate budgets,
failed-batch reuse, retained source/output refusal, metadata and picture wire
accounting, composition, and owned-string allocation retention.

Warning-denied all-target Clippy and rustdoc pass for native RTF and the facade.
The RTF rustdoc JSON glossary checks 52 banned names. Crate boundaries pass for
66 packages, the 36 Python checker tests pass, and the OfficeCLI capability
delta is unchanged. Benchmark library compilation also passes.

Cargo validation uses external copied locks (cached cc 1.6.0 for root checks).
All three repository lockfiles retain their previous SHA-256 values. The owned
0228 worktree/branch and temporary locks, build outputs, wrappers and logs are
removed after qualification. No binaries or large evidence are committed.

## Remaining work

Focused append/redaction vocabulary, unified RTF CRUD, remaining tier-3 formats,
Python/CLI integration and full performance/corpus acceptance remain open. This
wave makes no runtime performance acceptance claim; the full goal stays active.
