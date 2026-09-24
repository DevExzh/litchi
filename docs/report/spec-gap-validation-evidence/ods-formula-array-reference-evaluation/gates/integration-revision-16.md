# ODS value evaluator integration, revision 16

This batch replaces the shallow lazy-condition shape probe with an iterative
planner that discovers composed and nested condition shapes before selecting
branch metadata. Condition caches use normalized shape coordinates to avoid
repeated provider reads. Unselected branches remain unevaluated.

The public borrowed array and reference-list views now compare structurally.
Resolver coherence and reference geometry/read limits are documented explicitly.
`Evaluated::to_owned` adds fallible, budgeted lifetime-independent ownership for
scalar, array, reference and ordered reference-list results. Its retained memory
reservation outlives the copied storage; cancellation is checked before allocation
and before publication. Reference lexical markers and duplicate order survive
conversion.

## Verified gates

[The source-bound receipt](integration-revision-16.json) records all commands,
compiler identity, environment, and before/after hashes of 454 input files.
The isolated checkout matched the canonical checkout, and both remained unchanged.
The detached checkout HEAD is supplemental provenance; the file hashes identify
the tested production source.

- All-feature, all-target runtime tests: 1,018 passed, zero failed or ignored.
- Doctests: 4 passed.
- All-feature, all-target Clippy with warnings denied: passed.
- Rustdoc with warnings denied: passed.
- Formatting: passed.

The runtime total includes 46 value-evaluator integration tests, 16 worksheet
resolver tests, seven owned-value integration tests, and 393 library tests.
The five raw logs are retained in `integration-revision-16-logs.tar.gz`; the
receipt records the archive hash and each decompressed member hash.

The [ownership review](value-owned-review.md) clears the owned allocation and
cancellation boundaries. This batch supersedes the two enabled failures recorded
in checkpoint 14. It does not establish release performance acceptance or complete
the specification gap audit. Release scalar comparisons, value/worksheet scaling,
owned-copy measurements, and the broader remaining audit work are still pending.
