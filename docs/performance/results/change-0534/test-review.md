# 0534 CFB physical role/FAT paired-prefix test review

`scope: read-only test and differential-coverage review before implementation`

`revision: 7c1d3911da286b8a9dd0fde15475bbbeed2f84c9`

No 0534 candidate test or runtime source exists yet. This review identifies
the smallest behavior-focused coverage for the proposed rewrite of
`validate_physical_sector_layout`. No Rust source, build, test, benchmark, or
profiler was run here.

## Existing coverage

The private tests in [`file.rs`](../../../../crates/litchi-cfb/src/file.rs)
and `allocation_validation_tests.rs` already cover much of the surrounding
CFB pipeline:

| behavior | existing evidence | direct gap for 0534 |
| --- | --- | --- |
| Valid physical role/FAT reconciliation through a real open | `sample_file()` and the normal CFB/XLS/DOC/PPT quality suites | No direct assertion of every paired-array length shape. |
| Extra FAT entries beyond physical sectors | `tolerates_nonfree_fat_padding_beyond_the_physical_file` changes a padding marker to `ENDOFCHAIN` and expects `open` to succeed | No focused assertion that arbitrary extra markers are ignored by the private final pass. |
| Unclaimed marker refusal | Broad malformed CFB and allocation-validation tests exercise related errors | No direct exact-message assertion for each marker family in the final pass. |
| Short FAT table | Higher-level malformed-file checks cover truncated inputs | No direct exact missing-entry index or precedence assertion. |
| Claimed roles | `claim_sector` tests cover role publication and conflicts | No direct proof that this final pass skips marker checks for all six claimed roles. |
| Mutation and allocation behavior | The method is currently `&self`; the allocator lane covers complete operations | No direct snapshot assertion around the new loop, and no unit test should measure allocations. |

The existing long-padding integration test is important and must remain. It
does not make a direct private test redundant because it reaches the method
through all preceding loaders and cannot isolate short-prefix precedence.

## Smallest useful candidate-independent test

Add one table-driven private test beside `synthetic_fat_file` and the existing
0533 claim tests, for example
`validate_physical_sector_layout_preserves_paired_prefix_contract`. The test
should construct a minimal `OleFile<Cursor<Vec<u8>>>`, set a controlled
`sector_roles` vector and FAT vector, snapshot both vectors, call the private
method, and assert the exact typed result. The test remains valid if the
0534 runtime candidate is reverted because it calls the unchanged private
method rather than a new helper or implementation detail.

The cases should include these behavior families:

| family | representative setup | expected assertion |
| --- | --- | --- |
| empty/equal | roles `[]`, FAT `[]` | `Ok(())`, with both vectors unchanged |
| empty/long FAT | roles `[]`, FAT containing non-`FREESECT` values | `Ok(())`; all extra FAT padding is ignored |
| equal, all roles | one slot for each of `Unclaimed`, `Fat`, `Difat`, `Directory`, `MiniFat`, `MiniStream`, and `RegularStream`; use `FREESECT` for `Unclaimed` and arbitrary markers for claimed roles | `Ok(())`; this proves only the sentinel role is marker-checked |
| equal, bad sentinel | roles `[Unclaimed]`, FAT marker `0`, `ENDOFCHAIN`, `FATSECT`, `DIFSECT`, `MAXREGSECT`, and `u32::MAX - 1` in separate rows | exact `unclaimed physical sector 0 has FAT marker 0x{marker:08X}` |
| short, valid prefix | roles `[Unclaimed, Fat]`, FAT `[FREESECT]` | exact `FAT does not contain an entry for physical sector 1` |
| short, earlier bad marker | roles `[Unclaimed, Fat]`, FAT `[ENDOFCHAIN]` | marker error at sector zero wins before the missing-tail error |
| short with each role present | for every role, roles `[role, Unclaimed]` and one valid marker for the first role | exact missing error at sector one, including when the first role is claimed |
| long with arbitrary tail | roles containing all seven roles, FAT with valid common-prefix markers plus one or more non-`FREESECT` tail markers | `Ok(())`; no tail marker is inspected |

The all-role rows should be data-driven rather than a copied version of the
production loop. A small helper may compare `OleError::CorruptedFile` payloads
exactly, but it must not reproduce the candidate's length or marker logic.
After every row, compare `file.sector_roles` and `file.fat` with their
snapshots. The `&self` signature already proves the method cannot mutate
through this API; the snapshot check protects against an accidental interior
or future staged mutation and makes the intended boundary explicit.

The marker extremes include `u32::MAX`, which is the valid `FREESECT` marker,
and values adjacent to the reserved marker range. This checks formatting and
classification without pretending that a test can allocate a vector near
`usize::MAX`. The candidate performs no physical offset arithmetic, so
pointer-overflow safety is a source-level obligation: safe slice iterators,
`Vec::len()`, and no unchecked index or wrapping calculation must remain in
the implementation. Existing malformed-input no-panic tests remain the
appropriate broad guard for hostile `u32` values.

No test should assert a particular iterator shape, instruction count, branch
layout, or allocation count. Those are owned by the frozen performance and
assembly lanes. No new public API or cross-crate fixture is needed.

## Required broader gates and performance guards

The focused test must be accompanied by the frozen 0534 quality commands:

- both all-feature and no-default-feature `litchi-cfb` tests;
- all-feature `litchi-xls`, `litchi-doc`, and `litchi-ppt` tests;
- workspace check, warning-denied Clippy, rustdoc, crate-boundary, and strict
  performance-claim checks; and
- existing real DOC/XLS/PPT, CFB padding, malformed-input, stream-chain, and
  physical-ownership tests.

The native lane remains the authority for latency. All four primary XLS
workflows must improve p50 by at least 3% in both paired repeats. The 0534
plan also requires lower XLS-owned constructor `Ir` and lower physical-pass
exclusive `Ir` in the XLS-owned and CFB-few-large profiles, while CFB tiny and
many-small remain guards. Every mean, percentile tail, maximum, whole-child
RSS, same-build drift, and allocation-vector result remains visible. A
matched or same-build change over five percent is a review trigger; it is not
permission to discard a row.

The allocator lane should compare all four measured vectors—allocation calls,
reallocation calls, allocated bytes, and incremental region peak—without
using allocator-instrumented elapsed time as native evidence. The paired-prefix
rewrite should add no success-path allocation, but that expectation is not a
measurement result. Assembly and binary inspection must use the exact final
runtime-plus-test source that was measured; any changed rebuild requires a
fresh full ABBA.

The test review disposition is: **one focused private table-driven test is
required for exact marker, length, role, precedence, and no-mutation behavior;
the existing real-padding and malformed-CFB tests remain required; no
candidate-dependent test or implementation-shaped assertion is needed.**
