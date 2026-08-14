# Change 0095: XLS semantic same-length splice publication

Date: 2026-08-14

## Scope

The existing source-backed XLS comment and worksheet-visibility owners now
submit source-relative [`SameLengthStreamSplice`] values through the protected
`litchi-ole-common` publisher wrapper. The wrapper retains the validated CFB
topology, source/version and source/target fingerprint checks, complete
composed-artifact reopen, selected-owner readback, and signed/encrypted/DRM
refusals. No public archive type or physical CFB identifier is exposed.

This is a structural publication tranche, not a general XLS performance
claim. The default case count is unchanged. The existing eager/source-backed
comment and visibility selectors remain the matched controls; only the
source-backed path changed its publication handoff.

## Replacement-byte evidence

The values below are the exact source-relative replacement bytes submitted to
the splice plan, compared with the previous complete `Workbook` replacement
buffer. They are not claims about total CFB source I/O, candidate memory,
sequential output, or process allocations.

| Owner / workload | Splices | Replacement bytes | Previous full `Workbook` bytes |
|---|---:|---:|---:|
| Existing comment, one owner | 2 | 109 | 80,946 |
| Existing comments, 256-owner batch | 512 | 27,904 | 80,946 |
| Worksheet visibility, one owner | 1 | 1 | 18,166 |
| Worksheet visibility, 64-owner batch | 64 | 64 | 18,166 |

The comment owner emits separate NOTE/TXO-family ranges while the visibility
owner emits one-byte `BoundSheet8.hsState` ranges. All replacements retain
their source lengths; length-changing comment edits continue to use the
explicit eager path.

## Release ABBA observation

The comparison used CPU 2, release binaries, 10 warm-ups, 100 measured
samples per cell, and before-A/after-A/after-B/before-B order. The percentages
below are source-backed after-versus-before p50/p95 deltas in the two balanced
directions:

| Source-backed workload | p50 A / B | p95 A / B |
|---|---:|---:|
| Existing comment, one owner | +0.215% / +0.199% | +0.610% / +1.735% |
| Existing comments, 256-owner batch | +1.111% / +0.634% | +0.607% / -0.687% |
| Worksheet visibility, one owner | -0.440% / -0.173% | +0.039% / -0.697% |
| Worksheet visibility, 64-owner batch | +1.479% / +1.037% | +2.538% / +2.881% |

The eager controls drifted substantially between paired directions. These
results therefore do not establish a material latency improvement or
regression and no speedup is accepted. The accepted result is the checked
replacement-byte reduction above. Allocation, RSS, and physical/source-I/O
evidence remain open.

## Preservation and verification gates

The existing gates remain mandatory and untimed: complete semantic readback,
source/target fingerprints, exact changed-span checks, untouched worksheet and
opaque-stream preservation, patch replay/inverse, exact no-op identity,
stale-source handling, limits, partial-sink behavior, and signed/encrypted/
protected-source refusal. The source-backed owners still retain their complete
candidate snapshot; the splice plan reduces replacement staging, not the
candidate-memory bound.

The focused implementation/test surfaces are:

- `crates/litchi-ole-common/src/source_backed_overlay.rs`, including the
  protected wrapper's splice-plan regression;
- `crates/litchi-xls/src/comments/transaction.rs` and
  `crates/litchi-xls/tests/xls_comment_transactions.rs`;
- `crates/litchi-xls/src/sheet_visibility.rs` and its source-backed
  visibility regressions; and
- the existing eight `tools/perf-baseline` XLS comment/visibility selectors,
  which now record splice counts and replacement bytes without adding case
  names.

The compact release artifact is
[`xls-semantic-splice-abba-0107-summary.json`](../results/xls-semantic-splice-abba-0107-summary.json).
It should retain the binary/corpus hashes, ABBA protocol, source-backed
replacement-byte and splice-count arrays, paired distributions, and the
explicit claim boundary above. iWork remains deferred while the `iwa-*`
crates change separately.

[`SameLengthStreamSplice`]: ../../../crates/litchi-cfb/src/splice.rs
