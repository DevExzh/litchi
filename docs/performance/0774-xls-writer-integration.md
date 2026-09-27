# 0774 — integrate checked XLS writer fields

Status: integration validated. Current paired observations are scoped below;
no historical performance claim is adopted from the original branch.

The pending 0766 work applies cleanly to `25c3b27ba4` as three commits ending
at `b5739a89e8`; review corrections end at `836120efe7`. It addresses the XLS writer length-field work left open by
[0758](0758-owner-decisions-2026-09-24.md). The original 0766 worktree is
unchanged. [Evidence packet](results/change-0774/README.md).

The writer now checks selected BIFF length, count and index bounds when inputs
are registered, returning typed errors before changing retained state. The
change includes UTF-16 string lengths, fonts and formats, formula records,
filter values, names, pivot data and extension record indices. Some public
registration methods now return `Result`; callers must handle refusal instead
of receiving an index for unrepresentable data. This is a behavior/API change,
not a claim that every XLS writer surface has been audited.

Formula registration uses the writer's tokenizer and encoder to reject an
unwritable formula immediately. The static function-name table replaces a
freshly allocated map. The integration removes the branch's separate retained
formula-token map: it had no resource policy and kept token buffers plus a
high-water table allocation for the writer's lifetime. Temporary validation
tokens are released before staging; serialization re-encodes the formula.
This retains refusal semantics without increasing retained per-cell state.
Fresh measurements assess the additional validation work and the static lookup
together; they do not validate the removed cache's historical claims.

## Architecture

| Constraint | Evidence |
| --- | --- |
| ADR 0001/0006, correctness and typed refusal | Invalid registration precedes state mutation; adversarial boundary tests compare output with untouched control writers. |
| ADR 0012/0016, checked BIFF representation | Existing checked cell/location types remain; additional wire-field limits reject wrapping or truncation. |
| ADR 0002/0024, ownership | Changes remain in the XLS owner, without new production dependencies. |
| ADR 0005, performance evidence | Original timing claims are not imported; the paired probe measures the current source pair and identifies its limitations. |

## Verification

The complete XLS command passed 1,655 tests across 80 reported suites, with
one ignored doctest. Format, all-features/all-targets check, warning-denied
library Clippy, warning-denied rustdoc, the XLS-enabled facade check, and the
harness all-targets check and crate-boundary check passed. Source/configuration/lock hashes bind each
command. This is not a full facade or full harness test run.

The added length-field suite exercises boundaries and failure atomicity.
The property test uses 64 deterministic randomized workbooks by default and
compares accepted values through the reader, repeated output determinism and
an accepted-calls-only control. These checks do not establish native Excel
interoperability, crash durability or completeness of all malformed-input
handling. Existing durability tests also pass on the integrated writer.

Before review corrections, a separate run exercised 1,024 randomized workbooks with fixed seed
`0x07745eed00000000`; both tests in that target pass. The command, raw accepted/
refused tallies and source-stability check are retained. This extends the
existing property test rather than substituting a new self-mirroring test.

Final review confirms aggregate reference counting matches the stream writer:
worksheet/name/pivot activation, external sheets, add-ins and DDE/OLE links
preflight their contribution before mutation. The 1,369/1,370/1,371 boundary
test verifies refusal leaves serialized bytes unchanged. Dedicated near-limit
tests for each activation transition would extend this coverage; no claim of
exhaustive boundary coverage is made.

The same 1,024-case seed was rerun after both review corrections; both property
tests pass on final source `836120efe7` (`property-final-1024.json`).

## Fresh paired observations

On the recorded Linux host, Rust 1.95.0 release probes use LTO, CPU 12, identical
standalone dependency locks, four processes per case/leg in alternating order,
and nine samples after two warmups. Each sample constructs a writer, registers
20,000 cells and serializes to memory. Formula text is `A{row}+1`, using one-based
row text; numeric cells contain row plus one. This measures syntactic authoring
and serialization, not formula evaluation or native application behavior.
Writer destruction and output SHA checks are outside timing.

| Case | Median process p50, before → after | Median process peak RSS, KiB | Output bytes |
| --- | ---: | ---: | ---: |
| Formula | 20.071 → 14.131 ms | 13,420 → 13,636 | 1,159,680 |
| Numeric control | 5.819 → 5.786 ms | 11,566 → 12,030 | 817,152 |

Formula p50 falls 29.6% in this synthetic probe. Process-p50 spread is 1.77% /
2.31% before/after; numeric spread is 0.37% / 3.72%. Neither case crosses the 5%
latency or median peak-RSS regression flag. Numeric RSS has roughly 9.7%
process spread, so its 4.0% median increase is not a precise memory estimate.
Formula RSS rises 1.6%. All raw samples, ranges and flags remain in
`analysis.json`; p95/p99 from nine samples are maxima, not stable tail estimates.
Output lengths and SHA-256 digests agree in every before/after process.

A separate heaptrack lane runs one sample per process with no warmup. Counts
include startup, generated inputs, serialization, destruction and SHA; instrumented
timings are excluded from native results. Summed allocation-size histogram
counts equal the tool's allocation-call totals.

| Case | Allocation calls, before → after | Allocated bytes, before → after | Tool-reported peak heap |
| --- | ---: | ---: | ---: |
| Formula | 642,449 → 502,471 | 48,078,704 → 33,140,823 | 10.18M → 10.18M |
| Numeric control | 43,465 → 43,463 | 19,372,562 → 19,372,449 | 9.35M → 9.35M |

Formula temporary allocations increase from 119,639 to 219,630 consistent with validation
being separate from serialization; total calls and allocated bytes nevertheless
fall. Numeric temporary allocations are 649 in both legs. Peak heap values are
rounded tool output. The static table removes the repeated fourteen-name map,
while validation still encodes twice; no retained token cache is required for
these observed results. This is not an allocation-stack attribution study.

No claim extends to general workbook distributions, physical cold-cache,
remote sources, filesystem publication, parallel scaling, or Excel calculation.
The first probe build failed on an argument-parser formatting-string typo and
produced no native observations; its source/log is retained as `measure-0`.
The corrected source produced `measure-1` and the separate allocation lane.

## Completion and remaining scope

All final gates and the extra property run pass. Owned targets are removed
after executable identity checks, with cleanup receipts and a final packet seal.
The original 0766 worktree and unrelated main edits remain intact.

This closes the reviewed XLS integration batch, not the non-iWork program.
Legacy writer-wide resource budgets, broader corpus measurements, individual
near-limit activation tests, cold/range/scaling evidence and the full goal audit
remain separate work.
