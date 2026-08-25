# Change 0278: XLS source, FileSource, eager, and facade attribution

Date: 2026-08-25

Status: diagnostic attribution retained; freshness-session candidate selected
for a separate production batch; `performance_claim: none`

## Evidence boundary

This change adds the opt-in `xls_source_attribution` runner in
`tools/perf-baseline` without changing a production API, parser, selector, or
default benchmark case. The runner measures the same source-backed open, list,
and selected-cell projections through four matched positional sources:

- immutable owned bytes;
- an atomic-only positional file wrapper;
- the existing mutex/range-union diagnostic wrapper shape;
- production `litchi_core::FileSource` with counted and timed `len`, `read_at`,
  and `version` calls.

Eager XLS open/list/cell and facade open/list are retained as standalone
compatibility controls. They are not direct latency comparisons: source
construction is outside the source-family timer, while eager file/CFB/owner
construction and facade path detection/open are inside their respective
timers. Every report serializes its timing, counter, process, input-staging,
and oracle scope.

The original fixture is hashed and copied outside timing to a private,
read-only staged file. Every path-backed sample reopens that snapshot and the
runner verifies its size and SHA-256 again before publishing the report. This
warms the page cache and establishes logical-I/O attribution only; it provides
no cold-cache or physical-I/O evidence. Warmups and retained samples share one
child per mode/operation.

The harness revision is `b2c2260efa36421e446b52e0983f9bca8fd12ac3`.
Focused validation passed 9 unit tests, runner-only Rustfmt, and ten real-file
smokes. Clippy found no runner finding; strict package Clippy remains blocked by
22 pre-existing warnings in unrelated performance-tool files. Evidence review
found no P0-P2 staging, allocation, timing-scope, oracle, schema, or portability
defect.

## Retained protocol and identity

The runner was built in a clean detached worktree with Rust 1.95 release mode,
one build/benchmark worker, and CPU 2. The release binary is 8,417,112 bytes,
SHA-256
`9ce0514e5106f7d019b58d54d888ae39656f68eef04487f531b2f3f0783b72b3`.
Each of 17 mode/operation reports used 20 warmups and 100 retained warm-cache
samples in one child.

The corpus is `test-data/ole/xls/ConditionalFormattingSamples.xls`, 1,402,368
bytes, SHA-256
`d1942d857ffbd4d10ebca1745cd5d70c14af9d9f1388c91ed0a0800e31ad5ce7`.
Its Workbook stream is 1,314,225 bytes, SHA-256
`99305abd97f40bfc2fa4c052701bbebc971c1feb12278e8b76ecfbaca777676f`.

The source-backed family records the following `p50 / mean / p95 / p99`, in
nanoseconds. P50 is the conventional even-sample median; p95/p99 use nearest
rank.

| Mode | Open | List | One cell |
|---|---:|---:|---:|
| owned bytes | `218176 / 225507 / 252258 / 375017` | `218997 / 222767 / 253669 / 284492` | `285950 / 291378 / 327422 / 373776` |
| atomic file | `277687 / 281190 / 300514 / 321273` | `278709 / 280625 / 293485 / 301305` | `358024 / 361455 / 382528 / 401704` |
| tracked file | `1610331 / 1615730 / 1659518 / 1748292` | `1591230 / 1601736 / 1669163 / 1735923` | `1851158 / 1853657 / 1891729 / 1906460` |
| FileSource | `459173 / 462449 / 478650 / 500421` | `458813 / 461588 / 477119 / 496715` | `624941 / 627670 / 642025 / 669283` |

## FileSource freshness attribution

Atomic-file and FileSource have the same staged-file positional I/O boundary,
and every retained sample records exact-equal logical work:

| Operation | Reads / bytes | Version calls | Atomic -> FileSource p50 | Atomic -> FileSource mean | Extra mean version time | FileSource version share |
|---|---:|---:|---:|---:|---:|---:|
| open | `655 / 567685` | `1266` | `+181486 ns / +65.36%` | `+181259 ns / +64.46%` | `185463 ns` | `47.08%` |
| list | `655 / 567685` | `1266` | `+180104 ns / +64.62%` | `+180963 ns / +64.49%` | `183808 ns` | `46.83%` |
| one cell | `920 / 569391` | `1802` | `+266917 ns / +74.55%` | `+266215 ns / +73.65%` | `265068 ns` | `49.68%` |

The additional measured FileSource `version()` time closely explains the
entire central atomic-to-FileSource gap. This clears the prior investigation
threshold of 1% or approximately 50 microseconds for all three operations and
selects operation-scoped freshness work as the next bounded candidate.

The tracked wrapper is not a production proxy. Its range-union bookkeeping
uses 74.44%-77.86% of tracked mean elapsed time while reads, bytes, and version
calls remain exact. Its 565,201 open/list and 566,907 one-cell union bytes are
unique-range diagnostics, not additional I/O.

## Standalone compatibility controls

Eager p50 is 718,966 ns open, 709,929 ns list, and 730,863 ns one-cell; its XLS
owner phase is approximately 95% of p50. Facade p50 is 380,299 ns open and
358,294 ns list. Those values are reported only inside their own timing
families; facade counters are unavailable and no cross-family latency
attribution follows.

Source-backed and facade projections expose 16 worksheet names. Eager exposes
13 and omits `Quarters`, `Bike rating`, and `Compare to totals`; this known
compatibility difference is retained rather than normalized away.
`Products1!A2` projects as `string:4:Date` in source-backed and eager owners.
The projections are implementation-local compatibility oracles, not an
independent source-parser oracle or a source/eager equivalence assertion.

## Decision and claim boundary

Change 0278 is diagnostic single-revision evidence, not a control/candidate
comparison. `performance_claim: none`: no selector-wide, tail, cross-family,
cold-cache, physical-I/O, allocation/RSS, peak-memory, or broad XLS/CFB claim
follows. The selectable matrix remains **398 names** and the default remains
**36 cases / 198 records**.

The next production batch may implement only a private operation-scoped
CFB/XLS freshness session. It must preserve initial/final source fences,
`SourceChanged` and typed-error precedence, cancellation order, FILEPASS and
unknown-payload no-read behavior, limits, malformed tails,
STRING/CONTINUE/duplicate-last semantics, selected-sheet locality, and owned
result publication after the final fence. Keep it only if a clean
A1/B1/B2/A2 FileSource run improves p50 and mean in both directions by at least
1% or approximately 50 microseconds, remains inside drift gates, preserves
exact logical work and semantics, and has no reproducible adverse regression
above 5%.

The 17 raw reports, summary, and SHA/size manifest are retained in
[`results/0278-xls-source-attribution-20260825/`](../results/0278-xls-source-attribution-20260825/).
