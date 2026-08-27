# Change 0323: Reject redundant XLS freshness-probe removal

Status: rejected before implementation.

## Evidence

The control binary was built offline in release mode with serialized
`CARGO_BUILD_JOBS=1`. The standalone `tools/perf-baseline/Cargo.lock` needed a
one-line refresh before that build; this was reproducibility maintenance only.
Production behavior remains unchanged.

The retained control artifacts were:

- Binary size: `8,465,104` bytes
- Binary SHA-256: `143632ff666fae431ce9be36d9fcdddcad6b67cf899fb8de79da2a8038147e26`
- Corpus size: `1,402,368` bytes
- Corpus SHA-256: `d1942d857ffbd4d10ebca1745cd5d70c14af9d9f1388c91ed0a0800e31ad5ce7`

The six-selector A1 smoke used CPU 2, 20 warmups, and 50 samples. Its
retained `FileSource` control measurements were:

| Selector | p50 (ns) | Mean (ns) | Versions | Version time mean (ns) |
| --- | ---: | ---: | ---: | ---: |
| `open` | 472,385 | 489,578 | 1,266 | 225,371.98 |
| `list` | 457,245 | 461,164 | 1,266 | 215,948.12 |
| one-cell | 619,668 | 624,683 | 1,813 | 310,408.6 |

Semantic observations were stable: 16 sheets, exact sheet names, the cell
string `4:Date`, and `source_version_stable true`. Logical reads and bytes
matched the atomic control.

Source inspection proves the expected runner reductions if only the middle
`ensure_current_parts` probe were deleted:

- `open`: 3 fewer version calls.
- `list`: 3 fewer version calls; this runner has the same measured open/list
  total as `open`.
- one-cell: 5 fewer version calls.

Using the A1 average version-call costs, the measured attribution estimates are
approximately 534 ns (`open`, 0.109% of the mean), 512 ns (`list`, 0.111%),
and 856 ns (one-cell, 0.137%). The attribution estimate is far below the
predeclared acceptance gate of at least 1% or 50,000 ns. A candidate gain was
not directly measured, so the production edit and full ABBA run were declined
on expected ROI.

## Decision

No production candidate or B leg was implemented, and no full 12,000-process
ABBA run was started. The evidence is diagnostic only and makes no direct
candidate speed claim.

The smoke artifacts were retained under `/dev/shm` and were not committed.
The worktree contained protected unrelated dirty files, so no unrelated state
was modified.

This record makes no RSS, I/O, cold-start, or physical-storage claim.
