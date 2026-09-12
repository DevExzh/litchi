# 0543 individual hardware diagnostic review

All five flags are context-switch counts: one matched adverse increase and four
absolute same-build repeat changes. All other seven-counter/IPC/branch-rate
comparisons stay below their directional review thresholds. Coverage is complete;
these whole-child OS counters neither isolate planning nor establish a noise cause.
The candidate remains rejected by cap latency and Clippy.

| ID | Comparison | Before | After | Change |
| --- | --- | ---: | ---: | ---: |
| 0543-hardware-001 | matched dense-sparse 1 | 140 | 174 | +24.286% |
| 0543-hardware-002 | baseline dense-sparse r1→r2 | 140 | 189 | +35.000% |
| 0543-hardware-003 | baseline medium r1→r2 | 112 | 118 | +5.357% |
| 0543-hardware-004 | candidate dense-sparse r1→r2 | 174 | 131 | -24.713% |
| 0543-hardware-005 | candidate medium r1→r2 | 114 | 105 | -7.895% |

`python3 -B review_hardware.py` deterministically replays all derived flags and
their dispositions against the hash-bound hardware report.
