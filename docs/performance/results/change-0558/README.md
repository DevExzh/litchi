# change-0558 evidence packet

Change record: [`docs/performance/0558-ole2-single-read-fence.md`](../../0558-ole2-single-read-fence.md).
Disposition: retained. `performance_claim: none`; no claim-registry entry is added.

## Contents

| Path | What it is |
| --- | --- |
| `plan.json` | Frozen before the candidate build and before any capture. Declares the hypothesis, the acceptance gates, the CPU pin and the deviation reason. |
| `environment.json` | Host, toolchain, tool versions, load at capture, and the explicit statement that quiescence is not established. |
| `binary-identity.json`, `input-identity.json` | SHA-256 and size of both measured binaries and of the corpus. The driver refuses to run if the two binaries hash equal. |
| `run.py` | The capture driver. One fresh child per cell; retains a receipt, stdout report and stderr for every child. |
| `latency/` | 48 children: two repeats of A1/B1/B2/A2 over {file-source, owned-readat} x {open, list, one-cell}, 20 warmups and 100 samples each. |
| `syscalls/` | 12 `strace -f -c` children over the same matrix, 1 warmup and 5 samples each. |
| `guardrail/` | 56 case/corpus rows across 14 CFB/DOC/PPT/XLS/OLE-common selectors in all four ABBA legs. |
| `profiles/` | Callgrind branch-simulation dumps for `ole_common_one_edit_save` many-small, baseline and candidate. |
| `analyze.py`, `analysis.json` | Gate checks, deterministic counters, syscall counts, paired latency, same-binary and cross-repeat drift. |
| `analyze_guardrail.py`, `guardrail-analysis.json` | Per-case both-direction comparison and the explicit list of adverse-in-both-directions cells above 5%. |

## Replay

```sh
python3 -B docs/performance/results/change-0558/analyze.py
python3 -B docs/performance/results/change-0558/analyze_guardrail.py
```

Both scripts read only the retained children. `analyze.py` exits non-zero if any
identity gate fails.

## What is not here

No cold-cache, physical-device, remote/range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform capture. The measured
binaries are release builds retained outside the repository; their SHA-256
values are recorded in `binary-identity.json`.
