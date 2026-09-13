# change-0562 evidence packet

Change record: [`docs/performance/0562-zip-descriptor-read-once.md`](../../0562-zip-descriptor-read-once.md).
Disposition: retained. `performance_claim: none`; no claim-registry entry is added.

## Contents

| Path | What it is |
| --- | --- |
| `plan.json` | Frozen before capture: hypothesis, equivalence argument and gates. |
| `traces/` | Post-change `strace -f -e trace=pread64` captures for the same four cases change 0561 captured before the change. |
| `summarize_reads.py`, `read-repetition.json` | The same summarizer 0561 uses, so before and after are computed identically. |
| `latency/` | The A/G/G/A children for the OOXML file-source selectors. |
| `analysis.json` | Before/after read counts and every latency comparison in both directions. |

## Replay

```sh
python3 -B docs/performance/results/change-0562/summarize_reads.py \
  --traces docs/performance/results/change-0562/traces \
  --output docs/performance/results/change-0562/read-repetition.json
```

Compare against `docs/performance/results/change-0561/read-repetition.json`,
which was produced by the same script from captures of the same four cases.

## What is not here

No cold-cache, physical-device, remote/range-source, peak-RSS or allocation
capture. The traces are warm-cache and single-sample; they count syscalls, not
physical device I/O.
