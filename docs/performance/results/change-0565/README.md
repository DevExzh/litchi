# change-0565 evidence packet: one pass over the XLS globals

Change record:
[`docs/performance/0565-xls-globals-single-pass.md`](../../0565-xls-globals-single-pass.md).
Disposition: retained. `performance_claim: none`; no claim-registry entry.

## Contents

| Path | What it is |
| --- | --- |
| `plan.json` | Frozen before capture: hypothesis, the schedule model, the accepted contract changes and the gates. Carries an explicit `amendments` entry recording what changed after the baseline capture and before any candidate capture, and why. |
| `corpus-survey/` | The 104-fixture survey that changed the design, with its generator. Both the chosen and the rejected fill sizes replay from it. |
| `counters/` | Deterministic logical counters for both sides: reads, bytes, source observations, per operation and per source mode. |
| `syscalls/` | The `strace -f -c` isolation pairs at 1 and 11 samples for both sides, six operations each. |
| `latency/` | The 40 A1/B1/B2/A2 children, their receipts, `matrix.json` and `analysis.json`. |
| `capture_before.sh`, `summarize_before.py`, `run_abba.sh`, `analyze.py` | The capture and analysis drivers. |
| `baseline-identity.json` | The staged baseline binary hashes and the corrected toolchain. |
| `quality.md` | The gate results. |
| `decision.json` | The machine-readable disposition, including the falsified prediction and the accepted costs. |

## Replay

```sh
python3 -B -c "import json;print(json.load(open('docs/performance/results/change-0565/latency/analysis.json'))['summary'])"
```

The corpus survey reproduces byte-identically:

```sh
python3 -B docs/performance/results/change-0565/corpus-survey/survey.py \
  --repo . --out docs/performance/results/change-0565/corpus-survey
```

## Two things about the measurement setup

**The recorded toolchain was wrong and it was load-bearing.** The staged binary
identity first recorded rustc 1.98.1, which is the rustup default reported from a
directory the repository's `rust-toolchain.toml` does not govern. The build used
the pinned 1.95.0. This matters because `tools/perf_abba_summary.py` requires
every `environment` field except `git_revision` to match across legs, so a
candidate built with the default toolchain would have been refused outright.
`baseline-identity.json` carries the correction and three independent proofs.

**Neither leg ran from the working tree.** That same tool refuses any leg whose
worktree is dirty and requires the control and candidate revisions to differ, so
both legs ran from clean detached worktrees. Both carry the identical bounded
locality gate in `tools/perf-baseline`, so the legs differ only in the library
change.

## Noise floor

An A/A dry run on this host, one binary against itself at 20 samples, drifted
+1.7% to +3.2% at p50 and +11.0% at p99. The 5% review trigger is not comfortably
above the tail-statistic noise floor here, and movements below roughly 10% at p95
or p99 carry no information unless both directions agree.

## What is not here

No cold-cache, physical-device, remote or range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform capture. The counters are
logical calls over a warm, page-cached, immutable staged copy. The `xls-tiny`
corpora and the nanosecond-scale semantic selectors sit below this harness
clock's useful resolution.
