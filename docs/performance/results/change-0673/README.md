# Evidence packet — change 0673

The managed ZIP index path had already merged its first central-directory read
with the index scan in change 0632, but it still allocated a 64 KiB locator
scratch before probing the common terminal EOCD. This packet measures and
implements the bounded second stage from change 0651 row 18: a 46-byte stack
probe on the managed path, with the historical 64 KiB scratch allocated only
when that probe misses.

The result is retained in `soapberry-zip` with `performance_claim: none`.
The three-fixture probe records one fewer open allocation and 65,537 fewer
allocated bytes on each fixture, with identical open request counts, bytes,
and source-version observations. The 533-container open oracle produces
byte-identical 18,494-line reports.

## Provenance

| | |
| --- | --- |
| base commit | `5fa92d7ced5a78f3c2c84a6afe9d0ba404253717` (`0657`) |
| branch | `perf/0673-zip-locator-scratch` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-5fa92d7ce` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0673` |
| host | AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0 |
| probe pin | `taskset -c 8`; release builds, `CARGO_PROFILE_RELEASE_DEBUG=0`, jobs 2 |

The before and after legs use the same source fixtures and change 0632 probe
body. The oracle uses the same 533-container list as change 0632 and stages
each binary in a temporary target outside the worktree.

## Contents

| path | purpose |
| --- | --- |
| [`counts/summary.txt`](counts/summary.txt) | three-fixture allocation and request summary |
| `counts/counts-*.txt` | complete before/after probe output for XLSX, PPTX, and DOCX |
| `counts/binaries.sha256` | staged probe binary identities |
| `counts/provenance.txt` | the two trees and temporary build directory |
| [`probe/run.sh`](probe/run.sh) | reproducible paired allocation/request probe |
| `oracle/src/main.rs` | change 0632's open differential oracle |
| `oracle/run.sh` | builds both legs, runs all 533 containers, and compares reports |
| `oracle/summarize.py` | writes the retained cost table and report-difference summary |
| `oracle/corpus-open-costs.tsv` | per-container before/after open costs |
| `oracle/summary.txt` | report line count, cost totals, and difference classes |
| `oracle/reports.sha256` | hashes of the two temporary full reports |
| `oracle/report-identity.txt` | report identity result |
| `oracle/report-diff-classes.txt` | classified report differences |
| `oracle/binaries.sha256` | staged oracle binary identities |
| [`decision.json`](decision.json) | machine-readable decision and evidence record |
| [`log-sections.md`](log-sections.md) | four coordinator log sections |
| [`gates.txt`](gates.txt) | final focused verification commands and outcomes |

The full oracle reports are intentionally temporary; their hashes, line count,
per-container costs, and difference classification are retained in the packet.

## Reproduction

From the repository root:

```sh
docs/performance/results/change-0673/probe/run.sh \
  /home/zhuhe/code/litchi-worktrees/before-5fa92d7ce \
  /home/zhuhe/code/litchi-worktrees/0673

docs/performance/results/change-0673/oracle/run.sh \
  /home/zhuhe/code/litchi-worktrees/before-5fa92d7ce \
  /home/zhuhe/code/litchi-worktrees/0673
```

Both runners use offline Cargo resolution, jobs 2, and CPU 8 for the measured
process. They do not remove or modify repository target directories.

## What this packet does not claim

It does not claim a latency, throughput, cold-cache, physical-I/O, peak-RSS,
instruction-count, concurrency-scaling, remote-source, or cross-platform
result, and it does not add a claim-registry entry. The missed-probe fallback
keeps the old bounded 64 KiB allocation and search-window behavior.
