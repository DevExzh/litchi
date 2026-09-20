# Change 0723 evidence packet

Bounded XLS worksheet replay checkpoint pilot after baseline `45cb480eaa`.
`performance_claim: none`. The non-iWork performance program remains active.

The candidate preserves the worksheet-start and SST checkpoints and adds one
optional exact target-frame checkpoint per admitted worksheet index. The fixed
logical charge changes from 224 to 264 bytes. Full scan validation, original
source/execution fences, selected-frame decoding and duplicate ordering remain.
See `hypothesis.md`, `design-review.md` and the frozen `candidate-source/` files.

`build.py baseline|candidate` serially builds the unchanged standalone probes
from changes 0684 and 0686, then freezes each executable under
`/home/zhuhe/code/litchi-0723-bin/PHASE`. Build receipts bind source and binary
hashes. The initial candidate builds are retained separately because the public
integration tests were strengthened before final qualification and any paired
measurement; production code did not change for that correction.

`run-integration.py` runs formatting, checks, Clippy, CFB/XLS tests, facade tests
and rustdoc. `measure-corpus.py PHASE` compares 126 real XLS fixtures in both
source modes and the generated 70,001-cell fixture. `profile.py PHASE` records
a separate two-million-query owned 54016 first-cell profile. Its single total
and sample shares are attribution evidence, not paired latency proof.

The prospective `plan.json`, `pilot.py` and `analyze.py` define separate native,
repeat-process, allocator and counted-source matrices. Freeze them before
capturing the main matrix. The trace driver uses separate instrumented binaries;
its timings are never performance evidence. Do not pool diagnostic and native
captures or replace an adverse observation with a later sample.

All paths and CPU 12 affinity are explicit and must be adapted consistently on
another host. Source modes are owned bytes and files with warm OS caches.
Physical cold-device, remote, concurrency, cross-platform and native Office
performance are outside this packet. No registered claim or CRUD support
promotion follows.

## Disposition and results

**Rejected.** `analysis.json` records 14 failed native groups and four failed
repeated-query groups. Both late-target benefit gates, exact semantic checks,
allocator bounds and all capture bindings pass. The main report is
[0723](../../0723-xls-target-frame-chain-checkpoint.md); `measurements.md` and
`measurement-details.json` expose all central failures, tails and control drift.
Baseline production bytes are restored. The three candidate files remain under
`candidate-source/`; `candidate-production.patch` covers the two production files.
The focused integration test is archived with the candidate and is not installed
in the restored production tree.

The frozen native census is 24 groups × six legs × 100 measured owners × eight
queries, plus three warmup owners per group/leg. Repeated loops are 16 groups ×
six legs × nine processes × 50,000 measured queries. Allocator capture is 96
groups × two phases × three processes. Diagnostic runs are separate and never
contribute timing samples. Earlier failed diagnostic attempts are retained under
`initial-trace-failure/`, `initial-sequence-build-failure/`,
`initial-sequence-argument-failure/` and `trace-probe-prebuild-correction/`.

## Replay

Run from the repository root. A fresh measurement reproduction needs an isolated
checkout of baseline `45cb480eaa`, the archived candidate files, the existing
0684/0686 probes, and matching fixture bytes. Build baseline before installing
the three candidate files, build candidate, then freeze and capture with the
packet scripts. Use a fresh output directory; the driver refuses overwrites.
Do not treat another machine's results as replacements for this packet.

For analysis replay, temporarily install the three archived candidate files in
an isolated checkout with the packet and original fixtures. The frozen analyzer
requires the captured candidate source census even though production was
reverted. After binary cleanup, `cleanup.json` supplies exact executable identity
witnesses. Run:

```text
python3 -B docs/performance/results/change-0723/analyze.py --offline
python3 -B docs/performance/results/change-0723/report-tables.py
python3 -B docs/performance/results/change-0723/trace-analyze.py
python3 -B docs/performance/results/change-0723/audit-corpus.py
python3 -B docs/performance/results/change-0723/negative-checks.py
```

The primary analyzer's expected exit code is **1** because the candidate fails
retention. This is not a successful optimization gate. Do not change its frozen
thresholds. The independent audit distinguishes a verified rejection from a
retention PASS; see `audit-review.md`. Restore baseline sources after replay.
Generated timestamps may change on replay; do not overwrite the sealed packet
when testing reproduction. `restoration.json` records the final production
source identity, while `artifact-seal.py --check` verifies exact packet bytes.

The current restored checkout can run `restoration-check.py` and
`artifact-seal.py --check` directly without installing the candidate. The former
checks baseline source, archived candidate identities, unchanged constraints,
rejected disposition and removal of all five owned external roots. The full
independent audit's rejection report replays exactly after cleanup; its expected
exit code remains 1. `terminal-replay.json` records the analysis and correctness
commands and expected statuses.
