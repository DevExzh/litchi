# 0361fc — Resume constructor measurements after host interruption

- Parent: 0361fb; production source remains `536d7fdcca`.
- Status: protocol validated and recorded before any matched-v5 native attempt.

Unrelated builds repeatedly interrupted the matched-v4 runs before they
qualified any blocks. The previous runner restarted every route after an
interruption, even though its collector retained complete valid blocks. The
new versioned runner verifies all earlier attempts and starts at the first
unfinished route. Its manifest declares the remaining block schedule before
calibration. The source, probe, protected matched-v5 binaries and allocation
receipts remain unchanged.

Every route still needs exactly five complete ABBA blocks. Route order,
500 ms calibration target, 150 ms minimum measured leg, two warmup cycles,
CPU affinity, output checks and before/during/after host guards are unchanged.
Each resumed route calibrates and warms up again. Every completed block,
including an unfavorable ratio, contributes to its final median. Partial
blocks never contribute. An overlap stops the attempt and retains its files.
The next attempt's schedule depends only on verified completed block counts.

The collector rejects skipped routes, repeated completed blocks, missing
attempt numbers, inconsistent source binaries or script hashes, tampered
outputs and incomplete attempts claiming completion. Numeric attempt ordering
also prevents `native-a10` from sorting before `native-a2`. Existing v4
records stay with their original scripts and source; they are not recast as
v5 evidence. No native timing result is claimed by this protocol change.

The new scripts are [the runner](results/0361fc/constructor-abba.py) and
[the collector](results/0361fc/constructor-collect.py). Eight
[synthetic receipt tests](results/0361fc/test_constructor_resume.py) pass,
covering valid resume, overlap, output tampering and schedule rejection.
Python compilation passes. Running the new collector on the real matched-v5
directory verifies the existing source/build, eight output comparisons and
15 allocation routes, and correctly reports zero native attempts.

For attempt number N (the next consecutive number), run from this worktree:

```sh
python3 docs/unified-ops/results/0361fc/constructor-abba.py \
  (deleted-coord-artifact) \
  (deleted-coord-artifact) \
  (deleted-coord-artifact) \
  > (deleted-coord-artifact) 2>&1
```

The coordinator must obtain a quiet host window before starting. All scripts,
raw attempts and external artifacts retain their recorded hashes. Allocation
and latency acceptance, final integration and the workspace gate remain
required.

## Matched-v5 attempt a1

The first v5 attempt ran after all three workers acknowledged a heavy-command
hold and the coordinator's host preflight reported no blockers. An unrelated
`peculiar` release build started during the first route's calibration. The
during/after guards detected it and stopped the attempt before any complete
ABBA block. The [verified receipt](results/0361fc/evidence.json) retains the
manifest, runner log and external raw-artifact inventory. It reports one
interrupted attempt and zero qualified blocks. No native ratio is accepted.
The workers' hold was released after the interruption; the next native attempt
still requires a quiet host window.
