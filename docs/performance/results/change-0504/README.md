# ODG direct transition reuse evidence

See [the change record](../../changes/0504-odg-direct-transition-reuse.md) for
results, ADR constraints, preservation tests and measurement limits.

`source-manifest.json` identifies the clean control revision and the exact
`candidate.patch`. The probe and its locked dependencies are the committed
`docs/performance/probes/0502-odg-open` project. To reproduce, build that probe
at the recorded base revision, copy its executable as a control, apply the
candidate patch in an isolated checkout, then build and copy the candidate.
Use the same toolchain and command in `environment.json` for both builds.
No benchmark binaries or build directories are retained by this batch; a
rebuild at another absolute path may have a different binary identity, which
must be recorded by the new capture rather than relabeled as the old binary.

Run a new capture into a separate output directory:

```sh
python3 -B docs/performance/results/change-0504/capture.py \
  --before /absolute/control/odg-open-probe \
  --after /absolute/candidate/odg-open-probe \
  --output /absolute/new-capture
```

The retained capture can be summarized without any executable:

```sh
python3 -B docs/performance/results/change-0504/analyze.py
```

This validates raw sample-derived statistics and matching corpus/checksum
fields, then rewrites `summary.json` deterministically. It reuses the retained
0502 bootstrap implementation. `capture.json` records exact commands, hashes
and A1/B1/B2/A2 order. The `before/*-pilot.json` reports are the separate
pre-acceptance pilot and are excluded from formal comparisons.

`callgrind-*.out` and `heap-*.zst` retain the small raw profiles; the adjacent
text summaries retain symbolized attribution. Profiles include setup and
verification, use zero warmups and one measured sample, and are supplementary
to uninstrumented timing. No hardware counter or operation-local heap claim
follows. Profile command lines are present in their logs/raw headers.

`gates.json` records actual validation commands and exit status. `custody.json`
binds all retained evidence members except itself by SHA-256. Cleanup removes
only `/tmp/litchi-goal-0504`, after copying the intended reports and traces;
`cleanup.json` records the actual removal. Other workspace outputs are outside
this cleanup scope.
