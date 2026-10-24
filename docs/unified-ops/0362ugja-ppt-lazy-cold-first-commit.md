# 0362ugja — PPT lazy cold first-commit route

The frozen PPT lazy probe's normal timed route is deliberately warmed.  Each
timed child opens a fresh package, then `timed()` calls `run_batch()` for three
warmups before starting its measured sample.  `primary_noop_commit` therefore
does exercise a never-hashed package, but its first hash/block-root work is in
the warmup interval and is not a first-commit timing observation.

The supplemental consumer route is named `primary_first_commit_cold`.  The
public adapter `tools/ppt_lazy_cold_first_commit.py` now takes the repaired
full manifest and its explicit canonical build-input projection, validates both
endpoints through `ppt_lazy_native_qualification.validate_inputs`, and exposes
a bounded symmetric pair:

```text
ABBA = a1 baseline, b1 candidate, b2 candidate, a2 baseline
rounds = 2 (maximum), normal + allocator binary at every position
route = primary_noop_commit, warmups = 0, iterations = 1, samples = 1
CPU = 29 (reserved quiet CPU)
```

Every child runs through the public runner's target accounting, ownership,
quiet-host guards, executable recheck, and before/after endpoint lifecycle.
The aggregate contains only the probe-reported `elapsed_ns` operation values:
`operation_total_ns`, `operation_count`, and integer
`operation_mean_ns`.  Process setup/teardown wall time is not substituted, and
no median or calibration estimate is used.  The pair validator rejects missing
roles, reordered ABBA positions, non-one-operation samples, and aggregate
values that are not the exact sum/count of the operation observations.  Each
sample also retains the public runner's child-event provenance; the final
validator checks its route label, arguments, binary digest shape, guard,
target, and child-event fields.

The source-only contract tests are
`tools/test_ppt_lazy_cold_first_commit.py`.  The CLI can validate and emit a
plan without native execution:

```sh
python3 tools/ppt_lazy_cold_first_commit.py \
  --manifest /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-final-472b-repaired-20261024/ppt-lazy-edit-472b-repaired.json \
  --projection /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-final-472b-repaired-20261024/ppt-lazy-build-input-projection-472b-repaired.json \
  --out-dir /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-final-472b-repaired-20261024/cold-first-plan-v3-20261024
```

The source-only plan completed with `execution=not_run`; its external
`plan.json` is 2628 bytes, SHA
`4f3f4ce89cb5731fdd8266e2064a92805537e9e3dea97ab8209fab8129af306e`.

This route is supplemental and is not inserted into the frozen 28-route ABBA
manifest or its acceptance claim.  No native execution, calibration,
allocation measurement, timing, or performance acceptance was performed in
this packet.  The existing 0357b first-commit residual (1.105×, no budget)
remains the only accepted first-commit performance record until a separately
authorized quiet run produces this pair's evidence.
