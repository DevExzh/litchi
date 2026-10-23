# 0362yf — matched native ODP qualification harness

This slice supplies the missing native qualification infrastructure for the
ODP source-backed owner after the correctness and source-owner work recorded
by 0362y–0362ye. It is source and runner infrastructure only; no release
binary, timing result, allocator capture, Callgrind file, or performance claim
is included.

## Probe contract

`tools/perf-baseline/src/bin/odp_native_qualification_probe.rs` is a new
parameterized probe. Its fixture, appended title/body, route, batch size and
sample count all come from command-line arguments. The routes are:

| Route | Timed lifecycle |
|---|---|
| `open` | open a fresh positional source-backed presentation |
| `read` | open, read the slide count and complete visible text |
| `write` | open and write the unchanged source-backed presentation |
| `append` | open, commit one bounded tail append, and write the deferred owner |
| `replay` | create the append patch, open an equal independent source, replay it by content, and write |

`--mode verify` is untimed and must run before calibration. It checks exact
source/no-op bytes, semantic reading, append change, same-owner replay,
equal-content cross-copy replay, inverse restoration, and refusal of a changed
recipient. Every output is reported as a byte count and SHA-256 digest. The
timed mode only reports bounded sink lengths; its exact digest precondition is
the verify result.

`--mode calibrate` doubles the per-child operation batch until a caller-set
target duration is reached, with a finite maximum. `--mode timed` performs
warmups outside its measured interval and returns one elapsed batch sample per
request. Source byte cloning is setup work outside each operation timer, while
opening, semantic work, publication and replay remain inside their route's
timer.

## Runner contract

`tools/odp_native_qualification.py` accepts an execution manifest rather than
embedding a baseline or candidate revision. The manifest supplies, for each
endpoint, a source worktree, requested Git ref, executable path and external
target directory; it also supplies the fixture byte count/SHA-256/provenance,
CPU, load ceiling, warmups, blocks, calibration targets and timeout. At run
time the runner resolves both refs, requires clean worktrees, hashes both
executables, verifies the frozen fixture, and records all resolved identities.
No source-specific historical hash is compiled into the probe or runner.

The runner performs the following sequence:

1. Run both endpoint correctness preflights and require identical exact output
   digests, semantic-read receipts, patch proofs and all correctness flags.
2. Calibrate every route independently for both endpoints.
3. Run calibrated `A1 control, B1 candidate, B2 candidate, A2 control`
   blocks, pinned with `taskset` to the manifest CPU.
4. Before calibration and every ABBA leg, fail closed if `/proc/loadavg`
   exceeds the manifest ceiling or a compiler/profiler/benchmark process is
   present. Each timed child must return a positive batch no shorter than the
   manifest minimum.
5. Write raw child reports, correctness receipts, calibration, per-block
   checkpoints and the final report atomically under the caller's external
   artifact directory. Existing non-empty output directories are rejected.

Target and artifact directories below either source worktree are rejected.
The runner invokes children without a shell and supplies an external
`TMPDIR`; no Cargo command is run by the runner.

An execution is prepared with:

```text
python3 tools/odp_native_qualification.py \
  --manifest /external/odp-qualification/input.json \
  --out-dir /external/odp-qualification/run
```

The manifest is intentionally frozen at execution time. Its fixture
`provenance` field must identify the tracked test-data source or externally
approved fixture source; the runner records that text beside the verified
digest. The `--plan` mode validates the manifest shape and writes the complete
route/protocol plan without touching binaries or timing.

## Validation

The runner's standard-library contract tests pass:

```text
python3 -m py_compile tools/odp_native_qualification.py tools/test_odp_native_qualification.py
python3 -m unittest tools/test_odp_native_qualification.py
```

No Cargo/build/timing slot was used for this source-only harness commit. A
quiet-host execution against the owner's selected frozen refs remains the
qualification step; its external artifacts must be reviewed for exact output
digests before any timing ratio is accepted.
