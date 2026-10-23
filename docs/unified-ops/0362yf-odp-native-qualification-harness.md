# 0362yf — matched native ODP qualification harness

This slice supplies the missing native qualification infrastructure for the
ODP source-backed owner after the correctness and source-owner work recorded
by 0362y–0362ye. It is source and runner infrastructure only; no release
binary, timing result, allocator capture, Callgrind file, or performance claim
is included.

## Probe contract

`tools/perf-baseline/src/bin/odp_native_qualification_probe.rs` is the current
candidate's parameterized probe. The companion
`tools/perf-baseline/odp_native_qualification_e4a_probe.rs` is copied into the
pinned e4a checkout and exposes the same external-fixture command-line
protocol. The fixture, appended title/body, route, batch size and sample count
all come from command-line arguments. The genuinely shared timed routes are:

| Route | Timed lifecycle |
|---|---|
| `open` | open a fresh positional source-backed presentation |
| `read` | open, read the slide count and complete visible text |
| `write` | open and write the unchanged source-backed presentation |
| `append` | open, commit one bounded tail append, and write the deferred owner |

The current candidate also exposes `replay` for its untimed correctness proof:
it creates the append patch, opens an equal independent source, replays it by
content, and writes it. `replay` is deliberately not in the shared timing
route set because the e4a owner predates `Patch` and can only repeat an
independent `Edit::plan` publication.

`--mode verify` is untimed and must run before calibration. Both probes check
exact source/no-op bytes, semantic reading, append change, the plan/patch's
source `content.xml` hash against an independently read fixture member, and
appended output change. The legacy report has the separate
`litchi.odp.native-qualification.legacy-probe.v1` schema and explicitly marks
patch, inverse and scoped replay as unavailable; its `independent_append`
output is never called replay. The current report uses the
`litchi.odp.native-qualification.probe.v1` schema and additionally checks
same-owner replay, equal-content cross-copy replay, inverse behavior, changed
recipient refusal, and the scoped recipient with one changed unrelated archive
member. Applying the inverse directly to the committed owner must restore the
complete source archive byte-for-byte, while reopening the published output as
an independent recipient must restore `content.xml`. The reopened publication
may have different ZIP entry metadata because the replayed member is newly
published; that proof therefore compares the content scope and keeps opaque
members separate. Every output is reported as a byte count and SHA-256 digest.
Calibration and timed reports carry a route proof matching this preflight, so
timing validation does not rely on output length alone.

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
time the runner resolves both refs, binds each clean worktree `HEAD` to that
immutable commit, hashes both executables, verifies the archive and
`content.xml` fixture identities, and records all resolved identities.
No source-specific historical hash is compiled into the probe or runner.

The runner performs the following sequence:

1. Run both endpoint correctness preflights. Require identical exact output
   digests for the shared `source`, `no_op` and `append` publications,
   semantic-read receipts and source/target content proofs. Candidate-only
   patch/inverse/scoped flags are validated strictly but are not fabricated for
   the legacy endpoint.
2. Calibrate every route independently for both endpoints.
3. Run calibrated `A1 control, B1 candidate, B2 candidate, A2 control`
   blocks, pinned with `taskset` to the manifest CPU.
4. Before calibration and every ABBA leg, fail closed if `/proc/loadavg`
   exceeds the manifest ceiling or a compiler/profiler/benchmark process is
   present; inaccessible `/proc` reads are failures. Each timed child must
   return a positive batch no shorter than the manifest minimum and must repeat
   its exact route proof.
5. Write raw child reports, correctness receipts, calibration, per-block
   checkpoints and the final report atomically under the caller's external
   artifact directory. Existing non-empty output directories are rejected.

Target and artifact directories below either source worktree are rejected.
Fixture, worktree `HEAD`/ref, executable size/hash, and source cleanliness are
revalidated after calibration and after the ABBA routes. Filesystem, hashing,
JSON, and directory failures produce a structured `failure.json` receipt. The
runner invokes children without a shell and supplies an external `TMPDIR`; no
Cargo command is run by the runner.

An execution is prepared with:

```text
python3 tools/odp_native_qualification.py \
  --manifest /external/odp-qualification/input.json \
  --out-dir /external/odp-qualification/run
```

For an untimed correctness receipt without calibration or native timing, use
the same frozen manifest with `--correctness-only`. This mode does not reject
unrelated host work; strict quiet-host guards remain required for calibration
and ABBA timing. It writes the same endpoint-level correctness reports and a
`report.json` whose claim scope is explicitly untimed.

The manifest is intentionally frozen at execution time. Each endpoint must
declare a protocol profile: `legacy-e4a-plan` for the baseline and
`current-commit-patch` for the candidate. The runner maps those profiles to
the two probe schemas and rejects any other baseline/candidate pairing. Its fixture
`provenance` field must identify the tracked test-data source or externally
approved fixture source; the runner records that text beside the verified
digest. The `--plan` mode validates the manifest shape and writes the complete
shared-route/protocol plan without touching binaries or timing. The plan does
not include replay as a timed route.

## Validation

The runner's standard-library contract tests pass:

```text
python3 -m py_compile tools/odp_native_qualification.py tools/test_odp_native_qualification.py
python3 -m unittest tools/test_odp_native_qualification.py  # 15 tests
```

No Cargo/build/timing slot was used for this source-only harness commit. An
execution against the owner's selected frozen refs remains the qualification
step; its external artifacts must be reviewed for exact output digests before
any timing ratio is accepted. The correctness-only mode is available to
validate those proofs without a quiet-host requirement while timing remains
intentionally disabled.
