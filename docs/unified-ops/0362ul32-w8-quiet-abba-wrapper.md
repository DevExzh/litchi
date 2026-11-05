# 0362ul32 — W8 File quiet-host ABBA wrapper

This packet turns the V1lU2E W8 handoff into an executable, fail-closed
source-only lifecycle.  It does not build, invoke, time, profile, or modify
the frozen 109-route probe or the 0362h runner.  The native command remains
inside the generated no-argument launcher, so the runner's forbidden-process
guard sees the same command boundary it was reviewed against.

The wrapper is
`docs/unified-ops/results/0362ul32/w8-quiet-abba.py`.  The source-only tests
are `tools/test_w8_quiet_abba.py` and cover a normal guard state, failed and
timeout exits, an escaped descendant, oversized output, fake systemctl
success/failure, missing-control-group and cgroup-read errors, CPU/argv
launcher mutation, exact launcher file limits, a fake no-argument launcher,
escaped-child cleanup, unreadable-cgroup cleanup, analyzer-group cleanup,
separate-session cgroup-child cleanup, transient-scope GC, analyzer exit
and observed outside-scope-child paths, the real user-systemd GC shape, and
the cgroup/deadline command.  They ran as twenty-three
tests with:

```text
python3 -m unittest -v tools.test_w8_quiet_abba
```

The existing registered W8 policy module imports the same test case, so the
current gate invocation `python3 -m unittest tools.test_w8_projection_policy`
executes all twenty-three wrapper tests (40 tests total in that module); no frozen gate
source or native runner is changed.

The wrapper itself is 52,090 bytes with SHA-256
`2d1466a8c2825a97ab226b6178e5f2fe81d90424202069d17e0c012112e97df2`; the
generated no-argument launcher is separately hashed into the run
configuration before execution.

## Quiet-window command packet

All paths below are external scratch paths except the committed scripts.
`V1` must be the already accepted V1lU2E directory; `RUN_ID` must be unique.
`prepare` refuses existing output/TMP/config paths and revalidates the pair
receipt before writing the configuration.  It chooses and records
`min(os.sched_getaffinity(0))`; it does not silently choose another CPU at
run time.

```bash
ROOT=/home/zhuhe/code/litchi-worktrees/unified-ops-api
V1=/home/zhuhe/code/litchi-worktrees/.coord/w8-472b-native.V1lU2E
RUN_ID=20261024T000000Z-<unique>
SCRATCH=/home/zhuhe/code/litchi-worktrees/.coord/w8-quiet-abba.$RUN_ID
mkdir "$SCRATCH"

python3 "$ROOT/docs/unified-ops/results/0362ul32/w8-quiet-abba.py" prepare \
  --root "$ROOT" --v1 "$V1" \
  --config "$SCRATCH/config.json" \
  --abba-output "$SCRATCH/abba" \
  --tmp-dir "$SCRATCH/tmp" \
  --run-id "$RUN_ID"

python3 "$ROOT/docs/unified-ops/results/0362ul32/w8-quiet-abba.py" run \
  --config "$SCRATCH/config.json"
```

The `run` command has no binary paths in its argv: it receives only the
configuration path, and systemd receives only the generated launcher path.
The launcher sets `resource.RLIMIT_FSIZE` to exactly 67,108,864 bytes per
file, `RLIMIT_CORE` to zero, `TMPDIR` to the fresh scratch directory, and
executes the frozen runner with five outer blocks and the recorded CPU.

The systemd command recorded in the receipt is equivalent to:

```text
systemd-run --user --scope --unit=w8-file-abba-<RUN_ID>.scope
  --property=MemoryAccounting=yes
  --property=MemoryMax=20000000000
  --property=MemorySwapMax=0
  --property=TasksMax=4096
  --property=KillMode=control-group --
  /usr/bin/timeout --foreground --signal=TERM --kill-after=60s 43200s <launcher>
```

During the run the wrapper polls the transient unit, retains every observed
status, `MemoryCurrent`, and `MemoryPeak` sample, kills the cgroup on aggregate-output or
deadline breach, records the unit's process identities, and retains the last
unit record even if systemd garbage-collects the scope.  `--scope` is
foreground-waiting on systemd 259, so the command intentionally omits the
incompatible `--wait` option.  The aggregate output
cap is 8 GiB across the entire ABBA output and TMP lifecycle tree, both during
the native runner and after the analyzer.  It rejects any output or lifecycle
file above the exact file limit, including a synthetic oversized
`analysis.json`, nonzero/timeout/deadline exit, missing unit metadata or
`MemoryPeak`, surviving cgroup members or observed descendants, and post-run
forbidden processes.  `systemctl`/ControlGroup/cgroup-read errors are retained
in the receipt and fail closed; the only tolerated post-run systemctl error is
an explicitly recorded unit disappearance after a valid final sample.  The
accepted systemd GC representation is exactly inactive/dead/success with an
empty ControlGroup, `[not set]` MemoryCurrent/MemoryPeak, and omitted
MainPID/ExecMainStatus; other incomplete property sets remain refusals.
Every refusal writes
`tmp/w8-quiet-abba-receipt.json`; successful and failed runs therefore have a
raw receipt.  The final analyzer is invoked only after these checks and the
runner's 109-route/2,400-receipt manifest checks pass.  The native runner and
launcher are owned by the transient systemd cgroup; cleanup kills only the
recorded ControlGroup and its observed members.  The analyzer starts after
that scope exits, in its own process group outside the native cgroup; its
group handle remains owned through final cleanup, and descendants observed
through its parent chain are retained by start-time identity for cleanup.
No unrelated host process or unobserved foreign cgroup is searched or
signalled.

The configuration records the selected CPU and complete runner argv.  The
generated launcher embeds both as literals, and `run` parses and compares the
embedded values against the hashed configuration before launching.  Signal,
exception, deadline, output-cap, and analyzer-timeout paths stop and reap the
owned systemd scope/process group and observed escaped PIDs before writing the
raw receipt.

## Bound inputs

The generated configuration binds the following reviewed identities and
rehashes them before native execution:

```text
before binary  67879576  be7e3dbd9c31403d3bc017c07299f391e2b4b5b8f889107f9e0bfd686ada75c6
after binary   70239360  7c9061725c74c2fc255d33694cbbe1f6de0dd5f750cc8d1017dc82811bef58f6
before endpoint     3193  72080e5313b5ba183a45769d2e78a9116d8b3ea7527f7af8b9aae059f42e2633
after endpoint      3152  500c156c1fe1b9962a94ff0f5f8c03752ca1fdd08b3725438f2811c52466ab63
pair receipt        8949  f59c22c65737495b7adc88a5c0c6286680f92ed4e101009f66ca81e4c3d72275
policy              7609  e56d549c733629ac73492a3511a45a1ed6a4cec97f1a4705ceea4ba9b67d97d2
runner             39866  d8e708a0b0d01461e3c0d0597b312c7c514ea9369119c99ad0d67847f3296283
analyzer           30190  5e0f6ae124f70a1bf07dd0f9fb6b1450445182f77d15bc0a66d716e0bd77e4d6
pair tool          38986  121761a850a3c280cece5de4424ef3557ce1220994421c83c82e030ecbcca43a
f2 lock           139441  1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c
472b lock         139606  4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1
f2 source archive          6614528000  de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280
472b source archive        6634147840  44c11e933974c8fa3d6a67b46739bedeef866e60fbd4360ffe7f3bb8a4dc2274
```

The endpoint manifests and the explicit original-to-final policy continue to
own all 518 registry identities/checksums/edges and the six reviewed local
edge deltas.  The ordinary strict shared-identity runner path is untouched.

The final collector is the frozen command:

```bash
python3 -B "$ROOT/docs/unified-ops/results/0362h/analyze.py" "$SCRATCH/abba"
```

The wrapper performs this only on a successful guarded runner and records its
stdout/stderr and exit status in the raw receipt.
