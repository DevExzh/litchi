# 0362umf — OPC read-ahead timing guard and execution prefix

This finite return repairs the execution boundary for the latency-only OPC
managed read-ahead lane.  It adds a separate fail-closed Python timing driver
and records the build/correctness prefix for the final472b packet.  No Cargo
build, native correctness run, ABBA capture, allocator run, or performance
acceptance result was produced by this commit.

## Frozen build/correctness boundary

The preserved all-in-one recipe remains external and unchanged:

```text
script: /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/
        opc-read-ahead-final472b-prep/run-opc-read-ahead-final472b.sh
sha256: 5b20c3980e6892c5485d99b28dcc455e3e823c0d8190d0cabd87fee51c6a3ddb
```

The independently prepared prefix is:

```text
script: /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/
        opc-read-ahead-final472b-prep/build-correctness-prefix-final472b.sh
sha256: 9729d18cc321e8954305dd49022c913740474df6d11a2520530c2eb07e385266
bytes:  5786
diff:   /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/
        opc-read-ahead-final472b-prep/build-correctness-prefix-final472b.diff
diff sha256: 0bd44e98e06727469d1df861bccd7d97297381764e9777d47d41b4d5efba70f4
```

The prefix is the original recipe's lines 1–124, byte-for-byte, followed only
by a stop comment and `exit 0`.  It therefore performs the two locked release
builds and four correctness invocations, then stops before the original
`host_snapshot`, `run_legs`, and ABBA invocation.  Its exact command and cwd
are `cd` to the packet preparation directory followed by
`./build-correctness-prefix-final472b.sh`; it must use the packet's fresh
external run root and the environment recorded by 0362ume.  The prefix is an
execution input, not evidence that those commands have run.

The endpoints remain the exact clean source pair:

| endpoint | revision | tree | source checkout |
| --- | --- | --- | --- |
| baseline | `410bd99f00fdfb4161a94215710340f16ea536bb` | `41850bc694ccfda6f84510611d32fe452248773d` | `/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/opc-read-ahead-c796-prep/source-baseline-410` |
| candidate | `472b6705475b073e7fe08f6e11ee17fb1e29f838` | `ebc1cee05f124c0c63977aff07a23288c19e7d00` | `/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source` |

## Timing-driver repair

`docs/unified-ops/results/0362umf/opc_read_ahead_timing.py` is the separate
quiet-host ABBA driver (37,727 bytes; SHA-256
`eef0ef3edaf8a914d0243aaf412d4cc1796c715938cb49cce3dad41721a6d4a7`).  It
does not build or mutate either endpoint.  It requires both source paths,
revisions, and tree IDs and verifies `HEAD`, `HEAD^{tree}`, and a clean
`git status` before every child and again after it.  The report validator
checks the complete forward/exact physical range traces, source-read policy,
text/archive identities, source version, budget/disposal fields, and the
latency-only `allocation.sample` absence.

The driver holds an external nonblocking CPU lock for all eight legs.  Before,
during, and after every child it records full `/proc` snapshots and load; it
rejects unrelated Cargo/rustc/rustdoc/perf/valgrind/heaptrack work, selected
CPU use by another process, excessive load, source drift, surviving
descendants, target growth, output/report overflow, RSS over 20 GB, and the
deadline.  Every child is a new process group with selected-CPU affinity,
20-GB address-space and zero-core-dump limits.  Any violation terminates the
whole group, records the event, and fails closed.  The monitor's own PID is
excluded from selected-CPU ownership checks.

After the prefix has actually produced fresh binaries and correctness reports,
the timing command is, with the packet's paths made concrete:

```sh
python3 docs/unified-ops/results/0362umf/opc_read_ahead_timing.py \
  --baseline-binary "$RUN_ROOT/target-baseline/release/litchi-perf-baseline" \
  --candidate-binary "$RUN_ROOT/target-candidate/release/litchi-perf-baseline" \
  --baseline-revision 410bd99f00fdfb4161a94215710340f16ea536bb \
  --candidate-revision 472b6705475b073e7fe08f6e11ee17fb1e29f838 \
  --baseline-tree 41850bc694ccfda6f84510611d32fe452248773d \
  --candidate-tree ebc1cee05f124c0c63977aff07a23288c19e7d00 \
  --baseline-source /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/opc-read-ahead-c796-prep/source-baseline-410 \
  --candidate-source /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source \
  --target-baseline "$RUN_ROOT/target-baseline" \
  --target-candidate "$RUN_ROOT/target-candidate" \
  --output-dir "$RUN_ROOT/timing" \
  --tmpdir "$RUN_ROOT/timing-tmp" \
  --cpu-lock "$RUN_ROOT/opc-read-ahead.cpu.lock" \
  --cpu 60
```

`RUN_ROOT` is the fresh external
`/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/opc-read-ahead-final472b-run`.
The command is not a claim that timing has run.  The final receipt remains
latency-only; allocator qualification is out of scope and must not be
inferred from the ordinary binary.

Static validation in this commit is limited to Python compilation/help and
source-fence checks against both clean endpoints.  Native execution still
requires the coordinator's exclusive quiet-host window, fresh external
outputs, and the locked build/correctness prefix.
