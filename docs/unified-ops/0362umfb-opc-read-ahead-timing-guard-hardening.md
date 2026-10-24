# 0362umfb — OPC read-ahead timing guard hardening

This follow-up hardens the timing-only driver after the independent review of
`0362umf`.  It contains no Cargo build, native correctness run, ABBA capture,
allocator run, or performance result.  The successful build/correctness prefix
and its external receipt remain unchanged.

## Frozen implementation

The reviewed driver is:

```text
docs/unified-ops/results/0362umf/opc_read_ahead_timing.py
bytes: 57434
sha256: 0234ba04fdb7cebaaefc98b9681b2e8ce42d1ee928f856bc2bea02d2ce695318
```

The guard now:

- enables a Linux child subreaper and requires every child to verify its
  parent, create its own session/process group, and install `PDEATHSIG`; it
  retains start-time identities across reparenting and fails closed if group
  cleanup cannot prove that all owned processes are gone;
- checks the deadline immediately after each polling wait, before accepting a
  normal child exit, and reports capture-thread join or capture-I/O failures;
- treats normal `/proc` PID disappearance as expected, but fails closed on
  unexpected stat, identity, or RSS errors.  Kernel-thread/zombie missing RSS
  is the only explicit zero-RSS exception.  Restricted `exe`/`cwd` metadata is
  retained only after readable `comm` and `cmdline` prove the process is not in
  the guarded heavy set;
- rejects symlinked target entries and any target growth, checks the target
  cap throughout the child, and bounds each host snapshot/JSONL stream, event,
  and final receipt;
- binds each binary by byte count and SHA-256 before the ABBA sequence and
  before and after every leg, while retaining the source HEAD/tree/status
  fence before and after each leg;
- rejects a busy pre-run host using an absolute load-1 ceiling (default `1.0`)
  and a quiet interval.  During a child, the selected CPU is accounted with
  `/proc/stat` interval ticks minus the monitor and owned-process ticks, so the
  child workload itself is allowed while unrelated selected-CPU work is not.
  This is interval accounting, not a claim of PSR-only migration coverage.

The driver still requires a fresh external output root and an exclusive CPU
lock.  Its result remains latency-only; allocator qualification is outside the
scope of this lane.

Static checks for this return were:

```text
python3 -m py_compile docs/unified-ops/results/0362umf/opc_read_ahead_timing.py
python3 docs/unified-ops/results/0362umf/opc_read_ahead_timing.py --help
git diff --check
```

No timing command was started because the host was concurrently running
unrelated Cargo/rustc work and the absolute pre-run quiet fence correctly
rejects that state.
