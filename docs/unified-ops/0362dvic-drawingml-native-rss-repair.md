# 0362dvic — independent DrawingML native-child resource capture

- Date: 2026-10-24.
- Status: source/protocol repair committed; no Cargo build, native probe, or
  timing qualification is claimed.
- Scope: close the external-process resource-evidence gap in the paired
  DrawingML raw/owner protocol from 0362dvi.

## Corrected plan identity

The leaf initially bound the following external plan path and SHA-256.
Final composition supersedes this pin with the 135637Z / 662a9159 plan
recorded in [0362dvid](0362dvid-drawingml-protocol-composition.md):

```text
(deleted-coord-artifact)/drawingml-additional-qualification-plan-20261024T125856Z.json
cf5208bc30eb1b650cfc7f2c15a53d6ae4d0cb3c43999236eb83bb9e58cd4eeb
```

This supersedes the earlier `0b8e1744d1`/`d671c2db…` plan identity in the
preparation draft. A changed path or digest fails before a build or ABBA run,
and both values are retained in the top-level receipts.

## Independent child observation

The shared `SafeRunner` has one optional, backwards-compatible observer hook.
Its default is inert, so established W4b, W4, and W5 lanes do not intercept
or measure helper processes. The paired DrawingML runner supplies the observer
only for ABBA and receives the direct process object returned by its one
`Popen` call. No global `Popen` patch or child discovery scan is used.

For every direct probe child, the observer writes a fresh
`external-process-resource.jsonl` row after the child exits. The row contains:

- PID plus the kernel boot ID and `/proc/<pid>/stat` start-time ticks, so PID
  reuse is rejected;
- an independently computed NUL-joined argv SHA-256, cwd, and the source
  root/revision/identity reference from the runner event;
- sampled `VmRSS`, peak sampled RSS, and kernel `VmHWM`, with a positive sample
  count; and
- the child return code and a captured/unavailable status.

The journal is newline-delimited, fsynced, and bound into the generic ABBA
event by its exact line hash and byte count. The validator independently
recomputes the argv hash, checks the start identity, source and cwd bindings,
event identity, return code, sidecar line pointer, positive resource values,
and coverage of the probe's self-reported `VmHWM`. A qualified report without
all of these checks is rejected. A very short child may produce an explicit
`unavailable` row; that preserves the runner outcome but cannot become a
qualification claim.

The validator also requires the sidecar event set to be exactly the qualified
ABBA probe set. The generic event journal still records the validator control
process and fixture-preflight controls, but the external observer deliberately
excludes explicit `event_kind=control`/`measured=false` events plus the exact
validator and fixture-preflight IDs. This keeps the final sidecar hash
unchanged across control handoff and leaves its event set exactly equal to the
qualified probe set.

## Focused checks

These checks are source-only and do not invoke Cargo, Rust, a native probe, or
a timing workload:

```text
python3 -m py_compile tools/native_qualification_runner.py tools/native_qualification/drawingml_raw_paired_runner.py tools/native_qualification/drawingml_raw_validator.py tools/test_native_qualification_runner.py tools/test_drawingml_raw_qualification.py
python3 tools/test_drawingml_raw_qualification.py       # 9 passed
python3 tools/test_native_qualification_runner.py        # 23 passed
```

The DrawingML tests include mutation cases for argv identity, source binding,
and the sidecar line pointer, plus validator/control handoff and a
deterministic disappeared-PID/unavailable-child control. The unavailable-child
test simulates loss before start-identity capture; it is not a rapid-subprocess
measurement.
A separate Python child smoke check captured a positive independent RSS/VmHWM
sample through the observer. No performance or native acceptance result is
inferred from these checks.
