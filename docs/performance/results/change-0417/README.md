# Change 0417 representative CRUD baseline

This bundle measures the 30 existing representative selectors in the CRUD
coverage index at clean revision
`b10d6c25a13242ca260a8c897946f4d80ae06c61`. It is descriptive evidence from
one revision. It does not establish a speedup or completion of the non-iWork
performance program.

## Protocol and scope

The protocol was declared before capture in [protocol.json](protocol.json).
Each selector runs in two fresh normal processes, with 20 warmups and 500
retained samples per process. Separate allocator processes use 3 warmups and
30 retained samples. The first repeat follows [matrix.json](matrix.json);
the second reverses that order. All processes use CPU 2 and the one-worker
configuration. Warmups and samples share a process and generated in-memory
corpus. These are not cold-file, remote-source, concurrent, or scaling results.
Task-owned measurement, builds, tests and profiling are serialized; unrelated
background activity on the shared KVM host is uncontrolled.

Read [timing-boundaries.md](timing-boundaries.md) before interpreting a row.
Some selectors time an already-open query or commit only; others time a
specific authoring or publication lifecycle. The boundaries differ across
selectors. Full-output `CountingSink` retention and untimed expected artifacts
prevent interpreting a write-call/window limit as a whole-process memory bound.
The one-worker setting also does not exclude untimed branch-preparation threads.

Normal and allocator elapsed values are separate evidence streams. Allocation
attribution is emitted only by RTF streaming creation and DOCX story hyperlink
redaction; absence in the other 28 selectors is unavailable, never zero.
DOCX allocation includes package cleanup after its elapsed timer stops.
Peak allocator snapshots and `/usr/bin/time -v` RSS have their recorded scopes;
neither is an operation-local peak heap measurement.

The index still has 15 categories, with dynamic calculation/refresh unsupported,
and 30 selectors across the other 14 categories. Its 10 measured and 20
correctness-only statuses concern the existing default/full-run acceptance
contract. This separate opt-in baseline does not promote statuses or modify
the 425-selector registry or the 36-case / 198-row default matrix.

## Evidence and replay

- [build-identity.json](build-identity.json) binds the clean source revision,
  compiler flags, successful build command and both binary hashes.
- `runtime-source-identity.json` binds the runtime harness, locks and relevant
  production source files to that measured revision; `tool-source-identity.json`
  independently binds the committed replay helpers.
- [environment.json](environment.json) records the machine and toolchain.
- [capture.json](capture.json) retains exact commands, order, timestamps and
  exit codes. `normal/` and `allocator/` retain original reports, corpus catalogs
  and whole-process time/RSS logs; `preflight/` retains the one-sample checks.
- `inputs/` preserves the original taxonomy and index used to declare the
  matrix. That original index incorrectly admitted `medium` for XLS validation;
  the checked-in index and its validator now correctly admit `tiny` and `large`.
  The captured selector used `large`, so this correction does not change its
  input or measured source revision.
- `scripts/` retains the actual build, capture and profile launch scripts.
  Their absolute paths record this capture environment; rebuilding requires
  equivalent paths or deliberate path substitution in a new capture.

Raw `environment.rustc_version` reports 1.95.0 because the harness invokes
`rustc --version` at run time from the pinned checkout (`lib.rs:54889`). This
field describes the runtime command, not the compiler that built the binary.
The explicit Cargo 1.98.1 build and `RUSTUP_TOOLCHAIN=1.98.1` are retained in
the build journal; the separately captured verbose toolchain identity is in
`environment.json`. The original report fields are preserved.

Both normal and allocator ELF `.comment` sections confirm the build compiler
in `checks/*compiler-comment.txt`. Replay helper source hashes and their commit
are bound in `tool-source-identity.json`.

From the repository root, replay report/corpus identities, phase/order/argument
contracts, raw statistics and the retained summary with:

```sh
python3 docs/performance/results/change-0417/verify.py
```

The verifier supports a copied bundle and requires the repository's Python
validation tools. It does not require the historical temporary binaries or
worktree. After those temporary artifacts were removed, `negative-probes.py`
passed two isolated-export replays and rejected all 12 corruptions. Byte
integrity is independently checked with:

```sh
cd docs/performance/results/change-0417
sha256sum -c SHA256SUMS
```

The statistical summary retains both repeats, integer-nanosecond midpoint p50,
nearest-rank p95/p99, IID median order-statistic intervals and all absolute
repeat drift above 5%. IID intervals do not account for shared-host drift;
the repeat observations remain separate. No allocator latency ratio or causal
comparison between different selectors or revisions is accepted.

## CPU diagnostic

The separate 20-sample / 3-warmup PPTX profile has 33,347 `cycles:u` stacks,
zero lost samples, and 0.479% unresolved leaf weight. Deflate contributes
71.929% of whole-command leaf weight; SHA-256 contributes 20.66%. These
weights include setup, preflight, warmups, verification and metadata-reporting
descendants. They are not elapsed-time percentages for the timed phase sum.
The separate multiplexed PMU run has roughly 83% scheduling coverage,
whole-command IPC 2.384 and branch-miss rate 2.301%. L1 alias zeroes are
unvalidated and LLC events unsupported; neither provides cache-miss evidence.

`profile/` retains original command journals, report/catalog pairs, counter CSV,
readable self/inclusive reports, compressed `perf.data` and symbolized stacks.
Profile metadata identifies perf 7.0.14; the earlier environment snapshot recorded
7.0.12. Both observations remain visible. Regenerate the period attribution
from retained symbolized stacks without the old binary:

```sh
python3 docs/performance/results/change-0417/profile/summarize.py
```

This replay needs `zstd`. Regenerating symbolized views from raw `record.data.zst`
also requires rebuilding the exact profiled ELF at the recorded path (or mapping
it through perf's symbol lookup options); the original temporary ELF is removed
after capture. No allocation or speedup inference follows from the CPU profile.
