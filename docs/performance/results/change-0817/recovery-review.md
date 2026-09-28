# 0817 quality recovery review

This review covers the quality-attempt-1 recovery only. It does not amend
`protocol-review.md`, change the ordinary-save measurement plan, or authorize
any workload, profiler, binary, or Cargo action by this reviewer.

## Failure boundary

Quality attempt 1 passed formatting and all-feature/all-target checking. Its
all-feature test gate then recorded 640 passed tests and one ignored test in 26
successful suite summaries before the final integration suite failed. The
failure is retained in:

- `quality-1/02.log` (SHA-256
  `5b9a20d22afc72f938385c9e57ba9e74d2dfe39686384b184cef84fb5bed24fe`);
- `quality-1/checks.json` (SHA-256
  `04bf86e767819591ceab1dd461a3f0cb8312327ece185a396777f99d2fa17171`); and
- `quality-1/failure.json` (SHA-256
  `651890ba3c17df914460ba0b2c5c784aed950e36409ccaa1452c7ab56a5aa`).

The failing test was
`xlsx_source_cell_values_planning_allocations_are_scoped_and_aligned`. Under
the all-feature command, the normal report correctly identified
`ordinary_save_procfs_operation_scoped`, while the old assertion expected
`none`. This is a test expectation defect, not a production or runtime
harness failure. No ordinary-save workload or performance capture was run.

## Test-only repair

The proposed change is bounded to
`tools/perf-baseline/tests/xlsx_planning_allocations.rs`. The quality-1 frozen
copy has SHA-256
`c947077f9517a1d39003b4234e0eeff106652920b2aa615ab48884948bb88cf8`; the
repaired working file has SHA-256
`60b75c8f38d247fecb153e1ae87d569d18a79c15ea13f72d4550779cf57eac3d`.

It adds feature-conditional exact identities:

| Build features | Normal binary | Allocator binary |
| --- | --- | --- |
| `allocator-metrics` only | `none` | `system_allocator_operation_scoped` |
| `allocator-metrics,ordinary-save-process-metrics` | `ordinary_save_procfs_operation_scoped` | `ordinary_save_procfs_and_system_allocator_operation_scoped` |

The repair changes only those identity expectations. It retains the binary
identity checks, sample and warmup counts, elapsed/sample-order alignment,
allocation-vector presence and scopes, split-region reconstruction, live/peak
invariants, corpus/output/semantic identity, and cleanup. The labels match the
existing feature-gated implementation: the normal binary does not enable the
allocator wrapper, while the combined observer reports both procfs and system
allocator instrumentation.

The repair is admitted only after both focused variants pass:

```text
cargo test --offline --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --all-features --test xlsx_planning_allocations -- --test-threads=2

cargo test --offline --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --features allocator-metrics --test xlsx_planning_allocations -- --test-threads=2
```

The all-feature run exercises the failed branch. The allocator-only run
exercises the `none`/system-allocator branch and prevents the conditional from
being correct only because all features happen to be enabled.

## Frozen quality-1 archive

Before amending `plan.json` or any recovery driver, preserve an immutable
quality-1 archive. It must include the complete `quality-1/` directory,
including `02.log`, `checks.json`, `failure.json`, `source.json`,
`frozen-inputs.json`, and the full `input-snapshot/`. Record a sorted file
manifest containing relative path, byte length, and SHA-256 for every archived
file. Preserve the failed gate and its exit code; do not rewrite it as a
passing quality attempt.

The archive must also bind the pre-repair test source and the old plan/packet
hashes. The quality-1 input snapshot is the authority for the source state at
the failed run. Any amended plan gets a new hash and must reference the old
plan hash and archive manifest. The archived
`quality-1/input-snapshot/protocol-review.md` remains byte-identical and is the
immutable pre-recovery witness. The current `protocol-review.md` may carry a
separately reviewed recovery amendment after that archive is sealed; record
both old and current hashes and do not treat the current amendment as runtime
source evidence.

Do not use a live file in place of a missing quality-1 artifact. If an archive
member is missing or its digest cannot be verified, stop recovery and retain a
typed incomplete-archive result.

## Source-difference gate

The carry-forward claim is valid only if the archived quality-1 snapshot and
current workspace differ in exactly one runtime-relevant file:

`tools/perf-baseline/tests/xlsx_planning_allocations.rs`.

The recovery must prove byte/hash equality for every production file,
`tools/perf-baseline/src/**`, binary source, workspace manifest, both lock
files, corpus input, host/cgroup witness, and normative packet input outside
that named test. Packet or recovery-driver amendments are recorded as packet
changes and are not evidence that runtime suites are unchanged. Any additional
runtime or test-source difference invalidates selective reuse and requires a
fresh full all-feature test gate.

The source proof must be machine-readable and list the allowed old/new test
hash pair. It must also show that the quality-1 command used the same offline
locked harness graph and the same relevant environment contract.

## Bounded quality-2 composition

Quality attempt 2 may carry forward the 640 successful and one ignored
quality-1 test results only as explicitly labeled prior evidence after the
source-difference gate passes. It must not claim that the old failed command
passed. The carry-forward manifest should enumerate the 26 successful suite
summaries from `quality-1/02.log`, retain their old log/receipt digests, and
point to the source snapshot that proves each suite's runtime inputs were
unchanged.

The repaired integration test and documentation tests are fresh evidence:

1. run the corrected `xlsx_planning_allocations` integration test with all
   features;
2. run the same integration test with `allocator-metrics` only; and
3. run all-feature documentation tests with the exact offline locked manifest.

Record each command, target directory, environment, exit code, test summary,
log descriptor, and source/lock witness. The corrected integration result must
be counted separately from the carried-forward 640 results until the recovery
receipt assembles the final quality status.

Quality attempt 2 runs fresh formatting and all-feature/all-target checking
before the recovery test gate, then fresh warnings-denied all-feature/all-target
Clippy, warnings-denied all-feature documentation, and the crate-boundary
checker. The command rows and environment bind this order.
The final quality receipt must state which evidence was carried forward and
which gates were freshly executed. It may be accepted only when the fresh
targeted tests, documentation tests, and all five fresh non-test gates pass.

Avoid rerunning the approximately 13-minute untouched test suites only when
the exact source-difference proof and per-suite carry-forward manifest pass.
If either proof fails, rerun the complete all-feature test command and retain
the new failure or pass as a normal quality attempt.

## Recovery acceptance

Recovery is complete only when all of the following are retained:

- immutable quality-1 archive and failure receipt;
- exact one-file source-difference proof;
- passing all-feature and allocator-only corrected integration runs;
- passing all-feature documentation tests;
- fresh format, check, Clippy, rustdoc, and boundary gates;
- old and amended plan hashes with the recovery relation recorded; and
- a final quality receipt that excludes historical timing, optimization, or
  workload claims.

This repair restores test-label correctness for feature combinations. It does
not alter production code, ordinary-save runtime code, corpus admission,
timing boundaries, allocation semantics, or the 0817 sample plan.
