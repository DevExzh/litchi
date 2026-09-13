# 0548: OLE2 checkpoint collector rejected

The bounded checkpoint candidate fails the frozen admission gates. Production
is restored to the exact measured baseline; the public corruption benchmark,
candidate snapshot, tests and complete paired evidence are retained. No runtime
speedup is admitted by this batch.

| Primary XLS workflow | Repeat 1 p50 change | Repeat 2 p50 change |
| --- | ---: | ---: |
| `xls_source_backed_open` | -7.87% | -3.33% |
| `xls_source_backed_open_one_cell` | +2.12% | -3.70% |
| `xls_owned_source_open` | -7.97% | -3.13% |
| `xls_owned_source_open_one_cell` | -5.26% | -3.78% |

Negative values are improvements. Every primary case must improve at least 3%
in both repeats. Source-backed open-and-read-one-cell instead regresses 2.12%
in repeat 1. Seven of eight rows pass, which does not satisfy the all-row gate.

Independent profiles also reject adoption: XLS owned-source constructor
inclusive Ir rises 2.865–2.891%; collector self Ir rises 5.855% for XLS and
5.871% for CFB few-large in both repeats. These counts are scoped to five
positive timed constructor dumps per job, excluding setup and termination.
They do not establish operation-local hardware cycle or cache behavior.

Allocation calls, allocated bytes and incremental region peak pass the separate
gates. All 12,800 public malformed-guard samples match their exact oracle.
Invalid p50 and mean remain within the frozen envelopes: maximum 2.234x the
same-invalid baseline and 1.312x baseline-valid, below limits of 4x and 2x.
Passing envelopes do not erase the individually retained rejection regressions.

The common guard covers valid, root self-cycle, prefix cycle, later cycle,
early end, reserved marker, invalid index and excess-chain cases at 128 and
16,384 sectors. It validates the generated file independently before mutation.
Only `OleFile::open` is timed; oracle checks and returned-value destruction are
outside the clock. The guard has no allocation instrumentation.

The Rust candidate retains preflight, reservation order, labels and visited
zero-fill. Scalar Brent checkpoints guard speculative traversal; suspicious
chains replay the original visited-bit walk using prepared scratch. Existing
formatted error strings can still allocate. The exhaustive Rust differential
test covers 376,264 table/start/count combinations per feature configuration,
including exact ordered errors, scratch reuse/reset and checked counter limits.
Those candidate-only tests remain in the unapplied source snapshot after rejection.

The campaign retains 48,000 main native samples, 1,440 allocator samples,
16 profile children with 80 timed constructor dumps and 12 CFB setup dumps,
four whole-child hardware diagnostics, 64 measured guard children and 32
single-sample smoke children. Smoke and instrumented elapsed samples never
enter native admission. All measurement children run serially. Other work on
this shared host is not controlled; host compiler observations and all repeat
variation remain explicit. No stable-tail or broad provider/scaling claim follows.

Two guard compilation failures occurred before any measurement. Both source
versions and command/output receipts are retained in `initial-guard-build` and
`second-guard-build`. Warning-denied Clippy and all smoke oracles then passed
before each main build. Analysis inventories and execution scope are explained
in [analysis execution note](analysis-execution-note.md).

See [decision](decision.json), [quality summary](quality-summary.json),
[individual result review](results-review.md), [protocol](protocol.md),
[source review](source-review.md), and [ADR matrix](adr-compliance.md).
The final source manifest must exactly equal the measured baseline manifest.
All retained binaries are hash-checked before the sole owned target is removed;
strict sealed replay checks source, receipts, analysis, decision and cleanup.

Cargo-fuzz and a nightly toolchain were unavailable; no sanitizer campaign is
claimed. Historical 0546’s 1,306 tests were the XLSX owner suite, not workspace
tests. This batch reports actual OLE owner tests separately from workspace check.

OLE2 and OOXML performance remain the active priority. ODF is deferred until
that optimization goal completes; iWork is excluded. The next collector
experiment must reduce checkpoint state work under the same bounds and exact
error contract, and earn fresh matched admission. The broader CRUD, provider,
memory and scaling goals remain open.

[Next bounded opportunity](next-opportunity.md) also considers a checked combined
membership/mark operation, subject to proving that emitted work actually falls.

Final validation passes all 15 declared checks with **4,382 executed tests**: CFB
310 in each feature configuration, XLS 1,345, DOC 1,187 and PPT 1,230.
Workspace check, formatting, warning-denied Clippy/rustdoc, crate boundaries
and strict performance-claim registry checks pass. Candidate validation is
separately retained: 15 checks and 4,386 tests, including its two additional
Rust test functions in both CFB configurations.
