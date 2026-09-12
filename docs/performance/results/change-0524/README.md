# 0524: checked CFB chain visitation experiment

This batch tests private checked visited-bit lookup/set fusion in the reusable
CFB allocation validator. The fresh 0523 profile attributes about 40% of XLS
constructor instructions exclusively to chain collection. This is a current
matched experiment; 0523 is motivation, not a substituted control binary.

The baseline is revision `477281a2f83c3256bf3cc06fbc5c57724c9b6bbc` with its
existing CFB allocator instrumentation. Frozen `plan.json` declares the
hypothesis, matrix, order and admission rule. Before editing source, all 30
previously read accepted ADR/index hashes were revalidated unchanged.

## Reproduction and custody

Use fresh evidence and owned scratch/target paths for a rerun. The driver
refuses existing stage directories, receipts and logs. Two-job nonincremental
standalone release builds use Rust 1.95.0 and the harness's ordinary release
profile; root workspace LTO settings do not apply. CPU 2 is pinned; the host
is shared. `host.json` retains current environment/tool observations.

```sh
python3 -B docs/performance/results/change-0524/run.py freeze
python3 -B docs/performance/results/change-0524/run.py build-normal
python3 -B docs/performance/results/change-0524/run.py native --repeat 1
python3 -B docs/performance/results/change-0524/run.py profile
python3 -B docs/performance/results/change-0524/run.py hardware
python3 -B docs/performance/results/change-0524/run.py build-alloc
python3 -B docs/performance/results/change-0524/run.py alloc
# Apply candidate.patch and tests.patch, format, then test before freezing.
rustfmt --edition 2024 crates/litchi-cfb/src/file.rs
python3 -B docs/performance/results/change-0524/preflight.py
python3 -B docs/performance/results/change-0524/run.py freeze --stage candidate
python3 -B docs/performance/results/change-0524/run.py build-normal --stage candidate
python3 -B docs/performance/results/change-0524/run.py native --stage candidate
python3 -B docs/performance/results/change-0524/run.py native --stage baseline --execution-stage candidate --repeat 2
python3 -B docs/performance/results/change-0524/run.py profile --stage candidate
python3 -B docs/performance/results/change-0524/run.py hardware --stage candidate
python3 -B docs/performance/results/change-0524/run.py build-alloc --stage candidate
python3 -B docs/performance/results/change-0524/run.py alloc --stage candidate
python3 -B docs/performance/results/change-0524/checks.py
```

Native blocks follow A1/B1/B2/A2. The last control block executes the retained
baseline binary while the workspace contains frozen candidate source. Every
receipt therefore records both the binary's source manifest and the execution
workspace manifest. Before/after checks bind that workspace and actual binary;
private-index replay reconstructs each compiled source from the base commit
and retained stage patch. No historical report is modified.

## Boundaries and admission

Each stage has 24,000 native durations: nine fixed opaque-heavy XLS workflows
and CFB tiny/many-small/few-large, two children per group with 20 warmups and
1,000 samples per row. Each child shares a process and materialized input
across samples. Default filesystem/cache configuration does not activate
filesystem or fresh-process-per-sample measurement for these selectors.

Each stage separately retains 720 canonical allocator samples after three
warmups. Allocation elapsed times are excluded from native summaries. Regions
surround the existing operation clock, excluding input fixture construction,
post-operation oracles, report generation and returned object drop. Incremental
peak is region peak minus entry live bytes; absolute live values, deallocation
and whole-child RSS remain separate. CFB source counters are not applicable;
tracked XLS reports prove logical ordinary/opaque read ranges, not physical I/O.

Eight profile children per stage retain five selected constructor calls each.
CFB fixture-generation opens are separate numbered setup dumps. Positive
incoming edges establish timed scope; summaries equal constructor incoming Ir
and self plus direct Ir. XLS profiles end before the selected-cell query.
Collection-off child-call metadata is not a timed-call/allocation count. Raw
Valgrind warnings are retained. Two grouped hardware captures per stage cover
the whole process, including setup, clones, queries, oracles, drops and reports;
100% matched event runtime is required for grouped IPC claims.

The frozen admission rule requires at least 3% lower p50 in both paired repeats
for all four primary XLS source-backed/owned-source open and open-one-cell
workflows, lower constructor instructions in both repeats, no material
allocation or peak growth, and passing correctness/quality gates. Every
matched latency/RSS regression over 5% and every absolute same-build variation
over 5% is retained for review. Native within-child uncertainty is recomputed
from raw vectors; two children do not establish stable tails or cross-host
confidence. Instruction changes alone cannot override a failed native gate.

The candidate preserves bounds, cycle and marker validation, ownership claims,
physical reconciliation, fallible reservations and scratch reset. It adds no
public API, dependency, unsafe code, concurrency or provider behavior. Synthetic
in-memory evidence does not complete native-producer, physical-provider,
cold/range, fuzz, broad CRUD or scaling requirements. OLE2/OOXML remains first;
ODF is deferred until that optimization goal completes, and iWork is excluded.

## Preflight environment correction

The initial unchanged-source CFB library run compiled and passed 273 tests,
including all four new guards, but failed five existing filesystem tests.
Three explicitly reported `QuotaExceeded` on shared `/tmp`; two reported
unexpected filesystem failure states. The complete run is retained under
`preflight-initial/`. Repeating identical source with `TMPDIR` inside the owned
disk-backed target passes all 278 library tests. All subsequent quality gates
use that same explicit temporary directory. The five leftover artifacts from
the initial test process were identified by their test-owned names and receipt
time window, hashed and removed after the process was confirmed absent; the
record is `preflight-environment-adjustment.json`.

Three test-only rustfmt corrections were made before either preflight. The
final stage source patch and manifest bind the formatted compiled source.

## Rejected candidate and final source

The matched primary p50 rule fails. Production is restored, and the final
source retains only the independent collector differential and the existing
sequential-writer temporary-substitution test correction. `final/source.patch`
is the authoritative patch against the base revision; production prefixes
before the private test modules are byte-identical to the base.

The candidate quality run passed both formatting checks, then stopped at the
retained 277-pass/1-failure filesystem test result. The final source receives
the complete fourteen-gate quality matrix. The source and quality transition
is recorded in `quality-adjustment.json`. The temporary identity experiment
is separate filesystem evidence, not a benchmark or a retrospectively observed
inode from the failed Rust test.

After the measured candidate stage, restore its patch, apply the final patch
from this bundle, freeze a fresh final stage, and run `checks.py --stage final`.
Do not suppress or force historical environment-dependent failures when
reproducing a new campaign; retain the actual receipts of that run.

## Final validation

All fourteen final quality gates pass with 4,374 test executions; the valid
candidate preflight adds 278. Both failed runs remain explicitly separate.
Post-cleanup verification passes three source-stage replays, 59 serial
build/test/capture intervals, all eight numerical/profile/hardware and
comparison reports, 160 deterministic annotations, and four semantic negative
vectors. Both owned build paths and owned Python caches are absent. The
recursive evidence inventory is checked again after sealing.
