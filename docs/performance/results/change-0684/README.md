# Change 0684 evidence packet

Retained with scoped experimental evidence; `performance_claim: none`.
See [the change record](../../0684-xls-occurrence-query-index.md) for every
regression disposition and [independent review](review.md).
Baseline: `5805d54a1d8435e1bee96cde1303310703440d16`.

The candidate retains bounded occurrence locators after a complete successful
worksheet scan. It preserves target SST/formula resolution and source/execution
fences. Warm large-sheet gains coexist with construction cost, higher retained
memory, tiny file-backed overhead and repeated-refusal overhead. The initial
candidate's first-query regression was removed by ordinary-scan specialization;
its full measurements/source remain in `initial-candidate/`.

## Bindings and boundaries

`baseline.json` binds the goal and all unchanged accepted ADR files.
`corpus-manifest.json` binds all 126 XLS fixtures. Both source-mode corpus
reports match exactly: 119 opens, seven refusals, 349 successful walks,
15 refused walks and 4,027 selected/repeated checks per mode with no mismatch.
The corpus probe checks admitted visitor results against selected queries;
focused tests cover target-specific refused paths the visitor cannot admit.

`environment.json` records Rust 1.95.0, EPYC 9R45, Linux, CPU affinity and the
shared-host limitation. This batch serializes owned builds/measurements and
pins probes to CPU 12; unrelated host tasks can still run. File-backed reads
use warm OS caches, not a controlled physical cold-cache experiment.

`measurements/` contains raw source/probe/binary-bound corpus, logical I/O,
allocation and native timing reports. Baseline A/A and candidate A/B/B/A each
use 30 fresh-owner route samples after three warmups: 6,480 native route
records total. `comparison.json` retains all phase p50/mean/p95/p99, controls,
and 384 allocation observations across 64 groups. p99 is the sample maximum.
Each allocation group has three exactly agreeing repeats before and after.

Native timing uses core OwnedSource/FileSource. Counted-source timings are
only diagnostic. Input loading/source construction precedes open timing;
formatting/semantic projection is outside query timers. Never interpret
`total_elapsed_ns` as API latency. The allocation probe retains owner and
result through gauges and uses the same coordinate for both preparations;
the native prepared route's second query uses a separate selected coordinate.
Warm allocation deltas exclude the index already retained before the interval.

`diagnostics/` uses 10/1,010-query native processes, three repeats, perf stat
and process peak RSS; subtraction estimates extra-query hardware work.
Low-count negative estimates are preserved but receive no percentage claim.
`profiles-manifest.json` binds retained text profiles; baseline 10,000 and
candidate 1,000,000 iterations produce sufficient samples. Unequal iteration
counts preclude direct absolute profile-cycle comparison. Temporary perf.data
and build trees are deliberately not committed.

`budgets/` holds 24 candidate-only 0/1 byte/1 MiB/2 MiB controls and 120 outcomes.
Read counters demonstrate actual admission/fallback; wrapper timing is not a
native performance claim. A 1 MiB ceiling fails to admit 54016's geometrically
grown index and retries collection; zero disables the attempted builds.

## Reproduction

All four standalone probe manifests and lockfiles are retained. Recreate a
baseline worktree at the revision above, copy `probe`, `allocation-probe` and
`repeat-probe` to the same repository-relative paths, then build each baseline
and candidate manifest with separate target directories:

```sh
CARGO_BUILD_JOBS=2 CARGO_TARGET_DIR=/path/to/target \
  cargo build --release --offline --locked --manifest-path /path/to/probe/Cargo.toml
```

Offline builds require the locked dependencies to be locally available; omit
`--offline` on a fresh machine. The baseline owner build additionally used the
ignored workspace Cargo.lock bound by `workspace-lock.sha256`; standalone
measurement builds use their retained individual lockfiles.

Drivers accept explicit binaries but use recorded local baseline paths for
source bindings; adapt those paths when reproducing elsewhere. Choose an
available pinned CPU consistently in the drivers on other hosts.

```sh
python3 run-measurements.py baseline BEFORE AFTER ALLOC_BEFORE ALLOC_AFTER
python3 run-measurements.py candidate BEFORE AFTER ALLOC_BEFORE ALLOC_AFTER
python3 run-diagnostics.py baseline REPEAT_BEFORE
python3 run-diagnostics.py candidate REPEAT_AFTER
python3 run-budgets.py BUDGET_CANDIDATE
LITCHI_GATE_OUTPUT=docs/performance/results/change-0684/final-verified \
  python3 run-integration.py
python3 audit.py
python3 audit-diagnostics.py
```

Run commands from the repository root using each driver's full packet path.
Use a separate copy/output location to avoid overwriting retained evidence.
`audit.py` verifies live candidate source, immutable baseline git objects,
probe/corpus/raw bindings, quality, initial archive, budget controls, exact
semantic parity, and available binaries before cleanup. After cleanup, binaries
can be rebuilt from the bound sources; their hashes remain recorded.

## Validation and scope

`final-verified/` records six passing gates and source hashes: formatting,
all-feature/all-target XLS check, warning-denied Clippy, 1,455 owner tests
(one existing ignored), 61 facade tests, and warning-denied rustdoc. Earlier
failed check/focused-test/quality logs document resolved issues and are
superseded by these final gates. The small budget probe's initial format-string compile
error is likewise superseded by its successful build and 24 completed controls.

`evidence/` retains strict/structural claims, 50 claim-verifier tests, report
classification, CRUD coverage, non-iWork manifest and crate-boundary checks.
No registry claim or coverage promotion is made. OLE2/OOXML work remains active;
iWork is excluded. Physical cold-cache, remote latency, shared-budget throughput
scaling and cross-platform behavior remain outside this packet's claims.
