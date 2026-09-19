# Change 0694 — borrowed MCE element names

performance_claim: none

This packet retains the measurements for
[0694](../../0694-mce-borrowed-element-names.md): a private shared-MCE change
that defers owned expanded-name construction until a lookup needs it.
[Design and ADR constraints](design.md), [code review](code-review.md),
[independent evidence review](evidence-review.md), raw
samples, source bindings and explicit regression triggers accompany the result.
No scenario-coverage or registered-performance claim is promoted.

The baseline is `829bed696`. Native one/no-op/two-edit PPTX timings exclude
source materialization, target derivation, correctness checks and serialization.
The counting allocator is a separate executable. The marker-stripped archive
is a mechanism counterfactual, not a semantic-equivalence oracle. All native
measurements use CPU 12 on a warm shared host, without a quiescence guarantee.
The process invocation elapsed times in the differential oracle are bookkeeping,
not performance evidence; that correctness run overlapped integration tests.

## Reproduction

Use disposable baseline and candidate checkouts with this packet at the same
relative location. The baseline checkout must contain the exact revision above;
the candidate must contain the final two MCE Rust files. Install the recorded
Rust toolchain and dependencies. `workspace-Cargo.lock` retains the ignored
workspace lock used by the gates: copy it to the disposable checkout's root
`Cargo.lock`. Each standalone probe has its own retained lock. Adapt CPU 12 and
machine-specific absolute paths if necessary, and record the changed setup.
Do not overwrite existing measurements: copy the packet to a fresh checkout.

On the baseline sources, run:

```sh
python3 docs/performance/results/change-0694/prepare-control.py
python3 docs/performance/results/change-0694/check-control.py
python3 docs/performance/results/change-0694/build.py baseline
python3 docs/performance/results/change-0694/build-refusal.py baseline
python3 docs/performance/results/change-0694/measure.py baseline
python3 docs/performance/results/change-0694/measure-allocations.py baseline
python3 docs/performance/results/change-0694/measure-refusal.py baseline
python3 docs/performance/results/change-0694/profile.py baseline
```

Preserve the staged baseline executables and their receipts while switching to
the candidate sources. The scripts use the named `litchi-0694-bin` and
`litchi-target-0694` siblings of the checkout. First run `build-oracle.py baseline` on the candidate checkout; it temporarily
restores and then restores back the exact sources as described below. Run the
three build scripts with `candidate`; `measure.py` and `measure-refusal.py` with `compare`; and the
allocation/profile scripts with `candidate`. The compare drivers perform
A/B/B/A. Each main leg has 100 samples and five warmups.

The retained `build-oracle.py baseline` driver also supports this packet's
historical late baseline build: it archives the candidate codec/tests, restores
the exact baseline files temporarily, asserts the complete source map and
restores the candidate bytes in `finally`. Use a fresh restoration archive
when reproducing; an existing archive is accepted only if its bytes match.
The initial copied oracle lock refused `--locked` before compilation; the
successful baseline/candidate oracle builds share the offline-refreshed lock.
See `oracle-lock-initial` and `oracle-lock-refresh.log` for that setup history.

The compiled probe trees are byte-frozen, including their inherited README
files. The main executable is named `probe0694` and reports probe ID `0694`;
its edit marker intentionally remains `0691`. The reused refusal executable
is named `probe0694`, while its transcript identifier remains `0693-refusal`.
Some frozen sub-probe documentation refers to its original 0693 provenance;
the commands and paths in this README are authoritative for 0694.

Run the [exact-output oracle](oracle/README.md) without its optional timing
mode, pointing it at the two frozen `*-oracle` binaries and `test-data`:

```sh
python3 docs/performance/results/change-0694/oracle/corpus.py \
  /path/to/baseline-oracle /path/to/candidate-oracle \
  /path/to/checkout/test-data /path/to/packet/oracle-results
python3 docs/performance/results/change-0694/measure-oracle-controls.py
python3 docs/performance/results/change-0694/measure-oracle-real-controls.py
```

The two control scripts are secondary native measurements with ten warmups,
300 samples and A/A plus A/B/B/A. Their timers include processing and output
drop, with input loading, capability construction and output hashing outside.
They compare empty, one-name and 4,096-name extension profiles. Keep their
scoped regressions visible; do not substitute them for whole-workflow timings.

Derive tables with `summarize.py`, `summarize-allocations.py`,
`summarize-refusal.py`, `report-metrics.py` and `binary-sizes.py`. Run
`run-integration.py`, `quality-summary.py`, `run-evidence.py`, then `audit.py`.
The audit regenerates only deterministic summaries and verifies exact raw
bindings. `script-hashes.json` binds the final retained Python drivers.

## Closure

Final checks and cleanup are recorded in `integration/results.json`,
`quality-summary.json`, `evidence/results.json`, `cleanup.json` and
`final-validation.json`. `cleanup.py --apply` verifies all eight frozen
executables before deleting only the named target, binary and raw-profile
scratch directories and generated marker-control archive. It preserves the
workspace lock and shared targets. `seal.py` repeats the post-cleanup data audit
and final report/coverage/non-iWork/structural-claim gates.

The first focused test run exposed two new test expectations that omitted
filtered MCE directives from report counts. `tests-initial.log` records those
failures; corrected complete Reports pass in `tests-focused.log` and the final
owner/consumer suites. No production behavior was changed to satisfy them.
