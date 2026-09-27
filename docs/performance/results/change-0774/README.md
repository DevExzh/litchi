# Change 0774 evidence

See [the integration report](../../0774-xls-writer-integration.md).
This packet separates early branch validation, review corrections and final
source validation. Historical 0766 commit-message probe numbers are not adopted.

- `integration.json`: initial clean integration of the three original commits.
- `final-source.json`: final production source after review fixes.
- `review.md`: issues found and their disposition.
- `quality.py`, `quality-0/`: initial gates before review fixes.
- `quality-1/`, `quality.json`: final serial format/check/test/lint/doc/facade/harness/boundary gates.
- `property-1024.*`: additional fixed-seed property run on the initial integration.
- `property-final-1024.*`: same seed repeated on final production source.
- `function-table.json`: equality of the fourteen built-in function indices.
- `environment.json`, `workspace-Cargo.lock`: toolchain, host and root resolution.
- `probe-src/`: identical standalone probe template for both production trees.
- `measure-0/`: retained failed probe build, no native observations.
- `measure.py`, `measure-1/`: paired release build/run receipts and samples.
- `analyze.py`, `analysis.json`: offline statistics, parity and regression flags.
- `allocations.py`, `allocations-0/`, `allocation-analysis.json`: separate whole-process heaptrack capture and allocation-size histogram sums.
- `validate.py`: offline receipt/statistics/seal verification.

The probe creates a fresh writer for each sample, registers 20,000 formula or
numeric cells, and serializes to a memory buffer. Timing starts before writer
construction and ends after `write_to`; writer destruction, SHA-256 and output
checks are outside the timer. Process peak RSS includes all of that work.
Every output digest and length must agree within and across both legs.

Four processes per case and leg run serially on CPU 12 in alternating
before/after order. Each has two warmups and nine measured samples. Process
p50 values and spread are retained; nine-sample p95/p99 are merely the maximum
sample, not stable tail estimates. The host is shared, and no physical cold
cache or concurrency condition is claimed. Heaptrack runs separately with one
sample and no warmup; its instrumented times are excluded from native results.

To reproduce, make worktrees for the exact references in `final-source.json`,
copy the archived workspace lock to both, and adjust script root/target paths.
Use the archived standalone probe lock for identical dependency resolution.
Run Cargo/native/profiler workloads serially. Never overwrite a capture folder.
For offline checks, run `python3 -B docs/performance/results/change-0774/validate.py`.
