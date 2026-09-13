# 0557 metrics analyzer

`analyze.py` is a read only consumer for the 0557 XLSX provenance merge
campaign. It does not build Rust, launch a benchmark, edit a source manifest,
or choose a subset of samples. The driver owns child execution and every report,
catalog, receipt, binary, source manifest, host observation, and `/usr/bin/time`
sidecar remains receipt bound.

The matrix is four shapes (`medium`, `dense-sparse`, `noncompact`, and
`vendor-extension`) crossed with two primary source backed cases and six
controls. Native rows use two ABBA repeats with 20 warmups and 1,000 samples;
allocator rows use the same four legs with three warmups and 30 samples. The
baseline noise pilot contains only the eight primary case and shape rows and
must be complete before candidate output is observed.

The analyzer first checks the deterministic corpus identity against
`expected-primary.json`. Primary output hashes are checked as well; controls
must match their corpus identity and remain identical across the matched
stages. It then validates the report schema, singular case and shape, sample
count, sorted elapsed vector, sample order, statistics, confidence interval,
operation metrics, source vectors, catalog binding, receipt, binary custody,
source manifest, host record, and RSS sidecar. Missing eager source phases stay
explicitly unavailable with their source-phase scope and reason. They never
become zero-valued allocation evidence.

For a noise row, the relative p50 change is

```text
N_row = 100 * abs(noise_r2_p50_ns - noise_r1_p50_ns) / noise_r1_p50_ns
```

`N` is the largest `N_row`, with its case and shape retained. A value above
five percentage points makes the pilot too unstable for admission. Both
repeats remain in the report and no rescue repeat is permitted. The noise
report's per-row floors are provisional diagnostics only. For each matched
native primary row and repeat, the admission floor is recomputed from that
matched baseline p50:

```text
max(1.0, 3 * N, 100 * 50000 / matched_baseline_p50_ns)
```

When the pilot is above the five-point limit, ordinary analysis stops before
loading candidate rows and returns a stopped candidate gate. Candidate output
must therefore not be collected after an unstable pilot; the coordinator must
enforce that ordering when invoking the driver.

The final term is dimensionally a percentage: 50,000 nanoseconds is divided by
the baseline p50 in nanoseconds and multiplied by 100. Equivalently it is
`5_000_000 / matched_baseline_p50_ns` percentage points. The pure self test
covers one, five, and one-point-five percent floor examples and the zero
denominator rule.

Native comparisons retain p50, mean, p95, p99, min, max, standard deviation,
and both confidence interval endpoints. RSS retains all four sidecar values
for every case and shape; only `max_rss_kib` participates in its gate. Every
absolute candidate increase above five percent is listed as adverse. Every
same-build repeat drift above five percent is listed separately, including
zero-denominator changes.

Allocator vectors are checked only when the report supplies an operation
allocation sample. The existing combined
`commit_allocation_metrics` interval is retained alongside the new non-nested
`staging_allocation_metrics` and `commit_core_allocation_metrics` intervals.
The report includes calls, deallocations, reallocations, failed calls, byte
counters, live endpoints, process peaks, and the checked region peak. The
derived metric

```text
incremental_region_peak_live_bytes = region_peak_live_bytes - live_bytes_before
```

is calculated only after the sample bounds and live-byte balance pass. The
allocation gate compares every reported phase and metric's p50 and mean with
the same zero-safe five-percent rule. Primary and control elapsed p50/mean
gates use only the native lane. Allocator elapsed vectors remain diagnostics
and their adverse and same-build drift rows stay retained; allocator per-child
RSS and allocation phase metrics retain their own gates. Eager controls with
no provenance phase are recorded as explicit allocation unavailability and do
not satisfy a fabricated universal allocation gate.

Run the pure checks without evidence using:

```text
python3 -B docs/performance/results/change-0557/analyze.py --self-test
```

After the baseline pilot, `--noise-only` writes a deterministic noise report.
The ordinary invocation consumes the complete noise, native, and allocation
matrix. `--output` uses exclusive-create and identical-replay semantics.

This report supplies numerical evidence only. It makes no adoption, source
proof, preservation, correctness, mechanism, hardware, cold-cache, scaling,
native Office, ODF, or iWork claim.
