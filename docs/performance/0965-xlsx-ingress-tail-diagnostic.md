# 0965 — XLSX ingress-tail diagnostic

**Status: completed descriptive diagnostic; independent raw audit passed.**
This record follows the rejected 0964 direct-event-transport screen. It does
not reopen the 0964 decision, relax its `1.05` ingress-p99 safety bound,
authorize adoption, or claim a causal effect.

The valid 0964 epoch-2 raw reanalysis is retained as
[`results/change-0965/analysis/ingress-summary-v1.json`](results/change-0965/analysis/ingress-summary-v1.json),
443,033 bytes, SHA-256
`091a122bbb03f5ba65bca5964aa0659c2f2433cfd036d274c76c381b591495b8`. Its
absolute quantiles and descriptive comparisons are summarized in the
[result README](results/change-0965/README.md). The completed 0965 diagnostic
is reduced in the external
`/home/zhuhe/litchi-performance-0965/analysis/diagnostic-result-v1.json`,
284,168 bytes, SHA-256
`1eb904d55ad2ee4498da7ee1fd980fbd222d60e11c8cdcb5b4472f8047db2e10`. It
contains 24 fresh children and 6,144 measured rows. No rebuild or production
change was needed.

## Source-path finding

The source review is
[`results/change-0965/ingress-source-review.md`](results/change-0965/ingress-source-review.md).
It traces the frozen managed benchmark timer from
`SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context`
through `SourceBackedPackage::from_read_at_inner` and stops before selector
staging. The measured `ingress_ns` constructor performs source version/length
checks, exact managed source reads, ZIP indexing, and
`PackageReader::source_catalog` structural OPC content-types and relationship
cataloging. The constructor delegation is in
`crates/litchi-xlsx/src/cell_values/source.rs:425-445`; the managed OPC body is
in `crates/litchi-opc/src/source_backed.rs:6912-7221`; the timer is in
`tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs:416-425` in the
frozen external source tree.

The 0964 candidate's two changed loops require decompressed XLSX semantic XML:
`cell_values::validation::validate_xml` and raw worksheet
`parse_source_with_observer`. They are reached later by
`MultiSnapshot::load_source_backed`, workbook validation, and selected
worksheet validation/source parsing inside `edit_commit_ns`
(`crates/litchi-xlsx/src/cell_values/snapshot.rs:205-255,467-535`). They do not
execute inside `ingress_ns`. `source_stream_admission` names the later
worksheet semantic source-admission decision; it is distinct from the OPC
structural catalog admission that occurs during ingress. This rules out a
direct candidate-loop explanation for the ingress flags while leaving binary
layout, process-state and short-phase runtime effects as hypotheses to test.

## 0964 raw absolute quantiles

The JSON above is authoritative for every child and phase. The table gives
the median across six child-level absolute quantiles for each case and binary
role; it is not a ratio reduction and does not pool the 2,160 samples.

| case | role | ingress p50 (ns) | ingress p95 (ns) | ingress p99 (ns) |
| --- | --- | ---: | ---: | ---: |
| medium-one-percent | baseline | 80,995.5 | 94,815.5 | 102,635.5 |
| medium-one-percent | candidate | 87,405.5 | 102,000 | 116,061 |
| dense-sparse-one-percent | baseline | 93,230.5 | 106,675.5 | 117,965.5 |
| dense-sparse-one-percent | candidate | 94,110.5 | 107,545.5 | 119,781 |
| medium-one-edit | baseline | 81,940 | 90,015.5 | 96,365.5 |
| medium-one-edit | candidate | 83,190 | 92,205.5 | 110,546 |

The six adjacent-leg p99 ratios are retained in the JSON and README. Their
medians are `1.111553` for medium one-percent, `1.024725` for dense-sparse
one-percent and `1.107516` for medium one-edit, reproducing the two 0964
safety stops. Index-aligned per-sample ratios and pooled phase correlations
are descriptive only: sample index does not create a paired experimental unit,
and neither view supplies a confidence interval or causal decomposition.

## Completed matched diagnostic

The final plan is
`/home/zhuhe/litchi-performance-0965/protocol/plan-final-v2.json`, 39,436
bytes, SHA-256
`147631fc34eb9230cd002b7fd2ce9ae82c156943a502c558e9eb867e24b649b0`. The
packet uses the retained executables, one fresh serial process per child, one
managed worker, 20 warmups and 256 measured samples. Each case has one
adjacent `RR`, `RC`, `CC` and `CR` pair. Nearest-rank p99 is sorted rank
254 of 256, the third-highest observation; no
retries, replacements, adaptive extension or additional captures are allowed.

`R` is the retained 0957 baseline and `C` the 0964 candidate. Cross-order
ratios always mean candidate over baseline; same-binary magnitudes are
`max(q2/q1, q1/q2)`. The reduction classifies every case as `mixed`:

| case | RC p99 | CR p99 | CC magnitude | RR magnitude | label |
| --- | ---: | ---: | ---: | ---: | --- |
| medium-one-percent | 1.065359 | 0.897036 | 1.054635 | 1.042147 | mixed |
| dense-sparse-one-percent | 1.073303 | 0.984729 | 1.057456 | 1.012618 | mixed |
| medium-one-edit | 0.975812 | 1.076656 | 1.047029 | 1.045083 | mixed |

The two cross-order comparisons disagree in every case, and two `CC` controls
also exceed `1.05`. This packet shows no concordant candidate-specific p99
pattern and does not establish host noise as the cause. The original 0964 rejection remains binding;
the diagnostic has no adoption or statistical-effect outcome. Index-aligned
per-sample ratios, pooled correlations and order coefficients remain
descriptive, not paired experimental units or causal evidence.

The report-scope contract is
`/home/zhuhe/litchi-performance-0965/protocol/report-scope-contract-v1.json`,
6,792 bytes, SHA-256
`f4476e77d44cb31ada4981b182210a134cad0968121929472ae48ab5c410d686`. It
binds `ingress_ns` to the direct `SourceBackedEditor` constructor and keeps
the setup, phase and verification scopes byte-exact with the frozen 0964
harness. Source custody revalidation passed against the 10,933-file manifests;
its receipt is
`/home/zhuhe/litchi-performance-0965/receipts/source-revalidation-v1.json`.
Host-before and host-after snapshots are
`protocol/host-before-v2.json` (SHA-256
`6646310015b9e0acebb3db5aec2766fcf4716bd3936d5d1174ab2926fcfa514e`) and
`protocol/host-after-v1.json` (SHA-256
`9a8010731546d476cfcdf36be1b939fc4943baef0a31dc0c84e6389fd2e02989`); they
describe a shared host with fresh serial processes and a warm filesystem, not
an exclusive host or cold-cache experiment.

The independent raw audit is
`/home/zhuhe/litchi-performance-0965/review/independent-raw-audit-v3.json`,
240,496 bytes, SHA-256
`25353dd5e02720b7a125897a611993331b23feec8c8a96b208a5386743af97bd`.
It passes with zero mismatches across 120 phase-quantile sets, 12 pairs and
24 children; all scope strings and host-field mappings pass. Earlier v1/v2
failures (compact confirmation descriptor, then `time-v.txt` parsed as JSON)
remain preserved; v3 audited the same packet and reducer without a rerun. Do
not claim candidate ingress causality or proven host noise. A separate
measurement/attribution design is required before revisiting this question.

## Scope and custody

The retained raw helper and frozen JSON remain unchanged. This record and the
README bind the final plan, scope/host custody, diagnostic result and replay
commands. No source, benchmark or workload file is changed; the broader non-iWork Office performance
goal remains active.
