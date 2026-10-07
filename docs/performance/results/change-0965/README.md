# 0965 — ingress-tail evidence and matched diagnostic

This is a diagnostic follow-up to the rejected 0964 screen. It retains the
0964 epoch-2 raw reanalysis and records the completed 0965 native matched
diagnostic. The packet does not read allocator reports, pool samples for a
decision, claim causality, reopen the screen, or permit adoption. Its
independent raw audit passed.

The complete bounded result is
[`analysis/ingress-summary-v1.json`](analysis/ingress-summary-v1.json),
443,033 bytes, SHA-256
`091a122bbb03f5ba65bca5964aa0659c2f2433cfd036d274c76c381b591495b8`. It
contains absolute p50/p95/p99 values for every native child and every timing
phase, absolute quantiles and index-aligned raw-sample deltas/ratios for all
18 adjacent native legs, sample-window trends, launch-order trends, and
within-sample phase correlations. The replay helper is
[`tools/analyze_ingress_v1.py`](tools/analyze_ingress_v1.py), 28,336 bytes,
mode 0755, SHA-256
`b6dd1f72ffddb806188e4f7f665b64c70941a841cb8bf43dee48b288bdac240a`.
The external replay receipt is
`/home/zhuhe/litchi-performance-0965/receipts/ingress-replay-v1.json`, 1,859
bytes, SHA-256
`a442457c006bd438e7ada5dc84783d1e151f658e2fc08eafbf81704b02dc9ce5`.
An independent stdlib recomputation of all 180 child p99 values and 18 paired
ingress p99 ratios reports zero mismatches; its external receipt is
`/home/zhuhe/litchi-performance-0965/receipts/independent-quantile-check-v1.json`,
905 bytes, SHA-256
`af64fd56f05aa1504278780b636056cb1b1876859e7cada3ad975b7cc94277d7`.

The completed 0965 reduction is external at
`/home/zhuhe/litchi-performance-0965/analysis/diagnostic-result-v1.json`,
284,168 bytes, SHA-256
`1eb904d55ad2ee4498da7ee1fd980fbd222d60e11c8cdcb5b4472f8047db2e10`.
It covers 24 fresh children and 6,144 measured rows. The repository retains
byte-identical copies of the 0964 reanalysis summary and the new
[diagnostic reduction](analysis/diagnostic-result-v1.json), together with
replay helpers. Raw capture reports remain external.

## Custody and method

The analysis is bound to the authorized epoch-2 plan
`/home/zhuhe/litchi-performance-0964/protocol/plan-final-epoch2-v1.json`,
59,570 bytes, SHA-256
`dcfabed26427c6a117b41c2e4b6209baec40afe06c218054b2187ae9c49b0a64`, and to
the completed native packet
`/home/zhuhe/litchi-performance-0964/capture-native-epoch2-v1/completed-packet.json`,
14,839 bytes, SHA-256
`d9934f4ffbf4646953d18909de48ccea4db36aceaa1e5367c7c3d18a241fd69f`.
The packet's capture-start descriptor is 700 bytes, SHA-256
`7903939b5b13c2c597ae0a6732facede68b90fe3830236f350711649571d7cf1`; its
completed capture-terminal descriptor is 575 bytes, SHA-256
`16bb0261e04fdd8bd2208e0e6402a5aa9464d05ebbc023507018388e75e83bb3`.
The external raw reports remain outside the repository; the JSON result binds
all 36 report paths, sizes, and SHA-256 values.

There are 36 valid native children: three cases, 12 children per case, six
baseline and six candidate children per case, each with 10 untimed warmups and
60 measured samples. Adjacent A/B or B/A children within each frozen block
form the 18 paired legs. The helper validates the plan/packet bindings,
terminal-before-parse state, affinity witnesses, validation receipts, report
hashes, schemas, and sample sequence before computing any values.

Quantiles use nearest rank `ceil(q*n)`, one-indexed. With 60 measured samples,
p50 is rank 30, p95 is rank 57, and p99 is rank 60. Therefore each child's
p99 is its maximum measured sample; it is a tail-screening diagnostic rather
than a stable estimate of a population percentile. The result keeps child and
paired-leg values separate and does not pool the 2,160 samples into a gate.

## Findings

The table gives the median across six child-level quantiles for each role. It
is an absolute timing summary in nanoseconds, not the 0964 selector's ratio
table.

| case | role | ingress p50 | ingress p95 | ingress p99 |
| --- | --- | ---: | ---: | ---: |
| medium-one-percent | baseline | 80,995.5 | 94,815.5 | 102,635.5 |
| medium-one-percent | candidate | 87,405.5 | 102,000 | 116,061 |
| dense-sparse-one-percent | baseline | 93,230.5 | 106,675.5 | 117,965.5 |
| dense-sparse-one-percent | candidate | 94,110.5 | 107,545.5 | 119,781 |
| medium-one-edit | baseline | 81,940 | 90,015.5 | 96,365.5 |
| medium-one-edit | candidate | 83,190 | 92,205.5 | 110,546 |

The candidate-over-baseline ratios of the absolute p99 quantiles are spread
across all six legs in each case. Their values are `0.826486, 1.170127,
1.210526, 0.896504, 1.315015, 1.052980` for medium one-percent;
`1.238506, 1.140859, 0.980699, 0.771089, 1.068751, 0.890806` for
dense-sparse one-percent; and `1.080920, 1.101776, 1.128305, 1.013776,
1.258214, 1.113256` for medium one-edit. Their medians are respectively
`1.111553`, `1.024725`, and `1.107516`, reproducing the two 0964 safety stops
and the non-stop dense case. The JSON also keeps the distinct p99 of
index-aligned per-sample ratios for diagnostic review. A sample at index `i`
in one fresh child is not a paired experimental unit with sample `i` in the
other child; these ratios are descriptive and cannot support a confidence
interval or a causal claim. These values remain descriptive evidence for the
already rejected screen.

The six ten-sample ingress-window medians show no common monotonic warmup or
late-run rise. Candidate medians move from 89,410 to 89,265 ns for
medium-one-percent, 98,623 to 96,765 ns for dense-sparse one-percent, and
84,283 to 82,295 ns for medium-one-edit. The largest individual ingress
observations occur at varied measured-sample indices: the overall maximum is
141,151 ns in dense-sparse baseline block 1 position 2 at sample 43; the next
candidate dense-sparse maxima are 131,221 ns at sample 48 and 131,201 ns at
sample 31. Across the top 18 observations per case and role, sample-index
ranges are 0–56 / 9–58 for medium one-percent, 7–57 / 3–58 for dense-sparse
one-percent, and 2–50 / 1–49 for baseline/candidate respectively. This is
consistent with sparse tail observations distributed through the measured
sequence; it does not identify their cause.

Launch order is also descriptive. The strongest within-case child-order
association is dense-sparse baseline p99, Spearman `0.771429` across six
children; candidate dense-sparse p99 is `-0.314286`. The other p99
order coefficients are small or mixed. Six child values per role are too few
to interpret these coefficients as an order effect, but the full order and
per-child values are retained in the JSON for review.

Within-sample phase correlations show why ingress alone cannot explain every
total tail. The ingress/total Pearson correlations for baseline/candidate are
`0.7017629 / 0.4518979` for medium one-percent, `0.0529015 / 0.4491674` for
dense-sparse one-percent, and `0.6086344 / 0.5699614` for medium one-edit.
The complete five-phase matrices are in the result. The measured total is
close to, but not exactly, the sum of the four named phases; the observed
total-minus-phase-sum median residuals are 410–510 ns by case/role, with a
24,830 ns maximum in the medium-one-edit baseline observations. The
correlations and any pooled/index-aligned observations are descriptive
associations, not paired experimental units or a causal decomposition.

## Source boundary and completed diagnostic

The source review at
[`ingress-source-review.md`](ingress-source-review.md) establishes the exact
boundary. `ingress_ns` covers the `SourceBackedEditor` constructor and the
managed `SourceBackedPackage::from_read_at_inner` OPC path: source metadata,
ZIP indexing and `PackageReader::source_catalog` structural content-types and
relationship cataloging. The 0964 `Reader`/`NamespaceResolver` loops execute
after the timer, during XLSX workbook/worksheet validation and source parsing.
There is no constructor subphase through the worksheet reader to measure for
this candidate. `source_stream_admission` is worksheet semantic source
admission after OPC package construction; it is distinct from the OPC
catalog's structural admission and must not be used as a name for the latter.

Root's final confirmation selected
`/home/zhuhe/litchi-performance-0965/protocol/plan-final-v2.json`, 39,436
bytes, SHA-256
`147631fc34eb9230cd002b7fd2ce9ae82c156943a502c558e9eb867e24b649b0`. The
completed packet is
`/home/zhuhe/litchi-performance-0965/capture-ingress-diagnostic-v1/completed-packet.json`,
10,990 bytes, SHA-256
`a8be912089527c0f01a5c705ce036c81a117dd11419449d366c39816ab88a334`.
It has 24 fresh serial children, 20 warmups and 256 measured samples per
child, with one `RR`, `RC`, `CC` and `CR` adjacent pair per case. The p99
nearest-rank is sample 254, for 6,144 measured rows total.

| case | RC p99 | CR p99 | CC magnitude | RR magnitude | label |
| --- | ---: | ---: | ---: | ---: | --- |
| medium-one-percent | 1.065359 | 0.897036 | 1.054635 | 1.042147 | mixed |
| dense-sparse-one-percent | 1.073303 | 0.984729 | 1.057456 | 1.012618 | mixed |
| medium-one-edit | 0.975812 | 1.076656 | 1.047029 | 1.045083 | mixed |

All three cases are `mixed`: the two cross-order ratios disagree in every
case, and the `CC` control also exceeds `1.05` in the two one-percent cases.
This gives neither concordant candidate-ingress causality nor proven host
noise. The `0964` rejection remains unchanged. Index-aligned sample ratios,
pooled correlations and order coefficients are descriptive only; they are not
paired experimental units, confidence intervals or causal decomposition.

The final report-scope contract is
`/home/zhuhe/litchi-performance-0965/protocol/report-scope-contract-v1.json`,
6,792 bytes, SHA-256
`f4476e77d44cb31ada4981b182210a134cad0968121929472ae48ab5c410d686`. It
binds `ingress_ns` to the direct `SourceBackedEditor` constructor and keeps
the setup, phase and verification literals byte-exact with the frozen 0964
harness. Source revalidation passed against 10,933-file baseline and
candidate manifests; the receipt is
`/home/zhuhe/litchi-performance-0965/receipts/source-revalidation-v1.json`.
The root confirmation is
`/home/zhuhe/litchi-performance-0965/protocol/root-final-confirmation-v1.json`,
SHA-256
`b4a019abb2a1069ad62802644b4556d880aa36f5e6e2fa3f379bb6118e3f1f89`.

Host-before and host-after snapshots are
`protocol/host-before-v2.json` (SHA-256
`6646310015b9e0acebb3db5aec2766fcf4716bd3936d5d1174ab2926fcfa514e`) and
`protocol/host-after-v1.json` (SHA-256
`9a8010731546d476cfcdf36be1b939fc4943baef0a31dc0c84e6389fd2e02989`). They
describe a shared host with fresh serial processes and a warm filesystem, not
an exclusive host or cold-cache experiment. The independent raw audit is
`/home/zhuhe/litchi-performance-0965/review/independent-raw-audit-v3.json`,
240,496 bytes, SHA-256
`25353dd5e02720b7a125897a611993331b23feec8c8a96b208a5386743af97bd`.
It passes with zero mismatches across 120 phase-quantile sets, 12 pairs and
24 children; all five scope strings across 24 reports and the host-field mappings pass. Earlier v1 and
v2 attempts remain preserved as failures: v1 rejected a compact
root-confirmation descriptor, and v2 treated `time-v.txt` as JSON. V3 changed
only the reader and audited the same packet and reducer; no workload rerun or
capture extension occurred. No further captures or adaptive sample extension
are planned.

## Reproduction

The historical 0964 raw reanalysis remains replayable with a new output path:

```sh
R=/home/zhuhe/litchi-performance-0964
E=/home/zhuhe/litchi-performance-0965
PYTHONDONTWRITEBYTECODE=1 python3 "$E/tools/analyze_ingress_v1.py" \
  --root "$R" \
  --plan "$R/protocol/plan-final-epoch2-v1.json" \
  --output "$E/analysis/ingress-summary-replay.json"
```

The replay performs no build, workload launch, or allocator read. The retained
run completed with `py_compile` and the command above; the external and
repository copies of the helper and JSON result are byte-identical.

The completed 0965 packet reduction is reproducible from the sealed final plan
and packet without launching a workload:

```sh
E=/home/zhuhe/litchi-performance-0965
PYTHONDONTWRITEBYTECODE=1 python3 "$E/tools/reduce_ingress_0965_v1.py" \
  --plan "$E/protocol/plan-final-v2.json" \
  --packet "$E/capture-ingress-diagnostic-v1" \
  --output "$E/analysis/diagnostic-result-replay-v1.json"
```

The independent audit was run after the reduction with the same sealed packet:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 "$E/tools/audit_ingress_0965_v3.py" \
  --plan "$E/protocol/plan-final-v2.json" \
  --packet "$E/capture-ingress-diagnostic-v1" \
  --reduced "$E/analysis/diagnostic-result-v1.json" \
  --output "$E/review/independent-raw-audit-replay.json"
```

Neither command changes the sealed packet or permits adaptive captures. Do not
claim a candidate ingress cause from the unchanged callgraph; a separate
measurement or attribution design is required before revisiting the question.

## Artifact retention and cleanup

The [external inventory](external-evidence-inventory-v2.json) records retained
0965 evidence by file name, size, mode and SHA-256. Raw capture directories
remain under `/home/zhuhe/litchi-performance-0965`; retained executable and
source-manifest identities are in the final plan and the 0964 evidence packet.
No binaries are committed. Every committed evidence file is below 1 MiB.
Only generated Python cache files were removed; all measurements, audit
failures and reproducibility inputs were retained. No build cache was created.

Regenerate an inventory to a fresh output path with the reused inventory tool:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 "/home/zhuhe/litchi-performance-0964/tools/inventory_evidence_v1.py" --root "$E" --output "$E/external-evidence-replay.json" --capture-state stable
```
