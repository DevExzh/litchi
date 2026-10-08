# 0973 all-child instruction-profile metric proposal

Status: staging note only. It binds no 0973 capture and starts no workload,
build, decoder, or native measurement. The proposal is intentionally
descriptive across all twelve serial Float/e1/default children; it does not
select a fast or slow subset.

The retained 0970 report establishes that the Float/e1 report has one timed
region per sample and stable work/output structure, but no per-stage duration or
counter. Its 80 measured samples exclude the five warmups. The 0973 marker
covers 85 warmup-plus-measured calls, so marker-period metrics and report
timing metrics must retain those different denominators.

For child i, use the corrected ASLR-safe decoder and exact target
ELF/build-ID/map admission:

* D_i: sum of every retained instructions:u sample period in the child.
  Call this sampled instruction period, not an exact retired-instruction
  counter unless a separate counter is captured.
* M_i: sum of periods for samples whose physical chain contains the exact
  marker symbol doc_retention_probe_0942_cfb_profile::profiled_measured_case
  in the admitted target DSO.
* m_i: number of those marker-containing sample records.
* S_i = M_i / D_i: raw descriptive marker-period share.
* R_i = D_i - M_i: non-marker sampled period, with R_i / D_i retained as
  the residual share.
* T50_i and T95_i: p50 and p95 of the 80 measured harness samples only.
  Keep warmup count and marker call scope in the row; do not treat M_i / 85
  as a measured-call timing.

Report every child row and arm/cohort summaries for D, M, S, R, m, T50, and
T95: count, min, median, max, population standard deviation, and population CV.
Rank descriptive Pearson and Spearman correlations of each period metric and
share with T50 and T95 over all twelve children. If a child fails the frozen
loss/marker gate, retain its row and report the gate status; never remove it by
timing or period value. A gate failure blocks a qualification claim rather than
creating a post-hoc cohort.

Use the identity

Var(D) = Var(M) + Var(R) + 2 Cov(M,R)

as a descriptive variation ledger, with the residual and covariance terms
reported directly. Per-child period conservation must also show
D_i = M_i + non_marker_period_i. These identities distinguish variation in the
marker subset from variation outside it without calling either component a
causal phase cost.

Keep two ownership views separate:

1. **Inclusive ancestor tags.** For each exact physical or explicitly decoded
   tag such as the marker, run_case, measure_one, or a resolved stage symbol,
   count each sample at most once per tag and report its sample count and
   period. These are overlapping subsets. Show their period shares only against
   total or marker denominators and never add them.
2. **Exclusive deepest physical target-DSO proxy.** For each
   marker-containing sample, choose exactly one leaf-most resolved physical
   frame in the target DSO, using exact frame order and symbol boundaries.
   Assign unresolved or missing target frames to an explicit unknown bucket.
   Inline contexts are annotations and do not create another owner. Sum of the
   exclusive proxy buckets plus unknown must equal M_i; this is a
   single-sample proxy, not self-cost or owner attribution.

Retain marker records with unknown or unresolved frames, record duplicate-marker
anomalies, and bind zero lost events, raw sample count, total period, exact
ELF/build ID, load bias, and decoder version per child. The 0972 recovery
supports offline marker ancestry recovery but recovered 78/1387 marker chains,
below the historical 100-marker gate; it does not establish a native timing,
phase-owner, stack-depth, or upstream perf/libdw cause.

Inputs used only as design anchors:

* 0970 invariant report:
  /home/zhuhe/litchi-performance-0971/analysis/float-e1-work-v1.json
  (sealed report hash recorded in its receipt).
* 0972 corrected replay and independent review:
  /home/zhuhe/litchi-performance-0972/receipts/dwfl-replay-receipt-v6.json
  and
  /home/zhuhe/litchi-performance-0972/receipts/independent-offline-review-0972-final-v1.json.
