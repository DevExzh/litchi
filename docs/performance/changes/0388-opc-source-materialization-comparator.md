# Change 0388: operation-scoped allocator comparator for OPC materialization

Date: 2026-09-03

Status: implemented

`performance_claim: none`

`claim_authorized: false`

## Comparator boundary

The regression comparator now has an explicit allocator evidence scope. The
historical filesystem allocator mode remains the default for policies that do
not name a scope and still requires cache-state result keys, raw child-process
allocator samples, and its existing filesystem identity checks. The new
`allocator_evidence_scope: "operation"` mode is for non-filesystem selectors
such as `opc_source_materialize`.

Operation rows use only `(case, corpus)` identity. They require the allocator
binary and a paired `binary_identity`, reject `cache_state` and
`filesystem_evidence`, and validate one shared operation envelope: measured
system-allocator status/scope, all ten allocation vectors, sample cardinality,
and the elapsed-order sample-index permutation. No raw filesystem or child
process evidence is inferred.

## Checked 0387 comparison

[`perf-regression-policy-opc-source-materialize-allocator-v1.json`](../perf-regression-policy-opc-source-materialize-allocator-v1.json)
pins the three deterministic 0387 corpora and 15 retained samples per row,
and permits only non-increasing selected counters.
The checked
[`opc-source-materialization-shared-0388-comparison.json`](../results/opc-source-materialization-shared-0388-comparison.json)
was produced from the committed allocator-enabled control/candidate reports
in `results/change-0387/`. It contains 15 comparisons: allocation calls,
allocated bytes, logical source read calls/returned bytes, and materialized
Part count for each corpus. All three rows pass; source and materialization
work vectors are observed as invariant in this comparison, while allocation
vectors decrease in the candidate. The policy's zero-percent upper threshold
prevents increases; it does not require future work counters to remain
numerically equal.

The derivation is reproducible by the focused comparator test: it runs
`zstd -q -d -c` in memory over the six committed report files
(`control-*.json.zst` and `candidate-*.json.zst` for tiny, many-small, and
few-large), combines their one result row into a three-row report, removes
only the per-file parallel/catalog sidecars, and normalizes the combined shape
and payload arrays. The test then runs `compare_reports` with the checked
policy and asserts the complete object equals the checked JSON output. No
decompressed report or temporary file is checked in.

Run it with:

```sh
python3 -m unittest \
  tools.test_perf_compare.PerfCompareTests.test_checked_0387_operation_allocator_policy_and_comparison_are_scoped
```

Elapsed samples are validated but every latency comparison is withheld. The
control and candidate reports intentionally share the source revision because
the candidate source change was captured as a patch, and the policy therefore
does not claim distinct revisions or a clean candidate worktree. This output
is mechanism evidence only, not a latency, throughput, RSS, copied-byte,
decompression, physical-I/O, or broad OPC performance claim.
