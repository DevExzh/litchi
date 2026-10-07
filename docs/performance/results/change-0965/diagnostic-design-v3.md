# 0965 — ingress-tail cause diagnostic design

Status: prospective design only. No workload, build, source edit, or
production decision is authorized by this document.

## Decision and scope

The smallest useful next diagnostic is a native-only, fresh-process matched
comparison using the two retained 0964 executables. It should run paired
same-binary controls as well as baseline/candidate pairs:

* `R` is the retained 0957 baseline executable.
* `C` is the retained 0964 candidate executable.
* `RR` measures baseline-to-baseline variation.
* `RC` and `CR` measure the candidate over baseline in both launch orders.
* `CC` measures candidate-to-candidate variation.

Run these four two-child pairs for each of the three already qualified cases.
That is eight children per case and 24 serial children in total. Each child
uses 20 warmups and 256 measured samples. This is the upper end of the small
12–24-child diagnostic budget, but it is the minimum schedule that covers a
same-binary baseline, both cross-binary orders, and a same-binary candidate
control for every case. The dense-sparse case is retained as a sentinel: its
0964 ingress p99 did not trip the aggregate safety rule while both medium
cases did.

This experiment can expose whether the 0964 ingress-p99 pattern is concordant
across launch orders in this fresh packet, mixed across orders or controls, or
not reproduced above the existing bound. Those are packet-scoped diagnostic
patterns. They do not identify host, allocator, kernel, build, or source
causes. A clean independent rebuild would be required for the narrower
build-provenance question; that is outside this minimal diagnostic and must not
be inferred from it.

The controls also cannot eliminate every shared-host variable. ASLR and
per-process code/data placement, allocator arena state, kernel scheduling,
KVM placement, interrupts, and the shared filesystem/page cache can differ
between fresh children. A fresh process resets the benchmark's user-space
state, but not those host or kernel states. `RR`/`CC` and both launch orders
measure how large those effects are under this run; they do not establish a
causal decomposition of them.

The existing benchmark already has the required boundary. Its `ingress_ns`
timer starts at `SourceBackedEditor` construction, after the fresh source
clone, finite context, cache policy, and sink setup, and stops before selector
staging, edit/commit, publication, and commit teardown. A constructor-only
binary would alter precisely those setup and process-state conditions and
would require a new build. It would therefore answer a less faithful question
than the retained report's ingress vector. No constructor-only harness is
needed for this diagnostic.

0964 remains rejected throughout. The existing `1.05` ingress-p99 safety gate
is not relaxed, reinterpreted, or replaced by the classifications below. The
diagnostic has no adoption, confirmation, or statistical-effect outcome.

## Evidence motivating the diagnostic

The valid 0964 epoch-2 screen reports candidate-over-baseline ingress p99
medians of `1.111553` for medium one-percent, `1.024725` for dense-sparse
one-percent, and `1.107516` for medium one-edit. The two medium values caused
the hard screen rejection even though edit/commit and total p50/p95 medians
were below one. Each value is a median of six paired screening legs built from
60 samples per child; it has no confidence interval and is not a qualified
effect.

The concern is not new to 0964. The preceding records retain different
ingress-tail observations:

* 0960 records one-edit ingress p95 `1.0657` and p99 `1.1703` while its
  aggregate native medians did not trip the frozen safety rule.
* 0961 records dense-sparse ingress p95 `1.0594` and one-edit ingress p99
  `1.0532`; it separately rejected on allocator-process RSS and no signal.
* 0959 keeps all phase tails as review observations and records a one-edit
  total p99 ratio of `1.1039 [1.0674, 1.1438]`; it does not establish that
  ingress or any single phase caused that result.

The recurrence across trials is evidence for a control experiment, not proof
of a host artifact. The 0964 raw report's phase contract and source/build
bindings are the authorities for the new run:

The exact custody bindings for the new packet are:

| item | path | SHA-256 |
| --- | --- | --- |
| baseline source manifest | `/home/zhuhe/litchi-performance-0964/source-manifest-baseline-v2.json` | `74d3ead27a1518093be3f80fbe50cfa1280526e08957b1774dbc6fd891bc4eda` |
| candidate source manifest | `/home/zhuhe/litchi-performance-0964/source-manifest-candidate-v4.json` | `19e4bdf18aee484ade90067a98ec6ab0ff5f6d017105840df4030f697f25779b` |
| candidate allowlist | `/home/zhuhe/litchi-performance-0964/protocol/candidate-test-allowlist-v3.json` | `a1c62adbe848111d1370e223614473d8f1500797d52cce4a74f7f24eab045c97` |
| candidate freeze | `/home/zhuhe/litchi-performance-0964/receipts/candidate-freeze-v4.json` | `4f1df545ee500b404596a90068dba815f96cb810b315658798f6793a2fdc4335` |
| final quality receipt | `/home/zhuhe/litchi-performance-0964/receipts/quality-0964-final-v1.json` | `fd84a9dd71cf73691c091e00444a7649331da2ee0da045781484049a930a247e` |
| performance-harness lock | `/home/zhuhe/litchi-performance-0964/source-baseline/tools/perf-baseline/Cargo.lock` | `db62bdd9174484613acae1f49f8187cf105d0b4976e2fea4de91df7372a893d5` |

The exact retained executables are:

* baseline executable:
  `/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_benchmark.v2`,
  43,259,328 bytes, SHA-256
  `5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3`;
* candidate executable:
  `/home/zhuhe/litchi-performance-0964/executables/xlsx_managed_compression_benchmark.v3`,
  43,270,848 bytes, SHA-256
  `d21863a6613800391a01914865fc4eaecdf54ad3061447be8b6be7d47d0dc910`;
* native report schema:
  `litchi.xlsx-managed-compression-benchmark.v1`;
* exact phase and oracle contract:
  `/home/zhuhe/litchi-performance-0964/protocol/plan-final-epoch2-v1.json`;
* valid epoch-2 raw audit:
  `/home/zhuhe/litchi-performance-0964/review/independent-epoch2-raw-audit-v1.json`.

## Exact prospective schedule

The runner must use one fresh process per child, one managed worker, and the
same deterministic case arguments as 0964:

```text
--workers 1 --warmup 20 --samples 256
```

The three cases, in this fixed order, are:

```text
medium / one-percent
dense-sparse / one-percent
medium / one-edit
```

Every child uses the same in-memory source, context limits, expected output,
semantic hash, inverse proof, and sink contract bound in the final 0964 plan.
The source and candidate are not rebuilt. The workload command is otherwise
the existing benchmark invocation with a new output directory under a fresh
0965 packet.

Each pair is adjacent: the second child is launched only after the first has
written its terminal receipt and its process group has been reaped. No pair
runs concurrently, and no child is retried or replaced. The pair order is
rotated by case to avoid making one pair type always occupy the same wall-clock
position:

| case | pair order | children |
| --- | --- | --- |
| medium one-percent | `RR`, `RC`, `CC`, `CR` | `R R R C C C C R` |
| dense-sparse one-percent | `RC`, `CC`, `CR`, `RR` | `R C C C C R R R` |
| medium one-edit | `CC`, `CR`, `RR`, `RC` | `C C C R R R R C` |

Here `R` and `C` in the children column are binary roles, and each two-letter
entry is one adjacent pair. The schedule has exactly one `RC` comparison, one
`CR` comparison, one `RR` control, and one `CC` control per case. The launch sequence, UTC times,
binary hash, case, pair, and position must be recorded in an immutable child
receipt.

Use the corrected child-affinity method from 0964: set the child affinity in
the child pre-exec hook before the parent observes or waits, then verify the
child PID reports exactly `[12, 13, 14, 15]`. Retaining a `taskset` argument is
fine, but observing only the wrapper PID is invalid. Use `LANG=C`, `LC_ALL=C`,
`TZ=UTC`, the 0964 working directory, and the same `/usr/bin/time -v` resource
recording. Capture a fresh host snapshot before and after the packet, including
CPU/load, online affinity, memory, filesystem, and process scheduling
settings; do not alter governors or claim an exclusive host or cold cache.

The runner must stop on the first launch, terminal, oracle, affinity, or
binding failure. Terminal receipts must precede report parsing. A partial
packet is invalid for classification and must be sealed as such; it must not
be completed with replacement children.

The 256-sample choice is deliberate. The existing nearest-rank p99 rule then
uses the third-highest sample rather than the second-highest sample at 200
samples, while retaining the raw vector for an independent audit. It is still
a diagnostic tail estimate, not a confidence interval. No samples from 0964 or
the historical trials may be pooled with this packet.

The 20 warmups also differ from 0964's 10-warmup screen. This is a fresh
diagnostic packet; its magnitudes must not be presented as a direct numerical
replication of the 0964 magnitudes.

This is a whole-harness diagnostic. It cannot isolate the constructor from
state created by earlier warmups or from allocator/kernel state carried across
the process lifetime. That is intentional: the 0964 `ingress_ns` contract is
the constructor in the same finite transaction harness, and changing that
state boundary would test a different workload. Sample-index and top-five
descriptors can expose a warm-state pattern, but absence of such a pattern is
not proof that state is irrelevant.

## Reduction and independent audit

For each child, calculate nearest-rank ingress p50, p95, p99, and max from its
256 raw `ingress_ns` values. Retain the same four quantiles for
`edit_commit_ns`, `publication_ns`, `commit_drop_ns`, and `total_ns` as
context. Also record the mean of the top five ingress samples and the sample
indices of the p99 and top-five values; these are diagnostic descriptors of a
single spike or a sustained tail, not replacement gates.

For each adjacent pair, report the exact candidate-over-baseline ratio for the
one `RC` pair and the one `CR` pair. For a same-binary pair with first-child
quantile `q1` and second-child quantile `q2`, report the magnitude
`max(q2 / q1, q1 / q2)` for `RR` and `CC`; this is symmetric in launch order.
Pair ratios are computed only from children in the same fresh packet. Report
raw quantiles as well as ratios, and report a `>1.05` flag for every cross
order and same-binary magnitude. Also report the cross-order spread
`max(r_RC, r_CR) / min(r_RC, r_CR)` and the exact p95/top-five ratios. Do not
pool all measured iterations and do not bootstrap the one-pair units into a
false confidence interval. The raw independent audit must recompute all
quantiles, ratios, flags, child identities, and oracle/hash checks without
importing the reduction result.

The audit must additionally check:

1. all 24 children have exactly 20 warmups and 256 measured rows;
2. every child is fresh, serial, affinity-verified before wait, terminal before
   parse, reaped, and resource-clean;
3. each pair's source/output/semantic/inverse and deterministic case oracles
   match the final 0964 plan;
4. baseline and candidate binary hashes are exactly the two retained hashes
   above; and
5. no report path or sample vector is imported from 0964, 0961, 0960, or 0959.

## Predeclared interpretation

These are diagnostic labels, not performance decisions. The 0964 rejection
remains the recorded result in every branch.

* **Concordant.** Both `RC` and `CR` p99 ratios are above `1.05`, and both
  same-binary magnitudes are at or below `1.05`. Report the exact two ratios,
  their cross-order spread, all p95/top-five ratios, and all flags. This label
  describes the observed cross-order pattern in this packet; it is not a
  source, build, host, or causal attribution.

* **Nonreproduced.** Both `RC` and `CR` p99 ratios are at or below `1.05`, and
  both same-binary magnitudes are at or below `1.05`. Report the exact ratios
  and spread. This label says only that the 0964 cross-order p99 pattern did
  not recur above the bound in this packet.

* **Mixed.** The two cross-order flags differ, or either same-binary magnitude
  is above `1.05`. Report the exact values, flags, spread, and
  sample-index/top-five descriptors. Do not average mixed orders into a causal
  conclusion.

These labels are deliberately limited to `concordant`, `nonreproduced`, and
`mixed`. A failed custody, oracle, or control condition gives the packet status
`invalid` and receives no diagnostic label; it is sealed rather than repaired
or adaptively extended. No label changes the 0964 safety decision.

The same rules apply per case. A result in only one medium case is a scoped
diagnostic observation; it is not a general XLSX claim. A
`CC` flag is an observed same-binary pattern in this packet; it does not
identify a candidate executable, warm-state, compiler-layout, or source
difference.

## Follow-up boundary

Only if root elects to investigate after this diagnostic should a separate
prospective plan address build provenance. That plan would need at least two
independently built candidate executables from the frozen source, identical
toolchain/profile/lock bindings, and the same `RR`/`RC`/`CR`/`CC` controls. It
would be a new build and a new packet. It must not reuse this diagnostic's
classification as an adoption signal.

If the diagnostic is `nonreproduced` or `mixed`, stop this line and retain the
0964 rejection. If it is `concordant`, perform source-level attribution or a
new mechanism hypothesis before changing production; the label itself does not
attribute source, build, host, or allocator behavior. Ordinary, fallback,
malformed-input, and namespace/MCE/BOM/limit controls remain mandatory for any
future candidate qualification.

## No-work performed for this design

This document was written after reading `docs/GOAL.md`, the 0964 final trial,
the final epoch-2 plan and raw audit, and the 0959–0961 records. No benchmark
process, workload, source edit, or build was started while preparing it.
