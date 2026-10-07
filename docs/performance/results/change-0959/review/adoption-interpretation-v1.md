# 0959 Arc Number adoption interpretation

Status: **hold; do not adopt**. This is an independent read-only interpretation
of the sealed 0959 reduction. No query-control build or workload was launched.
The selected-cell control remains prospective and cannot change this decision
without a separately authorized, bound control run.

## Evidence binding

Primary reduction:

- `/home/zhuhe/litchi-performance-0959/analysis/comparison-v3.json`
- SHA-256 `9dbaefcd25b348c6b3f3c23ff4ed94b6e160fa7c95a5e37667a7efafa80242eb`
- schema `litchi.performance.0959.reducer.v3.final-plan-adapter`
- status `reduced`

The reduction binds plan
`/home/zhuhe/litchi-performance-0959/plan-final-v8.json`, SHA-256
`9964f2f3a86dd467de6d7840807400aa7468a48f7d993b50ea296c6e0909abf8`, and the
sealed capture terminals:

- native 72-child terminal SHA-256
  `b3876c4924f2aa88bbfab78a65e1522ec5227c5560efa5a2cd386cb4d2d8e8ab`
- allocator 18-child terminal SHA-256
  `f66b26f72bbb837fcfa8cba12a82860605b7f49fa84b7ea39f76090bc77674a2`

The native schedule is six blocks per case, 200 measured samples and 20 warmups;
the allocator schedule is three paired repeats per case, 10 samples and two
warmups. The reduction's bootstrap uses 10,000 resamples with seed 959.

## Practical result

The two one-percent cases pass the native safety bound but show no primary
improvement:

- `medium-one-percent`: edit-commit p50 candidate/baseline `1.020503`
  (CI upper `1.026544`), total p50 `1.009462` (CI upper `1.019175`).
- `dense-sparse-one-percent`: edit-commit p50 `1.030312` (CI upper `1.040188`),
  total p50 `1.012034` (CI upper `1.018060`).

The medium one-edit case fails the native safety bound in every primary area
that matters for adoption:

- edit-commit p50 `1.039509`, CI upper `1.052060`;
- edit-commit p95 `1.046893`, CI upper `1.061557`;
- total p50 `1.088054`, CI upper `1.136995`;
- total p95 `1.101930`, CI upper `1.145380`;
- whole-process max RSS `1.039052`, CI upper `1.061442`.

The phase controls show the cost is not limited to one noisy total. Medium
one-edit publication p50 is `1.152971` and commit-drop p50 is `1.073494`.
Commit-drop also rises consistently in the two one-percent cases: `1.119440`
for medium and `1.162973` for dense-sparse. Publication is slightly lower in
those one-percent cases, but that does not offset the one-edit failure.

Allocator counters show one possible mechanism benefit but not an adoption
benefit: allocation-call ratios improve to `0.913641`, `0.917987`, and
`0.912020` for medium one-percent, medium one-edit, and dense-sparse,
respectively. Allocated-byte ratios remain slightly higher (`1.005164`,
`1.006367`, and `1.008723`), and region peak live ratios are also higher
(`1.006753`, `1.002747`, and `1.011227`). The call-count reduction therefore
cannot be reported as lower allocation volume or lower latency. RSS is a
whole-process metric and remains separate from allocator-region peaks.

The reducer itself reports `gates.all_safety == false`; this is sufficient to
hold adoption. The result does not justify relaxing the gate or selecting only
the favorable one-percent routes.

## Control and attribution limits

The current formal cases are managed source-backed compression cases, not the
marker-free selected-cell query control. The pending control uses the existing
`xlsx_producer_medium_control_selected_cell` and
`xlsx_producer_dense_control_selected_cell` routes. It is useful for locating
selected-scanner behavior, but it is not a public-constructor benchmark and
cannot reverse the failed one-edit gate.

The source path has a known cost asymmetry. Baseline selected scanning transfers
an owned scratch `String` through `Number::new(value.into_boxed_str())`; the
candidate calls private `Number::from_lexical(&value)`, which copies the already
owned scratch string into a new `Arc<str>`. A selected-cell regression would be
consistent with that mechanism. Conversely, public `Number::new` still accepts
`Into<Box<str>>` and then copies into `Arc<str>`, so the candidate may impose a
second allocation/copy on external callers. No timed benchmark in the sealed
0959 harness exercises that public constructor; its existing calls are outside
the measured intervals.

Accordingly, the candidate has no demonstrated practical adoption benefit in
this packet. Keep production unchanged. If further investigation is warranted,
run only the separately bound selected-cell control described in
`query-control-protocol-v1.md`, then treat its result as attribution evidence,
not as a waiver of the failed medium one-edit safety gate.
