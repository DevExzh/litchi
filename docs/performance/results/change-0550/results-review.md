# 0550 XLSX source-backed commit attribution results review

`status: complete`  
`disposition: diagnostic-only; retain the measured source`  
`performance_claim: none`

The two canonical analyzers pass. The metrics analyzer covers 40 jobs: 8 preflight, 16 native, and 16 allocator jobs. It retains 480 native samples and 480 allocator samples. The paired runs have equal corpus, sink, source, output, semantic, and untouched-member identity for all eight case/shape groups. The report has 22 same-baseline repeat rows above the absolute 5% review threshold; every row is retained and individually interpreted in [`adverse-review.json`](adverse-review.json).

The profile analyzer also passes for eight one-percent profile jobs. It classifies 26 lifecycle owner dumps outside the selected operation and eight measured dumps for the exact `litchi_xlsx::cell_values::source::MultiSourceEdit::commit` owner. Its four shape-pair repeat comparisons contain six immediate-child rows above 5%; those six rows are retained in the JSON. No nested-target profile row crosses that threshold.

These are same-binary repeat observations. They are diagnostic evidence about the retained run and do not establish a speedup, regression cause, or general timing stability. There is no candidate comparison or admission gate in 0550.

## Evidence bindings

| Artifact | SHA-256 |
| --- | --- |
| `metrics-analysis.json` | 49e7978f1dd41b21031abe45cc2394a51085ca701d17fca79eda4ab5a8de14e4 |
| `profile-analysis.json` | 64ee681a71674696e65f3d36922a03dec950e8af5d2e49ecf00b1c65c5dfd0ae |
| `analyze_metrics.py` | 95cc7baadb0dcefff7c4e29bad84321d05ff71073b5f8522e26945622f359c0f |
| `analyze_profiles.py` | ac23bbe997c8753557c141752ea2bb2c3e59d634a549d75b6ecd517aabcd0fbc |
| `analysis-inputs.json` | 0878a2a8963c13d91fe1ff365f544eb5e3eb23350706fb5e6f21e9de3aaf8435 |
| `analysis-runs/metrics/receipt.json` | c49fe8ab3449754bfef21dd50d9e8e109900fabdaffdb42df92e4198be094cd2 |
| `analysis-runs/profiles/receipt.json` | 609326970ccd6efac0d7bbfdf02cd697ade55da1b329abc2eaa195549c79a5a3 |
| `plan.json` | 3eeb52cdad1312b3953006e538490d6523a707e9858354dcb6688b4a615471a7 |
| `run.py` | dcb20341f9558bd6de1befcfe0afe43d2045c90c6c7c5673529fe234d2e6d39f |
| `capture.py` | 4df2b10d6e884effcb160a07bc811d7b9d45b0d35c227ebf8d4c7227cf702463 |
| `baseline/source-manifest.json` | fa85f76972a37f52644c60b955c9125ed80e4d78e2434ec54b77c3cb19e64163 |
| `source-review.md` | 0bfb5748fe0b7d8a528e43bd1ce1d4a32a2962b45bb4595fa8f636a7a3dd7430 |
| `scope-review.md` | b97b942863dc657ddd62c7343e99630dda9b26b056b8aec0515414aaaa929f05 |
| `profile-review.md` | 7aec1262e21c5932e8ac2e0db11a0825ce06f4eac2e45f294e58f9bd4b910411 |
| `change-0549/next-target.md` | 5b9e75255cfe7010d4ca1bb9426ff389e235a76034d093fb5d14d7c58b930ae4 |
| `change-0512/scope-review.md` | 25f0a25cab7722fd154a1f7b201705401fd36fe7ccebcea353f7333aa25f8ead |

The metrics analyzer is the final SHA-256-pinned script `95cc7baadb0dcefff7c4e29bad84321d05ff71073b5f8522e26945622f359c0f`; its canonical report is `49e7978f1dd41b21031abe45cc2394a51085ca701d17fca79eda4ab5a8de14e4`. The profile report and analyzer are bound separately because the profile lane has its own scope and callgraph classification.

## The 22 metric drift rows

The threshold is applied to the absolute change between repeat 1 and repeat 2 of the same baseline binary. Positive and negative rows are retained alike. All 22 rows are native timing statistics; the allocator lane has zero rows above the threshold.

| # | Case | Shape | Metric | Repeat 1 | Repeat 2 | Delta | Change |
| ---: | --- | --- | --- | ---: | ---: | ---: | ---: |
| 1 | one-edit | dense-sparse | `timing.open.p99` | 106,550 | 116,411 | 9,861 | +9.254810% |
| 2 | one-edit | dense-sparse | `timing.reopen.p50` | 44,525,364 | 41,881,369 | -2,643,995 | -5.938177% |
| 3 | one-edit | dense-sparse | `timing.reopen.mean` | 44542536.13333332 | 42060433.19999999 | -2482102.93333333 | -5.572433% |
| 4 | one-percent | dense-sparse | `timing.open.p50` | 100,020 | 92,021 | -7,999 | -7.997401% |
| 5 | one-percent | dense-sparse | `timing.open.mean` | 100898.36666666667 | 92828.43333333335 | -8069.93333333332 | -7.998081% |
| 6 | one-percent | dense-sparse | `timing.open.p95` | 109,281 | 100,521 | -8,760 | -8.016032% |
| 7 | one-percent | dense-sparse | `timing.open.p99` | 114,001 | 101,870 | -12,131 | -10.641135% |
| 8 | one-percent | dense-sparse | `timing.reopen.p50` | 45,876,230 | 39,406,668 | -6,469,562 | -14.102209% |
| 9 | one-percent | dense-sparse | `timing.reopen.mean` | 45931694.86666665 | 39402987.00000001 | -6528707.866666645 | -14.213949% |
| 10 | one-percent | dense-sparse | `timing.reopen.p95` | 46,530,243 | 39,749,415 | -6,780,828 | -14.572948% |
| 11 | one-percent | dense-sparse | `timing.reopen.p99` | 46,614,203 | 39,979,655 | -6,634,548 | -14.232889% |
| 12 | one-percent | noncompact | `timing.open.mean` | 95427.23333333332 | 89261.8 | -6165.43333333332 | -6.460874% |
| 13 | one-percent | noncompact | `timing.open.p95` | 106,090 | 92,691 | -13,399 | -12.629843% |
| 14 | one-percent | noncompact | `timing.open.p99` | 114,491 | 99,920 | -14,571 | -12.726765% |
| 15 | one-percent | noncompact | `timing.publication.p50` | 6,707,634 | 6,066,781 | -640,853 | -9.554084% |
| 16 | one-percent | noncompact | `timing.publication.mean` | 6715497.933333333 | 6068999.066666667 | -646498.8666666653 | -9.626968% |
| 17 | one-percent | noncompact | `timing.publication.p95` | 6,768,389 | 6,097,307 | -671,082 | -9.914944% |
| 18 | one-percent | noncompact | `timing.publication.p99` | 6,791,430 | 6,109,257 | -682,173 | -10.044615% |
| 19 | one-percent | noncompact | `timing.reopen.p50` | 28,260,163 | 32,492,317 | 4,232,154 | +14.975689% |
| 20 | one-percent | noncompact | `timing.reopen.mean` | 28271223.6 | 32446247.26666667 | 4175023.666666668 | +14.767750% |
| 21 | one-percent | noncompact | `timing.reopen.p95` | 28,589,835 | 32,703,804 | 4,113,969 | +14.389621% |
| 22 | one-percent | noncompact | `timing.reopen.p99` | 28,897,736 | 32,795,384 | 3,897,648 | +13.487728% |

Rows 1–3 are the one-edit dense-sparse pair. Rows 4–11 are one-percent dense-sparse, rows 12–22 are one-percent noncompact. Open and publication changes are outside the exact owner; reopen is explicitly diagnostic and excluded from the workflow total. The JSON retains each canonical row, including its exact values, lane, case, shape, metric, repeat numbers, threshold, delta, and change.

## Profile drift and owner scope

| Shape | Function row | Repeat 1 | Repeat 2 | Change |
| --- | --- | ---: | ---: | ---: |
| dense-sparse | `__rustc::__rust_dealloc` | 255 | 140 | -45.098039% |
| dense-sparse | `alloc::raw_vec::RawVecInner<A>::finish_grow` | 382 | 744 | +94.764398% |
| medium | `alloc::raw_vec::RawVecInner<A>::finish_grow` | 413 | 489 | +18.401937% |
| medium | `core::ptr::drop_in_place<litchi_xlsx::cell_values::snapshot::SourceState>` | 2,189 | 2,304 | +5.253540% |
| noncompact | `alloc::raw_vec::RawVecInner<A>::finish_grow` | 373 | 353 | -5.361930% |
| vendor-extension | `core::ptr::drop_in_place<litchi_xlsx::cell_values::snapshot::SourceState>` | 1,987 | 2,189 | +10.166080% |

The six profile rows are immediate-child IR rows: two for dense-sparse, two for medium, one for noncompact, and one for vendor-extension. They are repeat drift in the same binary, with profile/native output identity equal. The named allocation and deallocation functions are Callgrind instruction rows, not allocator counts. The exact owner’s inclusive repeat changes remain separate, and nested validation/rewrite/merge rows are not summed.

The profile reset and dump boundaries are source-backed and shape-specific: medium, dense-sparse, and noncompact have three lifecycle `MultiSourceEdit::commit` calls before the measured call, so the measured dump is `.4`; vendor-extension has four lifecycle calls, so it is `.5`. The measured caller is `litchi_perf_baseline::run_xlsx_cell_values_edit_save`; lifecycle calls are from `run_xlsx_cell_value_lifecycle_gates`. The owner profile begins after staged `edit.set` work and excludes publication, reopen, and returned-commit destruction.

Native `commit_ns` includes edit staging plus `MultiSourceEdit::commit`; the workflow is open + planning + commit + publication, while reopen is diagnostic. Allocation metrics come from the separate `operation_global_system_allocator` binary: planning covers selector planning, the commit region includes staging plus commit, and instrumented elapsed is excluded. These scopes prevent the 22 timing rows or allocator vectors from being presented as owner-only latency or memory evidence.

## Semantic coverage

Both the one-edit and one-percent source-backed cases construct `MultiSourceEdit`; neither invokes the cell-values `SourceEdit::commit` path. The selected edits are scalar numeric replacements. The matrix does not measure `SourceEdit`, managed-budget execution, formulas or shared formulas, styles, shared strings, rich values, row/column edits, no-op performance, malformed output, or other content classes. Lifecycle gates exercise some correctness boundaries, but they do not expand this performance coverage.

Output identity equality proves matched bytes and semantic/untouched-member identity for the captured case/shape lanes. It does not prove broad Office compatibility, physical-cold behavior, provider behavior, concurrency, scaling, fuzz coverage, or a native-Office result.

## Next attribution target

The first-level owner partition consistently makes `rewrite_value_only_with_provenance` the largest disjoint child. The next narrow diagnostic should follow its direct `scan_with_limit` child. In medium repeat 1 that child is exactly 70,977,241 Ir; the corresponding repeat-1 child rows are 137,507,562 for dense-sparse, 91,137,124 for noncompact, and 71,003,018 for vendor-extension. These are nested direct-child attribution values, so they are not added to the owner partition or to `Snapshot::from_rewritten_value_source`.

The reduced-readback/parser source-level symbol is absent from the positive owner-reachable graph in the profile report. That absence may reflect inlining or codegen omission and is not a zero-cost observation. The `validate_xml` aggregate also includes workbook checks reached through invalidation-related calls, so it is not a worksheet-only cost. A future profile should split scan, writer, validation, reduced-readback/parser, and merge work with route, copy, and allocation counters before any production change is considered.

The current source is therefore retained and this batch closes as diagnostic-only. Preserve complete-output validation, reduced-readback/fallback, source lineage, atomic publication, and semantic/error contracts. OLE2/OOXML remains the active priority; ODF remains deferred and iWork is excluded.

The report does not make a blanket speedup or stability claim. The 0550 evidence is bounded to the exact source revision and captured matrix.

