# XLSX commit-local compact source-cell proof experiment

Status: candidate failed the frozen performance gates and baseline source is restored. All eleven final quality commands passed; no production speedup is claimed.

This experiment starts from `8aa0c5baf0616d16c79eba0c6c28dc1716338ad6` and follows the rejected planning-time proof in [change 0552](0552-xlsx-compact-source-proof.md). OLE2 and OOXML performance remain the priority; ODF optimization is deferred and iWork is outside this campaign.

The provisional candidate collects worksheet source spans only during an effective source-backed `MultiSourceEdit::commit`. It walks borrowed XML, checks source-cell order and addresses against the existing Store, and retains bounded row/cell spans until rewriting ends. It does not retain the proof in snapshots or construct another semantic Store. Uncertain structures, unsupported edits, malformed attributes, and resource refusals use the existing complete writer. Output validation, independent readback, and publication checks remain in place.

The [frozen plan](results/change-0553/plan.json) compares one-cell and one-percent workflows, including managed sources, across medium, dense-sparse, noncompact, and vendor-extension worksheets. Captures use baseline repeat one, both candidate repeats, then retained-baseline repeat two. Planning/refusal, cap, allocator, and RSS checks remain mandatory. Commit instruction profiles are conditional on the complete main, planning, and cap pilot passing. Final quality and individual adverse/drift review are required for disposition and both are complete.

The selected [draft07 binding](results/change-0553/candidate-binding.json) identifies the source and caps frozen before application. Targeted validation passed Clippy, 60 edit tests, and 74 public integration tests. The edit tests include direct compact/complete output comparisons, populated metadata-cap fallback, and test-only counters proving that planning, staging, empty commits, and effective same-value commits do not invoke the collector. A changed-value control observes one collector attempt and one accepted proof. The [independent source review](results/change-0553/collector-design-review.md) found no blocker in its five review areas.

Earlier candidate and check failures remain under [the evidence bundle](results/change-0553/README.md). The baseline's eleven quality commands and 1,313 passing tests are explicitly reused from change 0552 after source, lock, and artifact verification; they were not rerun as fresh 0553 baseline tests. Complete performance comparisons are retained. Fresh final-source validation on the restored baseline passed all eleven commands, including 1,313 tests in 59 groups.

## Measured outcome

All 119 receipts per stage completed successfully, and both frozen analyzers validated their full matrices. The candidate failed three required one-percent latency checks and one process peak-RSS check. Planning/refusal, cap, one-cell latency, workflow allocated bytes, and matched identity gates passed. The complete pilot failed, so instruction profiles were explicitly skipped; no gate was relaxed or timing capture rerun.

One-percent workflow p50 changes below are candidate relative to matched baseline; negative values mean less elapsed time. These are results for the rejected candidate, not retained library performance.

| Worksheet shape | Unmanaged R1 | Unmanaged R2 | Managed R1 | Managed R2 |
| --- | ---: | ---: | ---: | ---: |
| medium | -8.66% | -4.57% | -9.35% | -4.79% |
| dense-sparse | -3.55% | -4.07% | -5.21% | -4.12% |
| noncompact | -10.32% | -13.01% | -12.84% | -9.97% |
| vendor-extension | -2.56% | -1.53% | -5.55% | -3.88% |

The required improvement was at least 3% for every one-percent p50 and mean row. Unmanaged vendor-extension p50 improved only 2.56% and 1.53%; its repeat-two mean improved only 1.50%. Unmanaged dense-sparse one-percent repeat-two process peak RSS rose from 88,113,152 to 92,663,808 bytes, or 5.16%, exceeding the 5% ceiling. Process peak RSS is process-lifetime high water, not an operation-local allocation peak.

Workflow allocated-byte maxima fell 15.87–25.43% across the main cases and repeats. This did not override either failed gate. All 346 adverse/drift rows have individual retained reviews. Complete measurements are in [metrics-analysis.json](results/change-0553/metrics-analysis.json) and [guard-cap-analysis.json](results/change-0553/guard-cap-analysis.json); the [profile decision](results/change-0553/profile-decision.json) records the failed-pilot skip. The [ADR matrix](results/change-0553/adr-compliance.md) records the source constraints.

The next queued OLE2 mechanism is the [isolated name-only directory handoff](results/change-0553/ole2-name-candidate/source-note.md). Its patch passes static application checking but is unbuilt, unmeasured, and unadmitted. It still requires fresh attribution, correctness checks, and a separately frozen comparison.
