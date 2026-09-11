# 0513 ADR and harness boundary review

This review is tied to base revision `53a7c4a523e8518a7150d259ccf3435cb134e4a4` and the accepted manifest at [`adr-manifest.json`](adr-manifest.json). The manifest is the existing 30-file set: ADRs 0001 through 0029 plus [`docs/adr/README.md`](../../../adr/README.md). Its recorded SHA-256 is `bedcedc396b998354ade6a06a9864ad6e5978611dd391760e3b1ecd38056af34`; recomputing each of the 30 entries at review time matched the manifest. No accepted ADR or the ADR hierarchy was changed or superseded by 0513.

The repository ADR hierarchy continues to put correctness, lossless preservation, and safety ahead of performance; keeps the public facade concise; requires representative measurements for performance decisions; and treats production-readiness claims as evidence-gated. The accepted manifest therefore remains the governing design record for this evidence-only harness change.

## Change and ownership

The candidate source scope is limited to two files under `tools/perf-baseline`:

* `src/lib.rs` adds operation allocation observations to the existing XLSX commit and commit-plus-save cases. It also adds the `#[inline(never)]` `xlsx_commit_save_operation` helper, whose body delegates to the existing `Edit::commit()` and `Commit::workbook().write_to(sink)` APIs.
* `src/xlsx_commit_metrics_tests.rs` adds focused tests for observation alignment, normal-binary unavailable allocator status, exact save output, semantic commit behavior, and bounded-sink error propagation.

No `crates/` file, format implementation, public API, archive layer, parser, serializer, or workspace topology file is in this change. This is an OOXML/XLSX measurement enabler. OLE2 behavior is unchanged in this batch, and ODF work remains deferred until the OLE2/OOXML optimization goal is complete, as required by the plan.

## Boundary review

| Concern | Current boundary | ADR and semantic disposition |
| --- | --- | --- |
| XLSX commit timing | Existing setup (`Workbook::from_bytes`, update staging, and edit creation) remains before `Instant::now()`. The timed call remains `edit.commit()?`; verification, `black_box`, and `Commit` drop remain after the clock. | Preserves ADRs 0001, 0003, 0005, and 0006: transactional commit and validation behavior are still measured at the same operation boundary, while fixture construction and teardown stay out of it. |
| XLSX commit-plus-save timing | Expected-output construction, its commit, sink sizing/reservation, and other setup remain before the loop. The timed call is the helper; the helper performs exactly `edit.commit()` followed by `commit.workbook().write_to(sink)`. Sink equality, reopen, and cell verification remain after the clock. | Preserves the existing operation definition and the ADR 0010/0011/0017/0018 ownership split: the harness calls the public OOXML writer and does not move package, producer-template, or calculation-chain logic into the harness. |
| Allocation observations | `allocation_metrics::begin()` is immediately before the existing operation clock and `finish()` is immediately after elapsed time is taken. Only retained, post-warmup observations are published. | Satisfies ADR 0005’s measured-performance requirement without presenting setup, oracle work, or `Commit` destruction as operation allocation. Normal binaries report explicit unavailable status; they do not turn missing counters into zero allocations. |
| Callgrind target | `#[inline(never)] fn xlsx_commit_save_operation(...)` is the exact post-oracle boundary. The profile command collects only this symbol, with collection disabled at start. It includes the helper’s commit and sequential write, and returns the `Commit` before the caller drops it. | Gives a stable attribution target while preserving ADR 0011’s package ownership and ADR 0003’s commit lifetime/transaction semantics. Generator and expected-output work are excluded by construction and by the collection toggle. |
| Observation reporting | Commit-only rows use in-process elapsed/allocation observations without a sink. Save rows retain the deterministic `CountingSink` summary and combine it with aligned in-process observations. | Keeps the existing sink contract and output result shape; the added data is diagnostic evidence rather than a production API change. |
| Focused tests | The new test module exercises the four existing XLSX commit/save case selectors and the helper’s exact-output, semantic, and short-sink paths. | Supports ADR 0008’s verification gate and checks the measurement plumbing without changing format behavior. |

The allocator region is intentionally narrower than the full case invocation. It excludes source opening, edit staging, update generation, expected-output/oracle construction, sink reservation, post-operation comparison and reopening, semantic verification, and `Commit` destruction. Warmup regions may execute but are discarded; published allocator and elapsed vectors use the same retained sample set and sample indices. The save helper itself is only a named harness boundary: it introduces no new production operation and does not alter error propagation or output ordering.

## ADR applicability

The accepted decisions that directly govern this change remain satisfied:

* ADRs 0001, 0003, 0005, 0006, and 0008 continue to require correct transactional behavior, full validation before publication, scoped representative measurements, and evidence before support or speedup claims.
* ADRs 0010, 0011, 0017, and 0018 keep facade, OPC package, OOXML producer-template, and calculation-chain ownership in their existing layers. The helper only invokes the existing workbook writer after the existing commit.
* ADRs 0002 and 0024 keep the workspace and crate topology unchanged; test-only harness code does not create a new production crate or dependency edge.
* ADRs 0004, 0007, 0012–0016, and 0019–0022 concern semantic/public or other Office-format ownership that this source scope does not touch. ADR 0025 is likewise unchanged.
* ADRs 0026 and 0027 preserve OLE directory metadata and XLS sheet-anchor ownership. They remain part of the OLE2/OOXML priority, but no OLE2 or XLS implementation is modified by 0513.
* ADRs 0009 and 0023 preserve the ODF ownership and crate-split decisions. ODF remains deferred under the stated priority ordering.
* ADRs 0028 and 0029 concern the excluded iWork/IWA work and are unchanged.

There is consequently no ADR conflict, exception, or supersession to record. The manifest remains accepted evidence of the design baseline rather than a file being amended by this batch.

## Evidence limits

The native protocol reuses the current clocks and four existing XLSX commit/save selectors across the planned tiny, medium, and dense-wide corpus rows. It is a descriptive instrumentation-overhead check (serial ABBA, one worker, CPU 2, 100 retained samples and three warmups); it is not a registered before/after speedup claim. The allocator protocol uses the separate allocator binary and its own samples/repeats. Instrumented timing and RSS are excluded from allocator conclusions.

The allocator counters are process/global system-allocator observations collected only inside the region described above. They can reflect allocator activity from the instrumented process during that interval and the measurement mechanism can perturb execution. A normal binary has no active counter and therefore reports `unavailable`; absence of an observation is not evidence of zero allocation. These measurements do not establish per-function, per-thread, or production-wide allocation cost.

The Callgrind result is synthetic instruction attribution for three dense one-percent commit-plus-save helper calls with zero warmups. It excludes fixture generation, expected-output work, and caller-side teardown, but it is still Valgrind instrumentation data rather than wall-clock latency, hardware cycles, or a scaling result. Direct helper attribution is trustworthy for that stated region when the profile has the expected three helper calls and three writer calls; it must not be generalized to the complete runner or to other formats.

These limits are consistent with ADR 0005 and ADR 0008: 0513 closes an operation-observation gap and supplies a stable profiling boundary, but it does not claim an optimization or change production semantics. Any later OLE2/OOXML implementation optimization still needs matched native, allocator, and appropriately scoped profile evidence against the unchanged ADR manifest and guardrails.

## Review conclusion

0513 is accepted as a harness-only, operation-scoped evidence change under the unchanged ADR manifest. The existing commit and save work, timer semantics, sink behavior, oracle checks, validation, error paths, and object lifetimes remain intact. The new observations make their scope explicit and preserve unavailable status where the selected binary cannot measure allocations. OLE2/OOXML remains the active optimization priority; ODF is deferred, and no iWork claim is made.
