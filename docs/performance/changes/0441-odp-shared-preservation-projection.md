# 0441: Share ODP's immutable preservation projection

Starting an ODP transaction previously deep-copied its validated slides into
both the editable draft and a pristine comparison model. The latter now shares
the snapshot's immutable Arc projection. The draft stays detached; source-page
coverage refusal, origin mapping, semantic comparisons and exact source markup
retention remain unchanged. No public API or dependency changes.

The matched owned append experiment retains 24 reports, 720 samples and four
profiles across 64/4,096/8,192 source slides. All 60 allocator observations per
role and size agree. Medium/large peak above operation entry falls
19,993,648 → 18,027,568 bytes and 39,890,548 → 35,958,388 bytes, reductions of
9.834%/9.857%. Retained live delta is unchanged; source and commit remain live
at the endpoint and append still materializes the complete document.

**The original frozen practical gate failed.** Medium/large normal p50 is
0.664–1.279% slower; allocation calls fall only 1.398%/1.403%, and requested
bytes 1.708%/1.778%. The root omitted peak memory from its gate despite naming
peak in the hypothesis. The change is kept under an explicit
[post-hoc memory review](../results/change-0441/acceptance-review.md) against the
program's memory objective. The frozen protocol and failed gate are unchanged.
This is not a normal latency or RSS improvement claim.

No matched adverse change exceeds 5%. Baseline tiny normal p99 has a +5.807%
repeat flag. Whole-process cycles/instructions fall 0.713%/0.279%; setup and
oracle work are included, so these do not establish a timed-operation speedup.
Baseline/candidate symbolization retains 13/11 warnings per conversion. See
[measurements](../results/change-0441/measurements.md) for all values and scope.

All 352 ODP tests pass, including strengthened sharing, peer-isolation and
source-lifetime checks. The full harness, strict owner Clippy, rustdoc,
formatting, boundaries and portable evidence gates are recorded in the
[bundle](../results/change-0441/README.md). The two setup failures are retained;
no measured attempt was excluded or replaced.

The registry remains 436 selectors and the default matrix 36 cases. Repeated
staging XML scans are the next CPU hypothesis; one-shot caching, bounded
existing append, Part addition, repackaging, native breadth, cold/range and
scaling remain open under the unchanged non-iWork goal.
