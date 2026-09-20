# ODS lookup and reference-function disposition

Commit `6fa3b8af6a` implements ADDRESS, CHOOSE, HLOOKUP, INDEX, INDIRECT,
LOOKUP, MATCH, OFFSET, and VLOOKUP in the explicit, read-only ODS evaluator.
Semantic and resource review pass. The disposition accepts the implementation
with the measured performance limitations below; it does not claim an overall
speedup, complete OpenFormula support, or workbook recalculation.

The reviewed source is the 62-input freeze
`53261588afdad7e1efd29ac91c0f393f41e8ddbe1032f91ca9d60dcb0c86fa60`,
over baseline `635fd2e1348b621426b50909cbd5765c91837306`.
[The gate archive](gates/README.md) retains the exact selected inputs. Gates and
profiling use the retained isolated lockfile; the ambient root lockfile remains
unchanged and is not substituted into that evidence.

## Behavior and boundaries

Reference constructors retain bounded descriptors until a consumer requires
values. CHOOSE evaluates selected branches only, including projected matrix
demands. Search functions preserve complete data operands while scalar keys
and selectors retain coordinate sensitivity. Known invalid data shapes and
selectors refuse before the corresponding resolver reads. Known scalar key
errors retain their identity over formula-level data-shape refusals.

Exact search streams candidates with fixed comparison state. Approximate search
uses bounded probes and the specified duplicate/type rules. The selected result
is read only after a match. Resolver text remains borrowed, allocations and work
are charged, and typed provider/resource/cancellation/source failures remain
uncatchable by formula error handlers. Independent
[semantic](semantic-review.md) and [resource](resource-review.md) reports describe
the full scope and are bound to this freeze by [review-receipt.json](review-receipt.json).

Computed CHOOSE shape probes retain their complete selected results for the
specific projected demand. Widening a dynamic reference may revisit a
position-sensitive selector. The direct dynamic INDIRECT scaling evidence does
not establish a universal single-read guarantee for computed compositions.
Source-qualified external references remain inert. GETPIVOTDATA,
MULTIPLE.OPERATIONS, pivot calculation, dependency recalculation, and
formula-cache publication remain outside this batch.

## Validation

All seven isolated gates pass: 1,700 tests with no failures or ignored tests,
strict all-target Clippy, strict rustdoc, package and selected-source formatting,
crate boundaries, and diff checks. Source manifests are stable before and after
execution. Focused evidence contains 34 semantic tests, 9 resource tests, and
127 independently generated oracle observations. Native evidence retains 32
rows: 30 function observations (17 parity and 13 documented divergences) plus
two fixture observations. Native behavior does not override the reviewed ODF
contract.

All 121 performance-harness cases pass result and resolver-read preflight.
The retained pair contains 990 baseline and 3,630 candidate samples: 33 matched
controls, 88 added lookup cases, two phases, three warmups per child, and 15
fresh child samples per group. All five rejected preflight attempts and the
successful preliminary preflight are retained. No rejected attempt produced
timing samples, and no successful capture was replaced to improve its results.

## Performance acceptance and limitations

[The raw profile report](performance/results/performance-report.md) and
[independent retained audit](root-performance-audit.json) describe this one
capture. Source, harness, profile inputs, raw sample files, exact normalization,
and cleanup receipts are verified. Across all 66 matched case/phase groups,
allocation, requested/released bytes, peak live allocation, retained budget,
work, reads, output bytes, and checksums are unchanged.

There is a measured evaluate-only SUMIFS regression: median latency rises from
84.4805 to 95.1255 microseconds per evaluation, **+12.6005%**. A descriptive
bootstrap interval is +8.0045% to +17.2941%. The parse-evaluate lane changes by
-1.4629%, with interval -6.3439% to +7.9052%; that observation does not erase the
evaluate-only regression. The audit also flags 36 matched RSS increases above
5%, ranging from 184 to 536 KiB. These are process RSS observations, distinct
from the unchanged measured allocator accounting.

The SUMIFS/criteria modules and the common cell-work, resolver-read, and
borrowed-text conversion functions are unchanged from the baseline. That
source comparison does not establish the cause of the latency or RSS changes.
No timing change is dismissed as noise, attributed to a specific compiler
effect, or presented as a causal optimization result. These limitations are
accepted for this functionality batch and remain visible for subsequent
performance work.

The bootstrap settings are retrospective descriptive audit settings (seed
20260920, 10,000 resamples, 95% intervals), not a preregistered acceptance test.
Every matched comparison is retained, including all 37 latency/RSS review
flags. Measurements use a shared host and do not claim isolated-host timing or
end-to-end Office-document performance.

## Cleanup and final verification

The owned gate checkout and build target have been removed after verifying all
62 archived source inputs. [cleanup.json](cleanup.json) records this removal;
the profile retains its own target and baseline-worktree cleanup receipts.
Unrelated files and temporary directories remain untouched. The final
[aggregate verifier receipt](verification.json) reports `verified: true` from
retained files alone: nine function coverage groups, thirteen cross-cutting
requirements, source closure, reviews, oracle/native evidence, seven gates, and
all 4,620 performance samples. The broad spec-gap implementation goal remains
active after this batch.
