# 0703 — PPTX capture projection reuse diagnostic

Status: diagnostic complete; production unchanged. `performance_claim: none`.
Baseline revision: `f69d660800`.

Fresh traces confirm that changed opened-presentation commits repeat default
MCE processing of unchanged XML. For the real deck, 17 of 18 commit calls
match capture inputs and outputs after one edit; 16 match after two edits.
The capture produces 14 distinct owned projections with 543,070 bytes of
observed vector capacity. Reusing them across commit is a credible work-removal
candidate, but requires an explicit retained-memory policy. No cache or
production optimization is implemented in this batch.

## Scope and fresh results

The [packet](results/change-0703/README.md) runs two independent fresh-process
repeats of each source/workflow pair, with one transaction per process. Sources
are the 108,164-byte LibreOffice `slide-section-test.pptx` fixture and the
bounded `generated:12x8` recipe in the probe. Setup target discovery and
post-publication semantic verification are separate phases and excluded from
reuse counts. These instrumented runs contain no timers.

| Source / workflow | Capture calls | Distinct capture pairs | Commit calls | Commit calls matching capture | New commit pairs |
| --- | ---: | ---: | ---: | ---: | ---: |
| Real / no-op | 18 | 14 | 0 | 0 | 0 |
| Real / one edit | 18 | 14 | 18 | 17 | 1 |
| Real / two edits | 18 | 14 | 18 | 16 | 2 |
| Generated / no-op | 19 | 15 | 0 | 0 | 0 |
| Generated / one edit | 19 | 15 | 19 | 18 | 1 |
| Generated / two edits | 19 | 15 | 19 | 17 | 2 |

A pair binds raw SHA-256 and length, processed SHA-256 and length, ownership,
and the default profile/limits. Each source has four duplicate capture calls.
The real capture digests map uniquely to one presentation XML member and 13
slide XML members in the original archive. Presentation XML accounts for five
of the 18 calls. This member attribution comes from exact archive digest
matching, not from a caller or part-name field in the trace.

All workflows record one default MCE call during open and one during apply;
clone and edit record zero calls to this particular wrapper. The edit path
uses other processing APIs, so zero wrapper calls does not mean zero XML or
MCE work. Apply does not repeat the full capture sequence in these cases.
Exact no-op commits make no default MCE calls because they return the source
snapshot before recapture.

The real targets are zero-based slide 1 / shape 1 and, for the second edit,
slide 2 / shape 0. Generated targets are slide 0 / shape 0 and slide 1 / shape 0.
Both edited texts are checked after publication; no-op checks also require an
unchanged complete-package revision. All 1,112 default-wrapper calls in the
final 12 traces succeed. Phase counts, digest pairs, capacities and published
revisions match across fresh-process repeats. This is reproducibility evidence
for these fixtures, not a statistical performance sample.

## Retention cost and interpretation

The real capture's 14 unique raw payloads total 271,535 bytes. Its owned
processed content totals 312,859 bytes, while the observed output vectors have
543,070 bytes of capacity. Of that capacity, 509,058 bytes belong to distinct
capture pairs reused during the one-edit commit and 502,982 bytes during the
two-edit commit. The repeated presentation projection alone is 2,801 bytes of
content in a vector with 4,714-byte capacity.

These are sums of observed payload lengths and vector capacities. They exclude
cache metadata, retained source-owner overhead, allocator rounding, admission
bookkeeping, and transient overlap. They are neither a measured peak-memory
increase nor an RSS bound. Content digests and pointer matches do not prove
that a future address-keyed cache is safe; it would need a retained immutable
owner, an alias check for foreign `Part` implementations, complete profile
identity and an uncached fallback.

All 15 distinct generated capture results are borrowed. They total 56,323 raw
bytes and add no owned projection capacity. Retaining transformed output would
therefore have a very different value on this control. The generated source is
bound by the exact generator and observed XML digests; its original ZIP bytes
and archive digest are not retained. It is not claimed to be byte-identical
to earlier batches' generated archives.

## Decision and next work

Keep production unchanged. The [source review](results/change-0703/design-review.md)
identifies a small capture-local alternative: use the existing
`Presentation::catalog` memo for both the catalog limit check and slide capture.
That would avoid one duplicate catalog/MCE pass without retaining transformed
XML. Historical 0697 isolated measurements attributed only 3.32% of MCE sequence
time to all five presentation calls, so call counts alone do not establish this
small change as the highest-impact next optimization.

The larger measured opportunity is unchanged slide processing across a changed
commit. A later candidate must define and price aggregate cache admission,
source ownership, invalidation, fallback, and publication lifetime before
implementation. The existing `max_retained_candidate_bytes` limit specifically
governs serialized cross-slide-copy archives and cannot silently become a
transformed-XML cache allowance. This diagnostic establishes repeated work and
observed capacity; it does not establish acceptable memory tradeoffs or speedup.

Strict/foreign parts, error precedence, notes edits, custom capabilities,
malformed inputs, pressure fallback and changed-large-slide workloads remain
necessary production-candidate coverage. This packet traces successful
no-op/one-edit/two-edit workflows only; it is not typed-refusal evidence.

## Reproduction and verification

Run `build.py`, `run.py`, and `focused-summary.py` in the linked packet.
The build driver applies only the retained temporary codec patch, builds a
release probe with warnings denied, and restores exact production bytes in a
`finally` block before execution. The patch uses the existing core SHA-256
digest helper, records all default numeric MCE limits, and adds no dependency
to production. Ordinary production APIs and behavior remain unchanged.

The first probe execution passed, but Clippy found two single-pattern `match`
expressions. Those were changed to `if let`; the initial source, build, traces,
and failed lint log remain under `initial/`. The final probe was rebuilt and
all 12 processes were rerun. Probe formatting and Clippy pass. All six existing
evidence gates pass. The final audit independently recomputes focused results,
checks source/build/corpus bindings, verifies evidence receipts, and checks
owned scratch cleanup. No new fuzz campaign or full production test run is
claimed for this diagnostic-only change.
