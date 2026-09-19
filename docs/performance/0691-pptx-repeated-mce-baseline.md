# 0691 — refresh the opened-PPTX repeated-MCE baseline

Status: baseline and design investigation; no production optimization.
`performance_claim: none`. Baseline: `bd50cb0d5`.

The current opened-presentation edit still repeats markup-compatibility (MCE)
processing. A fresh native probe measures the real 13-slide deck at
**22.69–22.81 ms** median for capture, working clone, one shape-text edit,
commit and apply. The same-length namespace-marker counterfactual measures
**7.52–7.57 ms**. These are separate inputs, not an optimized implementation
or a before/after speedup. Capture and commit account for most of the gap.

The [evidence packet](results/change-0691/README.md) retains source, constraint,
corpus, probe, binary and raw-result bindings; reproducible drivers; native
profiles; separate allocator diagnostics; and temporary trace sources with
exact restoration records. All 33 prior constraint hashes remain unchanged.

## Native phase baseline

Each case has four independent process legs, five warmups and 100 fresh
packages per leg: 2,000 timed edits in total. Odd legs reverse case order.
The table reports the median of the four per-leg medians, in milliseconds;
the last column preserves the range of total medians. Raw samples, means,
nearest-rank p95/p99 and seeded within-leg median bootstrap intervals remain
in `native-summary.json`. Inter-leg total-median spread is 0.52–2.42%; this
does not establish a future noise bound or a population tail guarantee.

| Case | Capture | Clone | Set text | Commit | Apply | Total median range |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Real `slide-section-test` | 10.6023 | 0.0169 | 1.0893 | 10.9245 | 0.0958 | 22.689–22.806 |
| Real marker control | 3.3538 | 0.0165 | 0.5539 | 3.5153 | 0.0770 | 7.517–7.573 |
| Generated 12 slides × 8 text boxes | 0.7872 | 0.0118 | 0.1373 | 0.8802 | 0.0487 | 1.858–1.871 |
| POI notes `prProps` | 0.4065 | 0.0054 | 0.1368 | 0.4916 | 0.0299 | 1.070–1.096 |
| LibreOffice notes `tdf131082` | 0.5048 | 0.0051 | 0.3982 | 0.7936 | 0.0296 | 1.730–1.741 |

The enclosing total includes timer/check overhead between public calls. Source
file reading, generated corpus construction, initial package opening, target
discovery, save and reopen are outside the timers. Each process separately
validates the edited text through save/reopen and reports slide/shape semantic
digests, package revisions and output archive hashes. These values repeat
across all native legs. They are a bounded public-API projection and fixed
output oracle, not a new proof of every preservation contract.
The final oracle additionally asserts a changed revision, preserved slide and
shape counts, and identical part inventory, content types, relationships and
non-edited payload hashes before save and after reopen.

The control replaces 43 exact MCE namespace URIs across 103 ZIP members.
Member names, timestamps and uncompressed lengths are retained; ZIP
compression, flags and external attributes are regenerated on all 103 members
(see `control-metadata-check.json`). Its semantic slide/shape projection happens to match this real
deck, but changing namespace URIs is not a generally equivalent document
transformation. The two timed notes fixtures are marker-free. A third notes
fixture, `tdf89064`, admits capture but has no text shape admitted by the
probe's bounded edit-target search; its failed edit smoke and successful
capture trace are retained separately.

The older 0649 probe unconditionally counted allocator operations. Its phase
timings and profiles therefore are not native baselines and are not pooled
with this packet. The new probe also uses a distinct edit marker; no historical
latency improvement is claimed.

## Repeated work and costs

One fresh real capture makes **44 successful baseline MCE calls**, processing
**819,319 input bytes** from **14 distinct raw pointer/length identities**.
All calls produce owned XML. The presentation is processed five times and
each of the 13 slides three times. Trace URI/Arc observations distinguish the
equal-length slide buffers that the older length-only trace could not resolve.
The temporary instrumentation is diagnostic only and is restored byte-for-byte
before further checks. Setup captures used to discover an edit target must
remain separate from the capture/commit pair in the apply-prefix trace.

Native sampling of the **open-plus-capture prefix** attributes 20.07% of self
samples to MCE `start`, 4.79% to dropping its inherited context, 2.91% to
`write_start`, and 2.61% to `close`. Other XML, allocation and namespace work
is distributed across shared symbols; these percentages are not a complete
MCE inclusive attribution. Kernel symbols are unavailable; no samples were
reported lost. Raw reports and counter isolation runs remain in the packet.

The separate allocator companion reports identical per-phase request counts
across its three samples. The real edit performs 350,148 successful allocation
calls plus 9,030 reallocations, requesting 26,015,747 bytes including each
successful realloc's new size. Its observed peak is 463,159 live requested
bytes above the pre-capture baseline; net live growth at the post-apply boundary
is 195,730 bytes. Capture alone makes 157,951 allocation calls and commit
172,534. The control totals 94,311 allocation calls and 5,940,628 requested
bytes. Process allocator live bytes exclude allocator metadata and cannot be
equated with RSS or a production memory budget. The independent native
open-plus-capture child reports 5,664 KiB maximum RSS; there is no before/after
RSS comparison.

## Decision and next proof

The evidence supports eliminating repeated capture work before tuning XML
instructions. The [design review](results/change-0691/design.md) examines
short-lived scan reuse and the constraints on longer-lived XML retention.
Capture validates **all slide roots before any slide names**, then validates
notes. Any combined projection must preserve that error order, all limits,
relationship checks and publication validation.

A transaction-wide processed-XML cache is not implemented. It would need
runtime proof that `blob_arc()` aliases `blob()`, strong source ownership,
processing-profile separation, replacement invalidation, and an explicit
memory policy with fallback. The existing retained-candidate limit belongs
to serialized cross-slide-copy archives and cannot silently become an XML
cache budget. The next candidate should first price shorter-lived reuse of
the already required semantic results.

## Validation and limits

Both standalone builds pass with warnings denied; native format checking,
five-case smoke, four-leg semantic/output repeatability, corpus-control
verification and source/probe/raw-result audits pass. Initial probe-only
build failures (`unused_mut` under native cfg and a missing diagnostic `Debug`
derive) were corrected before measurements; their logs remain separate.
After restoring the instrumented source, all 587 PPTX library tests pass
in a warning-denied release build (zero failures or ignored tests). This is
focused baseline verification, not a rerun of every workspace feature gate.
Review prompted a stronger untouched-part oracle and enforced `--locked`
builds. The complete initial packet is retained under `initial/`; native,
allocation, profile and trace captures were repeated with the final probe.
The two captures are not an optimization comparison and are not pooled.

No production API, cache, retained state, validation behavior, dependency or
iWork source is changed. This packet measures owned archive bytes and warm
OS caches on one shared Linux x86-64 host. It does not measure file/range
sources, physical cold caches, concurrent scaling, full save workflows,
cross-platform behavior or native Office applications. It adds no registered
performance claim or CRUD coverage promotion. The broader GOAL remains active.
