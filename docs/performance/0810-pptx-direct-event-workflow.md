# 0810 — PPTX direct event handling workflow

Retained: the one-file candidate passes the frozen public workflow adoption policy, correctness gates, and source/profile review.

The fresh baseline is `3677e31be5`, which repaired the independent test-only
Clippy blocker from 0808. This trial reuses the reviewed candidate source
bytes, with fresh builds and samples. No historical timings are pooled.

The candidate changes only `crates/litchi-pptx/src/notes/codec.rs`. It reads
slice-backed XML events directly and uses quick-XML’s public namespace
resolver for the same scope transitions and resolution. Namespace pushes
precede node/depth validation; pending pops precede the next read. Existing
error conversion, bounded validation, attribute policy, and the two buffered
oracle functions are preserved. Three differential tests cover namespace
rebinding and refusal ordering. No public API or dependency changes.

The frozen protocol crosses six fixture shapes with capture, commit, and
lifecycle. Six native paired blocks use AB BA AB BA BA AB, CPU 12, thirty
samples and three warmups per process. Two allocation blocks use AB BA,
three samples and no warmup. Eighteen before-only qualification reports
precede candidate application. Four large-capture Callgrind reports complete
the protocol: 310 reports and 6,718 measured samples, excluding warmups.

Adoption requires a capture or lifecycle improvement of at least 3% with
bootstrap upper endpoint below 1. Any row above 1.05 with lower endpoint
above 1 vetoes adoption. Every paired allocation-block median must avoid
increases in calls, allocated bytes, net live bytes, and peak above entry.
Bootstrap uses 10,000 resamples of six process-p50 ratios, seed 810810.
Tail latency, RSS, and process spread remain separate review diagnostics.
Callgrind measures guest instruction attribution, with no native time,
phase-fraction, RSS, or universal speedup claim.

All six candidate production quality gates pass: formatting, all-targets
check, tests (1,241 passed, zero failed, three ignored across 85 groups),
warning-denied all-targets Clippy, warning-denied rustdoc, and repository
boundaries. Baseline qualification passes all eighteen sealed source,
output, fixture, text, and extension oracles before application.

[Protocol and raw evidence](results/change-0810/README.md).
The broader project goal remains active; iWork is excluded.

## Paired native results

Times below are medians of six process p50 values. Changes and confidence
intervals use the median of six paired after/before ratios, so they need not
equal the ratio of the displayed times. Negative change means faster.

| Shape | Workflow | Before ms | After ms | Paired change | Bootstrap ratio interval |
|---|---|---:|---:|---:|---|
| tiny | capture | 0.238466 | 0.232731 | -2.471% | 0.974423–0.976532 |
| tiny | commit | 0.213966 | 0.209956 | -2.025% | 0.976531–0.988357 |
| tiny | lifecycle | 1.416893 | 1.406527 | -0.806% | 0.990090–0.995460 |
| medium | capture | 0.461377 | 0.446333 | -2.918% | 0.957020–0.979407 |
| medium | commit | 0.298731 | 0.293581 | -1.967% | 0.978536–0.984600 |
| medium | lifecycle | 2.013041 | 1.999646 | -0.600% | 0.982747–0.996055 |
| large | capture | 18.810121 | 17.875001 | -4.975% | 0.943695–0.953076 |
| large | commit | 1.305292 | 1.280557 | -1.926% | 0.979613–0.983586 |
| large | lifecycle | 28.362532 | 27.769528 | -2.133% | 0.967758–0.984087 |
| vendor | capture | 0.549398 | 0.527748 | -4.073% | 0.954917–0.966248 |
| vendor | commit | 0.329842 | 0.323377 | -1.900% | 0.976638–0.984196 |
| vendor | lifecycle | 2.179366 | 2.150937 | -1.267% | 0.985953–0.993700 |
| unicode-vendor | capture | 0.552943 | 0.533898 | -3.803% | 0.958162–0.978841 |
| unicode-vendor | commit | 0.329962 | 0.323221 | -2.003% | 0.975995–0.981418 |
| unicode-vendor | lifecycle | 2.183797 | 2.160237 | -1.014% | 0.982840–0.991137 |
| valid-4attr | capture | 0.528873 | 0.506678 | -4.129% | 0.952842–0.968791 |
| valid-4attr | commit | 0.321901 | 0.314356 | -2.190% | 0.971923–0.980605 |
| valid-4attr | lifecycle | 2.143331 | 2.115242 | -1.277% | 0.980017–0.989423 |

Large, vendor, unicode-vendor, and valid-4attr capture meet the frozen benefit
gate. No latency veto occurs. All 144 allocation resource comparisons are
exactly equal between legs; allocation reduction is not part of the benefit.
All 306 main reports and 6,714 samples pass full output and semantic checks,
and two independent readers agree on the policy outcome.

The diagnostics retain 21 native spread flags above 5% (including RSS) and
three per-block p99 regression flags. Tiny capture block 0 has p99 +96.947%,
large commit block 3 +11.277%, and vendor commit block 4 +20.164%. Their paired
p99 medians are below 1, while their bootstrap intervals cross 1. These are
visible tail uncertainties, not discarded samples or a claim of tail-latency
improvement. Tiny capture after p99 spread reaches 102.004%. No native p50
spread exceeds 5%; allocation metrics have no spread flags. Native process
RSS ranges 4,792–18,828 KiB and allocation-process RSS 4,888–18,812 KiB;
process RSS is diagnostic only and does not establish memory improvement.

## Mechanism and review limits

All four profiles pass exact scope, one owner call, numbered publication,
empty termination, self-sum conservation, and output/semantic checks. Scoped
owner Ir is 549,292,102 → 537,148,235 and 549,348,645 → 537,183,147.
The scanner’s direct `NsReader::process_event` edge disappears; the
282,612 `Reader::read_event_impl` calls remain. Namespace resolver push self
Ir is unchanged at 13,019,166. Other `NsReader` callers remain (1,053 global
incoming calls after), as expected. These guest instruction counts support
the intended transport restructuring; they do not establish a native phase
fraction or an instruction-to-time conversion.

The first offline profile replay rejected stale schema/tool constants in the
reader. The reader was corrected to the exact frozen public-workflow 0806
schema, with full field and sealed semantic checks; all four original raw
captures then passed write/check replay. No sample, driver, or source changed,
and no measurement was repeated to obtain a favorable result.

Independent source review finds no namespace, error-ordering, limit-ordering,
borrowed-event, or oracle-integrity blocker. One non-blocking coverage note
remains: the new nested-scope test declares child default namespace rebinding,
but its observable result exercises a restored prefixed relationship binding.
The existing default-root cases and unchanged oracle cover the current
contract; future behavior depending on child default bindings needs an
observable assertion. See [source review](results/change-0810/source-review.md)
and [profile review](results/change-0810/profile-review.md).

Native/profile release builds retain eleven inherited unused-helper warnings;
the allocation build has none. Both all-features probe Clippy lanes and the
production Clippy gate deny warnings and pass. The repository boundary audit
passes 65 packages, 244 internal dependency declarations, and 11 explicit
debt items. No universal document, platform, tail-latency, memory, or cold-cache
improvement is claimed from these fixtures.

## Disposition and custody

The direct event transport is retained. Main analysis and an independent raw
reader agree exactly on all eighteen ratios and bootstrap intervals, all 108
native paired blocks, and all 144 allocation resource comparisons. The
9,196-file source census changes only the approved codec. Root lock/config
copies predate the first build, both six-file probes match sealed 0806 bytes,
and all 35 architecture inputs and three unrelated working-tree files retain
their prior identities.

The final aggregate audit checks stage order, immutable output identities,
quality receipts, all 310 reports/6,718 samples, final source disposition, and
six-binary cleanup. Artifact filesystem timestamps are retained as explicit
post-decision observations; workload stage times come from driver receipts.
Replay does not rely on timestamps assigned by a future checkout.

Aggregate replay also caught a tuple/list JSON roundtrip mismatch in the
independent reader’s qualification identity output. Making those eighteen
serialized identities explicit lists fixed replay without changing the
existing audit bytes or any numerical result. Both offline reader corrections
remain recorded in the execution notes.

Final aggregate replay passes after cleanup. The owned target removal covers
11,216 files and 4,153,982,362 logical bytes, with exact identities retained for
all six binaries. The retained codec and unrelated workspace files are unchanged
by cleanup. The completed batch is sealed for exact index and commit auditing.
