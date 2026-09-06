# 0440: Borrow ODP cached attribute namespaces

The private ODP attribute cache previously copied each resolved namespace URI
into a vector. It now borrows the URI from the XML reader, with lifetimes
preventing the cache from outliving the reader scope. The shared iterator,
lazy decoding, error order and drawing-attribute order remain intact. The
one-shot lookup path is unchanged. No public API or dependency changes.

The owned existing-append lifecycle from 0439 supplies the matched experiment:
24 reports and 720 samples in A1/B1/B2/A2 order, across 64, 4,096 and 8,192
source slides. Medium/large allocation calls fall 20.101%/20.163%, and
cumulative requested bytes fall 6.322%/6.562%, identically in both repeats.
Large calls change from 1,462,779 to 1,167,845; requested bytes change from
236,704,188 to 221,171,020. Exact semantic, package, preservation, patch,
no-op and sink identities pass the unchanged independent 0439 oracle.

The change is kept for allocation work only. Operation peak above entry and
retained live bytes are unchanged. Normal p50 changes range from −1.501% to
+1.743%; no latency or RSS improvement is claimed. Main large R2 p95/p99
increase 8.353%/8.138%. A fixed additional four-report, 120-sample large-normal
ABBA confirmation does not reproduce the adverse tails in either pair. All
main and repeat flags remain retained, including RSS and instrumented timing
variation. See the [measurements](../results/change-0440/measurements.md) and
[decision](../results/change-0440/decision.json).

Four selected whole-process profiles retain raw data and conversions.
Instructions fall 0.661% while cycles rise 0.594%; setup and oracle work are
included, so these do not establish a timed-operation speedup. Baseline and
candidate record conversions retain 15 and 13 symbolization warnings each.
The initial baseline record conversion overlapped the first confirmation
attempt through a root scheduling error. Both attempts remain retained and
explicitly excluded; serial replacements supply the selected evidence.

All 352 ODP tests pass, including three new borrowing and namespace-scope
checks; the complete harness scope passes 368 tests with one ignored. Strict
owner Clippy, rustdoc with warnings denied, formatting and crate boundaries
pass. Full dependency Clippy retains the existing common archive-reader
large-enum diagnostic. Source/build custody, independent report validation,
copied-bundle mutation probes and cleanup proof are in the
[evidence bundle](../results/change-0440/README.md).

The registry remains 436 selectors and the default matrix remains 36 cases.
This closes one private allocation-ownership experiment. One-shot cache cost,
repeated staging/validation, bounded existing append, Part addition,
repackaging, native breadth, cold/range I/O and scaling remain open under the
active non-iWork goal.
