# 0905 premeasurement method

0905 evaluates the bounded XML-audit summary candidate after the 0904
attribute census established corpus coverage. The implementation hypothesis is
separate from this protocol: the candidate must preserve the completed
`check_start` validation, aggregate and per-element limits, namespace and
`xml:space` behavior, error precedence, streaming/window accounting and the
multiple-attribute path. This packet makes no semantic or performance claim.

The root coordinator builds fresh baseline and candidate native/allocator
executables from isolated source roots using the current 0904 harness. The
baseline source manifest has the current 0904 harness inventory; the candidate
manifest differs at exactly `crates/xml-minifier/src/audit.rs`. The plan's
`source.file_count`, descriptors and source-root paths are filled only after
the final source and binaries are stable. The frozen candidate-file list has
one entry, so a second implementation or test file cannot enter the measured
delta accidentally.

Capture is serial, warm, pinned to CPU 12, and uses fresh processes under the
fixed child environment `PATH=/usr/bin:/bin`, `LC_ALL=C`, `LANG=C`, `TZ=UTC`.
The four cases are the two one-percent XLSX primaries, the medium one-edit
XLSX control and the medium semantic DOCX control. Qualification is 2/0;
native is six paired blocks of 100/10; allocator is two repeats of 3/2.
Qualification and allocator data remain diagnostic.

The native acceptance rule is two independent primary gates: paired p50 median
ratio at most 0.98 and a 10,000-draw bootstrap upper 95% bound below 1.0,
using seed 905009. Every p50/p95/p99 ratio above 1.05 and every role/case p50
spread above 1.20 remains visible as a review flag. No row may be replaced,
discarded or retried because it is slow.

The reducer is offline. It verifies plan/helper/binary/receipt/source custody,
report schema and statistics, exact case configuration, semantic identities,
source/output/sink invariants and paired block order before calculating any
ratio. `source_identity` receives the validated corpus worksheet count and
derives the one-edit unselected-read requirement from it. That avoids the
failed 0902 assumption that an untouched-read count is zero, while still
rejecting a report that silently omits compressed copies. DOCX identity fields
remain limited to what its harness actually emits.

The driver and reducer use create-only artifact writes. Failed commands remain
in place with all partial evidence. The root coordinator owns execution,
quality gates, any candidate restoration, final replay, report integration and
cleanup of marked temporary roots.
