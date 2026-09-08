# Reuse changed-worksheet compaction for empty web-binding validation

The retained 0468 profile attributes 10.3346% of inclusive commit sample weight
to `raw::web::read`, called after compaction and optional grid parsing. Its
namespace resolver copies the namespace and local name into separate vectors
for each structural event. The 0469 production change removed event ownership
from compaction but left this later read unchanged. The hypothesis for 0470 is
that a bounded proof during compaction can remove that later traversal and its
name allocations for ordinary worksheets without web-extension markup.

`changed_worksheet` attaches a finite `Probe` to private compacted bytes. It
observes only successfully emitted events; discarded inter-element whitespace
is absent from the proof input. The proof checks the SpreadsheetML root,
namespace resolution, depth, retained text and CDATA decoding, and absence of
DTDs. It falls back for SpreadsheetML `extLst`, x15 `webExtensions` and
`webExtension`, and xm `f` at every depth. Those triggers cover activation of
the web grammar and its unconditional orphan-end checks. General references
outside bindings remain ignored, matching the existing reader.

The writer copies raw attribute values into double quotes without escaping.
Single-quoted source attributes can therefore change the normalized XML's
meaning. Any apostrophe in raw attributes invalidates the proof. All unproven
states invoke the unchanged reader against the actual output. No error is
constructed or retained by the probe, and no binding vector or formula string
is built before the existing web validation phase. The 16 MiB limit applies
to compacted output, and the 128-element depth rule is unchanged.

Publication order remains compaction, optional grid parsing, web validation,
style validation, and change verification. Private output fields prevent using
a proof with mutated bytes. There is no additional retained worksheet-sized
structure, cache, public API, unsafe code, executor, or I/O. The existing
4,096-cell/1 MiB Store handoff and its byte/source lineage checks remain intact.

| ADR | Compatibility obligation |
| --- | --- |
| 0001, 0006 | Preserve bytes, semantics, typed refusals, limits and error priority; use the original reader for every unproven case. |
| 0002, 0010, 0011, 0024 | Keep worksheet grammar and its proof inside the XLSX owner; no archive dependency changes. |
| 0003 | Keep atomic publication, source-checked patches and exact no-op behavior. |
| 0005 | Finite temporary state and measured representative latency/allocation/RSS evidence; no ambient runtime or persistent cache. |
| 0008 | Focused differential tests plus existing XLSX preservation, adversarial, integration and documentation gates. |

An independent investigation also examined fusing the original eager semantic
parse and lossless snapshot scan. Their inputs differ whenever MCE processing
transforms XML: semantic parsing uses transformed bytes, while lossless spans
must address original bytes. A conditional fusion is possible only on the
borrowed preprocessing path, would require deferred snapshot errors to preserve
source-parser precedence, and may raise peak memory by retaining parser raw
cells and Layout together. That remains a larger follow-up requiring its own
allocation and RSS gate. It is not implemented or claimed by this batch.

Independent source review found no semantic blocker. It checked namespace scope
at `End`/`Empty`, the quote-normalization boundary, all unconditional web-reader
end checks, compacted-size precedence and transaction phase ordering. Probe
unit fixtures isolate generic and fallback cases; the compactor's differential
pipeline tests exercise the actual emitted-event callback and compare bytes,
bindings, error phase, debug variants and display messages. The first test build
found only redundant qualified paths rejected by the repository lint policy;
those were corrected and the failed build log is retained.

A second independent lifetime review investigated the initial normal RSS flag.
It found no candidate-only worksheet-sized retention: the callback stores no
borrow, the reader/writer/preserve stack drop before return, and the final tuple
moves the same compacted vector. The original post-edit vector remains live
around grid parsing in both versions. Explicitly releasing that now-unused
input immediately after successful compaction is a separate, testable memory
opportunity; it is not implemented here or used to explain the current RSS
measurements. Source inspection alone cannot select allocator page retention,
binary/runtime layout or process high-water timing as the explanation.
