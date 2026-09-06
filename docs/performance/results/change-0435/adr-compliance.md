# Batch 0435 ADR compliance

This note records the scope of the 0435 candidate `6fc317964` against the `7c25f299d`
baseline, against accepted ADR tree
`c950b6c8be822561b498d7bbe87c460873dcbf49`. It covers two related changes:
typed common publication/diagnostics and an ODT fresh-creation API for bounded
plain-text paragraphs. It records design and functional evidence only. No 0435
formal performance result or optimization claim is authorized here.

## Ownership and API fit

| Area | Current owner and contract | ADR fit |
| --- | --- | --- |
| Authored fixed XML | `litchi_odf_common::core::PackageWriter::add_authored_xml` validates the path, manifest binding, and authored XML, then publishes the caller's borrowed bytes with the sized Deflate path. It preserves comments and returns `PackageWriterError`, including acknowledged output and nested transport causes. The legacy `add_file` behavior is unchanged. | [ADR 0002](../../../adr/0002-crate-topology.md), [ADR 0004](../../../adr/0004-semantic-api-design.md), [ADR 0010](../../../adr/0010-facade-archive-ownership.md) |
| Generated XML attribution | `GeneratedXmlLimitExceeded` carries resource, observed value, and ceiling through the common I/O boundary; `PackageWriterError::xml_limit()` recovers it with a bounded source walk. Malformed XML keeps its existing refusal mapping. | [ADR 0004](../../../adr/0004-semantic-api-design.md), [ADR 0005](../../../adr/0005-io-memory-and-performance.md), [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md) |
| ODT semantic creation | `litchi-odt::streaming::{stream_plain_paragraphs_to,try_stream_plain_paragraphs_to}` owns the ODT paragraph grammar, whitespace and CR policy. The common crate owns lexical XML auditing and archive publication; the family crate supplies the ODT envelope and semantic source. | [ADR 0001](../../../adr/0001-priorities-and-api-layers.md), [ADR 0002](../../../adr/0002-crate-topology.md), [ADR 0023](../../../adr/0023-odf-family-crate-split.md), [ADR 0024](../../../adr/0024-current-topology.md) |
| Operation boundary | The new API is fresh, sequential creation of a plain-text ODT package. It does not advertise logical append, package-Part addition, arbitrary modification, or repackaging. Producer, XML, budget, cancellation, and sink failures retain typed causes; a failed package is not finalized. | [ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md), [ADR 0005](../../../adr/0005-io-memory-and-performance.md) |

The relevant implementation seams are [`add_authored_xml`](../../../../crates/litchi-odf-common/src/core/writer.rs#L1210),
[`add_generated_xml`](../../../../crates/litchi-odf-common/src/core/writer.rs#L1246),
the typed XML limit definitions in
[`generated_xml.rs`](../../../../crates/litchi-odf-common/src/core/generated_xml.rs#L23),
and the ODT streaming module's public entry points in
[`streaming.rs`](../../../../crates/litchi-odt/src/streaming.rs#L411).

## Limits and evidence status

The candidate keeps finite paragraph, text, XML, output, memory/work, and
cancellation limits explicit. A reusable fragment buffer is a bounded
publication window; it is not a claim about total allocator usage or whole
process RSS. ZIP framing is admitted before the first paragraph pull, so a
producer or sink failure can leave an incomplete sink and must poison/discard
the package while reporting only acknowledged bytes.

The ODT provider owns semantic whitespace and CR behavior. The common auditor
does not replace that responsibility. Fixed styles and metadata use the typed
authored publisher so their existing comments and bytes remain intentional;
generated paragraph fragments continue through the generated-XML audit.

The release and fixture checks described by the candidate are useful
correctness evidence, separate from the formal forward/reverse
capture, allocator/RSS evidence, scaling evidence, or native producer proof.
The native runtime probe is unavailable for this batch; separate fixture tests
must remain labeled as fixture coverage and cannot be promoted to native
compatibility evidence. The overall non-iWork goal therefore remains open.
