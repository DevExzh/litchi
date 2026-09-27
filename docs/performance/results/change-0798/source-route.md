# Counted owner and caller route

The measured module is `crates/litchi-opc/src/xml_attributes.rs`.
`crates/litchi-ooxml-common/src/xml/attributes.rs` re-exports its `BytesStartExt`
and `CheckedAttributes`, so instrumenting this owner also observes callers that
import the OOXML-common re-export. This includes PPTX notes validation, whose
`inspect_element` checks `attributes_raw().is_empty()` before constructing the
iterator and otherwise consumes `checked_attributes()` in a fallible loop.
Consequently, the census excludes precisely empty attribute tails skipped by
that caller. It is not a census of all XML elements or all document attributes.

The same public helper is also used by OPC package/content-type/relationship
readers and XML splicing. All invocations within the probe's operation region on
the caller thread are counted, irrespective of source file. The census does not
attribute a row to a unique Rust call site: tag identity and counts can identify
common shapes, but cannot separate two call sites handling the same tag.

The separate helper copies in OLE common, signing, XLDM, and XML-minifier are
not instrumented. Lenient `first_wins` and unchecked iteration are not counted.
Package ingress and output verification outside the region are excluded by the
explicit begin/finish boundary. The preserved generated fixture contract and
public call scopes remain authoritative; this census cannot represent other
producers, other formats, cold/range inputs, or concurrent workloads.
