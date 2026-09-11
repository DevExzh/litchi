# XLSX drawing anchor prerequisite

The ordinary drawing reader now retains numeric EMU geometry for two-cell,
one-cell, and absolute anchors using the shared spreadsheet drawing types.
`drawing_anchor()` is authoritative; the existing chart-style `anchor` field
remains a compatibility projection. Adding the public geometry field requires
external struct literals to initialize it.

The reader checks direct ownership, geometry order and duplicates, bounds,
XML numeric whitespace, and Strict/Transitional namespaces. MCE input and
expanded output are capped before construction. Context and object growth
reserve fallibly, and the parser borrows the active namespace resolver.

Independent review and root validation are bound to the hashes in
`review.json`. The integration suite also exercises opaque nested blips,
namespace expansion beyond the output cap, and NBSP rejection.

This is a prerequisite for the worksheet SVG lifecycle, which remains in
progress. The coordinate profile is numeric EMU only: the complete
`ST_Coordinate` union includes unit-bearing measurements, and lossless support
for those remains open. No rounding or full coordinate-union conformance is
claimed. No runtime performance or native Office acceptance is claimed here.
