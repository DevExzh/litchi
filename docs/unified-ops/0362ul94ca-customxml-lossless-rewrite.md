# 0362ul94ca — Custom XML lossless properties rewrite

This follow-up to 0362ul94c keeps the committed error classification and
source-freshness behavior while closing the fidelity gap in properties
publication.

`custom_xml::codec::rewrite_props` now scans the validated properties XML with
the namespace-aware reader and uses its source offsets to replace only the
expanded `itemID` and `uri` attributes.  Prefixes, quote style, whitespace,
comments, foreign attributes, and unrelated extension markup therefore remain
byte-for-byte unchanged.  Schema-reference insertion, removal, and empty
element expansion use bounded lexical edits; only opaque markup inside the
schema-reference container is refused when a topology change would make
lossless placement ambiguous.

The rewrite computes the exact output size before allocation, admits it against
`SerializedPropertiesBytes`, and uses `try_reserve_exact`.  Read-back errors
are propagated instead of being collapsed into `Internal`, so configured
limits and XML failures retain their original classification.  ID-only edits
use a root-only source-offset proof and skip the structural read-back; the
proof is limited by the same source event, depth, and attribute-byte profile.

The DOCX package consumers now classify missing, ambiguous, and occupied data
store IDs as typed refusals with fixed reasons.  Custom XML part-retention
checks also propagate malformed relationship targets instead of treating them
as unreferenced parts.

Regression coverage includes noncanonical namespace/prefix/quote spelling,
comments and foreign same-local-name attributes, preservation of extension
markup, the opaque topology refusal boundary, and exact versus one-byte-under
output limits.  This source-only change was prepared without Cargo, rustdoc,
or performance execution.

The 94cb correction tightens that boundary: unknown children directly under
`schemaRefs` remain in place, while a schema reference carrying foreign
attributes, comments, or non-whitespace content is refused only when that
reference would be removed.  MCE-selected branches now retain source-offset
provenance for both ID-only and schema-list edits, and generated references
are emitted from the requested vocabulary rather than cloning an existing
opaque reference.  The DOCX bibliography adapter classifies a missing store
or source tag as `RefusalKind::NotFound`, and malformed relationship targets
remain propagated errors in every data-store retention caller.
