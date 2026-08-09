# OpenDocument Formula Feature Matrix

This matrix records `litchi-odf-formula` support for packaged formula
documents and their MathML content. The crate is a bounded generic MathML tree
and package facade, not a complete MathML validator, OpenFormula spreadsheet
parser, or evaluation engine.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, generic, bounded, or whole-root only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Formula package open/create/save | ✅ | ✅ | ✅ | Formula/template MIME is validated; public package operations retain auxiliary members and can publish a replacement root. |
| Exact unmodified preservation | ✅ | ✅ | N/A | Original archive and `content.xml` bytes remain exact while unmodified. A changed root is canonically reserialized. |
| MathML XML parsing | 🟡 | 🟡 | 🟡 | Parser requires exactly one MathML `math` root, rejects duplicate expanded attributes and malformed XML, and enforces caller-selected depth/node/attribute/text ceilings within immutable hard maxima. It does not validate the MathML schema. |
| MathML semantic model | 🟡 | 🟡 | 🟡 | Thirty-three presentation/semantics kinds are classified; other elements, attributes, namespaces, and text remain a generic inert tree. Content models, arity, and most lexical values are not checked. |
| Common authoring constructors | 🟡 | N/A | 🟡 | Constructors cover common presentation structures and typed display/math-variant values. Generic mutation can construct shapes outside those helpers. |
| Compact XML serialization | ✅ | N/A | ✅ | Changed `content.xml` is emitted deterministically, including first-use foreign-namespace prefix assignment, without serializer-added indentation or formatting whitespace. Caller-supplied text nodes, including semantic whitespace, remain exact; this is not W3C XML Canonicalization. |
| Whole-root mutation | 🟡 | ✅ | ✅ | `edit()` provides immutable source snapshots, checked staging, atomic rebuild/reopen/readback, a named commit, exact-source stale checks, and reversible patches. It has no granular source map, durable wire envelope, history, or merge. |
| StarMath annotations | 🟡 | 🟡 | 🟡 | StarMath is retained and exposed only as annotation text; its syntax and functions are not parsed or validated. |
| OpenFormula spreadsheet syntax | ❌ | ❌ | ❌ | Cell references, operators, function catalogs, array expressions, and spreadsheet formula grammar are not implemented by this crate. |
| Permanent non-execution boundary | ✅ | N/A | N/A | MathML, StarMath annotations, entity references, and DTD declarations remain inert. No formula, entity/DTD, macro, VBA, control, action, DDE, or embedded-code evaluation or activation occurs. |
| Configurable resource limits | ✅ | ✅ | ✅ | Public checked limits cover package/XML bytes, element depth/count, per-element attributes, single-attribute bytes, and aggregate text; edits retain the source limits for candidate reopen. ZIP decompression is still delegated to the shared ODF package layer. |
| Evidence breadth | 🟡 | 🟡 | 🟡 | Synthetic integration tests cover round trips, malformed roots, exact limits, transaction no-op/change/stale/inverse behavior, auxiliary retention, and deterministic compact output; two LibreOffice `.odf` fixtures cover real producer ingress and byte-exact no-op preservation. Full MathML schema behavior remains outside the crate. |
