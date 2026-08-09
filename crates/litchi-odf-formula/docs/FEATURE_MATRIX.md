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
| MathML XML parsing | ✅ | ✅ | ✅ | The iterative parser requires one MathML `math` root, rejects malformed XML and duplicate expanded attributes, enforces finite ceilings, and validates the supported MathML 2 presentation schema content models while closing nodes. |
| MathML semantic model | 🟡 | ✅ | ✅ | Common MathML 2 presentation, semantics, table, multiscript, and action elements have checked content models and fixed arities. Recognized boolean, enumeration, integer, and length value domains are checked; free-text and foreign-namespace attributes remain inert. |
| Common authoring constructors | 🟡 | N/A | 🟡 | Constructors cover common presentation structures and typed display/math-variant values. Generic mutation can construct shapes outside those helpers. |
| Compact XML serialization | ✅ | N/A | ✅ | Changed `content.xml` is emitted deterministically, including first-use foreign-namespace prefix assignment, without serializer-added indentation or formatting whitespace. Caller-supplied text nodes, including semantic whitespace, remain exact; this is not W3C XML Canonicalization. |
| Semantic mutation and history | ✅ | ✅ | ✅ | `edit()` supports path-addressed insert/remove/replace, attribute, token-text, StarMath-source, and whole-root operations. Commits retain ordered reversible evidence; source-checked patches and patch chains have bounded durable sidecar envelopes. |
| StarMath annotations | 🟡 | ✅ | ✅ | Exact `StarMath 5.0` and `StarMath 6` annotation encodings are modeled, readable, authorable, and transactionally editable. Source remains permanently inert because StarMath grammar/evaluation is outside ODF and MathML schemas. |
| OpenFormula spreadsheet syntax | ❌ | ❌ | ❌ | Cell references, operators, function catalogs, array expressions, and spreadsheet formula grammar are not implemented by this crate. |
| Permanent non-execution boundary | ✅ | N/A | N/A | MathML, StarMath annotations, entity references, and DTD declarations remain inert. No formula, entity/DTD, macro, VBA, control, action, DDE, or embedded-code evaluation or activation occurs. |
| Configurable resource limits | ✅ | ✅ | ✅ | Public checked limits cover package/XML bytes, element depth/count, per-element attributes, single-attribute bytes, and aggregate text; edits retain the source limits for candidate reopen. ZIP decompression is still delegated to the shared ODF package layer. |
| Evidence breadth | ✅ | ✅ | ✅ | Synthetic malformed and fuzz-like tests cover content models, arity, value domains, limits, granular edits, durable history, stale/inverse behavior, and compact output. Checked-in LibreOffice MathML and `.odf` fixtures cover real producer ingress. |
