# ODM Feature Matrix

This matrix records the current public `litchi-odm` capability for
OpenDocument master documents. The crate is an immutable raw package shell;
its section and subdocument values are not connected to the package codec.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODM package snapshot | 🟡 | ✅ | N/A | Exact text-master MIME is required; original bytes, safe file names, raw XML, and projected metadata are exposed. |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | Snapshot and fresh publication require UTF-8, a 256 MiB family ceiling, the compatibility `office:text` marker, and bounded DTD-free XML with namespace-aware `office:document-content/office:body/office:text` placement. Duplicate `text:section` names and `xml:id` values are rejected; broader master-document schema semantics remain unvalidated. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates MIME, raw content, and manifest only; no opened-master save path exists. |
| Compact XML | 🟡 | N/A | 🟡 | Fresh input is validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Styles and metadata | 🟡 | 🟡 | ❌ | Styles are raw and metadata is projected through the common reader; authoring is absent. |
| Sections and subdocuments | 🟡 | 🟡 | ❌ | Content validation inventories section names only to enforce required names and uniqueness. Public section/subdocument values remain detached; there is no section tree, reference graph, cycle handling, or path validation. |
| Permanent external-resolution boundary | ✅ | N/A | N/A | Subdocument references are never opened, resolved, fetched, or recursively loaded. |
| Existing-package edits and patches | ❌ | ❌ | ❌ | No save, transaction, commit, source check, or reversible patch exists. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original archive bytes remain exact before mutation; no edited-package preservation path exists. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Active content | 🟡 | 🟡 | 🟡 | Scripts, macros, controls, actions, DDE, and embedded code remain unparsed inert bytes and are never executed. |
| Limits and evidence | 🟡 | 🟡 | 🟡 | Snapshot content has a 256 MiB family ceiling and depth-256/identity-count structural limits; authoring first applies shared 64 MiB and depth-256 compactness limits. Active tests use full namespace-bearing master structure and cover typed compactness rejection, semantic whitespace, duplicate identities, and exact opaque retention of a real LibreOffice master-document XML fixture. |
