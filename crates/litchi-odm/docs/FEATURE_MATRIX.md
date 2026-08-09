# ODM Feature Matrix

This matrix records the current public `litchi-odm` capability for
OpenDocument master documents. The crate is an immutable package snapshot with
a bounded semantic projection and one narrow, source-checked title-edit seam;
detached construction values remain separate from parsed package references.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODM package snapshot | ✅ | ✅ | N/A | Exact text-master MIME is required; original bytes, safe file names, raw XML, projected metadata, and title are exposed. |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | Snapshot and fresh publication require UTF-8, a 256 MiB family ceiling, the compatibility `office:text` marker, and bounded DTD-free XML with namespace-aware `office:document-content/office:body/office:text` placement. Duplicate `text:section` names and `xml:id` values are rejected; broader master-document schema semantics remain unvalidated. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates MIME, raw content, and manifest only; no opened-master save path exists. |
| Compact XML | 🟡 | N/A | 🟡 | Fresh input is validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Styles and metadata | 🟡 | 🟡 | 🟡 | Styles remain raw and metadata is projected through the common reader. Only `dc:title` can be set or cleared in an opened package; broader metadata authoring is absent. |
| Sections and subdocuments | ✅ | ✅ | ❌ | The snapshot projects every ordered `text:section-source` as its containing section name and an inert `Package` or `External` target. Section trees, cycle discovery, and path resolution remain intentionally absent. |
| Permanent external-resolution boundary | ✅ | N/A | N/A | Subdocument references are classified only; neither safe package paths nor external targets are opened, resolved, fetched, or recursively loaded. |
| Existing-package edits and patches | 🟡 | 🟡 | 🟡 | `Master::edit_title()` stages only `dc:title` set/clear operations. Commit patches an existing UTF-8 `meta.xml`, reparses and verifies typed readback, and returns an exact-whole-package-source reversible `title::Patch`; section references and other metadata remain read-only. |
| Untouched-byte preservation | ✅ | ✅ | ✅ | No-op title commits preserve the exact archive. Changed title commits replace only `meta.xml`; unrelated ZIP members and unknown metadata XML are retained, and the inverse patch restores the exact source package bytes. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Active content | 🟡 | 🟡 | 🟡 | Scripts, macros, controls, actions, DDE, and embedded code remain unparsed inert bytes and are never executed. |
| Limits and evidence | ✅ | ✅ | 🟡 | Snapshot content has a 256 MiB family ceiling, depth-256, identity/reference-count, and 16 KiB semantic-name/target/title limits. DTD and named entities are rejected. Authoring first applies shared 64 MiB and depth-256 compactness limits. Active tests cover real master packages, title set/clear/no-op/inverse behavior, stale-source rejection, unknown markup and auxiliary-member retention, classification, ordering, duplicate identities, entity rejection, and compact output. |
