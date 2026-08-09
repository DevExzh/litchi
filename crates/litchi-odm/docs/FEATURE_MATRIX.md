# ODM Feature Matrix

This matrix records the current public `litchi-odm` capability for
OpenDocument master documents. The crate is a cheaply cloneable immutable
package snapshot with bounded semantic projection and narrow, source-checked
title and linked-section edit seams; detached construction values remain
separate from parsed package references.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODM package snapshot | ✅ | ✅ | N/A | Exact text-master MIME is required; original bytes, safe file names, raw XML, projected metadata, and title are exposed. |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | Snapshot and fresh publication require UTF-8, the shared 64 MiB compact-XML ceiling, and bounded DTD-free XML with namespace-aware `office:document-content/office:body/office:text` placement. Prefix aliases are accepted. Duplicate `text:section` names and `xml:id` values are rejected; broader master-document schema semantics remain unvalidated. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates MIME, raw content, and manifest only; no opened-master save path exists. |
| Compact XML | 🟡 | ✅ | ✅ | Opened and freshly authored `content.xml` are validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Changed-package publication also requires every source XML member and changed XML part to pass the compactness gate. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Styles and metadata | 🟡 | 🟡 | 🟡 | Styles remain raw and metadata is projected through the common reader. Only `dc:title` can be set or cleared in an opened package; broader metadata authoring is absent. |
| Sections and subdocuments | ✅ | ✅ | 🟡 | The snapshot projects every ordered linked `text:section-source` as its containing section name, optional source-section/filter names, and an inert `Package` or `External` target. ODF 1.4's optional local-only source form is accepted without inventing a subdocument. Link sources must occur once as the first section child, `xlink:type`/`xlink:show` use their schema values, and existing `xlink:href` values are editable by exact section name or checked `Position`. Section trees, cycle discovery, and path resolution remain intentionally absent. |
| Permanent external-resolution boundary | ✅ | N/A | N/A | Subdocument references are classified only; neither safe package paths nor external targets are opened, resolved, fetched, or recursively loaded. |
| Existing-package edits and patches | 🟡 | ✅ | 🟡 | `Master::edit_title()` stages `dc:title` set/clear and `Master::edit_link()` stages one existing linked-section target. Both commits reparse and verify typed readback and return exact-whole-package-source reversible patches with applicability checks. Changed edits refuse signed packages, non-compact XML members, and stale sources; broader section structure and metadata remain read-only. |
| Untouched-byte preservation | ✅ | ✅ | ✅ | No-op title commits preserve the exact archive. Changed title commits replace only `meta.xml`; unrelated ZIP members and unknown metadata XML are retained, and the inverse patch restores the exact source package bytes. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Active content | 🟡 | 🟡 | 🟡 | Scripts, macros, controls, actions, DDE, and embedded code remain unparsed inert bytes and are never executed. |
| Limits and evidence | ✅ | ✅ | 🟡 | Snapshot content has the shared 64 MiB compact-XML ceiling, depth-256, identity/reference-count, and 16 KiB semantic-name/target/title limits. DTD and named entities are rejected. Active tests cover real master content, title and linked-section set/no-op/inverse behavior, exact-name and typed-position selection, stale-source and signed-package refusal, namespace aliases, ODF linked-source placement and attribute rules, unknown markup and auxiliary-member retention, classification, ordering, duplicate identities, entity rejection, compact input, and compact output. |
