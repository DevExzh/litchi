# ODM Feature Matrix

This matrix records the current public `litchi-odm` capability for
OpenDocument master documents. The crate is a cheaply cloneable immutable
package snapshot with bounded semantic projection and source-checked unified
title/linked-section transactions; detached construction values remain
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
| Styles and metadata | 🟡 | ✅ | 🟡 | Named `style:style` definitions from `content.xml` and `styles.xml` are projected with family, parent, and owning part. Common metadata remains projected through the shared reader; `dc:title` can be set or cleared atomically with link edits while broader metadata/style authoring remains absent. |
| Sections and subdocuments | ✅ | ✅ | 🟡 | The snapshot projects the complete ordered `text:section` tree with parent/child positions, style, `xml:id`, protection, and optional link position. Every linked `text:section-source` carries its containing section, optional source-section/filter names, and an inert `Package` or `External` target. ODF 1.4's optional local-only source form is accepted without inventing a subdocument. Link sources occur once as the first section child, `xlink:type`/`xlink:show` use their schema values, and existing targets are editable by exact section name or checked `Position`. |
| Package resource graph | ✅ | ✅ | N/A | Safe package members expose their declared media type and exact incoming linked-section positions; missing package targets are reported separately. External targets remain inert and are never resolved or fetched. |
| Permanent external-resolution boundary | ✅ | N/A | N/A | Subdocument references are classified only; neither safe package paths nor external targets are opened, resolved, fetched, or recursively loaded. |
| Existing-package edits and patches | ✅ | ✅ | 🟡 | `Master::edit()` stages one title plus any number of existing linked-section targets and publishes them with one full-package reopen and typed readback. Exact-source reversible patches support disjoint same-source merge with typed conflicts, bounded undo/redo history, deterministic JSON/content-addressed durable exchange, inverse application, and forward-only sealing. Legacy focused title/link seams remain available. Changed edits refuse signed/encrypted packages, non-compact XML/RDF members, and stale sources; broader section/style/metadata structure remains read-only. |
| Untouched-byte preservation | ✅ | ✅ | ✅ | No-op title commits preserve the exact archive. Changed title commits replace only `meta.xml`; unrelated ZIP members and unknown metadata XML are retained, and the inverse patch restores the exact source package bytes. |
| Encryption and signatures | 🟡 | 🟡 | ❌ | Password-encrypted bytes can be opened with an explicit password for inert inspection. Exact no-op transactions preserve signed/encrypted bytes; any changed transaction and credential-free durable publication is refused. Signature verification and encrypted authoring are absent. |
| Active content | 🟡 | 🟡 | 🟡 | Scripts, macros, controls, actions, DDE, and embedded code remain unparsed inert bytes and are never executed. |
| Limits and evidence | ✅ | ✅ | 🟡 | Snapshot semantic parts have explicit byte/depth/count and 16 KiB semantic-value bounds. DTD and named entities are rejected. Active tests cover a checked-in genuine LibreOffice `.odm`, section/style/resource projections, atomic title+link set/no-op/inverse behavior, durable canonical round trips, merge conflicts, bounded history, full reopen, signed/encrypted refusal, exact-name and typed-position selection, namespace aliases, ODF source placement and attribute rules, unknown markup retention, duplicate identities, malformed style/section values, compact input, and compact output. |
