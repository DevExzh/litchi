# ODM Feature Matrix

This matrix records the current public `litchi-odm` capability for
OpenDocument master documents. The crate is a cheaply cloneable immutable
package snapshot with bounded semantic projection and source-checked unified
transactions; detached construction values remain separate from parsed
package references.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODM package snapshot | ✅ | ✅ | N/A | Exact text-master MIME is required; original bytes, safe file names, raw XML, projected metadata, and title are exposed. |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | Original producer XML is accepted when it is UTF-8, bounded, DTD-free, well formed, and has namespace-aware `office:document-content/office:body/office:text` placement. Fresh authored XML remains subject to the shared compact-XML gate. Prefix aliases are accepted. Duplicate `text:section` names and `xml:id` values are rejected; broader master-document schema semantics remain unvalidated. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates MIME, raw content, and manifest only; no opened-master save path exists. |
| Compact XML | 🟡 | 🟡 | ✅ | Producer formatting is accepted on open and the unchanged source artifact is retained exactly. A changed transaction losslessly compacts regenerated core XML while preserving semantic text and `xml:space="preserve"` content; untouched auxiliary member payloads remain exact and need not be compact. Fresh authored XML is compact-validated. Space-only inter-element text is still accepted, so absolute minimality is not guaranteed. |
| Styles and metadata | 🟡 | ✅ | 🟡 | Named `style:style` definitions from `content.xml` and `styles.xml` are projected with family, parent, and owning part. Transactions can add minimal styles, rename definitions plus modeled references, or remove unreferenced styles. Common metadata is projected; title, author, subject, description, and keywords are atomically editable while other metadata is preserved. |
| Sections and subdocuments | ✅ | ✅ | 🟡 | The snapshot projects the complete ordered `text:section` tree with parent/child positions, style, `xml:id`, protection, and optional link position. Transactions can add an empty root section, rename a section plus modeled local-only references, or remove an unreferenced subtree. Every linked `text:section-source` carries its containing section, optional source-section/filter names, and an inert `Package` or `External` target. Existing targets are editable by exact section name or checked `Position`. |
| Package resource graph | ✅ | ✅ | 🟡 | Safe package members expose declared media type and exact incoming linked-section positions. Transactions can add, replace, remove, or transfer resources, with missing-target and incoming-reference closure checks. External targets remain inert and are never resolved or fetched. |
| Permanent external-resolution boundary | ✅ | N/A | N/A | Subdocument references are classified only; neither safe package paths nor external targets are opened, resolved, fetched, or recursively loaded. |
| Existing-package edits and patches | ✅ | ✅ | 🟡 | `Master::edit()` stages title/link, simple metadata, section-tree, style, and resource changes and publishes them with one full-package reopen and typed readback. Security policy can deny new external targets and bound resource size. Exact-source reversible patches support non-mutating three-way planning, typed conflicts, disjoint merge, bounded undo/redo history, deterministic durable exchange, stale-source refusal, inverse application, and forward-only sealing. Changed edits refuse signed/encrypted packages. |
| Untouched-byte preservation | ✅ | ✅ | ✅ | A no-op preserves the exact source archive. Changed publication preserves untouched auxiliary member payload bytes, but the rebuilt archive's physical ZIP records may be normalized. Exact-source inverse patches restore the complete original artifact byte for byte. |
| Encryption and signatures | 🟡 | 🟡 | ❌ | Password-encrypted bytes can be opened with an explicit password for inert inspection. Exact no-op transactions preserve signed/encrypted bytes; any changed transaction and credential-free durable publication is refused. Signature verification and encrypted authoring are absent. |
| Active content | 🟡 | 🟡 | 🟡 | Scripts, macros, controls, actions, DDE, and embedded code remain unparsed inert bytes and are never executed. |
| Limits and evidence | ✅ | ✅ | 🟡 | Snapshot semantic parts have explicit byte/depth/count and 16 KiB semantic-value bounds. DTD and non-predefined named entities are rejected. Tests directly open and edit a checked-in original LibreOffice `.odm` without fixture repacking, verify exact unchanged-source and inverse preservation, and cover atomic structure/style/resource/metadata edits, dependency closure, transfer, policy refusal, durable/stale exchange, three-way planning, full reopen, and compact changed output. |
