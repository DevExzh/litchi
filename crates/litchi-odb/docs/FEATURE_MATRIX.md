# ODB Feature Matrix

This matrix records the current public `litchi-odb` capability. The crate is
an immutable OpenDocument database package reader with a bounded inert schema
catalog, not a database runtime.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODB package snapshot | 🟡 | ✅ | N/A | Exact database MIME, UTF-8, a bounded namespace-aware `office:document-content/office:body/office:database` structure, and exactly one direct `db:data-source` are required. Original bytes, raw content/styles, projected metadata, and file names are exposed. |
| Raw `content.xml` | 🟡 | ✅ | ✅ | Existing package XML remains exact and is never reformatted. Opened-document edits splice only selected tags/subtrees, including formatted LibreOffice XML. Fresh authored XML remains compactness-checked. Unknown adjacent XML is not normalized or discarded. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates a compact, structurally valid inert data-source shell plus raw content and manifest. Existing snapshots publish through the shared provenance-bearing ODF XML splice boundary. |
| Compact XML | ✅ | N/A | ✅ | Fresh input and generated content are compact: no indentation or formatting whitespace is emitted. Semantic character data and `xml:space="preserve"` content remain exact. Existing source packages are preserved byte-for-byte. |
| Connections and queries | ✅ | ✅ | ✅ | `Database::catalog()` exposes inert file, resource, or server targets and stored query metadata. The unified transaction creates, replaces, removes, and composes targets/queries; scalar query updates preserve unknown attributes. Credentials, targets, and commands are never opened, fetched, or executed. |
| Schemas, tables, columns, keys, indices, forms, and reports | 🟡 | ✅ | ✅ | The unified transaction provides create/replace/remove operations plus dependency-aware table/column renames. Foreign keys have a first-class relation projection and rename closure. Forms/reports remain inert components. Whole-owner replacement intentionally replaces that owner's opaque internals; adjacent producer XML stays exact. |
| Existing-package edits and patches | 🟡 | ✅ | ✅ | Unsigned compact or formatted packages support ordered multi-owner transactions, atomic rebuild/full reopen, exact-source application, exact inverse restoration, semantic effect history, budgeted snapshot `History`, bounded semantic transfer, canonical durable JSON patches, deterministic disjoint joins, and non-mutating three-way conflict plans. |
| Untouched-byte preservation | ✅ | ✅ | ✅ | Original bytes are exact before mutation. Changed saves publish checked source ranges through the shared provenance-bearing ODF boundary; untouched package members remain exact and inverse patches restore the accepted source artifact exactly. |
| Encryption and signatures | 🟡 | 🟡 | 🟡 | Protection presence is inventoried. The explicit fail-closed policy preserves exact no-ops and refuses changed signed/encrypted publication; password opening/writing, signing, and verification are not exposed. |
| Permanent database non-execution boundary | ✅ | N/A | N/A | No driver, network, credential, query, refresh, or database execution path exists. External configuration remains inert text. |
| Active content | 🟡 | 🟡 | 🟡 | Macros, scripts, controls, actions, DDE, and embedded code are neither executed nor semantically inventoried. |
| Limits and evidence | ✅ | ✅ | ✅ | Snapshot content and changed output have 256 MiB ceilings; unified transactions cap operations at 65,536 and semantic/opaque values at 1 MiB; durable artifacts and wire JSON have separate finite limits. Tests cover composed CRUD, relation closure and dependency dispositions, raw authored-fragment negatives, durable/inverse/stale behavior, three-way planning, a signed real ODB, and formatted LibreOffice publication. |
