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
| Fresh package builder | 🟡 | N/A | 🟡 | Creates a compact, structurally valid inert data-source shell plus raw content and manifest; no existing package can be saved or rebuilt. |
| Compact XML | ✅ | N/A | ✅ | Fresh input and generated content are compact: no indentation or formatting whitespace is emitted. Semantic character data and `xml:space="preserve"` content remain exact. Existing source packages are preserved byte-for-byte. |
| Connections and queries | ✅ | ✅ | ✅ | `Database::catalog()` exposes inert file, resource, or server targets and stored query metadata. The unified transaction creates, replaces, removes, and composes targets/queries; scalar query updates preserve unknown attributes. Credentials, targets, and commands are never opened, fetched, or executed. |
| Schemas, tables, columns, keys, indices, forms, and reports | 🟡 | ✅ | ✅ | The unified transaction provides create/replace/remove operations plus dependency-aware table/column renames. Foreign keys have a first-class relation projection and rename closure. Forms/reports remain inert components. Whole-owner replacement intentionally replaces that owner's opaque internals; adjacent producer XML stays exact. |
| Existing-package edits and patches | 🟡 | ✅ | ✅ | Unsigned compact or formatted packages support ordered multi-owner transactions, atomic rebuild/full reopen, exact-source application, exact inverse restoration, semantic effect history, and budgeted snapshot `History`. Signed changed packages are refused. Durable cross-process patch serialization and three-way merge remain unsupported. |
| Untouched-byte preservation | ✅ | ✅ | ✅ | Original bytes are exact before mutation. Changed saves raw-copy every untouched ZIP member and byte-splice only selected `content.xml` ranges; inverse patches restore the accepted source artifact exactly. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password opening/writing and signing/verification are not exposed. |
| Permanent database non-execution boundary | ✅ | N/A | N/A | No driver, network, credential, query, refresh, or database execution path exists. External configuration remains inert text. |
| Active content | 🟡 | 🟡 | 🟡 | Macros, scripts, controls, actions, DDE, and embedded code are neither executed nor semantically inventoried. |
| Limits and evidence | ✅ | ✅ | ✅ | Snapshot content and changed output have 256 MiB ceilings; unified transactions cap operations at 65,536 and semantic/opaque values at 1 MiB. Tests cover composed CRUD, relation closure, exact no-op/inverse/stale behavior, and minimal editing of a real formatted LibreOffice package. |
