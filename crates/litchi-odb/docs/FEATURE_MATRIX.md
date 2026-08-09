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
| Raw `content.xml` | 🟡 | ✅ | 🟡 | Existing package XML remains exact and is never reformatted. Fresh compact input is checked before publication. Unknown XML is not normalized or discarded. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates a compact, structurally valid inert data-source shell plus raw content and manifest; no existing package can be saved or rebuilt. |
| Compact XML | ✅ | N/A | ✅ | Fresh input and generated content are compact: no indentation or formatting whitespace is emitted. Semantic character data and `xml:space="preserve"` content remain exact. Existing source packages are preserved byte-for-byte. |
| Connections and queries | 🟡 | ✅ | 🟡 | `Database::catalog()` exposes one inert file, resource, or server connection target plus bounded stored query names, command text, and `db:escape-processing`. A single-query transaction can replace command text and set or remove escape-processing metadata. Credentials and driver configuration are not modeled; targets and commands are never opened, fetched, or executed. |
| Schemas, tables, columns, keys, indices, forms, and reports | 🟡 | ✅ | ❌ | Source-bound `Catalog` exposes table presentation and schema declarations, validated ODF column types and constraints, keys with referential actions, indices, and inert linked or embedded form/report component declarations. Producer extensions and schema edits remain unsupported. |
| Existing-package edits and patches | 🟡 | ✅ | 🟡 | Compact unsigned packages support one source-bound stored-query edit, atomic rebuild and semantic readback, exact-source applicability checks, and inverse patches. Signed or noncompact rewrite sources are refused. Durable cross-process patch serialization is not implemented. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original archive bytes are exact before mutation; no mutation preservation claim is made. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password opening/writing and signing/verification are not exposed. |
| Permanent database non-execution boundary | ✅ | N/A | N/A | No driver, network, credential, query, refresh, or database execution path exists. External configuration remains inert text. |
| Active content | 🟡 | 🟡 | 🟡 | Macros, scripts, controls, actions, DDE, and embedded code are neither executed nor semantically inventoried. |
| Limits and evidence | ✅ | ✅ | 🟡 | Snapshot content has a 256 MiB family ceiling. Catalog discovery defaults to 64 MiB, 1,000,000 XML events, depth 512, 65,536 tables and queries, 1,000,000 columns, and 1 MiB semantic attributes; all are configurable through `Limits`. Active tests cover namespace aliases, real LibreOffice table/query packages, malformed family bodies, compact authoring, semantic whitespace, and exact inert snapshots. |
