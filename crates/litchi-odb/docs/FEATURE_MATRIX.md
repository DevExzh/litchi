# ODB Feature Matrix

This matrix records the current public `litchi-odb` capability. The crate is
an immutable OpenDocument database package shell with detached inert values,
not a semantic database codec or database runtime.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODB package snapshot | 🟡 | ✅ | N/A | Exact database MIME is required; original bytes, raw content/styles, projected metadata, and file names are exposed. |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | Snapshot opening checks UTF-8, a 256 MiB ceiling, and the literal `office:database` marker only. Fresh builds additionally require bounded, well-formed, DTD-free compact XML; database grammar remains unvalidated. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates MIME, raw content, and manifest only; no existing package can be saved or rebuilt. |
| Compact XML | 🟡 | N/A | 🟡 | Fresh input is validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Connections and queries | ❌ | ❌ | ❌ | Public connection/query values are disconnected from the package codec. Query command text is inert and is never executed. |
| Schemas, tables, columns, and relations | ❌ | ❌ | ❌ | No typed database schema traversal, data model, command model, or CRUD is available. |
| Existing-package edits and patches | ❌ | ❌ | ❌ | No transaction, commit, save, stale check, or reversible patch exists. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original archive bytes are exact before mutation; no mutation preservation claim is made. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password opening/writing and signing/verification are not exposed. |
| Permanent database non-execution boundary | ✅ | N/A | N/A | No driver, network, credential, query, refresh, or database execution path exists. External configuration remains inert text. |
| Active content | 🟡 | 🟡 | 🟡 | Macros, scripts, controls, actions, DDE, and embedded code are neither executed nor semantically inventoried. |
| Limits and evidence | 🟡 | 🟡 | 🟡 | Snapshot content has a 256 MiB family ceiling; authoring first applies shared 64 MiB and depth-256 compactness limits. Active tests include typed compactness rejection, semantic-whitespace preservation, and exact inert snapshots of two real LibreOffice database packages; no typed database conformance is claimed. |
