# ODB Feature Matrix

Feature checklist for OpenDocument database packages. ✅ Supported · 🟡 Partial · ❌ Unsupported ·
N/A Not applicable. The crate reads database packages; it is not a database runtime.

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODB package snapshot | 🟡 | ✅ | N/A | |
| Raw `content.xml` | 🟡 | ✅ | ✅ | |
| Fresh package builder | 🟡 | N/A | 🟡 | Inert data-source shell only |
| Compact XML | ✅ | N/A | ✅ | |
| Login, driver, table, application, and data-source settings | ✅ | ✅ | ✅ | Never used to connect |
| Connections and queries | ✅ | ✅ | ✅ | Queries never executed |
| Schemas, tables, columns, keys, indices, forms, and reports | ✅ | ✅ | ✅ | Form controls not modeled or executed |
| Existing-package edits and patches | 🟡 | ✅ | ✅ | |
| Untouched-byte preservation | ✅ | ✅ | ✅ | |
| Encryption and signatures | 🟡 | ✅ | 🟡 | No re-signing or re-encryption |
| Permanent database non-execution boundary | ✅ | N/A | N/A | |
| Active content | ✅ | ✅ | N/A | Inert; never executed |
| Limits and evidence | ✅ | ✅ | ✅ | |
