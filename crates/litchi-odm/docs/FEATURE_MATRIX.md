# ODM Feature Matrix

Feature checklist for OpenDocument master documents. ✅ Supported · 🟡 Partial · ❌ Unsupported ·
N/A Not applicable.

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODM package (`Package` owner, `Document` projection) | ✅ | ✅ | N/A | |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | |
| Fresh package builder | ✅ | N/A | ✅ | |
| Compact XML | 🟡 | 🟡 | ✅ | |
| Styles and metadata | 🟡 | ✅ | 🟡 | Only common metadata fields editable |
| Sections and subdocuments | ✅ | ✅ | 🟡 | Dependent body items cannot be transferred |
| Package resource graph | ✅ | ✅ | 🟡 | External sources refused |
| Permanent external-resolution boundary | ✅ | N/A | N/A | |
| Existing-package edits and patches | ✅ | ✅ | 🟡 | Signed or encrypted packages refused |
| Untouched-byte preservation | ✅ | ✅ | ✅ | |
| Encryption and signatures | 🟡 | 🟡 | ❌ | Password opening only; no verification or signing |
| Active content | 🟡 | ✅ | 🟡 | Inert; never executed |
| Limits and evidence | ✅ | ✅ | 🟡 | |
