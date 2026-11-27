# ODC Feature Matrix

Feature checklist for packaged and flat OpenDocument charts. ✅ Supported · 🟡 Partial ·
❌ Unsupported · N/A Not applicable.

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODC package snapshot | ✅ | ✅ | ✅ | |
| Flat ODC snapshot | ✅ | ✅ | ✅ | |
| Namespace-aware chart tree | ✅ | ✅ | ❌ | An opened tree cannot be written back |
| Typed chart views | 🟡 | 🟡 | N/A | Full typed definition only for lossless crate-authored charts |
| Fresh chart definition | 🟡 | N/A | 🟡 | |
| Compact XML | ✅ | N/A | ✅ | |
| Extensions and unknown XML | 🟡 | ✅ | 🟡 | Retained; no lossless edit round trip |
| Styles, metadata, and resources | 🟡 | 🟡 | 🟡 | Typed subset of chart properties |
| Existing-chart edits and patches | ✅ | ✅ | ✅ | |
| Formula and range handling | 🟡 | 🟡 | 🟡 | Validated, never evaluated or refreshed |
| Rendering and layout | ❌ | ❌ | ❌ | |
| Encryption, signatures, protection | 🟡 | 🟡 | ❌ | Inventory only; changed signed or encrypted charts refused |
| Resource limits | ✅ | ✅ | ✅ | |
| Test evidence | ✅ | N/A | N/A | |
| Permanent non-execution boundary | ✅ | N/A | N/A | |
