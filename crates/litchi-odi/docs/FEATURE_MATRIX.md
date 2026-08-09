# ODI Feature Matrix

This matrix records the current public `litchi-odi` capability for packaged
OpenDocument images and flat ODI snapshots. Both surfaces expose a bounded,
inert frame and source inventory; package edits remain deliberately narrow.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODI package snapshot | ✅ | ✅ | ✅ | Exact image MIME, namespace-aware family structure, original bytes, safe package names, inert frames, and package-local resources are exposed. |
| Raw `content.xml` | 🟡 | ✅ | 🟡 | Opening checks UTF-8, a 256 MiB ceiling, the namespace-aware document/body/image shape, and ODF 1.4's single frame/single image contract. Full Relax NG validation remains absent. |
| Fresh package builder | ✅ | N/A | ✅ | The deterministic baseline contains a valid embedded 1×1 PNG. Typed `frame` authoring supports linked/inline sources and accessible text; `resource` adds bounded package members with explicit manifest media types. Raw compact content remains available. |
| Compact XML | 🟡 | N/A | 🟡 | Fresh input is validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Styles and metadata | 🟡 | 🟡 | ❌ | Optional styles are raw text and metadata is projected through the common model; neither is writable. |
| Frames and image sources | ✅ | ✅ | ✅ | Packaged and flat ODI expose the one normative frame, source, XML identity, accessible title/description, declared media type, and lexical geometry without dereferencing links. Transactions edit existing `draw:name`, `xlink:href`, or `office:binary-data` sites. |
| Images, maps, geometry, and resources | 🟡 | ✅ | 🟡 | Geometry is inventoried losslessly. Package-local resources have typed present/missing state, safe resolved paths, manifest media types, lazy bytes, fresh authoring, and reversible add/replace/remove edits. Image maps, styling, rendering, and conversion remain absent. |
| Existing-package edits and patches | ✅ | N/A | ✅ | Flat and packaged ODI have source-bound transactions, exact no-ops, atomic validated commits, semantic readback, stale refusal, and inverse/apply patches. Package edits preserve untouched member payloads and support reversible resource CRUD. Any ODF `META-INF/*signatures*` member and encrypted/non-compact rewrite source is refused. |
| Flat ODI snapshots | ✅ | ✅ | 🟡 | `FlatImage` requires namespace-aware `office:document/office:body/office:image/draw:frame/draw:image` placement, rejects DTDs and excessive depth, inventories supported inert sources, preserves exact bytes, and edits only losslessly located metadata. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original package and flat XML bytes are returned exactly while no mutation is attempted. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signing/verification are not exposed. |
| Active content | 🟡 | 🟡 | 🟡 | Links and embedded bytes are never dereferenced, decoded, activated, or executed; opaque markup is not semantically inventoried. |
| Limits and evidence | 🟡 | ✅ | ✅ | Package/output and flat inputs have a 256 MiB family ceiling; authoring first applies shared compactness limits, resource count is capped at 100,000, and XML depth is capped at 256. Twenty active tests cover namespace aliases/spoofing, normative cardinality, excessive depth, minimal semantic authoring, package-resource CRUD/readback/preservation, generalized signature refusal, flat/package reversible edits, compactness, and semantic whitespace. No checked-in LibreOffice `.odi` artifact is available. |
