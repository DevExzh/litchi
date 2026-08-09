# ODI Feature Matrix

This matrix records the current public `litchi-odi` capability for packaged
OpenDocument images and flat ODI snapshots. Both surfaces expose a bounded,
inert frame, map, source, metadata, and package-resource inventory.

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
| Fresh package builder | ✅ | N/A | ✅ | The deterministic baseline contains a valid embedded 1×1 PNG. Typed frame and image-map authoring supports linked/inline sources, accessibility, styling, geometry, layers, and z-order. Compact `styles.xml`/`meta.xml` plus bounded typed package members are accepted. |
| Compact XML | 🟡 | N/A | 🟡 | Fresh input is validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Styles and metadata | ✅ | ✅ | ✅ | Optional style XML remains exact; frame graphic/text style references are typed and editable. Common title/author/subject/description/keyword metadata edits patch an existing compact `meta.xml` while preserving unknown nodes. |
| Frames and image sources | ✅ | ✅ | ✅ | Packaged and flat ODI expose source/image link semantics, XML identity, accessibility, style/layer/z-order/transform/anchor data, absolute and relative lexical geometry, and `draw:copy-of`. One shared edit trait covers both artifact kinds without dereferencing links. |
| Images, maps, geometry, and resources | ✅ | ✅ | ✅ | Rectangle/circle/polygon maps are bounded, namespace-aware, inert, accessible, and authorable. The resource graph includes referenced, missing, and unreferenced package members; resource bytes remain lazy and reversible CRUD preserves unrelated members. Rendering and conversion remain absent. |
| Existing-package edits and patches | ✅ | N/A | ✅ | Flat and packaged ODI have a shared frame-edit contract, exact no-ops, atomic semantic readback, stale refusal, inverse/apply patches, metadata patching, and package resource CRUD. A generic bounded exact-source history provides branch-safe undo/redo for both surfaces. Signed, encrypted, or non-compact rewrite sources remain refused. |
| Flat ODI snapshots | ✅ | ✅ | 🟡 | `FlatImage` requires namespace-aware `office:document/office:body/office:image/draw:frame/draw:image` placement, rejects DTDs and excessive depth, inventories supported inert sources, preserves exact bytes, and edits only losslessly located metadata. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original package and flat XML bytes are returned exactly while no mutation is attempted. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signing/verification are not exposed. |
| Active content | 🟡 | 🟡 | 🟡 | Links and embedded bytes are never dereferenced, decoded, activated, or executed; opaque markup is not semantically inventoried. |
| Limits and evidence | 🟡 | ✅ | ✅ | Package/output and flat inputs have a 256 MiB family ceiling; resource and image-map area counts are capped at 100,000, history at 1,024 states, and XML depth at 256. Tests cover malformed/spoofed maps, broader frame semantics, map/style/metadata authoring, graph preservation, shared edits, history, compactness, and existing regression cases. A search of checked-in corpora still found no producer `.odi` artifact. |
