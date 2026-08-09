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
| Styles and metadata | ✅ | ✅ | ✅ | Optional style XML remains exact; frame graphic/text style references are typed and editable. Flat `office:meta` and packaged `meta.xml` share title/author/subject/description/keyword edits that preserve unknown nodes. |
| Frames and image sources | ✅ | ✅ | ✅ | Packaged and flat ODI expose source/image link semantics, XML identity, accessibility, style/layer/z-order/transform/anchor data, absolute and relative lexical geometry, and `draw:copy-of`. One shared edit trait covers both artifact kinds without dereferencing links. |
| Images, maps, geometry, and resources | ✅ | ✅ | ✅ | Rectangle/circle/polygon maps support bounded typed replace/insert/remove while map extensions cause lossless refusal. The resource graph includes referenced, missing, and unreferenced members; safe-path auxiliary-member add/replace/remove preserves unrelated payloads. Rendering and conversion remain absent. |
| Semantic patches, merge, and transfer | ✅ | N/A | ✅ | Flat/package commits produce stable-key compare-and-set operations with exact source/target bytes, inverse and stale refusal. Join detects same-key divergence; non-mutating three-way plans report expected/actual/desired conflicts. Compatible frame/style/map/metadata operations transfer between flat and package artifacts; package-resource operations remain package-only. Publication checks the planned source again and fully reopens output bytes. |
| Existing-package edits and history | ✅ | N/A | ✅ | Flat and packaged ODI share frame and metadata edit contracts, exact no-ops, atomic semantic readback, inverse/apply patches, and safe package-resource CRUD. Commit-coupled history is source checked, branch safe, and bounded by both state count and exact serialized-byte budget. Signed, encrypted, or non-compact rewrite sources remain refused. |
| Flat ODI snapshots | ✅ | ✅ | ✅ | `FlatImage` requires namespace-aware `office:document/office:body/office:image/draw:frame/draw:image` placement, rejects DTDs and excessive depth, inventories supported inert sources, preserves exact bytes, and losslessly edits frame/style/geometry/source/map and embedded metadata sites. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original package and flat XML bytes are returned exactly while no mutation is attempted. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signing/verification are not exposed. |
| Active content | 🟡 | 🟡 | 🟡 | Links and embedded bytes are never dereferenced, decoded, activated, or executed; opaque markup is not semantically inventoried. |
| Security policy and limits | ✅ | ✅ | ✅ | Explicit policies bound operation count, retained patch bytes, resource/inline bytes, metadata text, map areas, external links, and package-member changes. Family inputs/output remain capped at 256 MiB, history at 1,024 states plus a caller byte budget, and XML depth at 256. Unsafe/reserved paths, stale plans, lossy map rewrites, signed packages, and hostile policy excesses fail before publication. |
| Evidence and provenance | 🟡 | ✅ | ✅ | Checked-in corpora were searched for both `.odi` and `.fodi`; no genuine producer artifact was present. `tests/fixtures/odf-1.4-normative-synthetic.fodi` is deliberately labeled hand-authored synthetic ODF 1.4 schema evidence, not producer provenance. Tests cover exact reopen/inverse/stale behavior, compatible joins, divergent conflicts, transfer, flat/package metadata, map/resource CRUD, byte-budgeted history, hostile policy limits, compactness, and prior regressions. |
