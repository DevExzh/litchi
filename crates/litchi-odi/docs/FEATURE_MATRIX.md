# ODI Feature Matrix

This matrix records the current public `litchi-odi` capability for packaged
OpenDocument images and flat ODI snapshots. Packaged images remain an immutable
shell; flat images expose a bounded, inert frame and source inventory.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODI package snapshot | 🟡 | ✅ | N/A | Exact image MIME is required; original bytes and safe package entry names are exposed. `Image::frames` also projects inert `content.xml` frame/source semantics. |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | Snapshot opening checks UTF-8, a 256 MiB ceiling, and the prefix-sensitive `office:image` marker only. Fresh builds additionally require bounded, well-formed, DTD-free compact XML; ODF namespace and schema semantics remain unvalidated. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates MIME, caller/default content, and manifest only. It cannot save or rewrite an opened package. |
| Compact XML | 🟡 | N/A | 🟡 | Fresh input is validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Styles and metadata | 🟡 | 🟡 | ❌ | Optional styles are raw text and metadata is projected through the common model; neither is writable. |
| Frames and image sources | 🟡 | ✅ | 🟡 | Packaged and flat ODI inventory source-order inert frame names and embedded or linked sources without dereferencing them. Flat transactions edit only existing `draw:name`, `xlink:href`, or `office:binary-data` sites. |
| Images, maps, geometry, and resources | 🟡 | 🟡 | ❌ | Flat ODI scans inert image frames only. Image maps, resource CRUD, styling, rendering, conversion, and packaged semantic traversal remain absent. |
| Existing-package edits and patches | 🟡 | N/A | 🟡 | Flat ODI has a source-bound transaction, exact no-op, atomic validated commit, semantic readback, stale refusal, and inverse/apply patch. Packaged ODI remains immutable. |
| Flat ODI snapshots | ✅ | ✅ | 🟡 | `FlatImage` requires namespace-aware `office:document/office:body/office:image` placement, rejects DTDs, inventories supported inert sources, preserves exact bytes, and edits only losslessly located metadata. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original package and flat XML bytes are returned exactly while no mutation is attempted. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signing/verification are not exposed. |
| Active content | 🟡 | 🟡 | 🟡 | Links and embedded bytes are never dereferenced, decoded, activated, or executed; opaque markup is not semantically inventoried. |
| Limits and evidence | 🟡 | 🟡 | 🟡 | Package and flat inputs have a 256 MiB family ceiling; authoring first applies shared 64 MiB and depth-256 compactness limits. Flat validation also caps XML depth and frame count. Active tests cover typed structure placement, exact flat bytes, compactness rejection, and semantic whitespace; real-file flat evidence is still absent. |
