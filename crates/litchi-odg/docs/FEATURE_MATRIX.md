# ODG Feature Matrix

This matrix records the current public `litchi-odg` capability for packaged
OpenDocument drawings and flat FODG snapshots. Packaged drawings remain an
immutable raw shell; flat drawings expose bounded pages, shapes, and text edits.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODG package snapshot | 🟡 | ✅ | N/A | `Drawing::open` and `from_bytes` require the exact graphics MIME and expose original bytes and safe entry names. |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | Snapshot opening checks UTF-8, a 256 MiB ceiling, and the literal `office:drawing` marker only. Fresh builds additionally require bounded, well-formed, DTD-free compact XML before the marker check; ODF namespace and schema semantics remain unvalidated. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates MIME, `content.xml`, and manifest only; it does not serialize drawing semantics or preserve an opened package. |
| Compact XML | 🟡 | N/A | 🟡 | Fresh input is validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Styles and metadata | 🟡 | 🟡 | ❌ | Optional styles are raw UTF-8; metadata is projected through the common model. Neither is writable here. |
| Pages and layers | 🟡 | 🟡 | ❌ | `FlatDrawing` inventories bounded FODG pages and their names; layers and packaged semantic traversal remain absent. |
| Shapes, text, geometry, and resources | 🟡 | 🟡 | 🟡 | Flat snapshots expose bounded shapes and paragraph text. A detached transaction can replace one lossless text span with typed readback; geometry, images, styles, links, forms, and resource CRUD remain absent. |
| Existing-package edits and patches | ❌ | ❌ | ❌ | Packaged ODG has no save or edit path. This does not describe the separately supported source-checked flat patch API. |
| Flat FODG snapshots and patches | 🟡 | ✅ | 🟡 | `FlatDrawing` validates namespace-aware document/body/drawing/page placement, preserves source bytes, commits bounded escaped text changes atomically, rejects patches for nonmatching source bytes, and provides an applicable exact-byte inverse. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Package snapshots and flat snapshots return their original bytes exactly. |
| Templates, encryption, signatures | ❌ | ❌ | ❌ | OTG, password operations, and signing/verification are not exposed. |
| Active content | 🟡 | 🟡 | 🟡 | Arbitrary markup may survive only as raw bytes. It is not inventoried or activated; scripts, controls, actions, DDE, links, and embedded code are never executed. |
| Limits and evidence | 🟡 | 🟡 | 🟡 | Package content has a 256 MiB family ceiling; authoring first applies shared 64 MiB and depth-256 compactness limits. Flat parsing caps depth, pages, shapes, and replacement text. Active tests cover source mismatch, applicable inversion, typed structure failures, compactness rejection, semantic whitespace, exact flat bytes, and opaque real resource XML. |
