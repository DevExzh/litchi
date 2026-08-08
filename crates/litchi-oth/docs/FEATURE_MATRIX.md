# OTH Feature Matrix

This matrix records the current public `litchi-oth` capability for
OpenDocument HTML templates. The crate is an immutable raw package shell; its
small text values are not connected to a document codec.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| OTH package snapshot | 🟡 | ✅ | N/A | Exact text-web MIME is required; original bytes, safe file names, raw XML, and projected metadata are exposed. |
| Raw `content.xml` | 🟡 | 🟡 | 🟡 | Snapshot opening checks UTF-8, a 256 MiB ceiling, and the literal `office:text` marker only. Fresh builds additionally require bounded, well-formed, DTD-free compact XML; ODF text-web grammar remains unvalidated. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates MIME, raw content, and manifest only; no opened-template save path exists. |
| Compact XML | 🟡 | N/A | 🟡 | Fresh input is validated before publication; indentation line breaks/tabs and padded markup return structured `XmlCompactness` errors. Accepted bytes, semantic character data, and `xml:space="preserve"` content remain exact. Space-only inter-element text is still accepted, so absolute minimality is not yet guaranteed. |
| Styles and metadata | 🟡 | 🟡 | ❌ | Styles are raw and metadata is projected through the common reader; authoring is absent. |
| Paragraphs and links | ❌ | ❌ | ❌ | Public value constructors are detached from parsing and serialization and do not establish semantic document support. |
| Text-web semantics and resources | ❌ | ❌ | ❌ | No typed text tree, styles, lists, bookmarks, resource graph, forms, or embedded-object model is connected. |
| Existing-package edits and patches | ❌ | ❌ | ❌ | No save, transaction, commit, source check, or inverse patch exists. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original bytes are returned exactly before mutation; auxiliary files cannot be copied into a fresh builder. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Links are never followed; scripts, macros, controls, actions, DDE, and embedded code are never executed or activated. Raw markup is not semantically inventoried. |
| Limits and evidence | 🟡 | 🟡 | 🟡 | Snapshot content has a 256 MiB family ceiling; authoring first applies shared 64 MiB and depth-256 compactness limits. Active synthetic tests include typed compactness rejection, semantic-whitespace preservation, and exact immutable snapshot bytes; no text-web semantic conformance is claimed. |
