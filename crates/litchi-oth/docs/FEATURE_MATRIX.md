# OTH Feature Matrix

This matrix records the current public `litchi-oth` capability for
OpenDocument HTML templates. The crate validates the ODF text-web package
envelope and exposes a narrow, inert paragraph-character-data projection.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| OTH package snapshot | ✅ | ✅ | N/A | Exact text-web MIME, bounded family-envelope validation, original bytes, safe file names, raw XML, and projected metadata are exposed. |
| Raw `content.xml` | 🟡 | ✅ | ✅ | Opening and fresh authoring require compact, DTD-free XML and exactly one namespace-aware `office:document-content/office:body/office:text` chain. Prefix aliases work; a wrong namespace, root, duplicate, or misplaced family container fails. This is not full Relax NG validation. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates a compact ODF 1.4 text-web envelope with MIME, raw content, and manifest. |
| Compact XML | ✅ | ✅ | ✅ | Every accepted `content.xml` is bounded and compact: no indentation/newline/tab-only formatting or padded markup. Semantic character data and `xml:space="preserve"` content remain exact. |
| Styles and metadata | 🟡 | 🟡 | ❌ | Styles are raw and metadata is projected through the common reader; authoring is absent. |
| Paragraph character data | 🟡 | ✅ | 🟡 | `Template::text_body()` projects ordered `text:p` character data, including inline character data, without field expansion or link following. An edit can replace only a paragraph with one direct XML text span; mixed, nested, CDATA, entity, and empty markup refuses rather than being reconstructed. It does not yet model ODF whitespace elements, headings, lists, bookmarks, styles, or formatting. |
| Text-web semantics and resources | ❌ | ❌ | ❌ | No typed text tree, styles, lists, bookmarks, resource graph, forms, or embedded-object model is connected. |
| Existing-package edits and patches | 🟡 | N/A | 🟡 | `Template::edit()` stages one bounded paragraph-text replacement selected by a zero-based `Position`. Commit rebuilds and reopens the package, checks semantic readback, preserves untouched member payloads, and returns an exact-source reversible patch. Signed or non-compact XML sources refuse mutation; exact no-ops share the source snapshot. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original bytes are returned exactly before mutation; auxiliary files cannot be copied into a fresh builder. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Links are never followed; scripts, macros, controls, actions, DDE, and embedded code are never executed or activated. Raw markup is not semantically inventoried. |
| Limits and evidence | ✅ | ✅ | ✅ | `content.xml` uses shared 64 MiB and depth-256 compactness limits plus an OTH-specific bounded expanded-name stack; paragraph replacements are capped at 16 MiB. Active local fixtures cover ODF 1.4 text-web opening, prefix aliases, wrong family namespaces, misplaced/duplicate content, DTD rejection, semantic whitespace, reversible semantic edits, and immutable byte preservation. No checked-in LibreOffice `.oth` corpus artifact is available. |
