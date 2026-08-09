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
| Fresh package builder | 🟡 | N/A | 🟡 | Creates a compact ODF 1.4 text-web envelope with MIME, raw content, and manifest. Existing-package mutation remains unavailable. |
| Compact XML | ✅ | ✅ | ✅ | Every accepted `content.xml` is bounded and compact: no indentation/newline/tab-only formatting or padded markup. Semantic character data and `xml:space="preserve"` content remain exact. |
| Styles and metadata | 🟡 | 🟡 | ❌ | Styles are raw and metadata is projected through the common reader; authoring is absent. |
| Paragraph character data | 🟡 | ✅ | ❌ | `Template::text_body()` projects ordered `text:p` character data, including inline character data, without field expansion or link following. It does not yet model ODF whitespace elements, headings, lists, bookmarks, styles, or formatting. |
| Text-web semantics and resources | ❌ | ❌ | ❌ | No typed text tree, styles, lists, bookmarks, resource graph, forms, or embedded-object model is connected. |
| Existing-package edits and patches | 🟡 | N/A | 🟡 | `Template::edit().commit()` is an exact no-op, failure-atomic lifecycle with a source-checked reversible patch. No content mutation is represented. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original bytes are returned exactly before mutation; auxiliary files cannot be copied into a fresh builder. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Links are never followed; scripts, macros, controls, actions, DDE, and embedded code are never executed or activated. Raw markup is not semantically inventoried. |
| Limits and evidence | ✅ | ✅ | ✅ | `content.xml` uses shared 64 MiB and depth-256 compactness limits plus an OTH-specific bounded expanded-name stack. Active local fixtures cover ODF 1.4 text-web opening, prefix aliases, wrong family namespaces, misplaced/duplicate content, DTD rejection, semantic whitespace, exact no-op source checks, and immutable byte preservation. |
