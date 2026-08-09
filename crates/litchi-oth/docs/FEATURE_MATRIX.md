# OTH Feature Matrix

This matrix records the current public `litchi-oth` capability for
OpenDocument HTML templates. The crate validates the ODF text-web package
envelope and exposes an inert, namespace-aware text-block projection.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, raw, detached, or preservation-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| OTH package snapshot | ✅ | ✅ | N/A | Exact text-web MIME, bounded family-envelope validation, original bytes, safe file names, raw XML, and projected metadata are exposed. |
| Raw `content.xml` | 🟡 | ✅ | ✅ | Opening accepts bounded ordinary or compact DTD-free XML and requires exactly one namespace-aware `office:document-content/office:body/office:text` chain. Fresh authoring remains byte-minimal. Prefix aliases work; a wrong namespace, root, duplicate, or misplaced family container fails. This is not full Relax NG validation. |
| Fresh package builder | 🟡 | N/A | 🟡 | Creates and reopens a compact ODF 1.4 text-web envelope with MIME, typed plain headings/paragraphs or raw content, manifest, and optional compact `styles.xml`, `meta.xml`, and `settings.xml` parts. Raw content and typed blocks cannot be mixed. |
| Compact XML | ✅ | ✅ | ✅ | Authored XML is bounded and byte-minimal: no indentation/newline/tab-only formatting or padded markup. Existing pretty-printed producer XML opens and remains byte-exact; preservation edits splice only selected character content. |
| Styles and metadata | 🟡 | 🟡 | 🟡 | Styles remain raw and metadata is projected through the common reader; compact raw style, metadata, and settings parts can be authored. |
| Text blocks and character data | 🟡 | ✅ | 🟡 | `Template::text_body()` projects ordered `text:p` and `text:h` blocks, style names, inert hyperlinks, and `text:s`/`text:tab`/`text:line-break` semantics without field expansion or link following. Paragraph edits preserve surrounding XML, support direct text/CDATA/entities and explicit or self-closing empty paragraphs, and refuse nested markup rather than dropping it. Lists, bookmarks, fields, and formatting runs are not typed. |
| Text-web semantics and resources | ❌ | ❌ | ❌ | No typed text tree, styles, lists, bookmarks, resource graph, forms, or embedded-object model is connected. |
| Existing-package edits and patches | 🟡 | N/A | 🟡 | `Template::edit()` stages multiple bounded paragraph-text replacements selected by zero-based `Position`s. Commit rebuilds and reopens once, checks every semantic readback, preserves untouched member payloads, and returns an exact-source reversible patch. Signed sources refuse mutation; exact no-ops share the source snapshot. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original bytes are returned exactly before mutation; auxiliary files cannot be copied into a fresh builder. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Links are never followed; scripts, macros, controls, actions, DDE, and embedded code are never executed or activated. Raw markup is not semantically inventoried. |
| Limits and evidence | ✅ | ✅ | ✅ | `content.xml` uses 64 MiB and depth-256 input ceilings plus bounded block/link/attribute projections; each paragraph replacement is capped at 16 MiB. Tests package and edit LibreOffice's checked-in Writer/Web template source, and cover prefix aliases, wrong family namespaces, misplaced/duplicate content, DTD/custom-entity rejection, ODF whitespace, multi-edit reversibility, and immutable byte preservation. |
