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
| Fresh package builder | 🟡 | N/A | 🟡 | Creates and reopens a compact ODF 1.4 text-web envelope with MIME, typed headings, paragraphs, and flat lists or raw content, manifest, and optional compact `styles.xml`, `meta.xml`, and `settings.xml` parts. Raw content and typed blocks cannot be mixed. |
| Compact XML | ✅ | ✅ | ✅ | Authored XML is bounded and byte-minimal. Existing pretty-printed producer XML opens byte-exact; changed XML parts are compacted without field evaluation, link following, or object activation before the strict package publication boundary. |
| Styles and metadata | 🟡 | ✅ | 🟡 | Named style declarations, family, parent, and declaring part are typed; metadata is projected through the common reader. Compact raw style, metadata, and settings parts can be authored. Style inheritance and granular metadata/style transactions remain future work. |
| Text blocks and character data | 🟡 | ✅ | 🟡 | `Template::text_body()` projects ordered `text:p` and `text:h` blocks, style names, inert hyperlinks, formatting ranges, common field families and stored values, bookmarks (including cross-block ranges), and ODF whitespace without evaluation or link following. Paragraph edits support direct text/CDATA/entities and empty paragraphs, and refuse nested markup rather than dropping it. |
| Lists, resources, forms, and objects | 🟡 | ✅ | 🟡 | Flat and nested list instances/items, numbering restarts, paragraph positions, images, embedded/linked object references, forms, and controls are inventoried as inert typed values. Fresh and source-tail list authoring is typed; resource resolution, activation, and granular form/object mutation are intentionally absent. |
| Existing-package edits and patches | 🟡 | N/A | ✅ | `Template::edit()` stages multiple bounded paragraph replacements and typed tail block/list additions. Commit compacts changed XML, rebuilds and reopens once, checks semantic readback, and returns an exact-source reversible patch. Independent disjoint edits join with typed overlap refusals; signed sources refuse mutation; exact no-ops share the source snapshot. |
| Undo/redo history | ✅ | N/A | ✅ | Format-owned `History` wraps the common finite step/byte budget over immutable OTH snapshots and supports checked record, undo, and redo. The in-memory exact-source patch is not yet a cross-process serialized patch envelope. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original bytes are returned exactly before mutation; auxiliary files cannot be copied into a fresh builder. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Links are never followed; scripts, macros, controls, actions, DDE, and embedded code are never executed or activated. Raw markup is not semantically inventoried. |
| Limits and evidence | ✅ | ✅ | ✅ | `content.xml` uses 64 MiB and depth-256 input ceilings plus bounded block/link/field/list/attribute projections; each paragraph replacement is capped at 16 MiB. A packaged fixture from LibreOffice's checked-in Writer/Web template tree is opened and edited. Tests cover rich text-web inventory, compact reopen/readback, join overlap and lineage refusals, bounded history, inverse patches, namespace/family failures, DTD/custom-entity rejection, and immutable source bytes. |
