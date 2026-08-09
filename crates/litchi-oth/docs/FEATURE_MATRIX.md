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
| Styles and metadata | 🟡 | ✅ | 🟡 | Named style declarations, family, parent, and declaring part are typed; metadata is projected through the common reader. Compact typed whole-part replacements participate in ordinary edits, inverse patches, join conflicts, and three-way plans. Style property editing, inheritance resolution, per-field operations, and deletion of an optional part as a composable semantic operation remain future work. |
| Text blocks and character data | 🟡 | ✅ | 🟡 | `Template::text_body()` projects ordered `text:p` and `text:h` blocks, style names, inert hyperlinks, formatting ranges, common field families and stored values, bookmarks (including cross-block ranges), and ODF whitespace without evaluation or link following. Paragraph and heading direct-text edits support character data/CDATA/entities and empty elements, and refuse nested markup rather than dropping it. Bookmark, field, run-formatting, and hyperlink mutation are not public. |
| Lists, resources, forms, and objects | 🟡 | ✅ | 🟡 | Flat and nested list instances/items, numbering restarts, paragraph positions, images, embedded/linked object references, forms, and controls are inventoried as inert typed values. Fresh and source-tail list authoring plus isolated top-level list replace/remove are typed; nested list structural edits are refused. Resource resolution/activation and resource, form, or object mutation are absent. |
| Existing-package edits and patches | 🟡 | N/A | ✅ | `Template::edit()` stages bounded paragraph/heading replacements, isolated list replace/remove, whole metadata/style replacement, and typed tail block/list additions. Commit compacts changed XML, rebuilds and reopens once, and checks semantic readback. Exact-source inverse application is byte-exact. Independent disjoint edits join with typed conflicts, and `Patch::plan_three_way` produces a non-mutating deterministic conflict plan before publication. Signed sources refuse mutation; exact no-ops share the source snapshot. |
| Durable patches | 🟡 | ✅ | ✅ | `Patch::to_bytes`/`from_bytes` use a deterministic bounded envelope, fully reopen embedded source and target snapshots, preserve exact stale-source application and inverse restore, and retain paragraph/heading/tail-append semantics. Decoded list changes are currently exact-apply-only rather than exposed as semantic list operations; this limits three-way composition after decoding. |
| Undo/redo history and transfer | 🟡 | N/A | 🟡 | Format-owned `History` wraps the common finite step/byte budget over immutable OTH snapshots and supports checked record, undo, and redo. No public cross-template transfer API exists. |
| Untouched-byte preservation | ✅ | ✅ | N/A | Original bytes are returned exactly before mutation; auxiliary files cannot be copied into a fresh builder. |
| Encryption and signatures | ❌ | ❌ | ❌ | Password operations and signature APIs are absent. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Links are never followed; scripts, macros, controls, actions, DDE, and embedded code are never executed or activated. `check_security` inventories embedded/external resources, forms, script members, and signatures and enforces a default-deny explicit policy while keeping permitted surfaces inert. |
| Limits and evidence | ✅ | ✅ | ✅ | `content.xml` uses 64 MiB and depth-256 input ceilings plus bounded block/link/field/list/attribute projections; each direct-text replacement is capped at 16 MiB and durable envelopes at 512 MiB. A packaged fixture from LibreOffice's checked-in Writer/Web template tree is opened and edited. Tests cover rich inventory, compact reopen/readback, joins and three-way plans, bounded history, deterministic durable/inverse/stale application, list replacement/removal, security policy, namespace/family failures, DTD/custom-entity rejection, and immutable source bytes. |
