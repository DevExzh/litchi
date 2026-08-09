# ODG Feature Matrix

This matrix records the public `litchi-odg` contract for packaged ODG and flat
FODG. iWork formats are not part of this crate's scope.

| Mark | Meaning |
|---|---|
| yes | Supported for the narrow scope in Notes |
| partial | Partial, preservation-only, or deliberately refused outside a safe closure |
| no | No public support |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODG package snapshot | yes | yes | N/A | `Drawing` and `PackageSnapshot` retain exact bytes and expose bounded pages, layers, shapes, metadata, styles, and safe member names. |
| Package semantic model | partial | yes | partial | Namespace-aware parsing requires one `office:document-content` / `office:body` / `office:drawing` chain. Pages expose identity, style, master, page-local layers, and bounded shapes. Global layers from `styles.xml`, layer policy, shape kind/name/layer/style/text style/z-index/geometry/text/accessibility, shared inert `drawing::Frame` context, and package-local image resources are typed without conflating page-local and global scope. Unknown markup remains source bytes. |
| Unified package transaction | partial | yes | yes | One transaction composes checked page, page-local-layer, shape/group, geometry/path/style, inert form-control reference, text/name/layer, and referenced package-resource operations. Group removal owns its complete XML subtree; layer removal is dependency checked. Commit validates compact whole-package output, performs a bounded full reopen/readback, and returns an exact-source reversible patch. Adjacent exact-lineage patches compose. |
| Durable patches and collaboration | yes | N/A | yes | Versioned reversible patches use deterministic canonical JSON, exact-source preconditions, bounded content-addressed blobs for structural/resource changes, inverse application, and full package reopen. Exact-lineage sub-edits join in stable identifier order only when effects are disjoint; overlap is reported structurally. Three-way planning is non-mutating and requires explicit conflict resolution. |
| History and transfer bounds | yes | N/A | yes | Snapshot history has explicit step/weight limits, deterministic eviction, undo/redo, and non-mutating over-budget refusal. Durable JSON, operation values, blob counts/sizes, composition effects, and conflicts all have finite bounds. |
| Rewrite refusal | yes | N/A | yes | Split/mixed/CDATA/entity text, noncompact source XML, signed packages, encrypted packages, and unsupported ownership return errors rather than silently rewriting. |
| Compact XML | yes | N/A | yes | Authored and edited XML output has no indentation, padded markup, DTD, or custom entities. Opened noncompact XML remains readable but cannot be republished by this API. |
| Flat FODG snapshots and patches | partial | yes | partial | `FlatDrawing` preserves source bytes, inventories bounded pages/shapes, and has an independent source-checked reversible text patch chain. |
| Geometry, paths, styles, resources, forms | partial | partial | partial | Existing four-attribute shape geometry, SVG path data, graphic-style references, and inert `draw:control` form references are losslessly editable. Detached path/control shapes participate in checked shape CRUD. Referenced package-local image/media members and manifest media types can be replaced or removed inertly. Arbitrary style-definition and form-model CRUD remain unsupported. |
| Active content | yes | yes | N/A | Controls, scripts, actions, DDE, links, and embedded payloads are retained inertly only and are never evaluated or executed. |
| Templates, encryption, signatures | partial | partial | partial | OTG template build/open/edit is explicit and preserves the template media type. Password writes and signing are not exposed. Signed and encrypted package rewrite is refused. |
| Limits and evidence | yes | yes | yes | Parsing caps depth, pages, layers, shapes, extracted text, replacement values/resources, and output. Tests cover a real LibreOffice `.odg`, real resource XML package preservation, styles-part layers, unified structural/geometry/path/style/form/resource publication, deterministic durable reopen/inverse/stale refusal, join conflicts/order, non-mutating three-way planning, bounded history, OTG preservation, dependency refusal, DTD refusal, misplaced-shape refusal, and noncompact rewrite refusal. |
