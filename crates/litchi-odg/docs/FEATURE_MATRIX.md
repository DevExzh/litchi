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
| Unified package transaction | partial | yes | yes | One transaction composes checked page, page-local-layer, shape/group, geometry, style, text/name/layer, and referenced package-resource operations. Group removal owns its complete XML subtree; layer removal is dependency checked. Commit validates compact whole-package output, performs a bounded full reopen/readback, and returns an exact-source reversible patch. Adjacent exact-lineage patches compose. |
| Rewrite refusal | yes | N/A | yes | Split/mixed/CDATA/entity text, noncompact source XML, signed packages, encrypted packages, and unsupported ownership return errors rather than silently rewriting. |
| Compact XML | yes | N/A | yes | Authored and edited XML output has no indentation, padded markup, DTD, or custom entities. Opened noncompact XML remains readable but cannot be republished by this API. |
| Flat FODG snapshots and patches | partial | yes | partial | `FlatDrawing` preserves source bytes, inventories bounded pages/shapes, and has an independent source-checked reversible text patch chain. |
| Geometry, styles, resources, forms | partial | partial | partial | Existing four-attribute shape geometry and graphic-style references are losslessly editable. Referenced package-local image/media members and manifest media types can be replaced or removed inertly. Arbitrary style-definition and form/control CRUD remain unsupported. |
| Active content | yes | yes | N/A | Controls, scripts, actions, DDE, links, and embedded payloads are retained inertly only and are never evaluated or executed. |
| Templates, encryption, signatures | no | no | no | OTG, password writes, and signing are not exposed. Signed and encrypted package rewrite is refused. |
| Limits and evidence | partial | yes | yes | Parsing caps depth, pages, layers, shapes, extracted text, replacement values/resources, and output. Tests cover a real LibreOffice `.odg`, real resource XML package preservation, styles-part layers, semantic parse, unified structural/geometry/style/resource publication, exact inverse, stale-source rejection, dependency refusal, DTD refusal, misplaced-shape refusal, and noncompact rewrite refusal. |
