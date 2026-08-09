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
| Package semantic model | partial | yes | partial | Namespace-aware parsing requires one `office:document-content` / `office:body` / `office:drawing` chain. Pages, declared layers, shape kind/name/layer/text, and shared inert `drawing::Frame` context are typed. Unknown markup remains source bytes. |
| Package shape transaction | partial | yes | partial | One `set_shape_text` or `set_shape_name` operation replaces only a losslessly addressable plain paragraph span or existing `draw:name` value. Commit validates compact whole-package output, reparses, verifies typed readback, and returns an exact-source reversible in-memory patch. |
| Rewrite refusal | yes | N/A | yes | Split/mixed/CDATA/entity text, noncompact source XML, signed packages, encrypted packages, and unsupported ownership return errors rather than silently rewriting. |
| Compact XML | yes | N/A | yes | Authored and edited XML output has no indentation, padded markup, DTD, or custom entities. Opened noncompact XML remains readable but cannot be republished by this API. |
| Flat FODG snapshots and patches | partial | yes | partial | `FlatDrawing` preserves source bytes, inventories bounded pages/shapes, and has an independent source-checked reversible text patch chain. |
| Geometry, styles, resources, forms | partial | partial | no | Frame occurrence context is read-only. Geometry/resource/form CRUD is not exposed. |
| Active content | yes | yes | N/A | Controls, scripts, actions, DDE, links, and embedded payloads are retained inertly only and are never evaluated or executed. |
| Templates, encryption, signatures | no | no | no | OTG, password writes, and signing are not exposed. Signed and encrypted package rewrite is refused. |
| Limits and evidence | partial | yes | yes | Parsing caps depth, pages, layers, shapes, extracted text, replacement text/name, and output. Tests cover real resource package preservation, semantic parse, atomic source-checked text and name patch/inverse behavior, stale-source rejection, DTD refusal, and noncompact rewrite refusal. |
