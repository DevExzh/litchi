# ODC Feature Matrix

This matrix records the current public `litchi-odc` capability for packaged and
flat OpenDocument charts. The crate supports immutable inspection, fresh
authoring, and lossless-or-refuse flat and packaged transactions.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, bounded, generic, or create-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODC package snapshot | ✅ | ✅ | ✅ | Exact chart MIME is required; original package bytes, raw styles, metadata, and file names are inspectable. OTC is not accepted. Package axis edits and explicit full chart-definition replacement are atomic and reversible. |
| Flat ODC snapshot | ✅ | ✅ | ✅ | Bounded immutable snapshots retain exact bytes. Axis name/style-reference plus controlled chart class/style/size, plot style/range/geometry, and series class/style/range/label/attachment commits use namespace-resolved exact spans, full-candidate validation, typed readback, source checks, and reversible patches; unsupported namespace insertion is refused. |
| Namespace-aware chart tree | ✅ | ✅ | ❌ | The common chart reader validates root structure and bounded XML resources while retaining unknown elements and attributes. It cannot write an opened tree back. |
| Typed chart views | 🟡 | 🟡 | N/A | Borrowed views cover the chart, first plot area/legend, axes, grids, series, domains, and points. Canonical crate-authored package charts additionally project into the complete typed definition only when byte-identical reserialization proves that projection is lossless; all other opened XML remains a generic retained tree. |
| Fresh chart definition | 🟡 | N/A | 🟡 | Builder supports titles, plot areas, axes, series, labels, trend/error elements, stock/wall/floor elements, cached tables, and all twelve ODF 1.4 §19.15 chart classes. Explicit bounded extension and unknown QName classes are inert and retain their supplied aliases. |
| Compact XML | ✅ | N/A | ✅ | Canonical typed chart publication emits ODF 1.4 content with no formatting whitespace between markup and passes caller-bounded compactness validation. Semantic text whitespace remains data and is preserved; opened non-minimal XML is never silently normalized by a lossless edit. |
| Extensions and unknown XML | 🟡 | ✅ | 🟡 | Opened extensions are retained in the immutable tree; fresh authoring can include caller-validated extension subtrees. No lossless edit round trip is claimed. |
| Styles, metadata, and resources | 🟡 | 🟡 | 🟡 | Styles remain raw and metadata is projected. Fresh and edited packages support namespace-aware, bounded whole-`styles.xml` validation and creation/replacement/removal plus inert package-resource create/update/delete with manifest media types, bounded inventory, readback, and inverse patches. `office:scripts` in styles is refused. No cascade or rendering resolver is provided. |
| Existing-chart edits and patches | ✅ | ✅ | ✅ | Flat and noncanonical packaged ODC retain source-bound exact-span axis, chart, plot-area, and series attribute edits without normalizing untouched XML. Canonical package charts additionally expose granular typed plot/axis/series/data/style edits in the same atomic commit as styles and resources. Exact-source patches compose, invert, serialize deterministically, fully reopen on decode, three-way join without mutating inputs, and report stable semantic conflicts. Definition and package patches transfer chart/style/data/resource dependencies onto independently evolved snapshots; noncanonical exact-span join/transfer proceeds only when replay proves the recorded summary reproduces the exact changed content. Package history is commit-coupled, bounded, and exact-byte undo/redo. Whole typed replacement remains explicit. |
| Formula and range handling | 🟡 | 🟡 | 🟡 | ODF 1.4 chart range-list syntax and inert cached-formula prefix/delimiter/reference structure are validated on open and publication. No formula evaluation, data-source resolution, refresh, schema lookup, or external fetch occurs. |
| Rendering and layout | ❌ | ❌ | ❌ | No chart renderer, layout engine, or style resolver is provided. |
| Encryption, signatures, protection | 🟡 | 🟡 | ❌ | Signature metadata and manifest encryption are inventoried. Changed publication refuses signed or encrypted envelopes before rewrite. Password opening/writing and signature creation/verification are not exposed by ODC. |
| Resource limits | ✅ | ✅ | ✅ | Snapshots retain caller-selected package/content bytes, XML depth, axes, series, expanded points, cached rows/cells, range-list items, resources, history, and scalar ceilings within hard safety bounds; builders, edits, patches, and reopen reuse them. |
| Test evidence | 🟡 | N/A | N/A | Crate-local tests cover structure, compactness, semantic whitespace, authoring, flat byte preservation, controlled noncanonical chart/plot/series exact spans and durable summaries, formulas and cell/row/column ranges, caller limits, opened-definition projection/editing, semantic/package join and conflict paths, chart/style/data/resource transfer, deterministic patch decode, commit-coupled history, style/script refusal, signature/encryption edit policy, full package reopen/readback, inverse/stale refusal, wrong-family refusal, malformed input, and full seed truncation/mutation sweeps. A checked-in corpus and history search found genuine LibreOffice chart subdocuments embedded inside `.fods`/`.fodt` producer files, but no standalone producer `.odc` or `.fodc`; those fragments are not relabeled as standalone ODC/FODC fixture evidence. See `PRODUCER_EVIDENCE.md`. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Extensions, formulas, and references remain inert. Macros, controls, actions, DDE, and embedded code are never executed or activated. |
