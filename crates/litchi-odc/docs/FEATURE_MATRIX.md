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
| Flat ODC snapshot | ✅ | ✅ | ✅ | Bounded immutable snapshots retain exact bytes. Axis-name commits use exact spans, typed readback, source checks, and reversible patches; unsupported namespace insertion is refused. |
| Namespace-aware chart tree | ✅ | ✅ | ❌ | The common chart reader validates root structure and bounded XML resources while retaining unknown elements and attributes. It cannot write an opened tree back. |
| Typed chart views | 🟡 | 🟡 | N/A | Borrowed views cover the chart, first plot area/legend, axes, grids, series, domains, and points; root and series `chart:class` values are typed, namespace-resolved, and lexical-alias preserving. Remaining content is generic retained XML. |
| Fresh chart definition | 🟡 | N/A | 🟡 | Builder supports titles, plot areas, axes, series, labels, trend/error elements, stock/wall/floor elements, cached tables, and all twelve ODF 1.4 §19.15 chart classes. Explicit bounded extension and unknown QName classes are inert and retain their supplied aliases. |
| Compact XML | 🟡 | N/A | 🟡 | The typed serializer adds no formatting whitespace and passes bounded compactness validation before publication; rejected candidates return structured `XmlCompactness`. Accepted semantic character data remains exact and generated manifest XML is compact. Space-only inter-element text is still accepted by the validator, so absolute minimality is not yet guaranteed. |
| Extensions and unknown XML | 🟡 | ✅ | 🟡 | Opened extensions are retained in the immutable tree; fresh authoring can include caller-validated extension subtrees. No lossless edit round trip is claimed. |
| Styles, metadata, and resources | 🟡 | 🟡 | 🟡 | Styles remain raw and metadata is projected. Fresh and edited packages support validated whole-`styles.xml` creation/replacement/removal plus inert package-resource create/update/delete with manifest media types, bounded inventory, readback, and inverse patches. |
| Existing-chart edits and patches | ✅ | ✅ | ✅ | Flat and packaged ODC support source-bound `chart:name` axis transactions, exact no-ops, typed readback, stale-source refusal, and inverse patches. Detached definitions add plot/axis/series/data/style CRUD, composable semantic patches, and bounded undo/redo history. Package patches compose across contiguous exact snapshots and whole typed replacement remains explicit. |
| Formula and range handling | 🟡 | 🟡 | 🟡 | ODF 1.4 chart range-list syntax and inert cached-formula prefix/delimiter/reference structure are validated on open and publication. No formula evaluation, data-source resolution, refresh, schema lookup, or external fetch occurs. |
| Rendering and layout | ❌ | ❌ | ❌ | No chart renderer, layout engine, or style resolver is provided. |
| Encryption, signatures, protection | ❌ | ❌ | ❌ | Password opening/writing, signing/verification, and protection are not exposed by ODC. |
| Resource limits | ✅ | ✅ | ✅ | Snapshots retain caller-selected package/content bytes, XML depth, axes, series, expanded points, resources, history, and scalar ceilings within hard safety bounds; builders, edits, patches, and reopen reuse them. |
| Test evidence | 🟡 | N/A | N/A | Crate-local tests cover structure, compactness, semantic whitespace, authoring, flat byte preservation, formula/range refusals, caller limits, granular definition CRUD/history/composition, package style/resource CRUD, full reopen, inverse/stale refusal, signed/non-compact rewrite refusal, wrong-family refusal, malformed input, and full seed truncation/mutation sweeps. Two genuine vendored LibreOffice chart subdocuments (LibreOffice 3.5 and 25.8 / ODF 1.4) are wired directly from the checked-in corpus. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Extensions, formulas, and references remain inert. Macros, controls, actions, DDE, and embedded code are never executed or activated. |
