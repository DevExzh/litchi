# ODC Feature Matrix

This matrix records the current public `litchi-odc` capability for packaged and
flat OpenDocument charts. The crate supports immutable inspection, fresh
authoring, and lossless-or-refuse flat axis-name transactions.

| Mark | Meaning |
|---|---|
| ✅ | Supported for the narrow scope in Notes |
| 🟡 | Partial, bounded, generic, or create-only |
| ❌ | No public support |
| N/A | The direction does not apply |

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODC package snapshot | ✅ | ✅ | N/A | Exact chart MIME is required; original package bytes, raw styles, metadata, and file names are inspectable. OTC is not accepted. |
| Flat ODC snapshot | ✅ | ✅ | ✅ | Bounded immutable snapshots retain exact bytes. Axis-name commits use exact spans, typed readback, source checks, and reversible patches; unsupported namespace insertion is refused. |
| Namespace-aware chart tree | ✅ | ✅ | ❌ | The common chart reader validates root structure and bounded XML resources while retaining unknown elements and attributes. It cannot write an opened tree back. |
| Typed chart views | 🟡 | 🟡 | N/A | Borrowed views cover the chart, first plot area/legend, axes, grids, series, domains, and points; root and series `chart:class` values are typed, namespace-resolved, and lexical-alias preserving. Remaining content is generic retained XML. |
| Fresh chart definition | 🟡 | N/A | 🟡 | Builder supports titles, plot areas, axes, series, labels, trend/error elements, stock/wall/floor elements, cached tables, and all twelve ODF 1.4 §19.15 chart classes. Explicit bounded extension and unknown QName classes are inert and retain their supplied aliases. |
| Compact XML | 🟡 | N/A | 🟡 | The typed serializer adds no formatting whitespace and passes bounded compactness validation before publication; rejected candidates return structured `XmlCompactness`. Accepted semantic character data remains exact and generated manifest XML is compact. Space-only inter-element text is still accepted by the validator, so absolute minimality is not yet guaranteed. |
| Extensions and unknown XML | 🟡 | ✅ | 🟡 | Opened extensions are retained in the immutable tree; fresh authoring can include caller-validated extension subtrees. No lossless edit round trip is claimed. |
| Styles, metadata, and resources | 🟡 | 🟡 | ❌ | Styles are raw and metadata is projected; fresh packages do not author styles, metadata, settings, media, or auxiliary resources. |
| Existing-chart edits and patches | 🟡 | ✅ | 🟡 | Flat ODC supports source-bound `chart:name` axis transactions and inverse patches. Packaged chart editing, history, and merge are not exposed. |
| Formula and range handling | 🟡 | 🟡 | 🟡 | Formula and range values are escaped inert strings; their grammar is not parsed or validated, and no data-source resolution, calculation, refresh, schema lookup, or external fetch occurs. |
| Rendering and layout | ❌ | ❌ | ❌ | No chart renderer, layout engine, or style resolver is provided. |
| Encryption, signatures, protection | ❌ | ❌ | ❌ | Password opening/writing, signing/verification, and protection are not exposed by ODC. |
| Resource limits | 🟡 | 🟡 | 🟡 | XML uses fixed internal ceilings rather than caller-selected profiles; fresh serialized content additionally has default 64 MiB and depth-256 compactness limits. The builder does not expose caller-selected compactness limits. |
| Test evidence | 🟡 | N/A | N/A | Crate-local tests cover structure, compactness, semantic whitespace, authoring, flat byte preservation, axis transactions, wrong-family refusal, malformed input, and full seed truncation/mutation sweeps. No real FODC fixture is available in the corpus. |
| Permanent non-execution boundary | ✅ | N/A | N/A | Extensions, formulas, and references remain inert. Macros, controls, actions, DDE, and embedded code are never executed or activated. |
