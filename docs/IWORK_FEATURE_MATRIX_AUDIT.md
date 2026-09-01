# iWork Feature-Matrix Audit

> Source audit starting at committed `b6e6ada83` (2026-09-01) and including the focused
> Percentage-format slice completed from that baseline. This document records the rationale and
> cross-suite gaps behind the three authoritative app matrices and does not replace them.

## Matrix state

The repository's matrix contract says detailed per-format matrices are the source of truth and uses `Feature | Status | Read | Write | Notes` with independent directions ([root matrix](FEATURE_MATRIX.md#detailed-matrices)). The new iWork matrices close the structural documentation gap; verification and certification gaps remain:

| Check | Result |
|---|---|
| Pages detailed matrix | ✅ [authoritative Pages matrix](../crates/litchi-pages/docs/FEATURE_MATRIX.md) |
| Keynote detailed matrix | ✅ [authoritative Keynote matrix](../crates/litchi-keynote/docs/FEATURE_MATRIX.md) |
| Numbers detailed matrix | ✅ [authoritative Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md) |
| Root index links | ✅ all three focused owners are linked from the [root index](FEATURE_MATRIX.md#detailed-matrices) |
| Current implementation certification | ❌ [FORMAT_IMPLEMENTATION_REVIEW.md](FORMAT_IMPLEMENTATION_REVIEW.md#scope-and-rubric) explicitly excludes all iWork |
| iWork matrix evidence links | ✅ the matrices contain direct focused-source, test, fixture, and audit-evidence links |
| Root index wording | ✅ stale matrix-count wording was removed when the iWork owners were registered |
| Current migration topology | ✅ 64 workspace packages, 237 internal dependency declarations, 226 canonical edges, 11 ordered debts, one migration host |

Status used below: ✅ fully supports the precisely stated bounded scope; 🟡 partial, metadata-only, preservation-only, host-only, or otherwise constrained; ❌ unsupported; N/A not applicable. A generated schema or detached value type alone receives no semantic-support credit.

The rows below are a conservative source-capability inventory, not a green test certification.
The focused Percentage owner gate is green at the synthetic/codec level. A disposable native
Numbers open/save/close/reopen probe plus a strict Rust semantic no-op readback is recorded in
[ADR 0008](adr/0008-migration-and-verification.md#2026-09-01-amendment-numbers-existing-cell-percentage-format-native-validation-record),
but formal native evidence promotion remains pending until the artifact and ledger are frozen.
Exact current and historical results are recorded in [IWORK_PROGRESS_AUDIT.md](IWORK_PROGRESS_AUDIT.md#current-worktree-verification).

## Shared IWA capabilities

| Feature | Status | Read | Write | Notes |
|---|---:|---:|---:|---|
| Pages/Keynote/Numbers detection | 🟡 | ✅ | N/A | Bytes are ZIP-only; path ingress accepts secure app-authored directories on Unix; Windows directory ingress fails closed |
| Direct IWA and legacy nested `Index.zip` | 🟡 | ✅ | 🟡 | Bounded decode; changed physical edits are limited to admitted direct exact sources/existing entries, while changed legacy nested catalogs are refused |
| Semantic directory snapshot | 🟡 | ✅ | ❌ | Secure frozen snapshot on Unix; not a complete package and cannot be exactly edited/reassembled |
| Exact package no-op | ✅ | ✅ | ✅ | Retained regular-ZIP source can be emitted byte-for-byte; a directory snapshot is not a complete package artifact |
| Focused changed transactions | 🟡 | ✅ | 🟡 | Selected owner transactions are source-bound, reversible, fail-closed, and reparse/reopen candidates; raw archive edits are not all reversible and this is not native-app verification |
| Unknown member/field preservation | 🟡 | ✅ | 🟡 | Strong for untouched exact artifacts; opaque members are not editable, and `roundtrip_iwork` can misroute compressed opaque bytes through the lossy Store-only writer |
| Durable patch serialization/history/merge | ❌ | N/A | ❌ | Focused owners retain process-local source/target artifacts or logical deltas; no durable/composable format |
| Focused/archive atomic durable filesystem save | 🟡 | N/A | 🟡 | `Package::save` in the Pages, Numbers, and Keynote owners delegates to the archive-owned atomic publication boundary; `write_to` remains a caller-owned streaming sink, and durable patch serialization/history/merge is still absent |
| Structured aggregate facade | 🟡 | ✅ | ❌ | Read-only tables/slides/sections/text; no rich app semantics or provenance |
| Encrypted/password packages | ❌ | ❌ | ❌ | Rejected; no password or decryption API |
| Cryptographic signatures | ❌ | ❌ | ❌ | No verification, certificate, invalidation, or re-signing model |
| External resources/actions | 🟡 | 🟡 | ❌ | Opaque/inert preservation only; no fetch, resolve, or execute path found |
| Resource limits | 🟡 | ✅ | 🟡 | Strong framing/wire/ZIP limits; output-overhead/compressed-size, graph/index, and aggregate live-memory/work ceilings remain incomplete |
| Fresh package creation in focused owners | ❌ | N/A | ❌ | Broad builders remain in the migration host |

Shared implementation references: [archive limits](../crates/litchi-iwa-archive/src/limits.rs#L5), [exact package reassembly](../crates/litchi-iwa-archive/src/package.rs#L1631), [directory boundary](../crates/litchi-iwa-archive/src/directory.rs#L81), [aggregate facade](../crates/litchi/src/iwork/mod.rs#L68).

## Pages focused owner

| Feature | Status | Read | Write | Notes |
|---|---:|---:|---:|---|
| Package open/preservation | 🟡 | ✅ | 🟡 | Regular ZIP/bytes for exact output; changed writes require retained physical provenance |
| Archive-free document/sections | ✅ | ✅ | N/A | ZIP or app-authored directory snapshot; intentionally no exact bytes or edits |
| Body/section text | 🟡 | ✅ | 🟡 | Existing storage, one bounded UTF-16 splice; dependent/reserved-marker cases refuse |
| Document/section/page settings | 🟡 | ✅ | 🟡 | Document options, name, pagination, settings, page layout, and none/solid background; no general layout graph |
| Headers and footers | 🟡 | ✅ | 🟡 | Existing logical slots and plain text only; no slot/template CRUD or formatting |
| Body footnotes | 🟡 | ✅ | 🟡 | Bounded read/insert/edit/remove/custom marker; not a general note/reference system |
| Body-table presentation properties | 🟡 | ✅ | 🟡 | Selector-driven appearance, dimensions, headers/freeze/repeat, lock, name, sort config, title; catalog/discovery remains limited |
| Table discovery/cells/formulas/merge/topology | ❌ | ❌ | ❌ | No public table catalog, values, formulas, merge, or row/column/table CRUD |
| Rich text styles/lists/links/revisions | 🟡 | 🟡 | ❌ | Text and ranges are visible; general style/annotation mutation is absent |
| Media assets | ❌ | ❌ | ❌ | Image/movie/audio modules are detached option values, not package asset APIs |
| Charts/shapes/drawings | ❌ | ❌ | ❌ | No focused package owner |
| Metadata | 🟡 | ✅ | ❌ | Canonical sidecars read only |
| Fresh package document/section structural authoring | ❌ | N/A | ❌ | Detached semantic builders exist; package creation and structural editing remain legacy-host-only |
| Collaboration/comments/coauthoring | ❌ | ❌ | ❌ | No focused semantic/package API |
| Render/PDF/HTML/RTF/automation | ❌ | ❌ | ❌ | Package output is not rendering/export/action execution |

Primary references: [package boundary](../crates/litchi-pages/src/package.rs#L456), [semantic directory boundary](../crates/litchi-pages/src/document.rs#L585), [section text transaction](../crates/litchi-pages/src/package/section_text.rs#L216), [table selector](../crates/litchi-pages/src/selector.rs#L131).

## Keynote focused owner

| Feature | Status | Read | Write | Notes |
|---|---:|---:|---:|---|
| Package open/preservation | 🟡 | ✅ | 🟡 | Exact retained-source output and focused patches; no general package authoring |
| Semantic show/slides/text | ✅ | ✅ | N/A | Read-only archive-free show, slides, notes, movies/build summaries, transitions |
| Existing placeholder text/notes | 🟡 | ✅ | 🟡 | Plain UTF-16 spans and existing storage only; no general rich-text/style graph |
| Skip/order/delete | 🟡 | ✅ | 🟡 | Existing slides; one move/delete/state transaction with topology refusals |
| Slide creation/duplication | ❌ | N/A | ❌ | No focused transaction |
| Show settings/transitions/backgrounds | 🟡 | ✅ | 🟡 | Selected existing graph and bounded fills/settings; no master/theme synthesis |
| Placeholder visibility | 🟡 | ✅ | 🟡 | Title/body/slide-number visibility only |
| Slide-table properties | 🟡 | ✅ | 🟡 | Appearance, dimensions, headers, lock, name, title, persisted sort; catalog/discovery remains limited |
| Table cells/formulas/formats/comments/topology | ❌ | ❌ | ❌ | No public focused package owner |
| Charts | 🟡 | 🟡 | 🟡 | Catalog/title/caption/axis-title/primary value-axis settings only; no data, series, type, legend, styles, or CRUD |
| Movies and soundtrack settings/order | 🟡 | 🟡 | 🟡 | Existing metadata, geometry, playback, title/caption, settings/order; no asset lifecycle; current soundtrack-order tests are red |
| Image/audio/movie bytes and asset CRUD | ❌ | ❌ | ❌ | Detached option values do not constitute package media support |
| Builds/animations | 🟡 | 🟡 | ❌ | Reduced build/effect summaries; unknown effects preserved, no edit transaction |
| Shapes/groups/lines/z-order | ❌ | ❌ | ❌ | No focused semantic/package surface |
| Masters/themes/layouts/guides | ❌ | ❌ | ❌ | No focused semantic/package surface |
| Metadata | 🟡 | ✅ | ❌ | Canonical sidecars read only |
| Fresh focused-package presentation authoring | ❌ | N/A | ❌ | Detached show/slide builders exist; package builder remains in the migration host |
| Comments/collaboration/render/export | ❌ | ❌ | ❌ | No focused owner or native renderer/exporter |

Primary references: [package boundary](../crates/litchi-keynote/src/package.rs#L406), [semantic slide model](../crates/litchi-keynote/src/slide.rs#L45), [reduced chart model](../crates/litchi-keynote/src/chart.rs#L1), [table semantic boundary](../crates/litchi-keynote/src/slide/table.rs#L1).

## Numbers focused owner

| Feature | Status | Read | Write | Notes |
|---|---:|---:|---:|---|
| Package/rooted workbook open | 🟡 | ✅ | 🟡 | Rooted sheets/tables, exact unchanged output, selected package transactions |
| Archive-free workbook/sheets/tables | ✅ | ✅ | N/A | Immutable semantic snapshot; auxiliary drawings are excluded |
| Scalar cells and bounded ranges | 🟡 | ✅ | 🟡 | Presence-preserving reads; admitted text/finite number/bool/date/duration/formula/clear writes |
| Formulas and cached values | 🟡 | 🟡 | 🟡 | Bounded AST and 13 authorable functions; no general parser, dependency engine, recalculation, or full registry |
| Sheet/table names and sheet order | 🟡 | ✅ | 🟡 | Rename/reorder existing objects only; not defined-name or lifecycle support |
| Sheet/table lifecycle and row/column topology | 🟡 | ✅ | ❌ | Existing sheet reorder is supported above; no focused add/delete/duplicate or row/column insert/delete transaction |
| Package merge state/editing | ❌ | ❌ | ❌ | Detached geometry vocabulary exists, but there is no native package merge reader or transaction |
| Table settings | 🟡 | ✅ | 🟡 | Appearance, dimensions, headers/freeze/repeat, lock, title, and persisted sort configuration; no physical sort |
| Cell controls/pop-up menus | 🟡 | 🟡 | 🟡 | Checkbox/star/slider/stepper/pop-up only, under strict graph profiles |
| Existing-cell Number/Percentage display formats | 🟡 | ✅ | 🟡 | Selector-first existing-cell Number and Percentage read/set/reset operations preserve the scalar value and unrelated bytes. Number has operation-specific native evidence; Percentage focused/synthetic coverage is green, and ADR 0008 records a disposable native Numbers open/save/close/reopen plus strict Rust semantic no-op readback. The artifact/ledger is not frozen, so Percentage E3/E4 remain pending. See the [Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md#cells-formulas-controls-and-annotations). |
| Other generic display formats and rich styles | ❌ | ❌ | ❌ | Outside the focused Number/Percentage operations above, no general package getter or formatting transaction is exposed; currency, date, duration, text, custom, and rich-style families remain unsupported at this owner boundary. |
| Cell comments/replies | 🟡 | 🟡 | 🟡 | Text-only, strict rooted/co-located ownership; no broad threads, authors, mentions, or attachments; current integration tests are red |
| Charts/shapes/text boxes/media | ❌ | ❌ | ❌ | No focused semantic/package owner |
| Filters/categories/groups/pivots | ❌ | ❌ | ❌ | Detection/refusal or schema presence only; no semantic CRUD |
| Conditional formatting/data validation | ❌ | ❌ | ❌ | No focused owner |
| Print/page setup | ❌ | ❌ | ❌ | Freeze/repeat settings are not a print model |
| Table CSV projection | 🟡 | ✅ | ❌ | Per-table export only; no import or workbook/package conversion |
| Metadata | 🟡 | 🟡 | ❌ | Sidecar support is not a complete public editable package metadata API |
| Fresh focused-package workbook authoring | ❌ | N/A | ❌ | Detached semantic builders exist; package builder remains in the migration host |
| Collaboration/external refresh/render/PDF/XLSX | ❌ | ❌ | ❌ | No focused implementation |

Primary references: [package boundary](../crates/litchi-numbers/src/package.rs#L457), [cell transaction boundary](../crates/litchi-numbers/src/package/table_cells.rs#L186), [formula limits](../crates/litchi-numbers/src/formula.rs#L18), [table projection](../crates/litchi-numbers/src/table.rs#L456).

## Legacy-host delta

The following capability exists broadly in `litchi-iwa`, but must be labeled **legacy host only** until it moves behind the concrete owner API and passes the app matrix's evidence gate:

| Capability family | Pages host | Keynote host | Numbers host |
|---|---:|---:|---:|
| Fresh source-free package builder | Present | Present | Present |
| Structural document/slide/sheet/table lifecycle | Broad | Broad | Broad |
| Rich text, text boxes, shapes, lines | Broad | Broad | Broad |
| Images, movies, audio | Broad | Broad | Broad |
| Charts | Broad | Broad | Broad |
| Table cells/formulas/rich formatting/topology | Broad | Broad | Broad |
| Focused-owner replacement complete | No | No | No |

“Broad” here is an inventory statement, not a completeness or native-acceptance grade. The host exposes native/raw identities, includes compatibility fallbacks, and is governed by the 11 ordered debts (orders `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`) in the deletion ledger ([host policy](../crates/litchi-iwa/README.md#legacy-and-raw-editing), [ADR 0028](adr/0028-iwa-monolith-exit.md)). The authoritative topology is 64 workspace packages, 237 internal dependency declarations, 226 canonical edges, and one migration host.

## Evidence grades used by the authoritative matrices

| Grade | Required evidence | Current suite posture |
|---|---|---|
| E0 | Type/schema/source exists | Extensive; not support by itself |
| E1 | Synthetic unit/integration test or Litchi self-roundtrip | Extensive for focused transactions and legacy builders |
| E2 | Checked-in Apple-produced fixture parses; exact no-op/readback | Present for one basic fixture per app |
| E3 | Litchi-mutated candidate opens in the native app without repair | Operation-specific ADR evidence exists, but no complete checked-in ledger/gate |
| E4 | Native app saves, closes, reopens; Litchi strict reread verifies semantics/locality | External/manual and incomplete; not a suite-wide CI result |

Every supported read/write cell should name its focused API owner, executable test, fixture provenance, evidence grade, native app/version when applicable, and refusal/limit boundary. Exact package preservation, semantic feature support, and native interoperability must remain separate claims.

The existing-cell Percentage row currently has green focused/synthetic evidence and a disposable
native Numbers probe recorded in ADR 0008. The probe includes strict semantic no-op readback of the
native-resaved artifact, but it must remain below formal E3/E4 promotion until that evidence is
attached to a frozen fixture/ledger. The checked native Number fixture only proves family-boundary
behavior and does not promote Percentage native support.

## Matrix maintenance requirements

1. Keep shared IWA as substrate rather than a fourth user format.
2. Retain evidence-grade/test/fixture links for every ✅ or 🟡 direction and explicit ❌ rows for unsupported families.
3. Record focused-owner and legacy-host status separately during migration.
4. Update the relevant app matrix in the same change that lands or removes a capability.
5. Promote native-interoperability claims only after the current red tests are fixed and operation-specific E3/E4 evidence is recorded.
