# iWork Feature-Matrix Audit

> Source audit starting at committed `d33f30f41` (2026-09-01) and including the focused
> Scientific- and Fraction-format owner work plus the current Keynote physical `Sort Now` owner
> work from that baseline. This document records the rationale and cross-suite gaps behind the three
> authoritative app matrices and does not replace them.

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
| Current migration topology | ✅ 64 workspace packages, 238 internal dependency declarations, 227 canonical edges, 11 development-only edges, 11 ordered debts `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, one migration host; no physical-sort debt/gate closure |

Status used below: ✅ fully supports the precisely stated bounded scope; 🟡 partial, metadata-only, preservation-only, host-only, or otherwise constrained; ❌ unsupported; N/A not applicable. A generated schema or detached value type alone receives no semantic-support credit.

The rows below are a conservative source-capability inventory, not a green test certification.
The focused Percentage owner gate is green at the synthetic/codec level. A disposable native
Numbers open/save/close/reopen probe plus a strict Rust semantic no-op readback is recorded in
[ADR 0008](adr/0008-migration-and-verification.md#2026-09-01-amendment-numbers-existing-cell-percentage-format-native-validation-record),
but formal native evidence promotion remains pending until the artifact and ledger are frozen. The
focused Currency owner and codec are now source-visible with the focused codec filter green
(4/4), the `litchi-numbers-wire` library is green at 32/32, and the latest package target is green
at 22/22 after native BNC flag handling became shape-dependent. The operation-specific native Numbers candidate opened cleanly and
survived native save/close/reopen with strict Rust no-op reread; exact hashes and settings are
recorded in ADR 0008. This is operation-specific E3/E4 evidence, not a broad-format claim.
The Scientific owner now has an archive-free value, selector-first transaction, strict Buffa
type-259 codec, green focused tests, bounded sanitizer smokes, and a frozen operation-specific
native ledger. Numbers 14.4 opened, saved, closed, and reopened the Rust candidate without repair;
strict reread verified the requested precision and emitted an exact no-op. This is narrowly scoped
E3/E4 evidence, not a broad-format claim.
The Fraction owner now has an archive-free value, a selector-first exact-source transaction,
strict native type-262 preflight with a lazy Buffa view, and focused source/build/test/fuzz
evidence. It covers all nine `FractionAccuracy` strategies. The optional native field
`requires_fraction_replacement` (field 20) is accepted and preserved when absent or canonically
encoded as `false`; a canonical `true` is rejected because replacement semantics are not owned by
this seam. Computer Use additionally established operation-specific native E3/E4 evidence for a
representative Eighths-to-Hundredths edit; source, candidate, and native-resaved hashes are frozen
in ADR 0008. This does not promote all nine native UI variants, arbitrary-producer parity, native
byte parity, or package-wide performance.
The Keynote physical `Sort Now` owner now has a selector-first exact-source transaction over an
existing table, but it exposes no public row/value reader. Admission is limited to the canonical
type-6001 table-model route and explicitly proven tile, data-list, header, UID, and empty pre-BNC
sentinel shapes. It moves admitted body-row envelopes, sparse row headers, and UID mappings using
the persisted order as input, with Rust lexical text comparison and `f64::total_cmp`-based
deterministic ordering for numeric-like values. Formula/error cells, rich text/comments,
merge/filter/group/category/pivot/spill/conditional/hidden/imported/provenance dependencies,
non-empty stroke, cross-tile/cross-bucket movement, unknown mutable state, and other unproven
row-affine state fail closed atomically. A disposable Computer Use run opened a pre-hardening
candidate in Keynote 14.4 and showed the expected order, but the current strict owner rejects the
app-authored source because it contains unproven model field 39. The run is external exploratory
evidence, not current-owner E3/E4 promotion; the checked-in native fixture has no table and the
checked-in evidence test records hashes without launching Keynote. Native acceptance remains
pending, and ADR 0008 records the scope and artifact hashes without broadening it.
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
| Slide-table properties | 🟡 | ✅ | 🟡 | Appearance, dimensions, headers, lock, name, title, and persisted sort configuration; catalog/discovery remains limited |
| Physical table row sorting (“Sort Now”) | 🟡 | 🟡 | 🟡 | Selector-first `Package::{execute_slide_table_sort_order,execute_slide_table_sort_order_to_rows}` over existing tables; no public row/value reader. Admission is limited to the canonical type-6001 table-model route and explicitly proven tile, data-list, header, UID, and empty pre-BNC sentinel shapes. Scalar text/number/boolean/date/duration keys use Rust lexical text ordering, `f64::total_cmp`-based deterministic ordering for numeric-like values, ordinary boolean ordering, and stable source-row ordering for duplicates. Body-relative ranges isolate headers/footers; admitted tile-row envelopes, sparse row headers, and UID mappings move together. Strict wire preflight precedes borrowed lazy Buffa views; formula/error, rich-text, comment, merge, filter/group/category/pivot/spill/conditional, hidden/non-positional, imported/provenance, non-empty stroke, cross-tile/cross-bucket, unknown mutable, and other unproven row-affine dependencies fail closed atomically. Exact-source patch/inverse, candidate reopen/readback, locality, and preview invalidation are source-level contracts. A disposable Computer Use run opened a pre-hardening candidate in Keynote 14.4, but the current strict owner rejects that app-authored source because it contains unproven model field 39; the run is external exploratory evidence rather than current-owner E3/E4 promotion. The checked-in native fixture has no table and the checked-in evidence test records hashes without launching Keynote, so native acceptance remains pending ([owner](../crates/litchi-keynote/src/package/slide_table_physical_sort.rs#L1), [semantic surface](../crates/litchi-keynote/src/slide/table/physical_sort.rs#L1), [tests](../crates/litchi-keynote/tests/slide_table_physical_sort.rs#L1), [ADR 0008](adr/0008-migration-and-verification.md#2026-09-02-amendment-keynote-physical-sort-focused-owner-verification-status)) |
| Table cells/formulas/formats/comments/topology | ❌ | ❌ | ❌ | No public focused cell model or general structural table owner; physical row sorting above is not cell CRUD |
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
| Table settings | 🟡 | ✅ | 🟡 | Appearance, dimensions, headers/freeze/repeat, lock, title, and persisted sort configuration; physical row sorting is tracked as a separate Keynote owner |
| Cell controls/pop-up menus | 🟡 | 🟡 | 🟡 | Checkbox/star/slider/stepper/pop-up only, under strict graph profiles |
| Existing-cell Number/Percentage display formats | 🟡 | ✅ | 🟡 | Selector-first existing-cell Number and Percentage read/set/reset operations preserve the scalar value and unrelated bytes. Number has operation-specific native evidence; Percentage focused/synthetic coverage is green, and ADR 0008 records a disposable native Numbers open/save/close/reopen plus strict Rust semantic no-op readback. The artifact/ledger is not frozen, so Percentage E3/E4 remain pending. See the [Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md#cells-formulas-controls-and-annotations). |
| Existing-cell Currency display formats | 🟡 | ✅ | 🟡 | Selector-first existing-cell Currency read/set/clear/reset operations expose checked currency code, decimal, negative, thousands-separator, and Standard/Accounting settings while preserving the scalar value and unrelated bytes. The focused codec filter passes 4/4, the `litchi-numbers-wire` library passes 32/32, and the latest package target passes 22/22 after native BNC flag handling became shape-dependent: plain no-secondary Currency uses `0x0802`, while a Currency record carrying a secondary Number ID uses `0x0803` (`0x0802 | EXPLICIT_DECIMAL_FORMAT`, with `EXPLICIT_DECIMAL_FORMAT = 0x0001`). Standard versus Accounting does not choose this flag. ADR 0008 records operation-specific E3/E4 evidence: a clean native open, native save/close/reopen with the same scalar/settings/text, strict Rust no-op reread, exact inverse restoration, and exact candidate/native-resaved hashes. Bounded fuzz evidence remains separately scoped; no broad-format or deletion-gate claim is made. See the [Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md#cells-formulas-controls-and-annotations). |
| Existing-cell Scientific display formats | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_scientific_format`](../crates/litchi-numbers/src/package/table_cell_scientific_format.rs#L382) and archive-free [`Scientific`](../crates/litchi-numbers/src/cell/data_format/scientific.rs#L1) cover one existing cell's fixed decimal precision, native minus-sign negatives, and hidden thousands separator with explicit-to-automatic reset semantics. The package target passed 19/19, strict Buffa type-259 codec 4/4, wire library 34/34, and filtered legacy bridge 6/6; bounded codec/package sanitizer smokes each completed 100 executions. Numbers 14.4 opened, saved, closed, and reopened the precision-7 candidate without repair while preserving B2 text and B3 scalar 42, and strict reread was byte-identical. [ADR 0008](adr/0008-migration-and-verification.md#2026-09-01-amendment-numbers-existing-cell-scientific-format-native-validation-record) freezes the hashes and limits this to operation-specific E3/E4 evidence. |
| Existing-cell Fraction display formats | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_fraction_format`](../crates/litchi-numbers/src/package/table_cell_fraction_format.rs#L389) and archive-free [`Fraction`](../crates/litchi-numbers/src/cell/data_format/fraction.rs#L1) expose all nine `FractionAccuracy` denominator strategies for one existing cell, with explicit-to-inherited clearing. The native type-262 codec performs strict preflight before a lazy Buffa view; field 20 (`requires_fraction_replacement`) is preserved when absent or canonical `false` and rejects `true`. The focused package integration run passed 22/22, library codec tests 4/4, direct codec tests 3/3, and wire tests 36/36; checked-in fuzz corpora contain 44 proto seeds and 18 package seeds, and both targets completed 100-run AddressSanitizer smokes. Computer Use recorded operation-specific native E3/E4 evidence for a disposable Eighths source and Eighths→Hundredths edit; before native opening, exactly two uncompressed members differed while entry names/order and every other member payload matched. Exact source/candidate/native-resaved hashes are in ADR 0008. This does not promote all nine native UI variants, arbitrary-producer parity, native byte parity after Numbers normalization, package-wide performance, or a deletion-gate claim. |
| Other generic display formats and rich styles | ❌ | ❌ | ❌ | Outside the focused Number/Percentage/Currency/Scientific/Fraction operations above, no general package getter or formatting transaction is exposed; date, duration, text, custom, and rich-style families remain unsupported at this owner boundary. |
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

“Broad” here is an inventory statement, not a completeness or native-acceptance grade. The host exposes native/raw identities, includes compatibility fallbacks, and is governed by the 11 ordered debts (orders `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`) in the deletion ledger ([host policy](../crates/litchi-iwa/README.md#legacy-and-raw-editing), [ADR 0028](adr/0028-iwa-monolith-exit.md)). The authoritative topology is 64 workspace packages, 238 internal dependency declarations, 227 canonical edges, 11 development-only edges, and one migration host.

## Evidence grades used by the authoritative matrices

| Grade | Required evidence | Current suite posture |
|---|---|---|
| E0 | Type/schema/source exists | Extensive; not support by itself |
| E1 | Synthetic unit/integration test or Litchi self-roundtrip | Extensive for focused transactions and legacy builders |
| E2 | Checked-in Apple-produced fixture parses; exact no-op/readback | Present for one basic fixture per app |
| E3 | Litchi-mutated candidate opens in the native app without repair | External exploratory evidence exists for a pre-hardening Keynote candidate; current-owner promotion remains pending because the app-authored source contains unproven model field 39 |
| E4 | Native app saves, closes, reopens; Litchi strict reread verifies semantics/locality | External/manual exploratory evidence only; the current strict owner does not admit that source, and this is not a suite-wide CI result |

Every supported read/write cell should name its focused API owner, executable test, fixture provenance, evidence grade, native app/version when applicable, and refusal/limit boundary. Exact package preservation, semantic feature support, and native interoperability must remain separate claims.

The existing-cell Percentage row currently has green focused/synthetic evidence and a disposable
native Numbers probe recorded in ADR 0008. The probe includes strict semantic no-op readback of the
native-resaved artifact, but it must remain below formal E3/E4 promotion until that evidence is
attached to a frozen fixture/ledger. The focused Currency row has source-visible owner/codec/tests
and operation-specific native E3/E4 evidence recorded in ADR 0008, including exact
candidate/native-resaved hashes and strict no-op reread. The checked native Number fixture only
proves family-boundary behavior and does not promote Percentage native support; Currency evidence
does not generalize beyond its stated existing-cell operation.
The Scientific row has focused build/provenance, test, bounded sanitizer-smoke, and operation-specific
native E3/E4 evidence. Numbers 14.4 retained the requested precision, text, and scalar through
open/save/close/reopen, and strict Rust reread produced an exact no-op. This promotes only the
recorded existing-cell Scientific operation; it does not establish arbitrary-producer parity,
generic formatting support, or any ADR 0028 deletion-gate closure.
The Fraction row has focused source/build evidence and the following bounded verification record:
22/22 package integration tests, 4/4 library codec tests, 3/3 direct codec tests, and 36/36 wire
tests. Its checked-in fuzz corpora contain 44 proto seeds and 18 package seeds, and both targets
completed 100-run AddressSanitizer smokes. The codec rejects
malformed/noncanonical field-20 encodings and `true`; it preserves an absent field 20 as absent and
canonical encoded `false` byte-for-byte. Computer Use established operation-specific E3/E4
evidence for the representative Eighths-to-Hundredths scenario; before native opening, exactly two
uncompressed members differed and every other member payload/name/order matched. The source,
candidate, and native-resaved hashes are recorded in ADR 0008. No ADR 0028 deletion-gate closure is
implied. This does not promote all nine native UI variants, arbitrary-producer parity, native byte
parity after Numbers normalization, or package-wide performance.

The focused Keynote physical `Sort Now` row has a separate operation-indexed external evidence
record, not a current-owner native E3/E4 certification. Computer Use opened a candidate produced
before strict model-field admission was hardened, but the current strict owner rejects the
app-authored source because model field 39 is unproven. ADR 0008 records the disposable artifact
sizes and hashes as exploratory evidence only; the checked-in native fixture has no table and the
checked-in evidence test does not launch Keynote. Native acceptance remains pending until a
current-admitted native source is available and the operation is rerun.

## Matrix maintenance requirements

1. Keep shared IWA as substrate rather than a fourth user format.
2. Retain evidence-grade/test/fixture links for every ✅ or 🟡 direction and explicit ❌ rows for unsupported families.
3. Record focused-owner and legacy-host status separately during migration.
4. Update the relevant app matrix in the same change that lands or removes a capability.
5. Promote native-interoperability claims only after the current red tests are fixed and operation-specific E3/E4 evidence is recorded.
