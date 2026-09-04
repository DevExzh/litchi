# Numbers (`.numbers`) Feature Matrix

This is the authoritative matrix for the public `litchi-numbers` crate and the Numbers companion
to the repository [feature-matrix index](../../../docs/FEATURE_MATRIX.md). It describes the bounded
typed API for Apple Numbers packages (ZIP containers with IWA/protobuf payloads), not complete
Numbers application compatibility, rendering, recalculation, or every undocumented producer
variant. The package format is proprietary; the source and fixtures linked below are the audit
evidence.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the precisely stated scope in `Notes` |
| 🟡 | Bounded, partial, metadata-only, preservation-only, or otherwise limited |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to that direction |

`Read` and `Write` are independent. A package-preservation result is not
semantic support for an unmodeled member. External resources and links remain
inert; focused formula, comment, and control owners provide only the bounded
operations listed below. Nothing is fetched, executed, or recalculated.

## Ownership and input boundaries

`Package` is the exact-source owner for an existing regular ZIP package and
selected source-bound transactions. Its [`write_to`](../src/package.rs#L698)
method streams retained bytes to a caller-owned sink and makes no durability
claim. [`Package::save`](../src/package.rs#L758) delegates filesystem publication
to the archive-owned atomic save boundary; platform-specific durability limits
remain explicit. `Document` is an eager, immutable archive-free projection from
bytes, a path, or a supported app-authored directory; it intentionally drops physical IDs, media, previews,
and unsupported members. [`Document::from_sheets`](../src/document.rs#L722),
[`SheetBuilder`](../src/sheet.rs#L255), and [`TableBuilder`](../src/table.rs#L720)
build semantic values, not `.numbers` packages. Broad source-free authoring in
`litchi-iwa` is recorded separately as legacy-host-only and is not credited to
this matrix.

## Package and semantic workbook

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open and classify an existing Numbers package | 🟡 | ✅ | 🟡 | [`Package::open/from_bytes`](../src/package.rs#L532) validates a rooted Numbers graph and retains an exact regular-ZIP source for no-op output and admitted edits; the package has no fresh writer. Native fixture and foreign-family checks: [`native_fixture.rs`](../tests/native_fixture.rs). |
| Archive-free workbook, sheets, tables, and selectors | ✅ | ✅ | N/A | [`Document`](../src/document.rs#L588) and [`Package`](../src/package.rs#L733) expose rooted sheets/tables by exact name or checked position; duplicate or ambiguous ownership is rejected. Coverage: [`document_table_selector.rs`](../tests/document_table_selector.rs#L13), [`package_sheet_selector.rs`](../tests/package_sheet_selector.rs#L13). |
| Directory-backed semantic read | 🟡 | ✅ | N/A | `Document` captures a bounded immutable snapshot from an app-authored directory on supported Unix paths; Windows directory ingress fails closed. It is not an exact package and cannot be written back. Coverage: [`document_reader.rs`](../tests/document_reader.rs#L441). |
| Exact no-op and untouched-member preservation | ✅ | ✅ | ✅ | Retained source bytes, unsupported ZIP members, and unmodeled fields survive `Package::write_to`; changed focused transactions preserve unrelated members but may remove canonical previews. Coverage: [`table_cells` no-op tests](../tests/table_cells.rs) and [`pop-up-menu` changed opaque-byte test](../tests/table_cell_pop_up_menu.rs#L1762). |
| Fresh focused `.numbers` package creation | ❌ | N/A | ❌ | Semantic builders exist, but no focused package constructor or package writer is exposed; fresh builders remain in the migration host. |
| Global compatibility table projection | 🟡 | ✅ | ❌ | [`extract_structured_tables`](../src/package.rs#L817) is an allocating compatibility view that can include detached legacy table models; it is not the rooted workbook. Coverage: [`compatibility_oracles.rs`](../tests/compatibility_oracles.rs#L237). |
| Per-table CSV projection | 🟡 | ✅ | ❌ | [`Table::to_csv`](../src/table.rs#L626) emits bounded RFC 4180-style text from a semantic table; there is no CSV import or package conversion. Unit coverage: [`table.rs`](../src/table.rs#L1232). |

## Sheets, tables, and presentation settings

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Sheet order | 🟡 | ✅ | 🟡 | Existing rooted sheets can be moved by selector and final position through [`sheet_order`](../src/package/sheet_order.rs#L557); changes are exact-source, reopened, reversible, and preview-aware. Coverage: [`sheet_order.rs`](../tests/sheet_order.rs#L21). |
| Existing table relocation between rooted sheets | 🟡 | ✅ | 🟡 | [`Package::edit_table_relocation`](../src/package/table_relocation.rs#L434), [`Package::move_table`](../src/package/table_relocation.rs#L501), and [`apply_table_relocation`](../src/package/table_relocation.rs#L512) (with `apply_table_move` as an alias) move one existing table selected by a source sheet/table selector to an existing destination sheet (appended at the destination). Same-sheet requests are exact no-ops; changed moves preserve table/model/unknown/unrelated bytes, rewrite only the owning Document/Tables components, remove canonical previews, reopen/read back, and retain an exact inverse. Exact package snapshots are authoritative to this semantic transaction, so focused refusals are terminal; only source-built snapshots may use the crate-private host physical bridge for historical storage outside the semantic projection. Coverage: [`table_relocation.rs`](../tests/table_relocation.rs#L511). |
| Sheet and table names | 🟡 | ✅ | 🟡 | Existing rooted names can be renamed with collision, NUL, size, dependency, lock, and source checks; no sheet/table lifecycle CRUD. API/tests: [`names`](../src/package/names.rs#L658), [`names.rs`](../tests/names.rs#L677). |
| Table dimensions and row/column sizes | 🟡 | ✅ | 🟡 | Reads and edits admitted existing row/column size payloads through [`table_dimension`](../src/package/table_dimension.rs#L250); this does not insert/delete rows or columns or resize table topology. Coverage: [`table_dimension.rs`](../tests/table_dimension.rs#L1). |
| Header/footer, frozen, and repeating rows/columns | 🟡 | ✅ | 🟡 | [`table_headers`](../src/package/table_headers/api.rs#L13) owns the bounded settings record and exact-source rewrite; rich header content and structural table CRUD are outside the owner. Coverage: [`table_headers.rs`](../tests/table_headers.rs#L1). |
| Table appearance | 🟡 | ✅ | 🟡 | Selector-first [`table_appearance`](../src/package/table_appearance.rs#L355) replaces the effective banding, row-sizing, and gridline value while retaining native style graphs as private implementation details. Coverage: [`table_appearance.rs`](../tests/table_appearance.rs#L1111). |
| Table title visibility and outline | 🟡 | ✅ | 🟡 | [`table_title`](../src/package/table_title.rs#L384) preserves optional-field presence and edits existing titles under style/height/lock prerequisites. Coverage: [`table_title.rs`](../tests/table_title.rs#L48). |
| Table lock state | 🟡 | ✅ | 🟡 | Existing table lock state can be read and changed with exact-source verification. The lock transaction can unlock a locked table; other table-property owners may refuse edits while it remains locked. API: [`table_lock`](../src/package/table_lock.rs#L397); coverage: [`table_lock.rs`](../tests/table_lock.rs). |
| Persisted sort configuration | 🟡 | ✅ | 🟡 | [`table_sort`](../src/package/table_sort.rs#L1005) reads/edits persisted rules only; it deliberately does not reorder physical rows, formulas, comments, or tiles. Exact package snapshots use this semantic owner and a focused refusal is terminal. Coverage: [`table_sort.rs`](../tests/table_sort.rs#L375). |
| Historical physical-only persisted-sort compatibility | 🟡 | 🟡 | 🟡 | Only source-built snapshots whose cell-storage projection is outside the semantic reader may use the doc-hidden [`__table_sort_order_*_for_compatibility`](../src/package/table_sort_compat.rs#L37) bridge for field 44. It is host migration compatibility, not public API or a second owner, and has no ADR 0028 debt/deletion-gate impact. |
| General cell-grid, merge, and row/column topology CRUD | ❌ | ❌ | ❌ | The focused package has no general table catalog, merge transaction, or structural row/column/table lifecycle. The bounded scalar/formula operations in the next section do not provide those general capabilities. Public merge/topology vocabularies are detached semantic values; see [`table::merge`](../src/table/merge.rs#L1) and [`table::topology`](../src/table/topology.rs#L1). |

## Cells, formulas, controls, and annotations

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Sparse scalar cells and bounded dense ranges | 🟡 | ✅ | 🟡 | [`table_cells`](../src/package/table_cells.rs#L193) preserves missing versus stored-empty state and supports selector/A1 reads plus atomic set/clear of text, Boolean, finite number, date, and duration values. Native/locality coverage: [`table_cells.rs`](../tests/table_cells.rs#L126). |
| Formula source and cached values | 🟡 | 🟡 | 🟡 | [`formula`](../src/formula.rs#L1) and [`Input::Formula`](../src/table/cells.rs#L89) support bounded expressions, references, ranges, binary operators, and 13 authorable functions (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `AND`, `OR`, `IF`, `IFERROR`, `NOT`, `ABS`, `ROUND`). Expressions cap at 1 MiB/65,536 nodes/depth 64/256 arguments; package edits require admitted ownership/dependency graphs. No general parser, evaluator, dependency engine, external reference resolution, or recalculation is provided. Coverage: [`formula_authoring.rs`](../tests/formula_authoring.rs#L761). |
| Existing-cell Checkbox, Star Rating, Slider, and Stepper controls | 🟡 | 🟡 | 🟡 | The unified [`table_cell_control`](../src/package/table_cell_control.rs#L291) owner exposes selector-first semantic reads and transactions for scalar controls, including admitted split-component graphs. It keeps native IDs and control/list metadata private, stages bounded copy-on-write rewrites, verifies candidate reopen/locality, and fails closed for malformed, aliased, cross-component-unowned, or locked graphs. This row excludes Pop-Up Menu lifecycle; native acceptance remains operation-specific and no generic display-format claim follows. Coverage: [`table_cell_control.rs`](../tests/table_cell_control.rs#L1). |
| Existing-cell Pop-Up Menu control/display format | 🟡 | 🟡 | 🟡 | [`table_cell_pop_up_menu`](../src/package/table_cell_pop_up_menu.rs#L389) and the unified [`table_cell_control`](../src/package/table_cell_control.rs#L291) facade own read/set/clear/reset for admitted rooted same- and split-component graphs. The private owner proves effective locators, member/list/BNC/metadata ownership, refcounts, copy-on-write, reuse/create/reset/final-cull, exact-source patch/inverse, candidate reopen, and physical locality; Wave90 provides operation-specific native semantic persistence evidence. Segmented, ambiguous, aliased, opaque-inbound, cross-component-unowned, and arbitrary producer graphs fail closed, and no general control or table-format authoring is implied. Verification now builds a fallibly reserved native-object index, sorts it once, and uses binary-search routing while charging index work, allocations, and retained memory to the transaction budget; this is an adversarial/bounded-topology safeguard, not a measured throughput or RSS claim. Coverage: [`table_cell_pop_up_menu.rs`](../tests/table_cell_pop_up_menu.rs#L1367), [`table_cell_control.rs`](../tests/table_cell_control.rs#L647), and the private [`NativeObjectIndex`](../src/package/table_cell_pop_up_menu_native.rs#L858). |
| Existing-cell Number display format | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_number_format`](../src/package/table_cell_number_format.rs#L384) owns one existing cell's explicit archive-free `Number` value and read/set/clear/reset/inverse transactions. `None` denotes inherited/automatic; another display or control family is a typed refusal. Strict wire preflight precedes the private lazy Buffa view, and exact source, bounded staging, candidate reopen/readback, scalar/unrelated-byte preservation, locality, and inverse checks are retained. The dedicated migration-host raw-ID Number convenience methods are retired; generic source-built `DataFormat::Number` mutation and attached-table helpers remain host-owned. Number's native E3/E4 operation is recorded in ADR 0008, without broad format or package-authoring claims. Coverage: [`table_cell_number_format.rs`](../tests/table_cell_number_format.rs#L1) and [`edit_table_cell_number_format`](../examples/edit_table_cell_number_format.rs#L1). |
| Existing-cell Percentage display format | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_percentage_format`](../src/package/table_cell_percentage_format.rs#L389) owns one existing cell's explicit type-258 `Percentage` value and read/set/clear/reset/inverse transactions. `None` denotes inherited/automatic; wrong families are typed refusals. The strict wire codec and lazy Buffa path, BNC/list refcount and copy-on-write checks, bounded candidate reopen/readback, exact locality, and atomic refusal behavior are focused-owner guarantees. The dedicated migration-host raw-ID Percentage convenience methods are retired; generic source-built and cross-format `DataFormat::Percentage` mutation remains host-owned. Focused coverage is 19/19 package and 4/4 codec; the disposable native probe is documented in ADR 0008 but formal E3/E4 promotion remains pending its frozen artifact/ledger. Coverage: [`table_cell_percentage_format.rs`](../tests/table_cell_percentage_format.rs#L1) and [`edit_table_cell_percentage_format`](../examples/edit_table_cell_percentage_format.rs#L1). |
| Existing-cell Currency display format | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_currency_format`](../src/package/table_cell_currency_format.rs#L389), [`currency`](../src/cell/data_format/currency.rs#L1), and [`transaction`](../src/cell/data_format/currency.rs#L12) expose checked CurrencyCode, DecimalPlaces, NegativeStyle, ThousandsSeparator, and CurrencyStyle values for one existing cell. `None` denotes inherited/no-explicit Currency while explicit automatic decimals remain `Some(Currency)`. The alternate-number BNC/native type-257 owner accounts for optional secondary Number references, copy-on-write/refcounts, exact-source patches/inverses, unknown-byte/locality preservation, strict wire preflight with a lazy Buffa view, bounded reopen/readback, and typed family refusal. Plain no-secondary Currency uses native flag `0x0802`; a Currency record carrying a secondary Number ID uses `0x0803` (`0x0802 | EXPLICIT_DECIMAL_FORMAT`, with `EXPLICIT_DECIMAL_FORMAT = 0x0001`). Standard versus Accounting does not choose this flag. Coverage: [`table_cell_currency_format.rs`](../tests/table_cell_currency_format.rs#L1) and [`edit_table_cell_currency_format`](../examples/edit_table_cell_currency_format.rs#L1). The focused codec filter passes 4/4, the [`litchi-numbers-wire`](../../litchi-numbers-wire/Cargo.toml) library passes 32/32, and the latest package target passes 22/22 after native BNC flag handling became shape-dependent. ADR 0008 records operation-specific E3/E4 evidence: native clean open, native save/close/reopen with B2 text and B3 scalar/settings preserved, exact inverse restoration, and strict Rust no-op reread; source/candidate/native-resaved hashes are recorded there. The dedicated migration-host raw-ID Currency convenience methods are retired; generic source-built and cross-format `DataFormat::Currency` compatibility remains host-owned. Bounded fuzz evidence remains separately scoped, and no broad-format or deletion-gate claim is made. |
| Existing-cell Scientific display format | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_scientific_format`](../src/package/table_cell_scientific_format.rs#L382) and archive-free [`Scientific`](../src/cell/data_format/scientific.rs#L1) cover one existing cell's fixed decimal precision, native minus-sign negatives, and hidden thousands separator, with explicit-to-automatic reset semantics. The strict type-259 lazy Buffa codec passed 4/4, the focused package suite passed 19/19, `litchi-numbers-wire` passed 34/34, and both codec/package fuzz targets completed 100 sanitizer runs. [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-01-amendment-numbers-existing-cell-scientific-format-native-validation-record) records clean native Numbers open/save/close/reopen, precision-7 readback, exact no-op/inverse hashes, and two-member Rust locality. This remains an existing-cell operation, not generic display-format or rich-style support. The dedicated migration-host raw-ID Scientific convenience methods are retired; generic source-built and cross-format `DataFormat::Scientific` compatibility remains host-owned. |
| Existing-cell Fraction display format | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_fraction_format`](../src/package/table_cell_fraction_format.rs#L389) and archive-free [`Fraction`/`FractionAccuracy`](../src/cell/data_format/fraction.rs#L1) expose all nine denominator strategies for one existing cell, including explicit-to-inherited clearing through the [`transaction`](../src/cell/data_format/fraction.rs#L13) namespace. The type-262 owner is source-bound and uses strict wire preflight before a lazy Buffa view, exact no-op/changed transactions, reversible patches, copy-on-write/refcount closure, candidate reread, and locality checks. The dedicated raw-ID `NumbersEditor` Fraction convenience API is retired; generic source-built and cross-format compatibility remains host-owned. The focused package integration run passed 22/22, library codec tests 4/4, direct codec tests 3/3, and wire tests 36/36; corpora contain 44 proto seeds and 18 package seeds, and both targets completed 100-run AddressSanitizer smokes. Computer Use established operation-specific native E3/E4 evidence for a disposable Eighths source and Eighths→Hundredths edit: Numbers preserved B2 text and B3 Actual `42`; before opening, exactly two uncompressed members differed and all other member payloads/names/order matched, and source/candidate/native-resaved byte sizes and SHA-256 hashes are recorded in [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-02-amendment-numbers-existing-cell-fraction-format-verification-record). Field 20 (`requires_fraction_replacement`) preserves absent/canonical `false` and rejects `true`. This does not promote all nine native UI variants, arbitrary-producer parity, native byte parity after Numbers normalization, package-wide performance, or a deletion-gate claim. |
| Existing-cell Text display format | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_text_format`](../src/package/table_cell_text_format.rs#L382) and archive-free [`Text`](../src/cell/data_format.rs#L89) expose read/set/clear/reset/inverse for one existing cell's admitted explicit Text format. The owner keeps native IDs and format-table records private, uses strict wire preflight plus lazy Buffa views, and bounds copy-on-write/refcount edits, exact-source patches, candidate reopen/readback, scalar/non-format-byte preservation, locality, and typed family refusal. A new explicit Text attachment publishes canonical marker `0x80`; an unchanged converted-Text source with marker `0x81` and retained Number provenance is preserved exactly, but numeric-to-Text conversion is not supported. No native E3/E4 acceptance claim is made. Coverage: [`table_cell_text_format.rs`](../tests/table_cell_text_format.rs#L1) and [`edit_table_cell_text_format`](../examples/edit_table_cell_text_format.rs#L1). |
| Existing-cell Date & Time display format | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_date_time_format`](../src/package/table_cell_date_time_format.rs#L1) and archive-free [`DateTime`](../src/cell/data_format/date_time.rs#L1) own one existing cell's bounded native date/time pattern string. The native type-261 adapter requires the explicit marker `0x0008`, kind `3`, and strict fields 1/14; it admits Empty/type-5 Date and the evidenced type-9 numeric/formula shape, while rejecting plain type-2, marker-zero, wrong/reserved, and ambiguous shapes. The native date/time pattern string is bounded to 4096 bytes; the owner checks its envelope and admitted fields only, and pattern grammar is not validated. Strict preflight precedes the lazy Buffa view, and the metadata-only setter preserves COW/refcounts, unknown/unselected bytes, the scalar value, inverse, and physical locality. Numbers 14.4 (build 7043.0.93, macOS 26.5.2) opened the Litchi candidate without repair/conversion, retained the A1 marker and B2 semantic/inspector values through native save/close/exact-path reopen, and kept the exact DateTime members in `Index/Tables/Tile.iwa` and `Index/Tables/DataList-904498-2.iwa` byte-identical across native normalization. This is operation-specific E3/E4 evidence for one existing type-9 cell/pattern; strict normalized reread reported `changed=false`, `touched_components=0`, `full_reparse=false`, and `scalar_value=untouched`, with a byte-identical 138,725-byte output (SHA-256 `4e489da196bac7416c8c6d827afb2a6264892d4856a5a33dc0a18c6c2d2b1b0c`). No broad DateTime/package/native-byte-parity claim follows. The dedicated raw-ID `NumbersEditor` Date & Time route is retired; generic source-built or cross-format `DataFormat::DateTime`, broad `TextDateTimeField` smart-field lifecycle, and attached Pages/Keynote table compatibility remain host-owned. See [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-existing-cell-date-time-native-verification-record) and the current-boundary amendment in [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-date-time-raw-id-host-route-retirement). |
| Other generic cell display formats and rich styles | ❌ | ❌ | ❌ | Outside the focused Number/Percentage/Currency/Scientific/Fraction/Text/Date & Time/Custom operations above, the public [`data_format`](../src/cell/data_format.rs#L1) vocabulary remains detached and useful for semantic construction, but no general package getter or formatting transaction is exposed. Duration has strict native type-268 codec/projection groundwork and native marker-shape provenance (`0x0004` primary-only; `0x0005` with a generic Number secondary), but remains unsupported at this owner boundary; rich-style families remain unsupported as well. The Custom owner is limited to the document-scoped existing-cell registry transaction above and does not establish generic style or package authoring. |
| Root cell comments | 🟡 | 🟡 | 🟡 | [`comments`](../src/package/comments.rs#L671) exposes only text and semantic selectors; native IDs, list keys, UUIDs, author records, protobuf values, and archive names stay private. Text reads resolve rooted cell ownership; root create/replace/clear writes are bounded to strict existing tile/row/table ownership. The current owner may materialize a missing comment list, author storage, or sparse BNC cell slot, but never a missing tile/row. Existing replacement/clear requires globally unique, unshared root storage; changed clears and rewrites require exact PackageMetadata/save-token attribution and a reply-free root. Segmented, aliased, cross-component, opaque, ambiguous, or otherwise unproven write graphs fail closed. Author identity, mentions, attachments, collaboration, revisions, and native UI acceptance are not modeled. Coverage: [`table_cell_comment_reply.rs`](../tests/table_cell_comment_reply.rs#L983). |
| Direct replies to root cell comments | 🟡 | ✅ | 🟡 | [`comments_reply`](../src/package/comments_reply.rs#L1570) exposes source-ordered text replies through the semantic [`CommentReplyIndex`](../src/package/comments_reply.rs#L40), collection edits, append/add, ordinal set, remove, A1 wrappers, exact patches, inverse, reopen, and locality checks. Reads perform package-wide ownership/canonical checks; writes are limited to an existing-root, same-member direct-leaf list. Shared roots use copy-on-write, while shared reply leaves, nested/cyclic/segmented lists, cross-member or cross-component ownership, unknown opaque inbound owners, and arbitrary producer graphs fail closed. Reply author identity and collaboration metadata remain private, and native direct-reply acceptance is explicitly absent from ADR 0008. Coverage: [`table_cell_comment_reply.rs`](../tests/table_cell_comment_reply.rs#L1764). |
| Charts, shapes, text boxes, and media assets | ❌ | ❌ | ❌ | No focused Numbers `Package` owner exposes rooted drawable enumeration, geometry/z-order/grouping, chart model/data, text-box content, image/movie/audio bytes, asset insertion/replacement/deletion, or media garbage collection. `Document` intentionally drops these members; detached semantic values and opaque byte preservation are not feature support. The broad legacy host inventory remains migration compatibility only and cannot be credited here ([legacy policy](../../litchi-iwa/README.md#legacy-editing)). |
| Filters, categories, groups, pivots, conditional formatting, and validation | ❌ | ❌ | ❌ | These families have no focused public package model or transaction; detection/refusal of an unsupported graph is not feature support. |
| Print/page setup and workbook export | ❌ | ❌ | ❌ | Freeze/repeat settings above are table presentation metadata, not print pagination or page setup. No PDF/XLSX/rendering path is exposed. |
| Source metadata | 🟡 | ✅ | ❌ | [`Document::metadata`](../src/document.rs#L985) reads canonical Numbers sidecars captured during semantic ingress; there is no focused metadata writer. This projection currently has source-level evidence but no dedicated native-fixture metadata assertion. |
| Existing-cell Custom display format | 🟡 | ✅ | 🟡 | Selector-first [`litchi_numbers::Package::{table_cell_custom_format, edit_table_cell_custom_format, apply_table_cell_custom_format}`](../src/package/table_cell_custom_format.rs#L319) resolves a `SheetSelector`, sheet-scoped `TableSelector`, and checked `CellPosition` for one existing rooted cell and exposes the archive-free [`Custom`](../src/cell/data_format/custom.rs#L594) value plus exact-source read/set/clear/reset/inverse transactions. The private document registry is rooted by `TN.DocumentArchive` field 9 and message type 222; custom archive discriminators are 270 (Number), 271 (Text), and 272 (Date & Time). Handwritten wire preflight precedes the private lazy Buffa view; deterministic source-built exact-source fixtures provide E1 evidence only. The transaction preserves unknown/unselected fields and members, validates format-list/refcount closure, keeps UUIDs private, reuses semantically equal registry entries, retains shared references, allocates replacements, and culls only unused entries; candidate reopen/readback, exact inverse restoration, physical locality, content-redacted diagnostics, and typed budget/refusal paths are part of the contract. The focused package suite passes 19/19, the strict custom-format codec passes 8/8, and both fuzz targets complete 100-run AddressSanitizer smokes. No Apple-authored fixture, native Numbers acceptance, native save/resave, or E2/E3/E4 evidence is claimed; no generic Custom-format authoring is implied, and no legacy `NumbersEditor` Custom route is retired. Host Custom compatibility remains migration-host-only. |

## Limits, security, and publication

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Bounded package and semantic resource handling | 🟡 | ✅ | 🟡 | Archive, wire, object, reference, output, and transaction budgets are enforced. Public hard caps include 1,000,000 objects/references, 4,096 sheets, 65,536 tables, 16,000,000 materialized cells, and 64 MiB output text ([`document` limits](../src/document.rs#L15), [`package` limits](../src/package/limits.rs#L10), [representative output-limit test](../tests/table_cell_pop_up_menu.rs#L1831)); formula limits are listed above. Aggregate graph/live-memory and output-overhead budgets remain narrower than full application-scale guarantees. |
| Encrypted or password-protected packages | ❌ | ❌ | ❌ | Encrypted ZIP input is rejected at the shared archive boundary; no password, decryption, or re-encryption API exists ([archive check](../../litchi-iwa-archive/src/package.rs#L1501)). |
| Digital signatures | ❌ | ❌ | ❌ | No Numbers signature verification, trust, invalidation, or re-signing model is exposed. |
| External resources, actions, macros, and database refresh | 🟡 | 🟡 | ❌ | Targets and opaque payloads may remain inert package data, but the crate never resolves, fetches, opens, activates, refreshes, or executes them. See the suite [safety audit](../../../docs/IWORK_PROGRESS_AUDIT.md#safety-and-integrity-findings). |
| Durable save and serialized/composable patches | 🟡 | N/A | 🟡 | [`Package::save`](../src/package.rs#L758) delegates to the archive-owned atomic filesystem publication boundary. Focused patches retain process-local exact source/target artifacts and can be inverted or source-checked, but no durable patch format, merge/history protocol, or portable cross-platform durability claim is provided. |

## Legacy migration-host delta (not focused-owner support)

An exact `Package` owner is authoritative for an exact package: a structural,
family, lock, budget, stale-source, or locality refusal is terminal and is not
silently retried through a generic raw-ID writer. Explicit compatibility is
limited to the legacy cases still named by the ADRs: source-free packages,
historical physical-only table relocation and persisted-sort bridges for
source-built snapshots whose semantic projection cannot admit the graph,
generic source-built/cross-format `DataFormat` mutation and attached-table
helpers retained for Pages/Keynote, and the remaining identity-bearing comment
compatibility paths. These are migration-host fallbacks, not focused-owner
support; exact snapshots never fall back to them.

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Source-free Numbers package builder | 🟡 | N/A | ✅ | `litchi-iwa::numbers::NumbersDocumentBuilder` creates and saves packages from typed/raw-IWA host objects; it is the explicit source-free fallback because focused `Package` has no fresh-package constructor or writer ([host guide](../../litchi-iwa/README.md#create-numbers-spreadsheets-from-scratch)). |
| Sheet/table/cell/formula lifecycle CRUD | 🟡 | ✅ | ✅ | Broad host editors retain source-free and legacy raw-ID lifecycle work not yet moved into `litchi-numbers`. Generic source-built/cross-format `DataFormat` and attached-table helpers remain host-owned. Dedicated NumbersEditor raw-ID Number, Percentage, Currency, Scientific, Fraction, and Date & Time convenience methods are retired; there is no dedicated host Text route and no dedicated host Custom retirement. Generic source-built/cross-format `DataFormat::DateTime` and `DataFormat::Custom`, broad `TextDateTimeField` smart-field lifecycle, and attached Pages/Keynote table compatibility remain host-owned. A focused exact refusal is not a fallback trigger ([legacy policy](../../litchi-iwa/README.md#legacy-editing)). |
| Historical physical-only table relocation compatibility | 🟡 | 🟡 | 🟡 | There is no public `NumbersEditor::move_table`. Exact snapshots stay on [`litchi_numbers::Package::move_table`](../src/package/table_relocation.rs#L501), where semantic refusal is terminal. Only source-built snapshots whose storage is outside the semantic projection use the doc-hidden [`__move_table_from_bytes_for_compatibility`](../src/package/table_relocation_compat.rs#L37) physical bridge; it performs legacy candidate readback, is not a second implementation, and has no ADR 0028 debt/deletion-gate impact. |
| Rich formatting, drawings, charts, images, movies, and audio | 🟡 | ✅ | ✅ | Host-only inventory includes these builders and media paths under the [legacy policy](../../litchi-iwa/README.md#legacy-editing); this does not certify complete native interoperability or focused-owner support. |
| Identity-bearing comment/reply compatibility | 🟡 | ✅ | 🟡 | Remaining `litchi-iwa` comment/reply adapters cover legacy/native-identity cases that the focused text-only owner rejects; root-comment clear has been retired from this surface, while compatibility read/replacement and reply fallback remain host-only. |
| Focused-owner replacement complete | ❌ | N/A | ❌ | The host remains on the ordered migration ledger; host breadth must not be promoted to this crate's matrix. |

## Evidence and current verification

Evidence grades used by the suite are E0 (typed source), E1 (synthetic or
self-roundtrip), E2 (checked-in Apple-produced fixture parse/no-op), E3 (native
app opens a Litchi-mutated candidate), and E4 (native save/close/reopen with a
strict reread). Unless a row says otherwise, linked focused transaction tests
are E1; the native-fixture row is E2 only for parse/read/exact-no-op behavior.
E3/E4 are operation-specific external evidence, not a suite-wide gate. The checked-in
[Numbers fixture](../../../test-data/iwork/numbers/basic.numbers) is intentionally
basic, and the native provenance is documented in
[`test-data/iwork/README.md`](../../../test-data/iwork/README.md).

The focused table-relocation owner covers selector admission, exact same-sheet
no-op, changed destination ordering, table/model/unknown/unrelated-byte
preservation, exact-source conflict handling, inverse restoration, candidate
reopen, and physical locality in the [`table_relocation` integration target](../tests/table_relocation.rs#L539). The ADR 0008 relocation record reports six
focused tests, a 2,000-run relocation fuzz smoke, the crate-private host
delegation regression, and operation-specific native Numbers open/save/close/
reopen evidence. This does not certify complete row-affine or table-lifecycle
authoring, arbitrary-producer parity, benchmark/RSS behavior, or host exit.
The Pop-Up Menu verifier similarly replaces repeated ownership scans with a
fallibly reserved index sorted once and binary-searched thereafter; its budget
and adversarial-index checks are bounded safety evidence, not a throughput/RSS
measurement.

The unified scalar-control owner and split Pop-Up Menu route retain their
focused integration coverage (27/27 and 19/19, respectively), including
copy-on-write, refcount, metadata, limit, locality, inverse, and reopen
checks. Wave90 provides native semantic persistence for the admitted Pop-Up
Menu graph; it does not promote arbitrary control graphs or general display
format support. Root comment creation/replacement/clear and direct-reply
append/set/remove have focused source-level lifecycle tests, but ADR 0008
explicitly withholds native direct-reply acceptance. No focused Numbers owner
currently provides rooted drawable or media enumeration/CRUD; the host's
source-built drawing/media inventory remains migration compatibility only.

Number, Percentage, Currency, Scientific, Fraction, Text, and Date & Time are separate public
focused owners even where they share private display-format plumbing. Document-scoped Custom is
an additional focused owner for existing-cell registry entries in its Number, Text, and Date &
Time variants; it does not establish generic display-format ownership. Each
owner accepts only its own explicit family on one existing cell and preserves
the same selector-first, lazy-Buffa, exact-source, bounded, reopen, locality,
inverse, and typed-family-refusal discipline; shared implementation does not
mean generic cross-family package ownership.

The Text owner provides the same bounded package discipline for an existing
cell's explicit Text format. Its archive-free marker has no configurable
fields; a new explicit attachment publishes canonical Text (`0x80`), while an
unchanged admitted converted-Text source (`0x81`) remains byte-exact. The owner
does not convert numeric values to Text, and no native E3/E4 acceptance claim
is made. Coverage is recorded in
[`table_cell_text_format.rs`](../tests/table_cell_text_format.rs#L1) and the
[`edit_table_cell_text_format`](../examples/edit_table_cell_text_format.rs#L1)
example.

The focused Date & Time owner provides the same package discipline for one
existing cell's explicit Date & Time display metadata. Its native type-261
adapter requires explicit marker `0x0008`, kind `3`, and strict fields 1/14;
admitted shapes are Empty/type-5 Date and the evidenced type-9 numeric/formula
shape, while plain type-2, marker-zero, wrong/reserved, and ambiguous shapes
are refused. The metadata-only setter carries a bounded native date/time
pattern string (4096-byte maximum) and performs strict preflight before a lazy
Buffa view, retaining unknown and unselected bytes, scalar value, exact
COW/refcounts, inverse, and locality. The owner checks the bounded string
envelope and admitted fields only; pattern grammar is not validated.
The dedicated raw-ID `NumbersEditor` Date & Time route is retired; generic
source-built or cross-format `DataFormat::DateTime`, broad
`TextDateTimeField` smart-field lifecycle, and attached Pages/Keynote table
compatibility remain host-owned. Computer Use recorded
operation-specific native app-cycle evidence for one existing type-9 cell (the
full source/candidate/native-resaved ledger is in [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-existing-cell-date-time-native-verification-record));
strict normalized reread of the native-resaved artifact reported `changed=false`,
zero touched components, `full_reparse=false`, and `scalar_value=untouched`,
with a byte-identical 138,725-byte output (SHA-256
`4e489da196bac7416c8c6d827afb2a6264892d4856a5a33dc0a18c6c2d2b1b0c`). This
does not make a broad DateTime, package, or native-byte-parity claim.

The focused Number owner has 14/14 package integration tests, 2/2 fixture
tests, and a 1/1 host delegation check. ADR 0008 records its operation-specific
native candidate open/save/close/reopen evidence (B3 rendered `42.000`) and
strict native-resaved no-op reread; this remains one existing-cell Number
scenario rather than generic display-format or package-authoring evidence.

The current focused Percentage owner gate is green: its integration target
passes **19/19** tests and the focused codec filter passes **4/4** tests,
including synthetic adversarial, locality, inverse, budget, and family-boundary
coverage. The checked-in Apple-produced Number fixture is used only to verify
selector behavior and refusal of the wrong format family. A disposable native
Numbers Percentage open/save/close/reopen probe plus strict Rust semantic no-op
readback is recorded in [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-01-amendment-numbers-existing-cell-percentage-format-native-validation-record),
but this row makes no formal Percentage E3/E4 claim until the artifact and ledger are frozen.
The focused Currency owner has source-visible integration/codec/fuzz coverage and
operation-specific native E3/E4 evidence recorded in ADR 0008. The codec filter is green (4/4),
the [`litchi-numbers-wire`](../../litchi-numbers-wire/Cargo.toml) library is green (32/32), and
the latest package run is green (22/22) after native BNC flag handling became shape-dependent.
Plain no-secondary Currency uses native flag `0x0802`; a Currency record carrying a secondary
Number ID uses `0x0803` (`0x0802 | EXPLICIT_DECIMAL_FORMAT`, with
`EXPLICIT_DECIMAL_FORMAT = 0x0001`). Standard versus Accounting does not choose this flag. The
checked-in Number fixture remains only a
wrong-family/refusal baseline for Currency; native Currency evidence does not generalize to other
display families or complete Numbers authoring.
The Scientific owner has focused build/provenance, 19/19 package, 4/4 codec,
34/34 wire, 6/6 legacy-host, two 100-run sanitizer smoke, and native Numbers
open/save/close/reopen evidence. ADR 0008 records the source, Rust candidate,
inverse, native-resaved, and exact no-op hashes. This evidence remains scoped
to one existing cell's fixed-precision Scientific format.
The Fraction owner has focused build/provenance, 22/22 package integration,
4/4 library codec, 3/3 direct codec, 36/36 wire, and two 100-run
AddressSanitizer smokes from 44 proto seeds and 18 package seeds. Computer Use also established
operation-specific native E3/E4 evidence for the representative
Eighths-to-Hundredths edit. Before native opening, exactly two uncompressed
members differed (`Index/Tables/Tile.iwa`, 228 bytes, and
`Index/Tables/DataList-904498-2.iwa`, 53 bytes); entry names/order and every
other member payload were identical. ADR 0008 records the exact
source/candidate/native-resaved byte sizes and SHA-256 hashes. This is not
evidence for all nine native UI variants, arbitrary-producer parity, native byte
parity after Numbers normalization, or package-wide performance.
Broader dirty-worktree failures and the historical
Numbers projection baseline remain tracked in the [current-worktree verification](../../../docs/IWORK_PROGRESS_AUDIT.md#current-worktree-verification)
and must be rerun on a frozen revision before release certification. Four native-
oracle tests remain ignored because their private external fixture is unavailable.

This matrix must be updated with the focused implementation and its executable
test whenever a capability moves out of `litchi-iwa`; host-only, opaque, and
native-app evidence must remain separate claims.
