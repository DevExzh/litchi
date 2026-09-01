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
semantic support for an unmodeled member. External resources, links, formulas,
comments, and controls remain inert data; nothing is fetched, executed, or
recalculated.

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
| Sheet and table names | 🟡 | ✅ | 🟡 | Existing rooted names can be renamed with collision, NUL, size, dependency, lock, and source checks; no sheet/table lifecycle CRUD. API/tests: [`names`](../src/package/names.rs#L658), [`names.rs`](../tests/names.rs#L677). |
| Table dimensions and row/column sizes | 🟡 | ✅ | 🟡 | Reads and edits admitted existing row/column size payloads through [`table_dimension`](../src/package/table_dimension.rs#L250); this does not insert/delete rows or columns or resize table topology. Coverage: [`table_dimension.rs`](../tests/table_dimension.rs#L1). |
| Header/footer, frozen, and repeating rows/columns | 🟡 | ✅ | 🟡 | [`table_headers`](../src/package/table_headers/api.rs#L13) owns the bounded settings record and exact-source rewrite; rich header content and structural table CRUD are outside the owner. Coverage: [`table_headers.rs`](../tests/table_headers.rs#L1). |
| Table appearance | 🟡 | ✅ | 🟡 | Selector-first [`table_appearance`](../src/package/table_appearance.rs#L355) replaces the effective banding, row-sizing, and gridline value while retaining native style graphs as private implementation details. Coverage: [`table_appearance.rs`](../tests/table_appearance.rs#L1111). |
| Table title visibility and outline | 🟡 | ✅ | 🟡 | [`table_title`](../src/package/table_title.rs#L384) preserves optional-field presence and edits existing titles under style/height/lock prerequisites. Coverage: [`table_title.rs`](../tests/table_title.rs#L48). |
| Table lock state | 🟡 | ✅ | 🟡 | Existing table lock state can be read and changed with exact-source verification. The lock transaction can unlock a locked table; other table-property owners may refuse edits while it remains locked. API: [`table_lock`](../src/package/table_lock.rs#L397); coverage: [`table_lock.rs`](../tests/table_lock.rs). |
| Persisted sort configuration | 🟡 | ✅ | 🟡 | [`table_sort`](../src/package/table_sort.rs#L1005) reads/edits persisted rules only; it deliberately does not reorder physical rows, formulas, comments, or tiles. Coverage: [`table_sort.rs`](../tests/table_sort.rs#L375). |
| General cell-grid, merge, and row/column topology CRUD | ❌ | ❌ | ❌ | The focused package has no general table catalog, merge transaction, or structural row/column/table lifecycle. The bounded scalar/formula operations in the next section do not provide those general capabilities. Public merge/topology vocabularies are detached semantic values; see [`table::merge`](../src/table/merge.rs#L1) and [`table::topology`](../src/table/topology.rs#L1). |

## Cells, formulas, controls, and annotations

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Sparse scalar cells and bounded dense ranges | 🟡 | ✅ | 🟡 | [`table_cells`](../src/package/table_cells.rs#L193) preserves missing versus stored-empty state and supports selector/A1 reads plus atomic set/clear of text, Boolean, finite number, date, and duration values. Native/locality coverage: [`table_cells.rs`](../tests/table_cells.rs#L126). |
| Formula source and cached values | 🟡 | 🟡 | 🟡 | [`formula`](../src/formula.rs#L1) and [`Input::Formula`](../src/table/cells.rs#L89) support bounded expressions, references, ranges, binary operators, and 13 authorable functions (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `AND`, `OR`, `IF`, `IFERROR`, `NOT`, `ABS`, `ROUND`). Expressions cap at 1 MiB/65,536 nodes/depth 64/256 arguments; package edits require admitted ownership/dependency graphs. No general parser, evaluator, dependency engine, external reference resolution, or recalculation is provided. Coverage: [`formula_authoring.rs`](../tests/formula_authoring.rs#L761). |
| Interactive cell controls | 🟡 | 🟡 | 🟡 | Existing checkbox, star-rating, slider, stepper, and pop-up-menu graphs have selector-first read/write paths through [`table_cell_control`](../src/package/table_cell_control.rs#L291) and [`table_cell_pop_up_menu`](../src/package/table_cell_pop_up_menu.rs#L384); control metadata and ownership are strict and bounded. Coverage: [`table_cell_control.rs`](../tests/table_cell_control.rs#L1). |
| Existing-cell Number/Percentage display formats | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_number_format`](../src/package/table_cell_number_format.rs#L384) and [`table_cell_percentage_format`](../src/package/table_cell_percentage_format.rs#L389) read, set, clear, and reset one existing cell's explicit decimal-family format while preserving the scalar value and unrelated bytes. Number retains operation-specific native evidence; Percentage focused/synthetic coverage is green (`table_cell_percentage_format` 19/19 and the codec filter 4/4), and [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-01-amendment-numbers-existing-cell-percentage-format-native-validation-record) records a disposable native Numbers open/save/close/reopen plus strict Rust semantic no-op readback. The artifact/ledger is not frozen, so Percentage E3/E4 remain pending. Coverage: [`table_cell_number_format.rs`](../tests/table_cell_number_format.rs#L1), [`table_cell_percentage_format.rs`](../tests/table_cell_percentage_format.rs#L1), and [`edit_table_cell_percentage_format`](../examples/edit_table_cell_percentage_format.rs#L1). |
| Other generic cell display formats and rich styles | ❌ | ❌ | ❌ | Outside the focused Number/Percentage operations above, the public [`data_format`](../src/cell/data_format.rs#L1) vocabulary remains detached and useful for semantic construction, but no general package getter or formatting transaction is exposed. Currency, date, duration, text, custom, and rich-style families remain unsupported at this owner boundary; control-specific display metadata does not establish generic style support. |
| Cell comments and direct replies | 🟡 | 🟡 | 🟡 | [`comments`](../src/package/comments.rs#L671) and [`comments_reply`](../src/package/comments_reply.rs#L1570) expose text-only cell annotation lifecycle and ordered direct replies for strict rooted/co-located graphs; broad author identity, mentions, attachments, and collaboration are not modeled. Coverage: [`table_cell_comment_reply.rs`](../tests/table_cell_comment_reply.rs#L1). |
| Charts, shapes, text boxes, and media assets | ❌ | ❌ | ❌ | No focused Numbers package owner exposes chart/drawing enumeration or image/movie/audio bytes and CRUD. Detached values or opaque preservation are not semantic support. |
| Filters, categories, groups, pivots, conditional formatting, and validation | ❌ | ❌ | ❌ | These families have no focused public package model or transaction; detection/refusal of an unsupported graph is not feature support. |
| Print/page setup and workbook export | ❌ | ❌ | ❌ | Freeze/repeat settings above are table presentation metadata, not print pagination or page setup. No PDF/XLSX/rendering path is exposed. |
| Source metadata | 🟡 | ✅ | ❌ | [`Document::metadata`](../src/document.rs#L985) reads canonical Numbers sidecars captured during semantic ingress; there is no focused metadata writer. This projection currently has source-level evidence but no dedicated native-fixture metadata assertion. |

## Limits, security, and publication

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Bounded package and semantic resource handling | 🟡 | ✅ | 🟡 | Archive, wire, object, reference, output, and transaction budgets are enforced. Public hard caps include 1,000,000 objects/references, 4,096 sheets, 65,536 tables, 16,000,000 materialized cells, and 64 MiB output text ([`document` limits](../src/document.rs#L15), [`package` limits](../src/package/limits.rs#L10), [representative output-limit test](../tests/table_cell_pop_up_menu.rs#L1831)); formula limits are listed above. Aggregate graph/live-memory and output-overhead budgets remain narrower than full application-scale guarantees. |
| Encrypted or password-protected packages | ❌ | ❌ | ❌ | Encrypted ZIP input is rejected at the shared archive boundary; no password, decryption, or re-encryption API exists ([archive check](../../litchi-iwa-archive/src/package.rs#L1501)). |
| Digital signatures | ❌ | ❌ | ❌ | No Numbers signature verification, trust, invalidation, or re-signing model is exposed. |
| External resources, actions, macros, and database refresh | 🟡 | 🟡 | ❌ | Targets and opaque payloads may remain inert package data, but the crate never resolves, fetches, opens, activates, refreshes, or executes them. See the suite [safety audit](../../../docs/IWORK_PROGRESS_AUDIT.md#safety-and-integrity-findings). |
| Durable save and serialized/composable patches | 🟡 | N/A | 🟡 | [`Package::save`](../src/package.rs#L758) delegates to the archive-owned atomic filesystem publication boundary. Focused patches retain process-local exact source/target artifacts and can be inverted or source-checked, but no durable patch format, merge/history protocol, or portable cross-platform durability claim is provided. |

## Legacy migration-host delta (not focused-owner support)

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Source-free Numbers package builder | 🟡 | N/A | ✅ | `litchi-iwa::numbers::NumbersDocumentBuilder` creates and saves packages, but uses the legacy host/raw-IWA boundary ([host guide](../../litchi-iwa/README.md#create-numbers-spreadsheets-from-scratch)). |
| Sheet/table/cell/formula lifecycle CRUD | 🟡 | ✅ | ✅ | Broad host editors cover capabilities not yet moved into `litchi-numbers`; native identifiers and migration compatibility seams remain host-owned ([legacy policy](../../litchi-iwa/README.md#legacy-and-raw-editing)). |
| Rich formatting, drawings, charts, images, movies, and audio | 🟡 | ✅ | ✅ | Host-only inventory includes these builders and media paths under the [legacy policy](../../litchi-iwa/README.md#legacy-and-raw-editing); this does not certify complete native interoperability or focused-owner support. |
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

The current focused Percentage owner gate is green: its integration target
passes **19/19** tests and the focused codec filter passes **4/4** tests,
including synthetic adversarial, locality, inverse, budget, and family-boundary
coverage. The checked-in Apple-produced Number fixture is used only to verify
selector behavior and refusal of the wrong format family. A disposable native
Numbers Percentage open/save/close/reopen probe plus strict Rust semantic no-op
readback is recorded in [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-01-amendment-numbers-existing-cell-percentage-format-native-validation-record),
but this row makes no formal Percentage E3/E4 claim until the artifact and ledger are frozen.
Broader dirty-worktree failures and the historical
Numbers projection baseline remain tracked in the [current-worktree verification](../../../docs/IWORK_PROGRESS_AUDIT.md#current-worktree-verification)
and must be rerun on a frozen revision before release certification. Four native-
oracle tests remain ignored because their private external fixture is unavailable.

This matrix must be updated with the focused implementation and its executable
test whenever a capability moves out of `litchi-iwa`; host-only, opaque, and
native-app evidence must remain separate claims.
