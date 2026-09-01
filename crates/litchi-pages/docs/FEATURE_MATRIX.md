# Pages (`.pages`) Feature Matrix

This document is the authoritative matrix for the public `litchi-pages` crate. It covers
Apple Pages packages built from the iWork ZIP/directory container and IWA component archives.
It describes the typed API exposed by this crate, not complete Pages application compatibility,
rendering fidelity, or every revision of Apple's private protobuf schema.

`Package` is the exact-source adapter: it accepts regular package files or ZIP bytes, retains the
physical artifact, and exposes selected source-bound transactions. `Document` is an immutable,
archive-free semantic snapshot: it can read a ZIP or an app-authored package directory, but it
does not retain exact bytes, media, previews, or unsupported members. A successful package write
is a stream to a caller-owned sink; callers must provide their own durable or atomic filesystem
publication policy.

The broad Pages builder under `litchi-iwa` is a migration host, not evidence for a focused
`litchi-pages` row. Its capabilities are listed separately below.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | The documented feature scope has a public typed implementation |
| 🟡 | Support is bounded, partial, metadata-only, preservation-only, or otherwise constrained |
| ❌ | No public typed support is currently available |
| N/A | The concept does not apply to the feature or direction |

`Read` and `Write` are independent. A `🟡` direction may mean a supported subset, an exact
preservation path, or a source-bound transaction. `E1` in a note means synthetic or focused
self-roundtrip tests; `E2` means a checked-in Apple-produced fixture has open/no-op/readback
coverage. Neither grade proves that a changed output is accepted by the native Pages app.

## Package and semantic reader

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Pages package ingress and application detection | 🟡 | ✅ | 🟡 | [`Package::open`/`from_bytes`](../src/package.rs) parse regular files and ZIP bytes through the bounded IWA archive adapter. [`Document`](../src/document.rs) also captures app-authored directories on supported Unix hosts. Wrong iWork applications, unsafe paths, malformed archives, and ambiguous sources are rejected. Evidence: [`document_reader` tests](../tests/document_reader.rs), [`native fixture`](../tests/native_fixture.rs) (E1/E2). |
| Exact package no-op and source preservation | 🟡 | ✅ | ✅ | `Package::write_to` streams the retained artifact unchanged, including unsupported ZIP members and unmodeled protobuf fields. Selected changed transactions retain unselected data but do not provide a general package editor. Evidence: [`Package` implementation](../src/package.rs), [`native fixture`](../tests/native_fixture.rs), and transaction tests below (E1/E2). |
| Archive-free semantic `Document` | ✅ | ✅ | N/A | ZIP, bytes, shared bytes, or directory input projects to immutable ordered [`Section`](../src/section.rs) values with text, headings, paragraphs, storages, names, and known page counts. Exact package bytes, native IDs, media, previews, and unsupported members are intentionally dropped. Evidence: [`Document` implementation](../src/document.rs), [`document_reader` tests](../tests/document_reader.rs) (E1/E2). |
| Section selection and plain-text projection | ✅ | ✅ | N/A | Exact name and checked position selectors are public; duplicate names report an error. Sections and the complete document can be rendered as plain text without exposing native identities. Evidence: [`selectors`](../src/selector.rs), [`section` model](../src/section.rs), [`document_reader` tests](../tests/document_reader.rs) (E1). |
| Package metadata sidecars | 🟡 | ✅ | ❌ | `Package::metadata` reads supported fields from `Metadata/Properties.plist`, `BuildVersionHistory.plist`, and `DocumentIdentifier` into `litchi-core::Metadata`. No public metadata write transaction exists. Evidence: [`metadata projection`](../src/package.rs), [`document settings tests`](../tests/document_settings.rs) (E1). |
| Validation and bounded resource accounting | 🟡 | ✅ | N/A | Physical archive/IWA limits are caller-selectable through [`Limits`](../src/package.rs); semantic defaults cap text at 64 MiB, sections and body storages at 4,096, and package object inspection at 1,000,000. Aggregate graph/live-memory accounting is not a complete global budget. `DocumentReadOptions` adds checked source and semantic limits. Evidence: [`document limits`](../src/document.rs), [`archive limits`](../../litchi-iwa-archive/src/lib.rs), [`document_reader` limit tests](../tests/document_reader.rs) (E1). |
| Unknown fields, opaque members, and preservation-only payloads | 🟡 | ✅ | 🟡 | Unselected physical members and unmodeled IWA/protobuf bytes remain available through exact-source output and selected rewrites. They have no semantic API; unsafe, ambiguous, dependent, or unsupported ownership graphs fail closed. Evidence: [`section-text tests`](../tests/section_text.rs), [`body-table name tests`](../tests/body_table_name.rs), [`header/footer tests`](../tests/header_footer_text.rs) (E1). |

## Existing-package text and document formatting

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Body and section text | 🟡 | ✅ | 🟡 | `Package` exposes selector-first text reads and one checked UTF-16 splice per transaction, with insert, replace, delete, set, and clear helpers. Surrogate splits, section breaks, footnote anchors, inline-object markers, dependent content, multiple staged operations, and non-exact sources are refused. Commits reopen and verify the candidate and return an exact-source inverse. Evidence: [`section text owner`](../src/package/section_text.rs), [`section text tests`](../tests/section_text.rs) (E1; native fixture read is E2). |
| Section settings | 🟡 | ✅ | 🟡 | The aggregate semantic settings value preserves optional name, four header/footer flags, and three pagination fields. Existing sections can be selected by name or checked position and edited through an exact-source transaction; structural section creation, deletion, and reordering are not exposed. Evidence: [`section settings owner`](../src/package/section_settings.rs), [`section settings tests`](../tests/section_settings.rs) (E1). |
| Section name and pagination convenience transactions | 🟡 | ✅ | 🟡 | Dedicated transactions expose section-name replacement and section-start, page-numbering, and starting-page-number edits while preserving field presence and unknown discriminants. They remain existing-section, exact-source operations. Evidence: [`section name owner`](../src/package/section_name.rs), [`section name tests`](../tests/section_name.rs), [`pagination owner`](../src/package/section_pagination.rs), [`pagination tests`](../tests/section_pagination.rs) (E1). |
| Document and footnote formatter settings | 🟡 | ✅ | 🟡 | The combined settings transaction reads and replaces supported document options and footnote/endnote kind, marker format, numbering, and gap. It validates canonical enum values, preserves optional presence, checks dependencies, and verifies/reopens changed candidates. It is not a general document-style editor. Evidence: [`settings value`](../src/document_settings.rs), [`package transaction`](../src/package/document_settings.rs), [`document settings tests`](../tests/document_settings.rs) (E1). |
| Section backgrounds | 🟡 | ✅ | 🟡 | Existing sections support `None` and validated solid RGBA fills. Gradients, image fills, media references, and other native fill variants are observable only as unsupported and cannot be authored. Evidence: [`background value and transaction`](../src/section/background.rs), [`background package owner`](../src/package/section_background.rs), [`background tests`](../tests/section_background.rs) (E1). |
| Page layout | 🟡 | ✅ | 🟡 | Existing package layouts expose presence-preserving page width/height, margins, scale, portrait/landscape orientation, and vertical-body settings. Finite scalar and margin invariants are checked; unsupported dependencies and noncanonical variants refuse changed writes. Evidence: [`layout value`](../src/page_layout.rs), [`layout transaction`](../src/package/page_layout.rs), [`layout tests`](../tests/page_layout.rs) (E1). |
| Headers and footers | 🟡 | ✅ | 🟡 | Rooted logical header/footer slots can be enumerated and selected by section, template role, kind, and slot. Existing slot text can be set or cleared with exact-source, alias, dependency, and candidate-reopen checks. Slot creation, template CRUD, rich formatting, and general header/footer layout are not exposed. Evidence: [`header/footer model`](../src/header_footer.rs), [`text transaction`](../src/package/header_footer_text.rs), [`header/footer tests`](../tests/header_footer_text.rs) (E1). |
| Body footnotes | 🟡 | ✅ | 🟡 | Rooted body footnotes expose checked positions, semantic text, and optional custom marks. Existing notes can be edited; bounded insert and remove operations update the owned graph and return reversible exact-source patches. Structural markers, shared or ambiguous ownership, and unsafe references are rejected. Evidence: [`footnote model`](../src/footnote.rs), [`body footnote owner`](../src/package/body_footnote.rs), [`footnote text owner`](../src/package/footnote_text.rs), [`body footnote tests`](../tests/body_footnote.rs), [`footnote text tests`](../tests/footnote_text.rs) (E1). |
| Endnote text lifecycle | ❌ | ❌ | ❌ | Formatter settings can select endnote kinds and numbering, but the focused crate has no rooted endnote-text read, insert, edit, or remove API. Evidence: [`footnote kind model`](../src/footnote.rs). |

## Body tables

Table support is intentionally a presentation-metadata surface. The focused crate does not
publish a general table object, cell grid, formula model, or structural table editor. All rows
below operate on an existing rooted body table selected by checked position or exact visible
name; ambiguous, locked, dependent, or unsupported table graphs refuse changed edits.

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Body-table selection and metadata discovery | 🟡 | 🟡 | 🟡 | Selector-driven resolution proves the private body, drawable, table-info, and model ownership chain. There is no public table catalog or general semantic table enumeration API. Evidence: [`body-table selectors`](../src/selector.rs), [`table ownership proof`](../src/package/table_lock.rs), [`table-lock tests`](../tests/table_lock.rs) (E1). |
| Table appearance | 🟡 | ✅ | 🟡 | Existing tables expose validated style, gridline, banding, and row-sizing settings; exact-source edits support the admitted appearance profile and refuse unsupported style graphs. Evidence: [`appearance value`](../src/table/appearance.rs), [`appearance transaction`](../src/package/body_table_appearance.rs), [`appearance tests`](../tests/body_table_appearance.rs) (E1). |
| Table dimensions | 🟡 | ✅ | 🟡 | Existing table width/height and row/column dimension values are read and edited with checked finite bounds. This changes presentation metadata, not the table's row/column topology. Evidence: [`dimension value`](../src/table/dimension.rs), [`dimension transaction`](../src/package/body_table_dimension.rs), [`dimension tests`](../tests/body_table_dimensions.rs) (E1). |
| Header, footer, freeze, and repeat settings | 🟡 | ✅ | 🟡 | Existing body-table header/footer counts and freeze/repeat flags are presence-preserving and editable within the validated native graph. Physical header-row data and table topology are not exposed. Evidence: [`header settings`](../src/table/headers.rs), [`header transaction`](../src/package/body_table_headers.rs), [`header tests`](../tests/body_table_header_settings.rs) (E1). |
| Table name | 🟡 | ✅ | 🟡 | Existing rooted table names can be read and renamed through selector-first, source-bound patches. Duplicate names, invalid names, locked tables, and foreign/ambiguous ownership fail closed. Evidence: [`name value`](../src/table/name.rs), [`name transaction`](../src/package/body_table_name.rs), [`name tests`](../tests/body_table_name.rs) (E1). |
| Persisted table sort configuration | 🟡 | ✅ | 🟡 | Existing sort rules, scope, direction, and column references can be read and replaced. This edits persisted configuration only; it does not sort or recalculate table cell data. Evidence: [`sort value`](../src/table/sort.rs), [`sort transaction`](../src/package/body_table_sort.rs), [`sort tests`](../tests/body_table_sort.rs) (E1). |
| Table title visibility and outline | 🟡 | ✅ | 🟡 | Existing table title visibility and outline settings are read and edited with presence-preserving values and exact-source verification. Evidence: [`title value`](../src/table/title.rs), [`title transaction`](../src/package/body_table_title.rs), [`title tests`](../tests/body_table_title.rs) (E1). |
| Table interactive lock state | 🟡 | ✅ | 🟡 | Existing body tables expose locked/unlocked state and support source-bound lock edits. Lock state is not authorization or encryption; changed edits are refused when the ownership proof is not unique. Evidence: [`lock value`](../src/table/lock.rs), [`lock transaction`](../src/package/table_lock.rs), [`lock tests`](../tests/table_lock.rs) (E1). |
| Table cells, values, formulas, formatting, and merges | ❌ | ❌ | ❌ | No focused public API enumerates or edits Pages table cells, formulas, rich cell formatting, merged ranges, or calculated results. The legacy host's table editor does not upgrade this row. |
| Table row/column and table structural CRUD | ❌ | ❌ | ❌ | No focused package API creates, deletes, duplicates, resizes structurally, or reorders Pages tables, rows, or columns. |

## Rich content and application features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Rich-text runs, styles, lists, links, and annotations | 🟡 | 🟡 | ❌ | Text storages and plain text/range values are readable through the semantic model. The focused package has no public style, list, hyperlink, revision, annotation, or rich-run mutation API. Shared text codecs are not a Pages feature claim. Evidence: [`section text model`](../src/section.rs), [`text storage API`](../src/lib.rs), [`section text tests`](../tests/section_text.rs) (E1). |
| Images, audio, movies, and media assets | ❌ | ❌ | ❌ | The public [`image`](../src/image.rs), [`audio`](../src/audio.rs), and [`movie`](../src/movie.rs) modules contain validated detached option values for migration-host builders; they are not package asset readers or resource CRUD APIs. |
| Shapes, drawings, text boxes, and z-order | ❌ | ❌ | ❌ | No focused Pages package model or transaction owns drawable creation, geometry, grouping, layering, or text-box content. |
| Charts and chart data | ❌ | ❌ | ❌ | No focused Pages chart catalog, data-series model, chart creation, or chart property transaction is public. |
| Collaboration, comments, revisions, and coauthoring | ❌ | ❌ | ❌ | No focused semantic or package lifecycle API exists for collaboration services, comments, revision resolution, locks beyond table presentation state, or coauthoring. |
| Rendering and export | ❌ | ❌ | ❌ | The crate does not render Pages, produce PDF/HTML/RTF, paginate layout, or automate the Pages application. Plain-text projection is not rendering. |
| Fresh package creation and structural document authoring | ❌ | N/A | ❌ | Detached `Document`, `Root`, `Body`, and `Section::Builder` values can form an archive-free semantic snapshot, but `litchi-pages` has no fresh `.pages` package writer or section/asset/shape structural authoring API. |

## Security, durability, and interoperability boundaries

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Password-protected or encrypted Pages packages | ❌ | ❌ | ❌ | Shared iWork detection and archive ingress reject encrypted/password-protected sources; no password, decryption, or re-encryption API exists. Evidence: [`detector error and limits`](../../litchi-iwa-detect/src/lib.rs), [`archive error`](../../litchi-iwa-archive/src/error.rs). |
| External resources, actions, macros, controls, and embedded code | 🟡 | 🟡 | 🟡 | Unsupported payloads may remain inert and opaque in an exact package artifact, but no external target is resolved or fetched and no action, macro, control, or embedded code is activated or executed. Changed transactions refuse unsupported dependency graphs. Evidence: [`package preservation boundary`](../src/package.rs), [`legacy host policy`](../../litchi-iwa/README.md). |
| Cryptographic signatures and trust | ❌ | ❌ | ❌ | Exact no-op output can preserve bytes that happen to be signed; the focused crate has no signature verification, certificate trust, invalidation, or re-signing model. |
| Durable patch serialization, composition, merge, and history | ❌ | N/A | ❌ | Focused patches retain process-local source/target artifacts or logical deltas; there is no stable patch format, composition/merge protocol, or persistent history API. |
| Durable atomic filesystem publication | ❌ | N/A | ❌ | Focused commits return a verified in-memory `Package`; `write_to` only streams bytes and does not flush, sync, rename, or atomically replace a path. The caller owns durable publication policy. |

## Legacy-host delta

`litchi-iwa` remains the migration host governed by [ADR 0028](../../../docs/adr/0028-iwa-monolith-exit.md).
Its source-free `PagesEditor` can build broader Pages graphs, including sections, text boxes,
shapes, media, charts, tables, and table cell/formula operations. Those APIs use legacy/raw IWA
identities and are not counted as focused `litchi-pages` support until their ownership and
evidence move to this crate.

| Feature | Status | Read | Write | Notes |
|-------------------|--------|------|-------|-------|
| Fresh source-free Pages builder | 🟡 | ✅ | 🟡 | Present in [`PagesEditor`](../../litchi-iwa/src/pages/editor.rs) and documented in the [legacy host README](../../litchi-iwa/README.md); migration-only and not a focused package writer. |
| Sections, rich text, shapes, images, movies, audio, charts, and tables | 🟡 | ✅ | 🟡 | Broad host modules exist under [`litchi-iwa/src/pages`](../../litchi-iwa/src/pages). They expose larger legacy graphs than the focused owner but carry no focused-owner or complete native-acceptance credit. |
| Legacy host as a replacement for the focused API | ❌ | N/A | ❌ | The host is not a substitute for the `litchi-pages` public contract; the deletion gate and remaining migration dependencies still apply. |

## Evidence and maintenance

The tracked native fixture is [`test-data/iwork/pages/basic.pages`](../../../test-data/iwork/pages/basic.pages),
with provenance and hashes in [`test-data/iwork/README.md`](../../../test-data/iwork/README.md).
Its focused test proves package open, semantic text/section readback, validation, exact no-op
streaming, and ZIP/bytes parity. Focused transaction tests are primarily synthetic or
self-roundtrip (`E1`). No complete checked-in `E3`/`E4` ledger currently proves that every changed
candidate opens in Apple Pages, survives a native save/close/reopen cycle, and then passes a
strict Litchi reread.

The current verification snapshot is not a clean workspace certification: the selected focused
Pages library and integration tests passed, while the crate-boundary checker is red because the
untracked legacy-host `crates/litchi-iwa/src/pages/editor/tables/lock.rs` restores retired methods.

When a public Pages capability changes, update this file in the same change. Keep package
preservation, semantic read support, focused transaction support, and native-app acceptance as
separate claims. The [root feature-matrix index](../../../docs/FEATURE_MATRIX.md) links this
authoritative matrix.
