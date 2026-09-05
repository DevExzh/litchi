# Pages (`.pages`) Feature Matrix

This document is the authoritative matrix for the public `litchi-pages` crate. It covers
Apple Pages packages built from the iWork ZIP/directory container and IWA component archives.
It describes the typed API exposed by this crate, not complete Pages application compatibility,
rendering fidelity, or every revision of Apple's private protobuf schema.

`Package` is the exact-source adapter: it accepts regular package files or ZIP bytes, retains the
physical artifact, and exposes selected source-bound transactions. `Document` is an immutable,
archive-free semantic snapshot: it can read a ZIP or an app-authored package directory, but it
does not retain exact bytes, media, previews, or unsupported members. `Package::write_to` streams
to a caller-owned sink, while `Package::save` delegates staged regular-file publication, sync,
identity checks, and atomic replacement to the archive-owned save boundary; the documented
platform and filesystem durability caveats still apply.

The broad Pages builder under `litchi-iwa` is a migration host, not evidence for a focused
`litchi-pages` row. Its capabilities are listed separately below.

The focused package currently owns selector-first, exact-source transactions for section
settings/name/pagination/background/layout, body text, header/footer text, body footnotes,
body-table presentation metadata, user-hidden axes, and body drawable stacking order. These owners
keep native identities and wire values private, reopen changed candidates before publication, and return
source-bound reversible patches. Rendering-affecting transactions intentionally invalidate the
canonical root preview members when their owner requires it; exact no-ops and inverse patches
remain source-preserving.

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
coverage. A current-producer native probe called out explicitly below is operation-specific and
does not upgrade the E1/E2 grade or establish package-wide native acceptance.

## Package and semantic reader

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Pages package ingress and application detection | 🟡 | ✅ | 🟡 | [`Package::open`/`from_bytes`](../src/package.rs) parse regular files and ZIP bytes through the bounded IWA archive adapter. [`Document`](../src/document.rs) also captures app-authored directories on supported Unix hosts. Wrong iWork applications, unsafe paths, malformed archives, and ambiguous sources are rejected. Evidence: [`document_reader` tests](../tests/document_reader.rs), [`native fixture`](../tests/native_fixture.rs) (E1/E2). |
| Exact package no-op and source preservation | 🟡 | ✅ | ✅ | `Package::write_to` streams the retained artifact unchanged, including unsupported ZIP members and unmodeled protobuf fields. Selected changed transactions retain unselected data but do not provide a general package editor; rendering-affecting owners may deliberately delete the three canonical root previews and report that invalidation. Exact no-ops and inverse patches restore the complete source artifact, including previews. Evidence: [`Package` implementation](../src/package.rs), [`native fixture`](../tests/native_fixture.rs), and transaction tests below (E1/E2). |
| Archive-free semantic `Document` | ✅ | ✅ | N/A | ZIP, bytes, shared bytes, or directory input projects to immutable ordered [`Section`](../src/section.rs) values with text, headings, paragraphs, storages, names, and known page counts. Exact package bytes, native IDs, media, previews, and unsupported members are intentionally dropped. Evidence: [`Document` implementation](../src/document.rs), [`document_reader` tests](../tests/document_reader.rs) (E1/E2). |
| Section selection and plain-text projection | ✅ | ✅ | N/A | Exact name and checked position selectors are public; duplicate names report an error. Sections and the complete document can be rendered as plain text without exposing native identities. Evidence: [`selectors`](../src/selector.rs), [`section` model](../src/section.rs), [`document_reader` tests](../tests/document_reader.rs) (E1). |
| Package metadata sidecars | 🟡 | ✅ | ❌ | `Package::metadata` reads supported fields from `Metadata/Properties.plist`, `BuildVersionHistory.plist`, and `DocumentIdentifier` into `litchi-core::Metadata`. No public metadata write transaction exists. Evidence: [`metadata projection`](../src/package.rs), [`document settings tests`](../tests/document_settings.rs) (E1). |
| Validation and bounded resource accounting | 🟡 | ✅ | N/A | Physical archive/IWA limits are caller-selectable through [`Limits`](../src/package.rs); semantic defaults cap text at 64 MiB, sections and body storages at 4,096, and package object inspection at 1,000,000. Aggregate graph/live-memory accounting is not a complete global budget. `DocumentReadOptions` adds checked source and semantic limits. Evidence: [`document limits`](../src/document.rs), [`archive limits`](../../litchi-iwa-archive/src/lib.rs), [`document_reader` limit tests](../tests/document_reader.rs) (E1). |
| Unknown fields, opaque members, and preservation-only payloads | 🟡 | ✅ | 🟡 | Unselected physical members and unmodeled IWA/protobuf bytes remain available through exact-source output and selected rewrites. They have no semantic API; unsafe, ambiguous, dependent, or unsupported ownership graphs fail closed. Evidence: [`section-text tests`](../tests/section_text.rs), [`body-table name tests`](../tests/body_table_name.rs), [`header/footer tests`](../tests/header_footer_text.rs) (E1). |

## Existing-package text and document formatting

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Body and section text | 🟡 | ✅ | 🟡 | `Package` exposes selector-first text reads and one checked UTF-16 splice per transaction, with insert, replace, delete, set, and clear helpers. Surrogate splits, section breaks, footnote anchors, inline-object markers, dependent content, multiple staged operations, and non-exact sources are refused. Commits reopen and verify the candidate and return an exact-source inverse. Evidence: [`section text owner`](../src/package/section_text.rs), [`section text tests`](../tests/section_text.rs) (E1; native fixture read is E2). |
| Section settings | 🟡 | ✅ | 🟡 | The aggregate semantic settings value preserves optional name, four header/footer flags, and three pagination fields, including native presence and unknown discriminants. Existing sections can be selected by name or checked position and edited through an exact-source transaction; the selected section component and native template/layout relationships are verified while structural section creation, deletion, and reordering remain unexposed. These edits preserve the derived layout cache and root previews. Evidence: [`section settings owner`](../src/package/section_settings.rs), [`section settings tests`](../tests/section_settings.rs) (E1). |
| Section name and pagination convenience transactions | 🟡 | ✅ | 🟡 | Dedicated transactions expose section-name replacement and section-start, page-numbering, and starting-page-number edits while preserving field presence and unknown discriminants. They remain existing-section, exact-source operations. Evidence: [`section name owner`](../src/package/section_name.rs), [`section name tests`](../tests/section_name.rs), [`pagination owner`](../src/package/section_pagination.rs), [`pagination tests`](../tests/section_pagination.rs) (E1). |
| Document and footnote formatter settings | 🟡 | ✅ | 🟡 | The combined settings transaction reads and replaces supported document options and footnote/endnote kind, marker format, numbering, and gap. It validates canonical enum values, preserves optional presence, checks dependencies, and verifies/reopens changed candidates. It is not a general document-style editor. Evidence: [`settings value`](../src/document_settings.rs), [`package transaction`](../src/package/document_settings.rs), [`document settings tests`](../tests/document_settings.rs) (E1). |
| Section backgrounds | 🟡 | ✅ | 🟡 | Existing sections support `None` and validated solid RGBA fills. Gradients, image fills, media references, and other native fill variants are observable only as unsupported and cannot be authored. Evidence: [`background value and transaction`](../src/section/background.rs), [`background package owner`](../src/package/section_background.rs), [`background tests`](../tests/section_background.rs) (E1). |
| Page layout | 🟡 | ✅ | 🟡 | Existing package layouts expose presence-preserving page width/height, margins, scale, portrait/landscape orientation, and vertical-body settings. Finite scalar and margin invariants are checked; unsupported dependencies and noncanonical variants refuse changed writes. Changed layout publication invalidates the derived layout cache and all three canonical root previews; no-ops preserve them. Evidence: [`layout value`](../src/page_layout.rs), [`layout transaction`](../src/package/page_layout.rs), [`layout tests`](../tests/page_layout.rs) (E1). |
| Headers and footers | 🟡 | ✅ | 🟡 | Rooted logical header/footer slots can be enumerated and selected by section, template role, kind, and slot. Existing slot text can be set or cleared with exact-source, alias, dependency, affected-slot, and candidate-reopen checks. Changed text publication invalidates canonical root previews; no-op and inverse paths preserve them. Slot creation, template CRUD, rich formatting, and general header/footer layout are not exposed. Evidence: [`header/footer model`](../src/header_footer.rs), [`text transaction`](../src/package/header_footer_text.rs), [`header/footer tests`](../tests/header_footer_text.rs) (E1). |
| Body footnotes | 🟡 | ✅ | 🟡 | Rooted body footnotes expose checked positions, semantic text, and optional custom marks. Existing notes can be edited; bounded insert and remove operations update the owned graph, invalidate canonical root previews, and return reversible exact-source patches, while text/mark-only edits preserve previews. Structural markers, shared or ambiguous ownership, and unsafe references are rejected. Evidence: [`footnote model`](../src/footnote.rs), [`body footnote owner`](../src/package/body_footnote.rs), [`footnote text owner`](../src/package/footnote_text.rs), [`body footnote tests`](../tests/body_footnote.rs), [`footnote text tests`](../tests/footnote_text.rs) (E1). |
| Body drawable stacking order | 🟡 | ✅ | 🟡 | Existing body drawables are exposed only as opaque, source-bound [`BodyDrawableHandle`](../src/drawable_order.rs) values and selected by checked position or handle. `DrawableLayerMove` and exact permutations reorder user drawables while preserving the native body-text storage reference, when present, at its original structural slot; the slot is never exposed as a drawable. A changed commit rewrites one z-order payload, deletes `preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg`, verifies/reopens the candidate, and returns an exact inverse. Drawable creation/removal, geometry, grouping, and text-box content remain outside this owner. Evidence: [`drawable-order owner`](../src/package/drawable_order.rs), [`drawable-order tests`](../tests/body_drawable_order.rs) (E1; operation-specific current-producer probe below). |
| Endnote text lifecycle | ❌ | ❌ | ❌ | Formatter settings can select endnote kinds and numbering, but the focused crate has no rooted endnote-text read, insert, edit, or remove API. Evidence: [`footnote kind model`](../src/footnote.rs). |

## Body tables

Table support is intentionally a presentation-metadata surface. The focused crate does not
publish a general table object, cell grid, formula model, or structural table editor. All rows
below operate on an existing rooted body table selected by checked position or exact visible
name; ambiguous, locked, dependent, or unsupported table graphs refuse changed edits. The private
rooted ownership proof keeps body, drawable, attachment, table-info, model, and storage slots
closed to callers. Name discovery uses a borrowed strict Buffa view and allocates only the
semantic name. Appearance, dimensions, header settings, names, titles, and hidden-axis edits
invalidate the three canonical root previews on changed publication; persisted sort and lock
edits preserve them.

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Body-table selection and rooted ownership discovery | 🟡 | 🟡 | N/A | Selector-driven resolution proves the private body, drawable, attachment, table-info, model, and storage ownership chain for each operation. There is no public table catalog or general semantic table enumeration API. Body-table name discovery uses the borrowed [`table_model_discovery_codec`](../src/package/body_table_name.rs) projection and allocates only the returned semantic name. Evidence: [`body-table selectors`](../src/selector.rs), [`table ownership proof`](../src/package/table_lock.rs), [`table-lock tests`](../tests/table_lock.rs), and [`body-table name tests`](../tests/body_table_name.rs) (E1). |
| Body-table hidden rows and columns | 🟡 | ✅ | 🟡 | `Package::{body_table_hidden_axes, edit_body_table_hidden_axes, apply_body_table_hidden_axes}` uses [`BodyTableSelector`](../src/selector.rs) (checked position or exact visible name) and archive-free [`table::hidden_axes::{AxisIndex, HiddenAxes}`](../src/table/hidden_axes.rs). The strict source-backed route validates one role-qualified `TableInfoArchive` (current type 6000 or qualified legacy type 6003) and `TableModelArchive` (current type 6001 or qualified legacy type 6000), the canonical type-6267 row/column UID map (qualified legacy type 6200 is admitted only through its legacy route), type-4008 formula-owner chain, active hidden-state UUID, and directional row/column extents; only user-hidden positions are projected. Bounds, duplicate positions, missing/duplicate owners, malformed wire, stale UUIDs, bad UID maps, unsupported pivot/dependency graphs, locked tables, and finite wire/archive limits fail closed before publication. On admitted existing-owner rewrites, non-user filtered/pivot markers and unknown model/info/owner fields remain preserved; unsupported pivot/dependency topologies refuse changed edits. An absent hidden-state owner reads as empty, an empty request is an exact source no-op, and a nonempty absent-owner request is unsupported and returns `UnsupportedDependency` or `UnsupportedSource` without publication. Existing-owner commits use copy-on-write, touch one selected component, invalidate the three canonical root previews, reopen/read back the candidate, and return exact source-bound apply/inverse patches. This is visibility metadata only: no cell/formula, filter/pivot CRUD, row/column topology, or sorting API is implied. Focused graph, codec, identity/COW, and concurrency tests cover the E1 source/self-round-trip contract. The checked-in [`body-table-visible.pages`](../../../test-data/iwork/pages/body-table-visible.pages) fixture is a native Pages 14.4 visible 5-by-4 body-table baseline with a body marker; disposable Computer Use copies were saved, closed, and reopened without repair, and the focused package save path produced byte-identical output for the visible-table/no-op operation. The checked-in fixture has no user-hidden axes, so it is native baseline/no-op evidence rather than positive hidden-axis E2 or E3/E4 mutation evidence. The registered fuzz target and corpus remain bounded evidence; current fuzz verification is tracked with ADR 0008. |
| Table appearance | 🟡 | ✅ | 🟡 | Existing tables expose validated style, gridline, banding, and row-sizing settings; exact-source edits support the admitted appearance profile, preserve unselected native fields/slots, refuse unsupported style graphs, and invalidate canonical root previews when changed. Evidence: [`appearance value`](../src/table/appearance.rs), [`appearance transaction`](../src/package/body_table_appearance.rs), [`appearance tests`](../tests/body_table_appearance.rs) (E1). |
| Table dimensions | 🟡 | ✅ | 🟡 | Existing table width/height and row/column dimension values are read and edited with checked finite bounds. This changes presentation metadata, not table row/column topology; changed publication invalidates canonical root previews while no-ops and inverses preserve the source artifact. Evidence: [`dimension value`](../src/table/dimension.rs), [`dimension transaction`](../src/package/body_table_dimension.rs), [`dimension tests`](../tests/body_table_dimensions.rs) (E1). |
| Header, footer, freeze, and repeat settings | 🟡 | ✅ | 🟡 | Existing body-table header/footer counts and freeze/repeat flags are presence-preserving and editable within the validated native graph. Physical header-row data and table topology are not exposed; changed partition edits invalidate canonical root previews and keep native ownership slots closed. Evidence: [`header settings`](../src/table/headers.rs), [`header transaction`](../src/package/body_table_headers.rs), [`header tests`](../tests/body_table_header_settings.rs) (E1). |
| Table name | 🟡 | ✅ | 🟡 | Existing rooted table names can be read and renamed through selector-first, source-bound patches. The strict borrowed discovery/rewrite preserves unknown model fields and native slots; duplicate names, invalid names, locked tables, and foreign/ambiguous ownership fail closed. Changed renames invalidate canonical root previews. Evidence: [`name value`](../src/table/name.rs), [`name transaction`](../src/package/body_table_name.rs), [`name tests`](../tests/body_table_name.rs) (E1). |
| Persisted table sort configuration | 🟡 | ✅ | 🟡 | Existing sort rules, scope, direction, and column references can be read and replaced. This edits persisted configuration only; it does not sort or recalculate table cell data, and the owner deliberately performs no root-preview deletion. Evidence: [`sort value`](../src/table/sort.rs), [`sort transaction`](../src/package/body_table_sort.rs), [`sort tests`](../tests/body_table_sort.rs) (E1). |
| Table title visibility and outline | 🟡 | ✅ | 🟡 | Existing table title visibility and outline settings are read and edited with presence-preserving values and exact-source verification. Changed title publication invalidates canonical root previews; native references and unknown fields remain preservation-owned. Evidence: [`title value`](../src/table/title.rs), [`title transaction`](../src/package/body_table_title.rs), [`title tests`](../tests/body_table_title.rs) (E1). |
| Table interactive lock state | 🟡 | ✅ | 🟡 | Existing body tables expose locked/unlocked state and support source-bound lock edits. Lock state is not authorization or encryption; changed edits are refused when the ownership proof is not unique. Lock transactions preserve root previews and do not expose the native lock or table identifiers. Evidence: [`lock value`](../src/table/lock.rs), [`lock transaction`](../src/package/table_lock.rs), [`lock tests`](../tests/table_lock.rs) (E1). |
| Table cells, values, formulas, formatting, and merges | ❌ | ❌ | ❌ | No focused public API enumerates or edits Pages table cells, formulas, rich cell formatting, merged ranges, or calculated results. Hidden-axis metadata does not expose cell data or formula dependencies; the legacy host's table editor does not upgrade this row. |
| Table row/column and table structural CRUD | ❌ | ❌ | ❌ | No focused package API creates, deletes, duplicates, resizes structurally, or reorders Pages tables, rows, or columns. Hidden-axis edits only replace existing visibility metadata and never insert, remove, or reorder axes. |

## Rich content and application features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Rich-text runs, styles, lists, links, and annotations | 🟡 | 🟡 | ❌ | Text storages and plain text/range values are readable through the semantic model. The focused package has no public style, list, hyperlink, revision, annotation, or rich-run mutation API. Shared text codecs are not a Pages feature claim. Rich formatting and comment-thread editing remain migration-host work. Evidence: [`section text model`](../src/section.rs), [`text storage API`](../src/lib.rs), [`section text tests`](../tests/section_text.rs) (E1). |
| Images, audio, movies, and media assets | ❌ | ❌ | ❌ | The public [`image`](../src/image.rs), [`audio`](../src/audio.rs), and [`movie`](../src/movie.rs) modules contain validated detached option values for migration-host builders; they are not package asset readers or resource CRUD APIs. |
| Drawable geometry, grouping, text boxes, and shape lifecycle | ❌ | ❌ | ❌ | No focused Pages package model or transaction owns drawable kind/geometry, grouping, creation/removal, or text-box content. The separate focused body drawable-order owner exposes only opaque order handles and does not imply shape support. |
| Charts and chart data | ❌ | ❌ | ❌ | No focused Pages chart catalog, data-series model, chart creation, or chart property transaction is public. The legacy host's broad chart graph remains migration-only. |
| Collaboration, comments, revisions, and coauthoring | ❌ | ❌ | ❌ | No focused semantic or package lifecycle API exists for collaboration services, comments, revision resolution, locks beyond table presentation state, or coauthoring. Legacy text-box/drawable and body-table-cell comment threads remain host-only and do not upgrade this row. |
| Rendering and export | ❌ | ❌ | ❌ | The crate does not render Pages, produce PDF/HTML/RTF, paginate layout, or automate the Pages application. Plain-text projection is not rendering. |
| Fresh package creation and structural document authoring | ❌ | N/A | ❌ | Detached `Document`, `Root`, `Body`, and `Section::Builder` values can form an archive-free semantic snapshot, but `litchi-pages` has no fresh `.pages` package writer or section/asset/shape structural authoring API. |

## Security, durability, and interoperability boundaries

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Password-protected or encrypted Pages packages | ❌ | ❌ | ❌ | Shared iWork detection and archive ingress reject encrypted/password-protected sources; no password, decryption, or re-encryption API exists. Evidence: [`detector error and limits`](../../litchi-iwa-detect/src/lib.rs), [`archive error`](../../litchi-iwa-archive/src/error.rs). |
| External resources, actions, macros, controls, and embedded code | 🟡 | 🟡 | 🟡 | Unsupported payloads may remain inert and opaque in an exact package artifact, but no external target is resolved or fetched and no action, macro, control, or embedded code is activated or executed. Changed transactions refuse unsupported dependency graphs. Evidence: [`package preservation boundary`](../src/package.rs), [`legacy host policy`](../../litchi-iwa/README.md). |
| Cryptographic signatures and trust | ❌ | ❌ | ❌ | Exact no-op output can preserve bytes that happen to be signed; the focused crate has no signature verification, certificate trust, invalidation, or re-signing model. |
| Durable patch serialization, composition, merge, and history | ❌ | N/A | ❌ | Focused patches retain process-local source/target artifacts or logical deltas; there is no stable patch format, composition/merge protocol, or persistent history API. |
| Durable atomic filesystem publication | 🟡 | N/A | 🟡 | [`Package::save`](../src/package.rs) delegates staged regular-file publication, flush/sync, identity checks, and atomic replacement to the archive-owned save boundary. Platform/filesystem durability caveats and committed post-replacement errors remain explicit; `write_to` is still only a caller-owned streaming sink. |

## Legacy-host delta

`litchi-iwa` remains the migration host governed by [ADR 0028](../../../docs/adr/0028-iwa-monolith-exit.md).
Its source-free `PagesEditor` still owns broader Pages graphs, including sections, rich text,
text boxes, shapes, body media, charts, tables, cell/formula data, and comment threads. Most
legacy operations still select native drawable/table/storage objects by raw IDs; the recent
typed `MediaAssetId` boundary applies only to selected host media-asset references and does not
make the focused package a media owner. None of these host capabilities is counted as focused
`litchi-pages` support until its ownership and evidence move to this crate.

| Feature | Status | Read | Write | Notes |
|-------------------|--------|------|-------|-------|
| Fresh source-free Pages builder | 🟡 | ✅ | 🟡 | Present in [`PagesEditor`](../../litchi-iwa/src/pages/editor.rs) and documented in the [legacy host README](../../litchi-iwa/README.md); migration-only and not a focused package writer. |
| Sections and rich text/text-box content | 🟡 | ✅ | 🟡 | Broad section, text-storage, text-box, style, list, hyperlink, and inline-object operations remain under [`litchi-iwa/src/pages`](../../litchi-iwa/src/pages). They use migration-host graph identities and do not transfer rich-content ownership to the focused package. |
| Body images, movies, and audio | 🟡 | ✅ | 🟡 | Host [`images`](../../litchi-iwa/src/pages/editor/images.rs), [`movies`](../../litchi-iwa/src/pages/editor/movies.rs), and [`audio`](../../litchi-iwa/src/pages/editor/audio.rs) retain discovery, creation, duplication, geometry/properties, replacement, playback settings, and removal. Drawable graph selection remains raw-ID based; focused `litchi-pages` has no package media reader or asset CRUD owner. |
| Body charts and chart properties | 🟡 | ✅ | 🟡 | Host [`charts`](../../litchi-iwa/src/pages/editor/charts.rs) retains inline chart discovery/creation/duplication/removal, data/kind/direction/geometry, and many chart-property mutations. It remains a raw-ID migration surface; focused `litchi-pages` has no chart or chart-data owner. |
| Body shapes, drawings, and drawable order | 🟡 | ✅ | 🟡 | Host [`body shapes`](../../litchi-iwa/src/pages/editor/body_shapes.rs) retains creation, duplication, removal, geometry, and raw-ID shape operations. The focused `litchi-pages` package owns the bounded opaque body-order transaction; no legacy host drawable-order editor remains, and shape lifecycle or geometry are not part of that owner. |
| Body tables, cells, formulas, and table comments | 🟡 | ✅ | 🟡 | Host [`tables`](../../litchi-iwa/src/pages/editor/tables/mod.rs) retains broad table topology/storage/value/formula operations and [`table comments`](../../litchi-iwa/src/pages/editor/tables/comments.rs) retains cell comment/reply CRUD. The Pages raw-ID hidden-axis convenience route remains compatibility debt until creation parity and native mutation gates pass; focused Pages owns the bounded selector-first existing-owner metadata transaction. Attached Numbers/Keynote hidden-axis compatibility also remains migration-host-only. No focused cell/formula/comment owner exists. |
| Drawable/text-box comments and replies | 🟡 | ✅ | 🟡 | Host [`PagesEditor` comment methods](../../litchi-iwa/src/pages/editor.rs) and the shared comment editor retain raw-ID comment/reply CRUD for drawable/text-box targets. This is explicitly a remaining host gap; focused `litchi-pages` has no comment-thread model or transaction. |
| Legacy host as a replacement for the focused API | ❌ | N/A | ❌ | The host is not a substitute for the `litchi-pages` public contract; the deletion gate and remaining migration dependencies still apply. |

## Evidence and maintenance

The tracked native fixture is [`test-data/iwork/pages/basic.pages`](../../../test-data/iwork/pages/basic.pages),
with provenance and hashes in [`test-data/iwork/README.md`](../../../test-data/iwork/README.md).
Its focused test proves package open, semantic text/section readback, validation, exact no-op
streaming, and ZIP/bytes parity. Focused transaction tests are primarily synthetic or
self-roundtrip (`E1`). No complete checked-in `E3`/`E4` ledger currently proves that every changed
candidate opens in Apple Pages, survives a native save/close/reopen cycle, and then passes a
strict Litchi reread.

An operation-specific current-producer probe is recorded in [ADR 0028](../../../docs/adr/0028-iwa-monolith-exit.md#2026-09-03-amendment-typed-media-boundaries-and-bounded-archive-verification):
Apple Pages authored a marker, square, and circle; Arrange > Send to Back saved without repair
or conversion; and the focused reader resolved two opaque drawable handles. The library reorder
preserved the native body-storage slot, removed all three canonical root previews, reopened the
candidate successfully, and returned an exact inverse. Pages then opened the library-produced
candidate with the marker and both shapes intact, without repair or conversion UI. This is native
evidence for that bounded drawable-order operation only, not a complete Pages E3/E4 ledger or
package-wide acceptance claim.

The focused Pages graph, codec, identity/COW, and concurrency tests cover E1
source/self-round-trip behavior. The checked-in [`body-table-visible.pages`](../../../test-data/iwork/pages/body-table-visible.pages)
fixture is a native Pages 14.4 baseline with a visible 5-by-4 body table and a body marker.
Disposable Computer Use copies were saved, closed, and reopened without repair UI; the focused
package save path produced byte-identical output for the visible-table/no-op operation. The
checked-in fixture was restored after those UI checks. Its hidden-axis API returns `InvalidSource`
for this native shape, and the fixture has no user-hidden axes, so it supplies native
visible-table/no-op baseline evidence rather than positive hidden-axis E2 or E3/E4 mutation
evidence. An older exploratory Pages 14.4 attempt logged an NSCocoa MissingObject/TSPersistence
Import document error for a generated candidate; save/close timed out and no reopen occurred.
AppleScript table creation also stalled without GUI inspection. That attempt remains historical
negative exploratory evidence only. The private [`pages_hidden_state_codec`](../../litchi-iwa-protos/src/pages_hidden_state_codec.rs)
keeps repeated state/extents on a bounded caller-owned wire walk and uses the narrow Buffa
projection only for singular envelopes. The registered Pages hidden-axis fuzz target and checked-in
valid, malformed, ownership, and limit corpus remain bounded evidence; current fuzz verification
is tracked with ADR 0008. No positive hidden-axis native mutation, native byte parity, or
package-wide hidden-state support claim follows.

Feature-level verification is recorded in [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-05-amendment-pages-body-table-hidden-axis-owner-and-retained-host-compatibility).
It does not certify the whole workspace or close the Pages native-mutation gate.

When a public Pages capability changes, update this file in the same change. Keep package
preservation, semantic read support, focused transaction support, and native-app acceptance as
separate claims. The [root feature-matrix index](../../../docs/FEATURE_MATRIX.md) links this
authoritative matrix.
