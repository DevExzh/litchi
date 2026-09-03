# Keynote Feature Matrix

This is the authoritative feature matrix for the public `litchi-keynote` crate and native
Keynote (`.key`) packages. It is the format-specific companion to the repository
[feature-matrix index](../../../docs/FEATURE_MATRIX.md). The matrix describes the focused owner
crate; capabilities that still live in the migration host are listed separately and do not count
as focused-owner support.

The audited source tree is a dirty development snapshot. The verification note below records the
2026-08-31 run at `a1dc2f8007ed1ea59216026d91e0a1f0d621ab6e`; uncommitted changes are work in
progress, not release evidence. Keynote's `.key` container and IWA/protobuf payloads are treated
as native application data. This document does not claim complete conformance with an unpublished
or reverse-engineered Keynote schema, rendering fidelity, or application compatibility.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | The documented feature scope has a public typed implementation |
| 🟡 | Support is bounded, partial, metadata-only, pass-through, preservation-only, or otherwise limited |
| ❌ | No public typed support is currently available |
| N/A | The concept does not apply to the format or direction |

`Read` and `Write` are independent directions. A `🟡` direction is not a claim of general CRUD:
it may describe a subset, an existing-object-only transaction, an inert serializer, or exact
preservation. A passing synthetic round trip does not establish native Keynote acceptance.

## Ownership and ingress boundaries

The focused owner is `litchi-keynote`. [`Package`](../src/package.rs) retains an exact
regular-ZIP artifact and exposes selected source-bound transactions. [`Document`](../src/document.rs)
is an archive-free semantic projection: its path ingress accepts a ZIP or app-authored package
directory, but it intentionally drops exact bytes, media, previews, and editing state. Package
directories therefore have semantic-read support only. Detached [`Show`](../src/show.rs#L334)
and [`Slide`](../src/slide.rs#L45) builders construct values, not writable `.key` packages.

Every changed focused transaction is expected to authorize the exact source, enforce finite
physical/wire/semantic limits, validate the selected ownership graph, reopen the candidate, and
read back the requested value. A transaction is still limited to the feature named in its row;
it is not a general Keynote package editor.

## Package and semantic projection

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Existing `.key` ZIP ingress from path or bytes | ✅ | ✅ | N/A | `Package::open` and `Package::from_bytes` validate the native root, retain the source artifact, and accept explicit archive/semantic profiles ([package ingress](../src/package.rs), [native fixture](../tests/native_fixture.rs)). |
| Exact retained-artifact stream to a caller sink | ✅ | N/A | ✅ | `Package::write_to` emits the retained bytes and unmodeled members/fields unchanged; it does not flush, sync, rename, or atomically publish a filesystem path ([writer](../src/package.rs), [exact no-op test](../tests/native_fixture.rs)). |
| Archive-free ZIP or package-directory semantic document | ✅ | ✅ | N/A | `Document::open` and `Document::from_bytes` retain only the bounded show, slides, text, source metadata, and statistics; directory input has no exact package artifact or edit path ([document boundary](../src/document.rs), [directory parity test](../tests/native_fixture.rs)). |
| Semantic show, slide order, text, notes, media/build summaries | 🟡 | ✅ | N/A | Rooted graph projection is typed and immutable, but the model intentionally exposes only selected content and no native IDs ([show model](../src/show.rs#L334), [slide model](../src/slide.rs#L45)). |
| Unknown ZIP members, IWA components, and protobuf fields | 🟡 | ✅ | 🟡 | Untouched exact artifacts are retained; focused rewrites preserve unselected records only within each transaction's locality proof. Opaque/unsupported compression or unmodeled ownership can cause a typed refusal rather than a semantic edit ([physical source](../src/package.rs), [transaction source checks](../src/package/edit.rs#L281)). |
| Focused source-bound transactions and reversible patches | 🟡 | ✅ | 🟡 | Selected owners expose one bounded operation per commit, exact-source authorization, inverse patches, and candidate rereadback; coverage is not a composable/durable patch history and the current soundtrack/ZIP verification failures below prevent a green suite claim ([base transaction](../src/package/edit.rs#L79), [slide-order transaction](../src/package/slide_order.rs#L151)). |
| Fresh `.key` package creation | ❌ | N/A | ❌ | Detached semantic builders do not emit a native package; package creation remains in the legacy host ([show builder](../src/show.rs#L484), [legacy builder](../../litchi-iwa/src/keynote/creation.rs#L1)). |
| Durable patch serialization, composition, merge, and history | ❌ | N/A | ❌ | Focused patches retain process-local source/target artifacts or logical deltas; there is no stable patch format, composition/merge protocol, or persistent history API. |
| Durable atomic filesystem save | ❌ | N/A | ❌ | `write_to` streams to a caller-owned sink only; callers must choose their own safe create/sync/rename workflow ([save boundary](../src/lib.rs#L141)). |
| Metadata and properties sidecar | 🟡 | ✅ | ❌ | `Package::metadata` reads semantic show data plus the canonical `Metadata/Properties.plist`; no focused metadata editor is exposed ([metadata reader](../src/package.rs), [metadata fixture assertion](../tests/native_fixture.rs)). |

## Slides, text, settings, and playback state

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Slide selection by checked position or unique navigator name | ✅ | ✅ | N/A | Selectors are semantic and reject empty/ambiguous names; native object identifiers remain private ([selector](../src/selector.rs#L1), [show selection](../src/show.rs#L375)). |
| Existing slide skip/include state | ✅ | ✅ | ✅ | One existing slide can be skipped or included with an exact-source patch and full candidate reopen; this does not add, duplicate, or reorder slides ([edit API](../src/package/edit.rs#L99), [skip tests](../tests/skip_state.rs#L1)). |
| Existing slide move/reorder | 🟡 | ✅ | 🟡 | One move per transaction over the admitted flat slide topology; destinations are final positions and unsupported/hierarchical topologies refuse ([slide-order API](../src/package/slide_order.rs#L166), [order tests](../tests/slide_order.rs#L1)). |
| Existing slide deletion | 🟡 | ✅ | 🟡 | Deletes one selected existing slide, refuses deletion of the final slide, preserves selected package data according to ownership checks, and provides an inverse; no general lifecycle manager ([deletion API](../src/package/slide_delete.rs#L619), [deletion tests](../tests/slide_delete.rs#L900)). |
| Slide insertion, duplication, or fresh slide creation | ❌ | N/A | ❌ | No focused package transaction creates a slide or transfers its dependency graph; corresponding broad operations remain legacy-host-only ([legacy slide creation](../../litchi-iwa/src/keynote/editor/slide_create.rs#L1)). |
| Existing title/body placeholder text | 🟡 | ✅ | 🟡 | Plain text and bounded UTF-16 replacement are supported for existing, exclusively owned title/body storages; slide-number placeholders are visibility-only and dependent/inline-marker cases refuse ([text API](../src/package/slide_text.rs#L739), [text tests](../tests/slide_text.rs#L832)). |
| Existing speaker notes text | 🟡 | ✅ | 🟡 | Reads and edits an existing notes storage by selector and text span; absent notes are not synthesized ([notes API](../src/package/slide_notes.rs#L575), [notes tests](../tests/slide_notes.rs#L408)). |
| Per-slide placeholder visibility | 🟡 | ✅ | 🟡 | Existing title/body/slide-number roles can be shown or hidden; missing roles remain distinct and this does not change the show-wide slide-number preference ([visibility API](../src/package/slide_placeholder_visibility.rs#L675), [visibility tests](../tests/slide_placeholder_visibility.rs#L188)). |
| Rich-text storage inspection and plain-text projection | 🟡 | ✅ | N/A | Semantic slides expose text storages/ranges and rooted plain text; this is not a complete style, list, link, revision, or inline-object model ([slide text values](../src/slide.rs#L121), [native text projection](../tests/rich_storage.rs#L1)). |
| Rich-text styles, lists, links, revisions, and inline-object CRUD | ❌ | 🟡 | ❌ | General formatting and inline-object mutation are not public focused-owner operations; title/body edits can refuse dependent markers or unsupported storage shapes ([text refusal boundary](../src/package/slide_text.rs#L140), [legacy rich editor](../../litchi-iwa/src/keynote/editor/named_paragraph_styles.rs#L1)). |
| Show dimensions and playback settings | 🟡 | ✅ | 🟡 | Typed size, mode, slide-number/loop/autoplay/idle settings support existing-source replacement with candidate verification; no playback engine or rendering behavior is provided ([settings API](../src/package/show_settings.rs#L416), [settings tests](../tests/show_settings.rs#L1067)). |
| Slide transitions | 🟡 | ✅ | 🟡 | Existing modern transition envelopes expose bounded effects, timing, direction, delivery, and opaque fields; clear writes a native no-effect value and does not synthesize legacy-only transitions ([transition API](../src/package/slide_transition.rs#L1150), [transition values](../src/transition.rs#L760), [transition tests](../tests/slide_transition.rs#L1062)). |
| Slide backgrounds | 🟡 | ✅ | 🟡 | Effective/inherited and direct overrides support none, solid, gradient, and bounded opaque native fills; image-background semantics and rendering are not modeled ([background API](../src/package/slide_background.rs#L489), [background values](../src/background.rs#L10), [background tests](../tests/slide_background.rs#L1)). |
| Builds and animation summaries | 🟡 | ✅ | ❌ | Slides expose bounded build/effect summaries and durations; there is no focused build/animation edit transaction or playback runtime ([build model](../src/build.rs#L205), [semantic decode](../src/package.rs#L1240)). |
| Shapes, groups, lines, transforms, and z-order editing | ❌ | ❌ | ❌ | Focused semantic slides do not expose a general drawable graph or package shape editor; legacy `litchi-iwa` owns broad shape operations during migration ([focused slide fields](../src/slide.rs#L45), [legacy shapes](../../litchi-iwa/src/keynote/editor/slide_shapes.rs#L1)). |
| Masters, templates, themes, layouts, guides, and inheritance authoring | ❌ | 🟡 | ❌ | Some inherited style/background effects are read while resolving selected owners, but there is no public focused master/theme/layout lifecycle API ([background inheritance](../src/package/slide_background.rs#L1), [legacy layout editor](../../litchi-iwa/src/keynote/editor/slide_layout_update.rs#L1)). |
| Comments, collaboration, and coauthoring | ❌ | ❌ | ❌ | No focused semantic or package API exposes comment threads, author identity, collaboration state, revision resolution, or coauthoring ([focused slide fields](../src/slide.rs#L45)). |

## Tables

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Table selection/discovery in an existing slide | 🟡 | 🟡 | N/A | Property owners use a checked positional [`TableSelector`](../src/slide/table.rs#L26) and private graph admission; there is no public table catalog or stable table identity. |
| Table name | 🟡 | ✅ | 🟡 | Existing table names have selector-driven source-bound read/edit/clear behavior; locked or unsupported ownership graphs refuse ([table-name API](../src/package/slide_table_name.rs#L379), [tests](../tests/slide_table_name.rs#L839)). |
| Table title visibility/settings | 🟡 | ✅ | 🟡 | Existing table title state can be read and replaced; this is presentation metadata, not table content ([table-title API](../src/package/slide_table_title.rs#L558), [tests](../tests/slide_table_title.rs#L271)). |
| Table appearance and style metadata | 🟡 | ✅ | 🟡 | Bounded appearance/banding/gridline/row-sizing values are edited through existing table style ownership; arbitrary style graph creation is not claimed ([appearance API](../src/package/slide_table_appearance.rs#L1007), [tests](../tests/slide_table_appearance.rs#L893)). |
| Table dimensions | 🟡 | ✅ | 🟡 | Existing row/column point-size dimensions have checked edits; no row/column insertion or deletion ([dimension API](../src/package/slide_table_dimension.rs#L871), [tests](../tests/slide_table_dimension.rs#L879)). |
| Header/footer/freeze/repeat settings | 🟡 | ✅ | 🟡 | Existing table header settings expose bounded presentation and repeat/freeze metadata; this is not print layout or cell topology ([headers API](../src/package/slide_table_headers.rs#L763), [tests](../tests/slide_table_headers.rs#L1042)). |
| Table lock state | 🟡 | ✅ | 🟡 | Existing table lock state is readable and source-bound editable; a locked table can reject unrelated property edits ([lock API](../src/package/slide_table_lock_state.rs#L787), [tests](../tests/slide_table_lock_state.rs#L709)). |
| Persisted table sort configuration | 🟡 | ✅ | 🟡 | Reads/writes persisted sort rules and ranges only; it does not invoke Keynote's physical “Sort Now” executor or move row data. Physical row movement is the separate focused owner below ([sort values](../src/slide/table/sort.rs#L1), [sort transaction](../src/package/slide_table_sort_order.rs#L564), [sort tests](../tests/slide_table_sort_order.rs#L1)). |
| Physical table row sorting (“Sort Now”) | 🟡 | 🟡 | 🟡 | The existing-table owner executes persisted order through selector-first `Package::{execute_slide_table_sort_order,execute_slide_table_sort_order_to_rows}` and the exact-source `edit/apply` transaction; no public row/value reader is exposed. Admission is limited to the canonical type-6001 table-model route and explicitly proven tile, data-list, header, UID, and empty pre-BNC sentinel shapes. Sort keys are finite scalar text, number, boolean, date, and duration values: text uses Rust lexical ordering, numeric/date/duration values use `f64::total_cmp`-based deterministic ordering except that signed zeroes compare equal and retain source order, booleans use their ordinary ordering, and mixed scalar domains fail closed. Duplicate keys retain deterministic source-row order. Strict wire preflight precedes private borrowed Buffa lazy views; unknown/unselected bytes remain source-authoritative. Formula/error, rich-text, comment, merge, filter/group/category/pivot/spill/conditional, hidden/non-positional, imported/provenance, cross-tile/cross-bucket, non-empty stroke, unknown mutable, and other unproven row-affine dependencies refuse atomically. Exact-source patches/inverses, candidate reopen/readback, bounded locality, and preview invalidation are source-level contracts. A disposable Computer Use run opened a pre-hardening candidate in Keynote 14.4 and showed the expected order, but the current strict owner rejects that app-authored source because model field 39 identifies an unowned conditional-style CalculationEngine dependency graph. The run therefore remains external exploratory evidence, not current-owner E3/E4 promotion; the checked-in native fixture has no table and the checked-in evidence test records hashes without launching Keynote. Native acceptance remains pending ([semantic values](../src/slide/table/physical_sort.rs#L1), [physical owner](../src/package/slide_table_physical_sort.rs#L1), [integration tests](../tests/slide_table_physical_sort.rs#L1), [ADR 0008](../../../docs/adr/0008-migration-and-verification.md#2026-09-02-amendment-keynote-physical-sort-focused-owner-verification-status)). |
| Table cells, formulas, formatting, merges, and row/column topology | ❌ | ❌ | ❌ | Physical row movement above is not a public cell model. Formula and dimension vocabularies are detached semantic values, not focused package cell CRUD; there is no cell authoring, formula recalculation, merge, or general structural table transaction ([detached formula values](../src/slide/table/formula.rs#L1), [private table graph](../src/package/slide_table_core.rs#L1)). |

## Charts

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Existing slide chart catalog and visible titles | 🟡 | ✅ | 🟡 | A selected slide exposes chart summaries by source order/title and an existing chart title transaction; no chart creation or graph transfer ([chart catalog API](../src/package/slide_chart_title.rs#L509), [chart values](../src/chart.rs#L130), [tests](../tests/slide_chart_title.rs#L245)). |
| Existing chart captions | 🟡 | ✅ | 🟡 | Existing chart captions can be read/replaced/cleared through source-bound transactions; captions are not chart data or styling ([caption API](../src/package/slide_chart_caption.rs#L1062), [tests](../tests/slide_chart_caption.rs#L1346)). |
| Existing chart axis titles | 🟡 | ✅ | 🟡 | Existing chart axis titles have bounded selector-driven read/edit behavior for modeled axes ([axis-title API](../src/package/slide_chart_axis_title.rs#L1411), [tests](../tests/slide_chart_axis_title.rs#L24)). |
| Existing primary value-axis settings | 🟡 | ✅ | 🟡 | Bounded bounds, steps, and scale settings are source-bound edited and reread; this is not a full chart axis/style model ([value-axis API](../src/package/slide_chart_value_axis.rs#L1518), [tests](../tests/slide_chart_value_axis.rs#L1)). |
| Existing chart legend visibility (current WIP) | 🟡 | 🟡 | 🟡 | The dirty worktree exposes a bounded visibility getter and source-bound edit/apply transaction for existing charts. The source and integration suite are uncommitted, were not part of the recorded verification snapshot, and are not release-certified ([WIP owner](../src/package/slide_chart_legend.rs), [WIP tests](../tests/slide_chart_legend.rs)). |
| Chart data, series, types, legend layout/styles, formulas, and chart CRUD | ❌ | 🟡 | ❌ | No focused chart data graph or create/replace/remove transaction; broad chart authoring remains in the migration host ([legacy charts](../../litchi-iwa/src/keynote/editor/slide_charts.rs#L1)). |

## Movies, images, audio, and soundtrack

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Existing movie/audio drawable inventory and metadata | 🟡 | 🟡 | N/A | Semantic slides expose source-ordered `MovieInfo` values with kind, geometry, natural size, and playback metadata but no package asset bytes ([media values](../src/slide/media.rs#L504), [movie projection](../src/slide.rs#L139)). |
| Existing movie title and caption | 🟡 | ✅ | 🟡 | Existing file-backed movie title/caption fields have selector-driven exact-source transactions; no media resource lifecycle ([title API](../src/package/slide_movie_title.rs#L421), [caption API](../src/package/slide_movie_caption.rs#L598), [caption tests](../tests/slide_movie_caption.rs#L741)). |
| Existing movie geometry and playback | 🟡 | ✅ | 🟡 | Existing movie position/size/transform and bounded playback settings can be edited and reread; rendering and playback are not performed ([geometry API](../src/package/slide_movie_geometry.rs#L822), [playback API](../src/package/slide_movie_playback.rs#L660), [geometry tests](../tests/slide_movie_geometry.rs#L1)). |
| Image/movie/audio bytes and asset insertion/replacement/deletion | ❌ | ❌ | ❌ | Focused `image`/`audio`/`movie` modules provide detached option/value types only; no package resource API exists ([image options](../src/slide/image.rs#L1), [legacy media authoring](../../litchi-iwa/src/keynote/editor/slide_images.rs#L1)). |
| Existing soundtrack playback settings | 🟡 | ✅ | 🟡 | Reads and edits existing volume/mode settings while preserving the existing media collection; it never creates a soundtrack ([settings API](../src/package/soundtrack_settings.rs#L152), [tests](../tests/soundtrack_settings.rs#L92)). |
| Existing soundtrack media-reference order | 🟡 | ❌ | 🟡 | A bounded transaction can move existing in-package references only after strict root/show/soundtrack ownership checks; there is no public getter for the current reference order, and the transaction does not create or reorder media assets. The current `soundtrack_order` integration suite is red in the audited dirty snapshot ([order implementation](../src/package/soundtrack_order.rs#L634), [order tests](../tests/soundtrack_order.rs#L392)). |
| Soundtrack audio item lifecycle | 🟡 | ✅ | 🟡 | `soundtrack::items` owns bounded read/add/insert/replace/remove transactions with opaque source-bound handles; focused tests cover exact apply/inverse, conflicts, shared occurrences, malformed references, and aggregate-only preservation, and a focused fuzz target/corpus is present. The duplicate legacy item API is retired. One genuine replacement candidate passed Keynote save/close/reopen without repair, but this is representative replacement evidence rather than operation-wide native certification; soundtrack creation and general asset CRUD remain unsupported ([semantic surface](../src/soundtrack/items.rs), [package owner](../src/package/soundtrack_items.rs), [lifecycle tests](../tests/soundtrack_items.rs), [example](../examples/edit_soundtrack_items.rs)). |

## Security, limits, and explicit non-goals

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| ZIP/IWA/wire/semantic resource ceilings | 🟡 | ✅ | 🟡 | Physical archive limits and semantic ceilings cover input bytes, entries, component bytes, objects, slides, references, text storages/fragments, text bytes, wire fields/depth/work, and focused transaction output; aggregate graph/live-memory accounting is not a complete global budget ([limits](../src/package/limits.rs#L1), [semantic preflight](../src/package.rs)). |
| Safe regular-file path ingress | ✅ | ✅ | N/A | `Package::open` documents descriptor-safe regular-file handling, no-follow behavior, mutation checks, and a bytes-ingress fallback; platform profiles fail closed when unavailable ([path ingress](../src/package.rs), [directory reader boundary](../src/document.rs)). |
| Encrypted/password-protected packages | ❌ | ❌ | ❌ | Password/decryption support is not implemented; encrypted sources are rejected by the shared iWork detector ([detector boundary](../../litchi-iwa-detect/src/lib.rs#L1)). |
| Digital signatures and trust verification | ❌ | ❌ | ❌ | No Keynote signature verification, certificate trust, invalidation, or re-signing API is exposed. |
| External links/resources, actions, and playback targets | 🟡 | 🟡 | ❌ | Exact package preservation can retain opaque/inert values, but the focused crate does not fetch, resolve, follow, play, or execute external targets ([package preservation](../src/package.rs)). |
| Macros, scripts, controls, and embedded code execution | ❌ | 🟡 | ❌ | No execution or activation path is provided; any opaque payload is data only. See the suite [safety audit](../../../docs/IWORK_PROGRESS_AUDIT.md#safety-and-integrity-findings). |
| Rendering, layout calculation, playback, PDF/HTML export, and automation | ❌ | ❌ | ❌ | The crate edits package data and semantic metadata; it does not render or drive Keynote ([semantic document scope](../src/document.rs)). |

## Legacy migration-host delta

`litchi-iwa` contains broad Keynote construction and mutation APIs, including fresh packages,
slide lifecycle, text boxes, shapes, images, movies, audio, charts, tables, layouts, and builds.
Those APIs are explicitly a compatibility/migration host under [ADR 0028](../../../docs/adr/0028-iwa-monolith-exit.md)
and are not evidence that the focused owner supports the same feature.

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Fresh package / presentation authoring | 🟡 | N/A | 🟡 | Present only in the [legacy creation owner](../../litchi-iwa/src/keynote/creation.rs); it is not a focused-owner write claim. |
| Slide lifecycle and dependency graph transfer | 🟡 | ✅ | 🟡 | The focused owner has partial existing-object edits; the [host slide creator](../../litchi-iwa/src/keynote/editor/slide_create.rs) is broader. Full migration and native reopen evidence remain open. |
| Rich text, shapes, groups, lines, and layouts | 🟡 | ✅ | 🟡 | The host has broad [shape](../../litchi-iwa/src/keynote/editor/slide_shapes.rs) and [paragraph-style](../../litchi-iwa/src/keynote/editor/named_paragraph_styles.rs) operations; focused support is bounded text editing. |
| Images, movies, audio, and media assets | 🟡 | ✅ | 🟡 | The host has broader [image](../../litchi-iwa/src/keynote/editor/slide_images.rs), [movie](../../litchi-iwa/src/keynote/editor/slide_movies.rs), and [audio](../../litchi-iwa/src/keynote/editor/slide_audio.rs) operations; focused support includes semantic media summaries and bounded playback edits for existing file movies and slide-owned audio, while graph creation/removal, replacement, and geometry remain host-owned. |
| Charts and table cells/formulas | 🟡 | ✅ | 🟡 | The host has broader [chart](../../litchi-iwa/src/keynote/editor/slide_charts.rs) and [table](../../litchi-iwa/src/keynote/editor/slide_tables.rs) operations; the focused owner supports selected chart/table metadata only. |
| Builds and animation mutation | 🟡 | ✅ | 🟡 | The focused owner exposes read summaries, while mutation remains a [legacy-host capability](../../litchi-iwa/src/keynote/editor/builds.rs). |
| Focused-owner replacement complete | ❌ | N/A | ❌ | Host breadth does not count as focused support; migration and native-acceptance gates remain open. |

## Evidence and current verification blockers

Evidence levels used by the repository audit are:

| Level | Meaning | Keynote posture |
|-------|---------|-----------------|
| E0 | Typed source/schema exists | Extensive |
| E1 | Synthetic unit/integration test or Litchi self-roundtrip | Extensive for focused transactions |
| E2 | Checked-in Apple-produced fixture parses and exact no-op/readback works | Present for one basic fixture |
| E3 | Litchi-mutated candidate opens in native Keynote without repair | External exploratory evidence exists for a pre-hardening candidate; current-owner promotion is pending because the app-authored source contains unproven model field 39 |
| E4 | Native save/close/reopen followed by strict Litchi reread | External/manual exploratory evidence only; the current strict owner does not admit that source, and this is not suite-wide or automated |

The checked-in [native fixture](../../../test-data/iwork/README.md#native-iwork-fixtures) contains
one basic slide and no real movie, rich table, chart, build, comment, or soundtrack collection.
The [native fixture tests](../tests/native_fixture.rs) prove path/bytes parsing, semantic
parity, limits, and exact no-op output. Most feature tests synthesize IWA graph changes or verify
Litchi self-roundtrips; they do not by themselves prove a changed package is accepted by native
Keynote.

The audited dirty snapshot was not green:

| Check | Result |
|-------|--------|
| `cargo test -p litchi-keynote --all-features --tests --no-fail-fast` | ❌ Unit tests: 153 passed/1 failed; `soundtrack_order`: 3 passed/5 failed; other selected integration targets passed. |
| ZIP preservation unit test | ❌ `selected_zip_suffix_and_central_records_allow_only_reassembly_fields` failed its central-record preservation assertion ([test](../src/package/soundtrack_order.rs#L3286), [ZIP verifier](../src/package/soundtrack_order.rs#L1931)). |
| Soundtrack-order integration | ❌ Current failures cover changed-order verification/fixture permutations; the suite itself documents synthetic media construction and the expected exact/locality checks ([fixture setup](../tests/soundtrack_order.rs#L26), [order cases](../tests/soundtrack_order.rs#L392)). |

These failures are recorded as audit evidence, not silently downgraded to unsupported features.
They must be fixed and rerun on a frozen revision before the affected write cells are promoted to
native interoperability claims. The focused physical-sort owner has source-level test/fuzz
evidence and an external exploratory probe, but no current-owner E3/E4 certification: the current
strict admission rejects the app-authored probe's unproven model field 39. A current-admitted
native source and richer checked-in fixtures are still required before the Keynote matrix can make
a native physical-sort or general authoring claim.
