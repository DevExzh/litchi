# iWork Feature-Matrix Audit

> Source audit starting at committed `d33f30f41` (2026-09-01) and including the focused
> Scientific-, Fraction-, Text-, Date & Time-, Duration-, and Custom-format owner work, the current Keynote physical `Sort Now` owner,
> and the bounded Keynote existing slide-media content/poster owner
> work, the bounded Pages body-table hidden-axis owner, and the focused merged-cell readers through 2026-09-12. This document records the rationale and cross-suite gaps behind the three
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
| Current migration topology | ✅ 64 workspace packages, 241 internal dependency declarations, 11 ordered debts `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, one migration host; no physical-sort debt/gate closure |

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
The Date & Time owner now has a selector-first package transaction for one
existing cell's bounded native date/time pattern string. The type-261 adapter requires
explicit marker `0x0008`, kind `3`, and strict fields 1/14; it admits Empty/
type-5 Date and the evidenced type-9 numeric/formula shape, and rejects plain
type-2, marker-zero, wrong/reserved, and ambiguous shapes. The metadata-only
setter performs strict preflight before a lazy Buffa view, bounds the native
date/time pattern string to 4096 bytes, and preserves COW/refcounts,
unknown/unselected bytes, scalar value, inverse, and locality. The owner checks
the bounded string envelope and admitted fields only; pattern grammar is not
validated. Computer Use recorded operation-specific native
E3/E4 app-cycle evidence for one existing type-9 cell in Numbers 14.4 (build
7043.0.93, macOS 26.5.2): the candidate opened without repair/conversion,
retained the A1 marker and B2 value/settings through native save/close/
exact-path reopen, and the exact DateTime members in the Tile/DataList files
remained byte-identical across native normalization. The strict normalized
reread of the native-resaved artifact passed as `changed=false`, with zero
touched components, `full_reparse=false`, scalar value untouched, and a
byte-identical 138,725-byte output (SHA-256
`4e489da196bac7416c8c6d827afb2a6264892d4856a5a33dc0a18c6c2d2b1b0c`). This
evidence is limited to that existing type-9 cell/pattern; it is not a broad
DateTime/package or native-byte-parity claim. The dedicated raw-ID
`NumbersEditor` Date & Time route is retired. Generic source-built or
cross-format `DataFormat::DateTime`, broad smart-field paths, and attached
Pages/Keynote table compatibility remain host-owned. The source/candidate/
native-resaved ledger is in ADR 0008.

The focused Custom owner now exposes
`litchi_numbers::Package::{table_cell_custom_format, edit_table_cell_custom_format,
apply_table_cell_custom_format}` for one existing rooted cell selected by
`SheetSelector`, sheet-scoped `TableSelector`, and checked `CellPosition`. Its
archive-free `Custom` value covers the document-scoped Number, Text, and Date &
Time registry entries admitted through `TN.DocumentArchive` field 9/message
type 222 or `TN.super` field 8 → `TSA.custom_format_list` field 12; custom
archive discriminators are 270 (Number), 271 (Text), and 272 (Date & Time).
Handwritten wire preflight precedes the private lazy Buffa view. Deterministic
source-built exact-source fixtures provide E1 evidence; native Numbers
replacement and clear cycles provide operation-specific E3/E4 evidence.
Exact source-bound transactions preserve unknown/unselected fields and
members, enforce format-list/refcount closure, keep UUIDs private while
reusing equal entries, retaining shared references, allocating replacements,
and culling only unused entries, and verify candidate reopen/readback, inverse,
and physical locality. Focused integration validation passes 37 tests: 19
owner tests, 7 native Custom Number tests, 6 native Custom Text tests, and 5
native Custom Date & Time tests; boundary verification passes 927 policy tests.
The three dedicated raw-ID `NumbersEditor` Custom conveniences are retired;
generic source-built/cross-format `DataFormat::Custom` compatibility and
private attached Pages/Keynote table adapters remain host-owned. Focused
refusals remain terminal, and this does not claim generic Custom-format
authoring or broad native parity.

The focused Duration owner now exposes
`litchi_numbers::Package::{table_cell_duration_format,
edit_table_cell_duration_format, apply_table_cell_duration_format}` for one
existing rooted cell selected by `SheetSelector`, sheet-scoped `TableSelector`,
and checked `CellPosition`. Its archive-free `Duration` value admits the three
native styles and automatic or custom ranges over Weeks through Milliseconds.
The private route strictly admits native type `268`, fields 1/7/15/16/40, BNC
value type 7/kind 4, and marker `0x0004` for a primary-only reference or
`0x0005` when a valid generic Number secondary is retained. Only a true
marker/kind/reference-free absence reads as `None`; marker-zero tuples that
retain Duration kind/reference metadata are ambiguous inherited state and fail
closed. Explicit writes emit the marker matching that secondary edge.
Source-bound set/clear/reset transactions preserve scalar and
formula/cache values, unknown and opaque bytes, refcounts, locality, and exact
inverse, with strict family/dependency/lock/budget refusals. See the [owner](../crates/litchi-numbers/src/package/table_cell_duration_format.rs#L384),
[semantic value](../crates/litchi-numbers/src/cell/data_format/duration.rs#L1),
and [focused tests](../crates/litchi-numbers/tests/table_cell_duration_format.rs#L1).

Duration evidence includes E1 synthetic/self-round-trip coverage from the strict
codec, source-built fixtures, and focused package transactions. An automated
AppleScript-driven Numbers 14.4 open/save/close/reopen probe successfully
round-tripped Litchi candidates for both admitted marker forms: the primary-only
`0x0004` shape and the retained-generic-Number-secondary `0x0005` shape, with no
reported error or repair/conversion indication. No GUI repair-dialog inspection
was performed. Strict post-native rereads were semantic no-ops and exact inverse
restoration held; the recorded candidate/native-resaved and Duration Tile/DataList
member hashes were exact across each native resave ([ADR 0008](adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-existing-cell-duration-owner-and-raw-id-host-route-retirement)). This is disposable,
operation-specific E3/E4 evidence only. The Apple-authored source/probe artifacts
are provenance, not a checked-in E2 parse/no-op fixture, and do not establish broad
native acceptance, arbitrary-producer parity, native byte parity, or package-wide
performance. The
dedicated raw-ID `NumbersEditor` Duration routes are retired, while generic
source-built/cross-format `DataFormat::Duration` and private attached
Pages/Keynote table adapters remain host-owned. Focused refusals are terminal;
no topology, debt, or deletion-gate claim follows.

Numbers persisted-sort compatibility is narrower than the semantic owner. Exact
package snapshots use the focused semantic persisted-sort transaction, so a
structural, family, lock, budget, stale-source, or locality refusal is terminal.
Only historical source-built snapshots may use the crate-private physical
bridge for field 44; it is a migration compatibility seam, not a public owner,
second implementation, or ADR 0028 debt/deletion-gate change.
The table-relocation seam follows the same split: exact snapshots use the
focused semantic transaction and treat refusal as terminal, while only
source-built snapshots may use the doc-hidden physical bridge. Neither bridge
changes the focused public owner or any migration gate.
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
| Body-table hidden rows and columns | 🟡 | ✅ | 🟡 | Public `Package::{body_table_hidden_axes, edit_body_table_hidden_axes, apply_body_table_hidden_axes}` takes a checked position or exact visible name through `BodyTableSelector` and returns archive-free `AxisIndex`/`HiddenAxes`. The strict source-backed route validates one role-qualified `TableInfoArchive` (current type 6000 or qualified legacy type 6003), one role-qualified `TableModelArchive` (current type 6001 or qualified legacy type 6000), the canonical type-6267 UID map (qualified legacy type 6200 is admitted only through its legacy route), type-4008/type-6204/type-6220 ownership, active UUID, directional extents, UID-map bounds, duplicate/malformed-wire checks, lock/pivot/dependency refusals, and finite limits before publication; on admitted existing-owner rewrites, non-user filtered/pivot markers and unknown model/info/owner fields remain preserved, while unsupported pivot/dependency topologies refuse changed edits. An absent hidden-state owner reads as empty, an empty request is an exact source no-op, and a nonempty absent-owner request returns `UnsupportedDependency` or `UnsupportedSource` without publication. Existing-owner source-bound patches are copy-on-write, candidate-reopened, component-local, preview-invalidating, and exactly invertible. This is user-hidden visibility metadata only, not cell/formula/filter/pivot/sort/topology CRUD. Focused graph, codec, identity/COW, and concurrency tests cover the E1 source/self-round-trip contract. The checked-in [`body-table-visible.pages`](../test-data/iwork/pages/body-table-visible.pages) fixture is a native Pages 14.4 visible 5-by-4 body-table baseline with a body marker; its native hidden-state envelope has one owner with empty row and column state lists. Disposable Computer Use copies were saved, closed, and reopened without repair, and the focused package save path produced byte-identical output for the visible-table/no-op operation. Because no hidden positions are present, this fixture remains native baseline/no-op evidence rather than positive hidden-axis E2 or E3/E4 mutation evidence. The registered fuzz corpus remains bounded; current fuzz verification is tracked with ADR 0008. An older exploratory Pages 14.4 generated-candidate attempt logged an NSCocoa MissingObject/TSPersistence Import document error, timed out on save/close, and never reopened; AppleScript table creation stalled without GUI inspection. |
| Body-table merged-cell geometry | 🟡 | ✅ | ❌ | Selector-first `Package::body_table_merges` returns bounded rectangles; merge/unmerge and structural table mutations remain host-owned. See the authoritative Pages matrix for the current table surface. |
| Rich text styles/lists/links/revisions | 🟡 | 🟡 | ❌ | Text and ranges are visible; general style/annotation mutation is absent |
| Media assets | ❌ | ❌ | ❌ | Image/movie/audio modules are detached option values, not package asset APIs |
| Charts/shapes/drawings | ❌ | ❌ | ❌ | No focused package owner |
| Metadata | 🟡 | ✅ | ❌ | Canonical sidecars read only |
| Fresh package document/section structural authoring | ❌ | N/A | ❌ | Detached semantic builders exist; package creation and structural editing remain legacy-host-only |
| Collaboration/comments/coauthoring | ❌ | ❌ | ❌ | No focused semantic/package API |
| Render/PDF/HTML/RTF/automation | ❌ | ❌ | ❌ | Package output is not rendering/export/action execution |

The Pages raw-ID hidden-axis host route remains compatibility debt until creation parity and
native mutation gates pass; the focused route has no fallback to that host. The registered
`pages_body_table_hidden_axes` fuzz target and checked-in valid/malformed/ownership/limit seeds
remain bounded evidence, with current fuzz verification tracked in ADR 0008. The checked-in
Pages 14.4 fixture establishes a visible 5-by-4 body-table baseline, an empty read under the
admitted native visible profile, and exact `Package` no-op verification; a changed hidden-axis
request is refused as `UnsupportedDependency` before publication. It has no user-hidden axes. Its
native hidden-state envelope has one owner with empty row and column state lists, and therefore
does not promote hidden-axis E2 or E3/E4 mutation evidence. The
older failed Pages 14.4 probe is retained as historical exploratory negative evidence only.
Attached Numbers/Keynote hidden-axis compatibility stays migration-host-only. This owner changes
no workspace topology, ordered debt, migration-host, or ADR 0028 deletion-gate status.

Primary references: [focused hidden-axis package owner](../crates/litchi-pages/src/package/body_table_hidden_axes.rs#L479), [semantic directory boundary](../crates/litchi-pages/src/document.rs#L585), [section text transaction](../crates/litchi-pages/src/package/section_text.rs#L216), [table selector](../crates/litchi-pages/src/selector.rs#L131), [hidden-axis value](../crates/litchi-pages/src/table/hidden_axes.rs), [strict hidden-state codec](../crates/litchi-iwa-protos/src/pages_hidden_state_codec.rs), [hidden-axis source tests](../crates/litchi-pages/tests/body_table_hidden_axes.rs), [hidden-axis identity/COW audit](../crates/litchi-pages/tests/body_table_hidden_axes_cow_audit.rs), [hidden-axis concurrency tests](../crates/litchi-pages/tests/body_table_hidden_axes_concurrency.rs), and [hidden-axis fuzz target](../crates/litchi-pages/fuzz/fuzz_targets/pages_body_table_hidden_axes.rs).

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
| Physical table row sorting (“Sort Now”) | 🟡 | 🟡 | 🟡 | Selector-first `Package::{execute_slide_table_sort_order,execute_slide_table_sort_order_to_rows}` over existing tables; no public row/value reader. Admission is limited to the canonical type-6001 table-model route and explicitly proven tile, data-list, header, UID, and empty pre-BNC sentinel shapes. Scalar text/number/boolean/date/duration keys use Rust lexical text ordering, `f64::total_cmp`-based deterministic ordering for numeric-like values, ordinary boolean ordering, and stable source-row ordering for duplicates. Body-relative ranges isolate headers/footers; admitted tile-row envelopes, sparse row headers, and UID mappings move together. Strict wire preflight precedes borrowed lazy Buffa views; formula/error, rich-text, comment, merge, filter/group/category/pivot/spill/conditional, hidden/non-positional, imported/provenance, non-empty stroke, cross-tile/cross-bucket, unknown mutable, and other unproven row-affine dependencies fail closed atomically. Exact-source patch/inverse, candidate reopen/readback, locality, and preview invalidation are source-level contracts. A disposable Computer Use run opened a pre-hardening candidate in Keynote 14.4, but the current strict owner rejects that app-authored source because it contains unproven model field 39; the run is external exploratory evidence rather than current-owner E3/E4 promotion. The checked-in native fixture has no table and the checked-in evidence test records hashes without launching Keynote, so native acceptance remains pending ([owner](../crates/litchi-keynote/src/package/slide_table_physical_sort.rs#L1), [semantic surface](../crates/litchi-keynote/src/slide/table/physical_sort.rs#L1), [tests](../crates/litchi-keynote/tests/slide_table_physical_sort.rs#L1), [ADR 0008](adr/0008-migration-and-verification.md#2026-09-02-amendment-keynote-physical-sort-now-evidence-status)) |
| Table cells/formulas/formats/comments/topology | ❌ | ❌ | ❌ | No public focused cell model or general structural table owner; physical row sorting above is not cell CRUD |
| Charts | 🟡 | 🟡 | 🟡 | Catalog/title/caption/axis-title/primary value-axis settings only; no data, series, type, legend, styles, or CRUD |
| Movies and soundtrack settings/order | 🟡 | 🟡 | 🟡 | Existing metadata, geometry, playback, title/caption, settings/order, and bounded media-data access/replacement; item lifecycle is tracked separately and current soundtrack-order tests remain a distinct verification concern |
| Soundtrack audio item lifecycle | 🟡 | ✅ | 🟡 | The focused `soundtrack::items` owner provides bounded read/add/insert/replace/remove with opaque source-bound handles. Focused lifecycle, closure, malformed-reference, aggregate-only, exact apply/inverse, conflict, and fuzz gates pass, and the duplicate legacy item API is retired behind a boundary ratchet. One genuine replacement candidate passed Keynote save/close/reopen without repair; operation-wide native certification, soundtrack creation, general media-asset CRUD, and every whole-monolith deletion gate remain open. |
| Existing slide-media content and movie posters | 🟡 | ✅ | 🟡 | Selector-first `Package::{slide_media_data, edit_slide_media_data, apply_slide_media_data}` reads borrowed bytes and replaces content for an existing movie/audio drawable or the poster of an existing file movie. Rooted ownership, strict SHA-1/length witnesses, shared-record propagation, finite limits, exact no-op/inverse/stale-patch behavior, candidate reread, locality, and preview invalidation are source-level contracts; native IDs, `DataInfo` records, and ZIP paths remain private. The checked-in [`media-replacement-native.key`](../test-data/iwork/keynote/media-replacement-native.key) is a 752,060-byte Keynote 14.4 Audio/Audio/File/File baseline (SHA-256 `f5763984974612f078486cb2310f408cbcb7cd06ef6adca494aae19eaf62609d`) reopened without error. The combined content/poster candidate opened without repair and survived native save/actual-close/exact-reopen with all four objects and exact replacement payloads intact. Strict reread of the native-resaved artifact passed six shared-record reads and six exact no-op writes, establishing operation-specific E4; the poster result certifies byte-preserved storage/metadata, not an arbitrary visual preview match. A bounded 256-run ASAN smoke passed with harness assertions covering all three changed paths and exact inverses; this is bounded E1 fuzz evidence, not exhaustive coverage. |
| Generic image/audio/movie bytes and asset CRUD | ❌ | ❌ | ❌ | The focused owner above replaces selected existing records only. General insertion, duplication, removal, drawable properties, resource lifecycle, and media decoding remain outside the focused package owner; detached option values do not constitute generic package media support |
| Builds/animations | 🟡 | 🟡 | ❌ | Reduced build/effect summaries; unknown effects preserved, no edit transaction |
| Shapes/groups/lines/z-order | ❌ | ❌ | ❌ | No focused semantic/package surface |
| Masters/themes/layouts/guides | ❌ | ❌ | ❌ | No focused semantic/package surface |
| Metadata | 🟡 | ✅ | ❌ | Canonical sidecars read only |
| Fresh focused-package presentation authoring | ❌ | N/A | ❌ | Detached show/slide builders exist; package builder remains in the migration host |
| Comments/collaboration/render/export | ❌ | ❌ | ❌ | No focused owner or native renderer/exporter |

Primary references: [package boundary](../crates/litchi-keynote/src/package.rs#L406), [semantic slide model](../crates/litchi-keynote/src/slide.rs#L45), [reduced chart model](../crates/litchi-keynote/src/chart.rs#L1), [table semantic boundary](../crates/litchi-keynote/src/slide/table.rs#L1).

The canonical `basic.key` fixture used by the physical-sort probe has no
table. A separate checked-in [`table-discovery.key`](../test-data/iwork/keynote/table-discovery.key)
sample contains a native Plain 5-by-4 table for the bounded migration-host
name/dimension discovery regression. It does not promote physical sorting,
Litchi table mutation, native byte parity, or broader table support.

The separate [`media-replacement-native.key`](../test-data/iwork/keynote/media-replacement-native.key)
fixture covers the focused existing slide-media owner. It is a native Keynote
14.4 Audio/Audio/File/File source with shared audio and movie records, a
752,060-byte closed artifact, and SHA-256
`f5763984974612f078486cb2310f408cbcb7cd06ef6adca494aae19eaf62609d`. It was
saved, actually closed, and reopened from the exact path without error. The
combined replacement candidate opened without repair and survived native
save/actual-close/exact-reopen with all four objects and exact replacement
payloads intact. Strict post-native reread passed six shared-record reads and
six exact no-op writes, establishing operation-specific E4; the poster result
certifies byte-preserved storage/metadata, not an arbitrary visual preview
match. A bounded 256-run ASAN smoke passed with harness assertions covering all
three changed paths and exact inverses; this is bounded E1 fuzz evidence, not
exhaustive coverage. The focused owner does not claim general media asset CRUD.

## Numbers focused owner

| Feature | Status | Read | Write | Notes |
|---|---:|---:|---:|---|
| Package/rooted workbook open | 🟡 | ✅ | 🟡 | Rooted sheets/tables, exact unchanged output, selected package transactions |
| Archive-free workbook/sheets/tables | ✅ | ✅ | N/A | Immutable semantic snapshot; auxiliary drawings are excluded |
| Scalar cells and bounded ranges | 🟡 | ✅ | 🟡 | Presence-preserving reads; admitted text/finite number/bool/date/duration/formula/clear writes |
| Formulas and cached values | 🟡 | 🟡 | 🟡 | Bounded AST and 13 authorable functions; no general parser, dependency engine, recalculation, or full registry |
| Sheet/table names and sheet order | 🟡 | ✅ | 🟡 | Rename/reorder existing objects only; not defined-name or lifecycle support |
| Sheet/table lifecycle and row/column topology | 🟡 | ✅ | ❌ | Existing sheet reorder is supported above; no focused add/delete/duplicate or row/column insert/delete transaction |
| Package merged-cell geometry/editing | 🟡 | ✅ | ❌ | Selector-first `MergeReader::table_merges` reads rooted Numbers table rectangles without semantic cell materialization; `Package::table_merges` serves already-materialized packages. Both use the shared bounded Buffa codec. The metadata reader defers cell validation; host cache handoff and merge/unmerge transactions remain migration work. See the authoritative Numbers matrix for source and native regression evidence. |
| Table settings | 🟡 | ✅ | 🟡 | Appearance, dimensions, headers/freeze/repeat, lock, title, and persisted sort configuration; physical row sorting is tracked as a separate Keynote owner |
| Historical physical persisted-sort compatibility | 🟡 | 🟡 | 🟡 | Exact Numbers package snapshots use the semantic persisted-sort owner, and structural/family/lock/budget/stale-source/locality refusals are terminal. Only historical source-built snapshots may use the crate-private physical bridge for field 44; it is not a public focused owner, a second implementation, or an ADR 0028 debt/deletion-gate change. |
| Cell controls/pop-up menus | 🟡 | 🟡 | 🟡 | Checkbox/star/slider/stepper/pop-up only, under strict graph profiles |
| Existing-cell Number/Percentage display formats | 🟡 | ✅ | 🟡 | Selector-first existing-cell Number and Percentage read/set/reset operations preserve the scalar value and unrelated bytes. Number has operation-specific native evidence; Percentage focused/synthetic coverage is green, and ADR 0008 records a disposable native Numbers open/save/close/reopen plus strict Rust semantic no-op readback. The artifact/ledger is not frozen, so Percentage E3/E4 remain pending. See the [Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md#cells-formulas-controls-and-annotations). |
| Existing-cell Currency display formats | 🟡 | ✅ | 🟡 | Selector-first existing-cell Currency read/set/clear/reset operations expose checked currency code, decimal, negative, thousands-separator, and Standard/Accounting settings while preserving the scalar value and unrelated bytes. The focused codec filter passes 4/4, the `litchi-numbers-wire` library passes 32/32, and the latest package target passes 22/22 after native BNC flag handling became shape-dependent: plain no-secondary Currency uses `0x0802`, while a Currency record carrying a secondary Number ID uses `0x0803` (`0x0802 | EXPLICIT_DECIMAL_FORMAT`, with `EXPLICIT_DECIMAL_FORMAT = 0x0001`). Standard versus Accounting does not choose this flag. ADR 0008 records operation-specific E3/E4 evidence: a clean native open, native save/close/reopen with the same scalar/settings/text, strict Rust no-op reread, exact inverse restoration, and exact candidate/native-resaved hashes. Bounded fuzz evidence remains separately scoped; no broad-format or deletion-gate claim is made. See the [Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md#cells-formulas-controls-and-annotations). |
| Existing-cell Scientific display formats | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_scientific_format`](../crates/litchi-numbers/src/package/table_cell_scientific_format.rs#L382) and archive-free [`Scientific`](../crates/litchi-numbers/src/cell/data_format/scientific.rs#L1) cover one existing cell's fixed decimal precision, native minus-sign negatives, and hidden thousands separator with explicit-to-automatic reset semantics. The package target passed 19/19, strict Buffa type-259 codec 4/4, wire library 34/34, and filtered legacy bridge 6/6; bounded codec/package sanitizer smokes each completed 100 executions. Numbers 14.4 opened, saved, closed, and reopened the precision-7 candidate without repair while preserving B2 text and B3 scalar 42, and strict reread was byte-identical. [ADR 0008](adr/0008-migration-and-verification.md#2026-09-01-amendment-numbers-existing-cell-scientific-format-native-validation-record) freezes the hashes and limits this to operation-specific E3/E4 evidence. |
| Existing-cell Fraction display formats | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_fraction_format`](../crates/litchi-numbers/src/package/table_cell_fraction_format.rs#L389) and archive-free [`Fraction`](../crates/litchi-numbers/src/cell/data_format/fraction.rs#L1) expose all nine `FractionAccuracy` denominator strategies for one existing cell, with explicit-to-inherited clearing. The native type-262 codec performs strict preflight before a lazy Buffa view; field 20 (`requires_fraction_replacement`) is preserved when absent or canonical `false` and rejects `true`. The focused package integration run passed 22/22, library codec tests 4/4, direct codec tests 3/3, and wire tests 36/36; checked-in fuzz corpora contain 44 proto seeds and 18 package seeds, and both targets completed 100-run AddressSanitizer smokes. Computer Use recorded operation-specific native E3/E4 evidence for a disposable Eighths source and Eighths→Hundredths edit; before native opening, exactly two uncompressed members differed while entry names/order and every other member payload matched. Exact source/candidate/native-resaved hashes are in ADR 0008. This does not promote all nine native UI variants, arbitrary-producer parity, native byte parity after Numbers normalization, package-wide performance, or a deletion-gate claim. |
| Existing-cell Text display formats | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_text_format`](../crates/litchi-numbers/src/package/table_cell_text_format.rs#L382) and archive-free [`Text`](../crates/litchi-numbers/src/cell/data_format.rs#L89) expose an explicit marker for one existing cell with set/clear/reset and exact inverse behavior. The owner keeps native IDs, format-table records, and wire values private; strict preflight, lazy Buffa inspection, bounded copy-on-write/refcount edits, candidate reopen/readback, scalar/unrelated-byte preservation, locality, and typed family refusal remain part of the focused contract. New explicit attachments use canonical marker `0x80`; an unchanged admitted converted-Text source with marker `0x81` and retained Number provenance is preserved exactly. Evidence is E1 plus checked-in native-fixture E2 read/no-op only; no native E3/E4 acceptance claim is made. See the [Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md#cells-formulas-controls-and-annotations). |
| Existing-cell Date & Time display formats | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_date_time_format`](../crates/litchi-numbers/src/package/table_cell_date_time_format.rs#L1) and archive-free [`DateTime`](../crates/litchi-numbers/src/cell/data_format/date_time.rs#L1) own one existing cell's bounded native date/time pattern string. Type 261 requires explicit marker `0x0008`, kind `3`, and strict fields 1/14; admitted shapes are Empty/type-5 Date and the evidenced type-9 numeric/formula shape, while plain type-2, marker-zero, wrong/reserved, and ambiguous shapes fail closed. The native date/time pattern string is bounded to 4096 bytes; the owner checks its envelope and admitted fields only, and pattern grammar is not validated. Metadata-only writes perform strict preflight before lazy Buffa and preserve COW/refcounts, unknown/unselected bytes, scalar value, inverse, and locality. Numbers 14.4 (build 7043.0.93, macOS 26.5.2) opened the candidate without repair/conversion and retained the A1 marker and B2 value/settings through native save/close/exact-path reopen; the DateTime members in `Index/Tables/Tile.iwa` and `Index/Tables/DataList-904498-2.iwa` remained byte-identical across native normalization. This is operation-specific E3/E4 app-cycle evidence for one existing type-9 cell/pattern. Strict normalized reread of the native-resaved artifact passed as `changed=false`, with zero touched components, `full_reparse=false`, scalar value untouched, and byte-identical 138,725-byte output (SHA-256 `4e489da196bac7416c8c6d827afb2a6264892d4856a5a33dc0a18c6c2d2b1b0c`). The dedicated raw-ID `NumbersEditor` Date & Time route is retired; generic source-built or cross-format `DataFormat::DateTime`, broad smart-field paths, and attached Pages/Keynote table compatibility remain host-owned. See [ADR 0008](adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-existing-cell-date--time-native-verification-record) and the current-boundary amendment in [ADR 0008](adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-date--time-raw-id-host-route-retirement). |
| Existing-cell Duration display formats | 🟡 | ✅ | 🟡 | Selector-first [`table_cell_duration_format`](../crates/litchi-numbers/src/package/table_cell_duration_format.rs#L384) and archive-free [`Duration`](../crates/litchi-numbers/src/cell/data_format/duration.rs#L1) own one existing rooted cell's bounded native Duration value. The strict type-268 route admits fields 1/7/15/16/40, styles `0/1/2`, unit bits `1/2/4/8/16/32`, BNC value type 7/kind 4, and marker `0x0004` for a primary-only reference or `0x0005` when a valid generic Number secondary is retained. Only a true marker/kind/reference-free absence reads as `None`; marker-zero tuples that retain Duration kind/reference metadata are ambiguous inherited state and fail closed. Explicit writes preserve a valid secondary edge. Source-bound set/clear/reset transactions preserve scalar and formula/cache values, unknown and opaque bytes, refcounts, exact inverse, candidate readback, and locality, with strict family/dependency/lock/budget refusals. Focused package and codec tests use deterministic source-built fixtures and provide E1 synthetic/self-round-trip evidence. An automated AppleScript-driven Numbers 14.4 open/save/close/reopen probe successfully round-tripped Litchi candidates for both admitted marker forms (`0x0004` primary-only and `0x0005` with retained generic Number secondary), with no reported error or repair/conversion indication; no GUI repair-dialog inspection was performed. Strict post-native rereads were semantic no-ops, exact inverse restoration held, and recorded candidate/native-resaved and Duration Tile/DataList member hashes matched across each native resave ([ADR 0008](adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-existing-cell-duration-owner-and-raw-id-host-route-retirement)). This is disposable, operation-specific E3/E4 evidence only; Apple-authored source/probe artifacts are provenance, not a checked-in E2 fixture, and do not establish broad native acceptance, arbitrary-producer parity, native byte parity, or package-wide performance. The dedicated raw-ID `NumbersEditor` Duration route is retired; generic source-built/cross-format `DataFormat::Duration` and private attached Pages/Keynote table adapters remain host-owned, and focused refusals are terminal. See [focused tests](../crates/litchi-numbers/tests/table_cell_duration_format.rs#L1) and the [Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md#cells-formulas-controls-and-annotations). |
| Existing-cell Custom display formats | 🟡 | ✅ | 🟡 | Selector-first [`litchi_numbers::Package::{table_cell_custom_format, edit_table_cell_custom_format, apply_table_cell_custom_format}`](../crates/litchi-numbers/src/package/table_cell_custom_format.rs#L319) exposes one existing rooted cell's archive-free `Custom` value and exact-source read/set/clear/reset/inverse transactions. The private registry is admitted through `TN.DocumentArchive` field 9/message type 222 or `TN.super` field 8 → `TSA.custom_format_list` field 12; custom archive discriminators are 270 (Number), 271 (Text), and 272 (Date & Time). Strict wire preflight precedes a private lazy Buffa view; deterministic source-built exact-source fixtures provide E1 evidence, while native Numbers replacement and clear cycles provide operation-specific E3/E4 evidence. The transaction preserves unknown/unselected fields and members, enforces format-list/refcount closure, keeps UUIDs private while reusing equal entries, retaining shared references, allocating replacements, and culling only unused entries, and verifies candidate reopen/readback, exact inverse, physical locality, content-redacted diagnostics, and typed budget/refusal paths. Focused integration validation passes 37 tests (19 owner, 7 native Custom Number, 6 native Custom Text, and 5 native Custom Date & Time), and boundary verification passes 927 policy tests. The three dedicated raw-ID `NumbersEditor` Custom conveniences are retired; generic source-built/cross-format `DataFormat::Custom` compatibility and private attached Pages/Keynote table adapters remain host-owned. See the [Numbers matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md#cells-formulas-controls-and-annotations). |
| Other generic display formats and rich styles | ❌ | ❌ | ❌ | Outside the focused Number/Percentage/Currency/Scientific/Fraction/Text/Date & Time/Duration/Custom operations above, no general package getter or formatting transaction is exposed. Rich-style families remain unsupported. The Custom owner is limited to the document-scoped existing-cell registry transaction above and does not establish generic style or package authoring. |
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
| Table cells/formulas/rich formatting/topology | Broad | Broad except physical `Sort Now` | Broad |
| Focused-owner replacement complete | No | No | No |

“Broad” here is an inventory statement, not a completeness or native-acceptance grade. The host exposes native/raw identities, includes compatibility fallbacks, and is governed by the 11 ordered debts (orders `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`) in the deletion ledger ([host policy](../crates/litchi-iwa/README.md#legacy-editing), [ADR 0028](adr/0028-iwa-monolith-exit.md)). The dedicated raw-ID `NumbersEditor` Date & Time and Duration routes are retired; generic source-built/cross-format `DataFormat::DateTime` and `DataFormat::Duration`, broad smart-field lifecycle, and attached Pages/Keynote table compatibility remain host-owned. The authoritative topology is 64 workspace packages, 238 internal dependency declarations, 227 canonical edges, 11 development-only edges, 11 ordered migration debts, and one migration host.
This API-boundary retirement does not alter the workspace topology, ordered-debt count, migration-host count, or deletion-gate status.

The focused Numbers Custom transaction retires the three dedicated legacy host
Custom conveniences. Generic/source-built `DataFormat::Custom` compatibility
and private attached Pages/Keynote table adapters remain in `litchi-iwa`, and a
focused-owner refusal is not a host fallback or a generic formatting claim.

For Keynote specifically, physical `Sort Now` is focused-owner support rather
than a broad host capability. The host may retain explicitly source-built
compatibility and other table/cell graph operations, but exact-package focused
sort refusals are terminal and do not fall back to a host physical executor.

## Evidence grades used by the authoritative matrices

| Grade | Required evidence | Current suite posture |
|---|---|---|
| E0 | Type/schema/source exists | Extensive; not support by itself |
| E1 | Synthetic unit/integration test or Litchi self-roundtrip | Extensive for focused transactions and legacy builders |
| E2 | Checked-in Apple-produced fixture parses; exact no-op/readback | Present for one basic fixture per app |
| E3 | Litchi-mutated candidate opens in the native app without repair | Current records cover operation-specific Numbers candidates for Date & Time, Currency, Scientific, Fraction, and Duration in their stated existing-cell shapes; the Duration automated probe reported no error or repair/conversion indication, but had no GUI repair-dialog inspection. This summary is non-exhaustive. The pre-hardening Keynote physical-sort candidate remains exploratory because current admission rejects model field 39 |
| E4 | Native app saves, closes, reopens; Litchi strict reread verifies semantics/locality | Current records cover operation-specific Numbers save/close/reopen and strict semantic/locality rereads for Date & Time, Currency, Scientific, Fraction, and Duration in their stated existing-cell shapes, with exact candidate/native-resaved or relevant-member hashes where recorded. This summary is non-exhaustive and is not a suite-wide CI result; the Keynote physical-sort probe remains pending |

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
