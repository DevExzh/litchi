# iWork Progress Audit

> Audit baseline: 2026-09-01 at committed `d33f30f41` (`feat(numbers): own currency cell formats`). This update evaluates the focused Scientific-, Fraction-, Text-, Date & Time-, Duration-, and Custom-format owner work, plus the bounded Pages body-table hidden-axis owner, from that baseline; it remains scoped verification rather than release certification.

## Conclusion

The iWork implementation is substantial but is **not suite-complete or release-certified**.

- Pages, Keynote, and Numbers have credible bounded semantic readers.
- The focused owner crates have strong exact-source, fail-closed transactions for selected existing objects.
- Broad fresh creation, structural editing, rich formatting, charts, shapes, and general media asset lifecycle remain largely in the legacy `litchi-iwa` migration host; Keynote now has bounded focused owners for replacing bytes and for selector-first duplication/removal of selected existing slide media records.
- Three authoritative Pages, Keynote, and Numbers matrices now report focused-owner, legacy-host, preservation, and evidence boundaries; the existing implementation review still explicitly excludes iWork.
- Native fixtures establish basic parse/no-op fidelity, but there is no suite-wide automated proof that Litchi-modified files are accepted, saved, closed, and reopened by all three native applications.
- The focused Percentage owner and codec checks are green at the synthetic/source level. The focused Currency codec filter passes 4/4, the `litchi-numbers-wire` library passes 32/32, and the package target passes 22/22 after native BNC flag handling became shape-dependent. ADR 0008 records operation-specific native Currency E3/E4 evidence with exact candidate/native-resaved hashes, inverse restoration, and strict no-op reread. Percentage remains below formal E3/E4 promotion; Currency native evidence remains limited to its stated existing-cell operation.
- The focused Scientific owner now has an archive-free value, selector-first exact-source transaction, strict Buffa type-259 route, bounded sanitizer smokes, and operation-specific native E3/E4 evidence. The package target passed 19/19, strict codec filter 4/4, wire library 34/34, and legacy bridge filter 6/6. Numbers 14.4 opened, saved, closed, and reopened the precision-7 candidate without repair; strict reread was an exact semantic no-op. This evidence remains limited to the recorded existing-cell operation.
- The focused Fraction owner now has an archive-free value, selector-first exact-source transactions, strict native type-262 preflight followed by a lazy Buffa view, and a boundary-enforced retirement of the dedicated raw-ID `NumbersEditor` convenience API. Generic source-built `DataFormat::Fraction` mutation and Pages/Keynote compatibility helpers remain in the migration host. All nine `FractionAccuracy` strategies are covered. The frozen evidence records 22/22 package integration tests, 4/4 library codec tests, 3/3 direct codec tests, and 36/36 wire tests, plus 44 proto fuzz corpus seeds and 18 package fuzz corpus seeds; both fuzz targets completed 100-run AddressSanitizer smokes. Field 20 (`requires_fraction_replacement`) is preserved when absent or canonical `false`; `true` is rejected. Computer Use established operation-specific native E3/E4 evidence for a representative Eighths-to-Hundredths edit; exact source/candidate/native-resaved hashes are recorded in ADR 0008. This does not promote all nine native UI variants, arbitrary-producer parity, native byte parity, or package-wide performance.
- The focused Text owner now provides an archive-free marker and selector-first exact-source transactions for one existing cell. Strict source preflight precedes a private lazy Buffa view; explicit set/clear/reset, copy-on-write/refcount closure, candidate reopen/readback, scalar and unrelated-byte preservation, locality, inverse, and typed family refusal are covered. New explicit attachments use canonical marker `0x80`; unchanged admitted converted-Text sources with marker `0x81` and retained Number provenance remain byte-exact. The checked-in native fixture supplies E2 read/no-op evidence, while focused tests supply E1 evidence; no native E3/E4 acceptance claim is made and numeric-to-Text conversion is unsupported.
- The focused Numbers Custom owner now provides selector-first `litchi_numbers::Package::{table_cell_custom_format, edit_table_cell_custom_format, apply_table_cell_custom_format}` transactions for one existing rooted cell selected by `SheetSelector`, sheet-scoped `TableSelector`, and checked `CellPosition`. The archive-free `Custom` value covers document-scoped Number, Text, and Date & Time entries admitted through `TN.DocumentArchive` field 9/message type 222 or `TN.super` field 8 → `TSA.custom_format_list` field 12, with custom archive discriminators 270, 271, and 272. Strict wire preflight precedes a private lazy Buffa view; deterministic source-built exact-source fixtures provide E1 evidence. Source-bound set/clear/reset edits preserve unknown/unselected fields and members, enforce format-list/refcount closure, keep UUIDs private while reusing equal entries, retaining shared references, allocating replacements, and culling only unused entries, and verify candidate reopen/readback, exact inverse, physical locality, content-redacted diagnostics, and typed budget/refusal paths. Focused validation passes 37 integration tests: 19 owner tests, 7 native Custom Number tests, 6 native Custom Text tests, and 5 native Custom Date & Time tests. Numbers 14.4 accepted the focused replacement and clear candidates without repair; the checked-in [`custom-retirement-resaved.numbers`](../test-data/iwork/numbers/custom-retirement-resaved.numbers) and [`custom-retirement-cleared.numbers`](../test-data/iwork/numbers/custom-retirement-cleared.numbers) fixtures record the native cycles. The three dedicated raw-ID `NumbersEditor` Custom conveniences are retired; generic source-built/cross-format `DataFormat::Custom` compatibility and private attached Pages/Keynote table adapters remain host-owned, and focused refusals are terminal. Boundary verification passes 927 policy tests; no topology/debt/deletion-gate change follows.
- The focused Numbers Duration owner now provides selector-first, archive-free `litchi_numbers::Package::{table_cell_duration_format, edit_table_cell_duration_format, apply_table_cell_duration_format}` transactions for one existing rooted cell. The strict native route admits type-268 fields 1/7/15/16/40, styles `0/1/2`, unit bits `1/2/4/8/16/32`, BNC value type 7/kind 4, and marker `0x0004` for a primary-only reference or `0x0005` when a valid generic Number secondary is retained. Only a true marker/kind/reference-free absence reads as `None`; marker-zero tuples that retain Duration kind/reference metadata are ambiguous inherited state and fail closed. Explicit writes preserve a valid secondary edge. Source-bound set/clear/reset edits preserve scalar and formula/cache values, unknown and opaque bytes, refcounts, candidate readback, locality, and exact inverse, with strict family/dependency/lock/budget refusals. Focused package and codec tests use deterministic source-built fixtures and provide E1 synthetic/self-round-trip evidence. An automated AppleScript-driven Numbers 14.4 open/save/close/reopen probe successfully round-tripped Litchi candidates for both admitted marker forms (`0x0004` primary-only and `0x0005` with retained generic Number secondary), with no reported error or repair/conversion indication; no GUI repair-dialog inspection was performed. Strict post-native rereads were semantic no-ops, exact inverse restoration held, and the recorded candidate/native-resaved and Duration Tile/DataList member hashes matched across each native resave. This is disposable, operation-specific E3/E4 evidence only. The Apple-authored source/probe artifacts are provenance, not a checked-in E2 fixture, and do not establish broad native acceptance, arbitrary-producer parity, native byte parity, or package-wide performance. The dedicated raw-ID `NumbersEditor` Duration route is retired, while generic source-built/cross-format `DataFormat::Duration` and private attached Pages/Keynote table adapters remain host-owned; focused refusals are terminal. No topology/debt/deletion-gate change follows.
- The focused Date & Time owner now provides a selector-first, archive-free transaction for one existing cell's bounded native date/time pattern string. Type 261 requires explicit marker `0x0008`, kind `3`, and strict fields 1/14; Empty/type-5 Date and the evidenced type-9 numeric/formula shape are admitted, while plain type-2, marker-zero, wrong/reserved, and ambiguous shapes reject. Strict preflight precedes a lazy Buffa view, the native date/time pattern string is bounded to 4096 bytes, and exact COW/refcounts, unknown/unselected bytes, scalar value, inverse, and locality are retained. The owner checks the bounded string envelope and admitted fields only; pattern grammar is not validated. A real Numbers 14.4 (build 7043.0.93, macOS 26.5.2) candidate opened without repair/conversion, retained A1 marker text and B2 `2026/09/04 12:34:56` through native save/close/exact-path reopen, and preserved the DateTime members in `Index/Tables/Tile.iwa` and `Index/Tables/DataList-904498-2.iwa` byte-for-byte across native normalization. Strict normalized reread then reported `changed=false`, `touched_components=0`, `full_reparse=false`, and `scalar_value=untouched`; its 138,725-byte output was byte-identical to the native-resaved artifact (SHA-256 `4e489da196bac7416c8c6d827afb2a6264892d4856a5a33dc0a18c6c2d2b1b0c`). This is operation-specific E3/E4 evidence for this one existing type-9 cell/pattern only; no broad DateTime/package/native-byte-parity claim follows. The dedicated raw-ID `NumbersEditor` Date & Time route is retired, while generic source-built/cross-format `DataFormat::DateTime`, broad smart-field paths, and attached Pages/Keynote table compatibility remain host-owned.
- The focused Pages body-table hidden-axis owner now provides selector-first `Package::{body_table_hidden_axes, edit_body_table_hidden_axes, apply_body_table_hidden_axes}` transactions over archive-free `table::hidden_axes::{AxisIndex, HiddenAxes}` values for existing hidden-state owners. `BodyTableSelector` accepts checked positions or exact visible names; the strict source-backed route validates one role-qualified `TableInfoArchive` (current type 6000 or qualified legacy type 6003), one role-qualified `TableModelArchive` (current type 6001 or qualified legacy type 6000), the canonical type-6267 UID map (qualified legacy type 6200 is admitted only through its legacy route), type-4008/type-6204/type-6220 ownership, active hidden-state UUID, and directional extents before projecting only user-hidden row/column positions. Bounds, duplicates, malformed/ambiguous ownership, stale UUIDs, bad UID maps, unsupported pivot/dependency graphs, locked tables, and finite wire/archive limits fail closed. On admitted existing-owner rewrites, non-user filtered/pivot markers and unknown model/info/owner fields remain preserved; unsupported pivot/dependency topologies refuse changed edits. An absent hidden-state owner reads as empty, an empty request is an exact source no-op, and a nonempty absent-owner request returns `UnsupportedDependency` or `UnsupportedSource` without publication. Existing-owner edits are copy-on-write, candidate-reopened, component-local, preview-invalidating, and exactly invertible. Focused graph, codec, identity/COW, and concurrency tests cover the E1 source/self-round-trip contract. The checked-in [`body-table-visible.pages`](../test-data/iwork/pages/body-table-visible.pages) fixture is a native Pages 14.4 visible 5-by-4 body-table baseline with a body marker; disposable Computer Use copies were saved, closed, and reopened without repair, and the focused package save path produced byte-identical output for the visible-table/no-op operation. The fixture has no user-hidden axes, so it does not provide positive hidden-axis E2 or E3/E4 mutation evidence. The registered fuzz target and corpus remain bounded evidence; current fuzz verification is tracked with ADR 0008. An older exploratory Pages 14.4 generated-candidate attempt logged an NSCocoa MissingObject/TSPersistence Import document error, timed out on save/close, and never reopened; AppleScript table creation stalled without GUI inspection. The raw-ID `PagesEditor` route remains migration-host compatibility, with no fallback from the focused route; private attached Numbers/Keynote compatibility remains host-owned, and this slice changes no topology/debt/deletion gate.
- Numbers persisted-sort and table-relocation compatibility remain narrower than the focused semantic owners. Exact package snapshots use the focused semantic transaction and any structural, family, lock, budget, stale-source, or locality refusal is terminal. Only historical source-built snapshots may use the respective crate-private physical bridge (field 44 for persisted sort or physical table relocation); these bridges are migration compatibility seams, not public owners or second implementations, and have no ADR 0028 debt/deletion-gate impact.
- The focused Keynote physical `Sort Now` owner now has a selector-first, source-bound transaction with strict wire admission, lazy Buffa inspection, bounded row-envelope/UID movement, exact-source inverse/locality checks, and atomic refusal of unproven row-affine state. The 26-case package suite and bounded fuzz targets are source-level evidence only. A disposable Computer Use run opened a candidate produced before the latest strict model-field admission, but the current owner rejects the app-authored source because model field 39 identifies an unowned conditional-style CalculationEngine dependency graph; that run is exploratory evidence, not current-owner E3/E4 promotion. Native acceptance remains pending.
- The focused Keynote `soundtrack::items` semantic/package owner now handles bounded read/add/insert/replace/remove operations while keeping native IDs, media topology, package bytes, and wire values private. Its focused lifecycle, shared-occurrence, malformed-reference, aggregate-only, exact apply/inverse, conflict, and fuzz gates pass, so the duplicate legacy item API, eager wire helper, example, and legacy-only tests are retired. One genuine replacement candidate passed Keynote save/close/reopen without repair; this is representative replacement evidence, not operation-wide native certification, soundtrack creation, broad media CRUD, or a monolith-exit claim.
- The focused Keynote slide-media owner now provides selector-first `Package::{slide_media_data, edit_slide_media_data, apply_slide_media_data}` access for existing movie/audio content and file-movie posters. It validates the rooted drawable/media closure, strict metadata digest/length witnesses, shared-record ownership, finite limits, exact no-op/inverse/conflict behavior, candidate readback, and preview invalidation while keeping native IDs, `DataInfo` records, and ZIP paths private. The three narrow host wrappers `replace_slide_movie_data`, `replace_slide_movie_poster`, and `replace_slide_audio_data` are removed; generic host `replace_media` and broad media creation/duplication/removal/property lifecycle remain. The combined candidate opened in Keynote 14.4 without repair and survived native save/actual-close/exact-reopen with all four objects and three replacement payloads intact. Strict reread of the native-resaved artifact passed six shared-record reads and six exact no-op writes, establishing operation-specific E4; the arbitrary poster is certified as byte-preserved storage, not a visual preview match. Focused validation passes 23 owner/budget cases, neutral metadata validation passes 33 module/integration cases, and boundary verification passes 933 policy units across 64 crates, 238 internal edges, and 11 explicit debt items. A bounded 256-run ASAN smoke passed with harness assertions covering all three changed paths and exact inverses; this is bounded E1 fuzz evidence, not exhaustive coverage. Commit `f57f5e74c` passed all normal workspace hooks: formatting, manifest sorting, lint, library/integration tests, and doctests.

The detailed bounded coverage assessment is in [IWORK_FEATURE_MATRIX_AUDIT.md](IWORK_FEATURE_MATRIX_AUDIT.md).

## Keynote media lifecycle follow-up (2026-09-06)

Permanent native-authored Keynote 14.4 oracles record expected lifecycle
behavior. [`media-lifecycle-duplicate-native.key`](../test-data/iwork/keynote/media-lifecycle-duplicate-native.key)
is 752,730 bytes with SHA-256
`e9a9d2779749252861a9fc592a3af5020b92ca3cb9adf14a03d155606bb7c276`; it
contains five media objects, a duplicate offset by `(10,10)`, and an appended
caption. [`media-lifecycle-remove-native.key`](../test-data/iwork/keynote/media-lifecycle-remove-native.key)
is 697,237 bytes with SHA-256
`88813924afc1566dcffa92a8a59ab33e65b919bcef83dc3352cab3782bf95d13`; its
final movie removal leaves two audio objects and culls the movie/poster assets.
Both artifacts were saved, actually closed, and reopened from their exact
paths in Keynote 14.4.

These are native-authored expected-behavior oracles only; no Litchi-mutated
candidate or focused lifecycle E3/E4 claim follows. The neutral prerequisite
implementation is now landed: the core clone path preserves raw headers and
source bytes, stages source-atomic remaps, sorts exact remap scratch
deterministically, and requires explicit consistent remaps for known self-references. The neutral map decoder
uses a bounded two-pass traversal, one temporary allocation for nonempty O(n log n)
ordering, lazy Buffa views with canonical `int32` fields, and reference-parity
checks. The metadata codec atomically removes final owners and their `DataInfo`
records. The native baseline test removes four owner references and two
`DataInfo` records. Mapped or ambiguous records and surviving empty or
versioned component references refuse removal before publication.

The existing focused media-replacement closure delegates map decoding to the
neutral reader and has an unknown-extension read-preservation regression. Core
validation passes 6 cases, neutral map validation passes 22 unit plus 18
integration cases (40 total), Keynote validation passes 22 replacement plus 4
native-oracle integration cases (26 total), and boundary verification passes
936 units. Workspace strict linting passes, and both clone/remap and metadata fuzz
targets pass 256 AddressSanitizer runs. The full boundary scanner passes (64 packages, 238 internal dependency
declarations, 11 explicit debts). Commit `55499d9c5` passed normal formatting,
manifest sorting, strict workspace lint, all-feature workspace library and
integration tests, and documentation tests. At this Sep 6 snapshot the
selector-level lifecycle owner and further host-route retirement remained
pending; the Sep 7 follow-up below supersedes that status.

## Keynote audio lifecycle native-oracle follow-up (2026-09-06)

The native matrix now includes dedicated audio duplicate/removal oracles.
[`media-lifecycle-audio-duplicate-native.key`](../test-data/iwork/keynote/media-lifecycle-audio-duplicate-native.key)
is 752,241 bytes with SHA-256
`4613f7a275388a1d053407f849a3845de2789acdfe5c36139892d20a17e57c3c`; it
contains three audio objects and two captioned movie objects, retaining the
WAV, movie-content, and poster records. [`media-lifecycle-audio-remove-native.key`](../test-data/iwork/keynote/media-lifecycle-audio-remove-native.key)
is 559,085 bytes with SHA-256
`052e6389af8719e2e6ffadf2d5fbae0c5275983b94c5ce9efd811f59cf1c1bfb`; it
contains no audio objects, retains the two captioned movies, removes WAV
`DataInfo` 9075, and retains movie-content/poster `DataInfo` 9085 and 9086.
Both permanent artifacts were authored, saved, actually closed, and reopened
from their exact paths in Keynote 14.4.

The first-audio-removal intermediate snapshot retained the shared WAV through
the remaining audio occurrence. It is a temporary, uncommitted 751,692-byte
diagnostic artifact with SHA-256
`7e39344f54222786ae3241ef4f14ebcefdfc96bb0c54a7c8401348a71318f36e`; it did
not receive the native close/reopen gate. These are native-authored
expected-behavior oracles only, with no Litchi-mutated candidate or focused
lifecycle E3/E4 claim. At this Sep 6 snapshot the selector-level lifecycle
owner and its wire, clone-payload, and metadata-adapter pieces remained under
implementation; the Sep 7 owner and E4 section below supersedes that pending
status.

## Keynote focused media lifecycle owner and native E4 follow-up (2026-09-07)

The focused `litchi-keynote` owner now provides
`Package::{duplicate_slide_media, remove_slide_media}` plus typed movie/audio
aliases. `SlideSelector` and source-order `MovieSelector` are the public
selection surface; native IDs, UUIDs, component paths, `DataInfo` keys, and ZIP
members remain private. `SlideMediaLifecyclePatch` is authorized to its exact
source, supports `inverse()` and `is_noop()`, and is replayed through
`Package::apply_slide_media_lifecycle` only after bounded candidate readback.
The owner clones or removes the selected slide/build/build-chunk closure,
preserves shared data until the final owner, and uses one `LifecycleBudget`
across graph, clone-payload, metadata, lazy wire, archive, ZIP, and readback
work. The five-message
[`KNMediaLifecycleArchive.proto`](../crates/litchi-iwa-protos/src/buffa-projections/KNMediaLifecycleArchive.proto)
projection is 957 bytes and keeps borrowed snapshots and source-preserving
rewrites out of generated repeated views.

Computer Use verified six operation-specific Keynote 14.4 native E4 profiles.
The source candidates were under
`/private/tmp/litchi-media-lifecycle-20260906r/candidates`, and the
native-saved receipts were under
`/private/tmp/litchi-media-lifecycle-20260906r/native-saved`. Each temporary
source candidate was opened at its exact path, saved with
Cmd-S, actually closed until the theme chooser appeared, reopened at its exact
path, checked for expected counts, text, and captions without alerts, and
closed. The native-saved receipts are temporary and were not copied into the
permanent fixture set:

| Receipt | Expected profile | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| `focused-native-duplicate-movie.key` | 3 movies, 2 audio | 752,707 | `bb2f3d0318a311ba0cfdf4fa1d8b88488f8e0172ba02fda73d689ea15ae0344a` |
| `focused-native-duplicate-audio.key` | 2 movies, 3 audio | 752,189 | `39eb5ba491acbdf2bc7295a2c549520ed538e1ca73f5b5e109b0cb74d587be38` |
| `focused-native-remove-movie-shared.key` | 1 movie, 2 audio | 745,166 | `16d1241c50fb4ad093e10c2befebac1303a9442e4efda22ef6355c6d2b933c91` |
| `focused-native-remove-audio-shared.key` | 2 movies, 1 audio | 751,698 | `452af453a6b453aa454319e1033a1b5be0389a8954c245d2b789bee310c97170` |
| `focused-native-remove-movie-final.key` | 0 movies, 2 audio | 698,799 | `9e3b40157610ded1f618d64112c4e964c6f03089a5ceabb9fb7899eca5320b99` |
| `focused-native-remove-audio-final.key` | 2 movies, 0 audio | 559,060 | `1477986335aca0e093b404a259bba9d866d7083a101aeeb1b2b1509982f03ece` |

The lifecycle integration target passes 21 cases: the original 18 plus three
header-reference refusal regressions. With
`LITCHI_KEYNOTE_MEDIA_LIFECYCLE_NATIVE_SAVED_DIR` enabled, the direct
`native_saved_candidates_are_read_back_without_rewriting_them` integration test
strictly rereads all six native-saved candidates without rewriting them. Four
focused native lifecycle tests pass for the six exports, establishing
operation-specific lifecycle E4 only.

The neutral identity codec passes 45 unit cases, including the versioned
component accounting regression; the new lazy lifecycle codec passes 10 cases;
and strict protobuf and Keynote Clippy pass. The private lifecycle unit slice
passes 11 cases, while existing replacement validation passes 28 cases (22
replacement, 2 budget, and 4 native-oracle cases). The hardening fixes cover
the actual `ShapeInfoArchive` reference edge, exact versioned-component work
charging, source-relative core-header precharge with atomic byte/event
accounting, and serialization precharge before candidate allocation.
Regenerating all six focused candidate inputs produces bytes exactly equal to
the native-verified originals. Fresh-export equality and strict native-saved
readback both pass. The bounded lifecycle AddressSanitizer campaign passed 256 runs with no
findings (244 MiB peak RSS); boundary verification passes 942 unit tests.
Implementation commit `182f31566` passed all normal repository hooks:
formatting, manifest sorting, workspace library Clippy, all-feature
library/integration tests, and doctests; no host
lifecycle route was retired.

The final boundary scan passed for 64 workspace packages, 238 internal
dependency declarations, and 11 explicit debt items. Cleanup ran `cargo clean`
on the isolated sanitizer target (2,203 files, 2.1 GiB), then removed the 94
remaining turn-temporary files, generated fuzz artifacts, and boundary
bytecode caches. The native-saved files above are historical receipts; the
two newly authored audio oracle fixtures remain checked in.

## Current maturity

| Layer | What is solid | What still lacks | Assessment |
|---|---|---|---|
| Root `litchi::iwork` facade | Immutable, bounded Pages/Keynote/Numbers projection; host-free dependency path | Read-only and text-oriented; no package/edit, metadata, media, chart, or formula editing/recalculation surface | 🟡 focused read facade |
| `litchi-pages` | ZIP/directory semantic read; exact retained-package output; selected text, layout, footnote, header/footer, table-property, and existing-owner user-hidden-axis transactions | Fresh package creation, section/table structural CRUD, table cells/formulas, rich formatting, drawings, charts, media, collaboration, export, nonempty absent-owner hidden-axis requests, and successful native hidden-axis validation | 🟡 bounded reader/editor |
| `litchi-keynote` | Show/slide read; existing-slide state, text, notes, settings, background, transition, selected chart/movie/table properties, bounded existing slide-media content/poster replacement, selector-first existing slide-media duplication/removal with exact-source patches and native E4 receipts, native-authored media lifecycle oracle fixtures, a source-level physical table-sort owner, and bounded soundtrack-item read/add/insert/replace/remove | Slide creation/duplication, public table row/value reads, broader generic asset lifecycle, native-certified physical sort or operation-wide soundtrack-item lifecycle, soundtrack creation, full table data, chart data/series/CRUD, generic media asset insertion/duplication/removal/CRUD beyond the focused owner, full build/animation editing, shapes/groups, masters/themes, collaboration, rendering/export | 🟡 bounded reader/editor |
| `litchi-numbers` | Rooted workbook read; selected scalar cells, formulas, controls, table settings, names/order, comments/replies, and bounded existing-cell Number/Percentage/Currency/Scientific/Fraction/Text/Date & Time/Duration/Custom display-format transactions | Sheet/table lifecycle, row/column topology, full formula engine, other generic formats/rich styles, generic Custom-format authoring, package merge editing, charts/media/drawables, filters/pivots/categories, print/page setup and workbook export | 🟡 bounded reader/editor |
| `litchi-iwa` | Broad source-free authoring and native mutation across all three applications; extensive tests/examples | It is explicitly a compatibility/migration host, exposes native/raw seams, and still owns most rich/structural authoring | 🟡 broad but non-canonical |
| Shared IWA crates | Bounded Snappy/wire/archive parsing, package preservation, detection, focused codecs, exact artifacts, COW state, and archive-owned durable publication | Concrete-owner index adoption, aggregate graph/memory budgets, directory write parity, durable patch serialization/history, encryption/signatures | 🟡 mature substrate with open boundaries |

This assessment deliberately does not infer semantic support from generated protobuf availability or raw public declaration counts.

## Architecture and migration

ADR 0028 defines `litchi-iwa` as the sole legacy migration host and makes deletion conditional on closing every recorded edge and passing native reopen gates ([ADR 0028](adr/0028-iwa-monolith-exit.md#deletion-gate)). The authoritative current topology is **64 workspace packages, 238 internal dependency declarations, 227 canonical edges, 11 development-only edges, and 11 ordered migration debts**, with one migration host. The debt orders are `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]` ([boundary ledger](../tools/crate_boundaries.json)); these are ordered host-exit conditions, not a count of all IWA-related crates.

Positive progress:

- The root facade depends on the three concrete owners rather than `litchi-iwa` ([feature wiring](../crates/litchi/Cargo.toml#L60), [facade](../crates/litchi/src/iwork/mod.rs#L1)).
- Recent waves moved narrow Pages table/text/hidden-axis owners, Numbers cells/settings/comments/display-format owners, and Keynote chart/movie/table owners into focused crates.
- The Keynote lifecycle wave now owns selector-first existing slide-media duplication/removal, exact-source inverse patches, shared-data final-owner reclamation, and operation-wide resource accounting; generic host lifecycle routes remain until their own gates close.
- `litchi-iwa-index` is a neutral leaf consistent with ADR 0029.
- No `TODO`, `FIXME`, `todo!`, or `unimplemented!` stubs were found across the iWork crates; incomplete work is represented by typed refusal paths and migration debt.

Open migration facts:

- The concrete app crates do not yet consume the neutral index directly; debt 008 remains.
- Generated protobuf modules remain a public compatibility surface, while production semantic coverage uses a much narrower set of focused codecs.
- The host still joins focused values to legacy objects through native IDs and still coordinates package reassembly and broad mutation.
- Focused owners couple exact-source authorization/reassembly to process-local logical COW patches. The logical patches are archive-neutral and do not contain full package bytes; the archive now provides durable owner `Package::save` publication, but neither layer provides durable, serializable, composable, mergeable patch history.

## Evidence and release posture

| Area | Evidence present | Gap before a strong completion claim |
|---|---|---|
| Native fixtures | One checked-in basic native fixture per app plus package-directory fixtures and a bounded Keynote slide-media source; native provenance and hashes are documented in [`test-data/iwork`](../test-data/iwork/README.md) | The basic fixtures remain intentionally narrow, and the Keynote media source covers one existing Audio/Audio/File/File graph only. Rich tables, formulas, charts, comments, builds, themes, and malformed/edge producer variants are not represented broadly; the changed media candidate has operation-specific E4 evidence from native open/save/close/reopen plus strict post-native reread, while the remaining native matrix is not suite-wide |
| Keynote media lifecycle native oracles | The permanent duplicate/removal fixtures cover both movie-final-removal and audio-final-removal expectations in Keynote 14.4, including shared WAV retention while another audio occurrence remains and final WAV `DataInfo` reclamation when no audio object remains; both new audio artifacts were saved, actually closed, and reopened from their exact paths | These are native-authored expected-behavior oracles, not Litchi-mutated candidates. The intermediate first-audio-removal snapshot has no close/reopen gate. The owner-pending wording in this historical Sep 6 row is superseded by the Sep 7 owner and E4 rows below |
| Keynote selector-first lifecycle E4 receipts | The focused owner now provides exact-source selector transactions with graph closure, shared-data retention/final-owner GC, the shared `LifecycleBudget`, and a five-message, 957-byte lazy projection. The lifecycle integration target passes 21 cases (18 existing plus 3 header-reference refusal regressions); four native-focused tests cover six temporary native-saved outputs, and strict direct reread passes all six without rewriting them | This is operation-specific lifecycle E4 evidence. Neutral identity passes 45 units including versioned accounting, the new lazy codec passes 10, and strict protobuf and Keynote Clippy pass. The private lifecycle unit slice passes 11 and existing replacement validation passes 28 (22 replacement, 2 budget, 4 native-oracle); the actual `ShapeInfoArchive` reference fix, exact versioned work charge, source-relative core-header precharge, atomic bytes/events, and serialization precharge are recorded. Regenerated candidate bytes equal the native-verified originals; bounded lifecycle ASAN passed 256 runs and boundary units passed 942; implementation commit `182f31566` passed all normal workspace hooks |
| Package fidelity | Native parse, semantic parity, exact no-op output, inverse/locality tests, and substantial synthetic transaction suites | Self-roundtrip does not prove native acceptance of changed output; native verification for generated packages is explicitly external/opt-in |
| Pages body-table hidden axes | Focused selector/value/transaction, identity/COW, and concurrent snapshot coverage exercise existing role-qualified type-6000/type-6001 owners, canonical type-6267/type-4008 ownership, UID/extents, unknown-field/member preservation, exact patches, locality, preview invalidation, malformed/budget/lock/pivot/dependency refusals, absent-owner empty reads/no-ops, and nonempty absent-owner refusal; current Pages and protobuf suites plus scoped all-target strict Clippy are green, and the registered `pages_body_table_hidden_axes` fuzz target has bounded descriptors with checked-in valid/malformed/ownership/limit seeds | The checked-in [`body-table-visible.pages`](../test-data/iwork/pages/body-table-visible.pages) is a native Pages 14.4 visible 5-by-4 baseline with a body marker; its native hidden-state envelope has one owner with empty row and column state lists. Disposable UI copies were saved, closed, and reopened without repair, and the focused package exact `Package` no-op verification succeeds. It has no user-hidden axes, so there is no positive hidden-axis E2 or E3/E4 mutation evidence. Workspace/all-features lint now passes under the unchanged strict policy; current native-profile admission remains separately scoped. The older generated-candidate Pages 14.4 attempt remains historical negative evidence; current fuzz verification is tracked in ADR 0008. The raw-ID Pages host route remains until creation parity and native mutation gates pass, with no focused-route fallback |
| CI | Ubuntu workspace all-feature lib/integration/doc tests; all targets compile | iWork is absent from macOS/Windows test matrices; no leaf-feature isolation, native-app runner, fuzz/sanitizer, Miri, or coverage gate; iWork fixture paths do not trigger Rust CI |
| Ignored evidence | Four Numbers native-oracle tests exercise formula-rich cases | They depend on a private `/private/tmp` oracle and normal CI cannot reproduce them |
| Publish | Package manifests and protobuf provenance guards exist | Publish workflow does not install `protoc`; no clean unpacked-`.crate` test proves repository-external fixture references work |
| Documentation | Three authoritative app matrices, root-index links, extensive ADR wave log, and crate rustdocs | Root README/rustdocs understate focused writes while parts of `litchi-iwa` docs overstate canonical support; the implementation review still excludes iWork |

The prior [format implementation review](FORMAT_IMPLEMENTATION_REVIEW.md#scope-and-rubric) expressly excludes Pages, Keynote, Numbers, and shared IWA, so its pass result cannot certify this suite.

## Current-worktree verification

These results describe the focused pre-commit slice snapshot, not a release artifact. A current
fresh locked all-features rerun passed 460 Pages library/integration tests across 25 binaries and
802 protocol library/integration tests (764 library and 38 integration); 170 generated protocol
doctests were ignored. Clippy, boundary, migration-host, sibling, and sanitizer status remain separate gates
and are not inferred here.
The Scientific rows were rerun for this slice, and the Fraction, Date & Time, and Duration rows use the
focused source/build/test evidence for this update. The Duration row also records disposable,
operation-specific E3/E4 evidence from an automated AppleScript-driven Numbers 14.4
open/save/close/reopen probe for both admitted marker forms; it reported no error or
repair/conversion indication, but no GUI repair-dialog inspection was performed. The Percentage and Currency rows retain
their previously recorded results. The Pages hidden-axis row below records E1 source-backed and
concurrency coverage plus the checked-in native visible-table baseline. The fixture's Pages 14.4
5-by-4 body table, body marker, disposable-copy save/close/reopen without repair, exact `Package`
no-op, and the native visible-profile empty read do not demonstrate user-hidden axes or a changed
hidden-axis native cycle. A separate disposable profile copy was edited in Pages with
`Native profile read back` in A2, saved, closed, and reopened with the text intact and no repair
prompt; its post-close SHA-256 is recorded in the fixture README. This is native authoring/save/reopen
evidence only. The native hidden-state envelope is ownerful with one state, but both row and column
state lists are empty; no hidden positions are present for the focused route to project. The older
generated-candidate attempt remains negative exploratory history;
current fuzz verification is tracked with ADR 0008. The Keynote physical-sort checks below are source-level and
exploratory-native diagnostics; they do not certify current-owner native interoperability:

| Check | Result |
|---|---|
| `python3 tools/check_crate_boundaries.py` | ❌ Three errors: untracked `crates/litchi-iwa/src/pages/editor/tables/lock.rs` restores the retired `body_table_lock_state` and `set_body_table_lock_state` host surface |
| `cargo test -p litchi-numbers --test table_cell_percentage_format --no-fail-fast` | ✅ 19/19 focused Percentage owner tests passed, including selector, family-boundary, malformed-input, locality, inverse, concurrency, and budget cases |
| `cargo test -p litchi-iwa-protos percentage_format --no-fail-fast` | ✅ 4/4 focused Percentage codec tests passed, covering native domains, canonical framing, unknown preservation, and Number/Percentage separation |
| `cargo test -p litchi-numbers --test table_cell_currency_format --no-fail-fast --locked --offline` | ✅ 22/22 focused Currency owner tests passed after native BNC flag handling became shape-dependent, including refcount and secondary-reference cases |
| `cargo test -p litchi-iwa-protos currency_format --lib --no-fail-fast` | ✅ 4/4 focused Currency codec tests passed, covering native domains, canonical framing, unknown preservation, and lazy-view separation |
| `cargo test -p litchi-numbers-wire --lib` | ✅ 32/32 focused Numbers wire library tests passed |
| Focused Currency fuzz targets | 🟡 Source-visible with bounded synthetic/adversarial corpus coverage; sanitizer smoke evidence remains separately scoped and is not used to broaden the native claim |
| Numbers Currency native validation | ✅ Operation-specific E3/E4 evidence: Numbers 14.4 opened the Rust candidate without repair/recovery/conversion, preserved B2/B3 text/scalar semantics and requested Currency settings through save/close/reopen, and strict Rust no-op reread plus exact inverse restoration matched the recorded hashes in ADR 0008 |
| Scientific focused owner/build/test/fuzz/native validation | ✅ Package 19/19, strict Buffa codec 4/4, wire 34/34, and filtered legacy bridge 6/6 passed. Targeted Scientific boundary audits reported zero violations. Sanitizer-backed codec/package smokes completed 100 executions from 33/10 corpus files. Numbers 14.4 opened the precision-7 Rust candidate without repair, retained text/scalar/format through save-close-reopen, and strict reread produced an exact no-op; ADR 0008 freezes source, candidate, native-resaved, inverse, and locality hashes. |
| Fraction package integration tests | ✅ 22/22 focused existing-cell Fraction owner tests passed, covering selector-first exact-source transactions, all nine accuracies, set/clear/reset, inverse/no-op, COW/refcounts, preservation/locality, budgets, conflicts, and fail-closed refusal paths |
| Fraction library codec tests | ✅ 4/4 strict native type-262 codec tests passed, covering canonical framing, all nine accuracies, lazy Buffa parity, unknown preservation, and strict field-20 handling |
| Fraction direct codec tests | ✅ 3/3 direct codec tests passed, including absent/canonical-`false` field-20 preservation and canonical-`true` rejection |
| Fraction wire tests | ✅ 36/36 focused Numbers wire tests passed |
| Fraction fuzz targets | ✅ Checked-in corpora contain 44 proto seeds and 18 package seeds; the protocol and package targets each completed a 100-run AddressSanitizer smoke from isolated corpus/build/artifact roots with no crash or invariant failure. These bounded runs are not exhaustive input coverage. |
| Numbers Fraction native validation | ✅ Operation-specific E3/E4 evidence: Computer Use created B3 Fraction/Eighths from a disposable copy, Numbers opened the Rust Eighths→Hundredths candidate without repair/recovery/conversion, preserved B2 text and B3 Actual `42`, then saved/closed/reopened a duplicate with the same Fraction/Hundredths readback and no prompt. Before native opening, exactly two uncompressed members differed while names/order and every other member payload matched. Rust no-op/inverse and native-resaved name-selector rereads were exact; source/candidate/native-resaved hashes and byte sizes are recorded in ADR 0008. This represents Eighths/Hundredths only, not all nine native UI variants or arbitrary-producer/native-byte parity after Numbers normalization. |
| Numbers Date & Time native validation | ✅ Operation-specific E3/E4 evidence for one existing type-9 cell/pattern: Numbers 14.4 (build 7043.0.93, macOS 26.5.2) opened the Litchi candidate without repair/recovery/conversion, retained A1 `Litchi DateTime producer — 東京😀`, B2 `2026/09/04 12:34:56`, Actual `9/4/2026 12:34:56 PM`, and inspector Date & Time/date `2026/01/05`/time `19:08:09` through save/close/exact-path reopen. Native normalization changed document, calculation, stylesheet, metadata, view-state, and preview entries, but the exact DateTime members in `Index/Tables/Tile.iwa` and `Index/Tables/DataList-904498-2.iwa` stayed byte-identical to the candidate. Strict normalized reread reported `changed=false`, `touched_components=0`, `full_reparse=false`, and `scalar_value=untouched`; its 138,725-byte output was `cmp=0` with the native-resaved artifact (SHA-256 `4e489da196bac7416c8c6d827afb2a6264892d4856a5a33dc0a18c6c2d2b1b0c`). This remains scoped to that one existing type-9 cell/pattern; no broad DateTime/package/native-byte-parity claim follows. The dedicated raw-ID `NumbersEditor` Date & Time route is retired as a production API boundary; generic source-built/cross-format `DataFormat::DateTime`, broad smart-field lifecycle, and attached Pages/Keynote table compatibility remain host-owned. |
| Pages body-table hidden-axis focused owner | 🟡 Selector-first `Package::{body_table_hidden_axes, edit_body_table_hidden_axes, apply_body_table_hidden_axes}` over archive-free `AxisIndex`/`HiddenAxes` for existing hidden-state owners; strict source-backed ownership/UID/extent validation, unknown/native preservation, exact no-op and apply/inverse patches, locality, preview invalidation, bounded limits, concurrent COW snapshots, absent-owner empty reads/no-ops, and nonempty absent-owner refusal are covered by focused source tests. Admitted existing filtered/pivot markers are preserved, while unsupported pivot/dependency topologies refuse changed edits. Current Pages/protobuf focused suites and scoped all-target strict Clippy are green. The checked-in native fixture is a visible Pages 14.4 5-by-4 body-table baseline with a body marker; its native hidden-state envelope has one owner with empty row and column state lists, and the native profile reads `HiddenAxes::empty()` while accepting an exact empty no-op. Disposable UI copies were saved, closed, and reopened without repair, and exact `Package` no-op verification succeeds. A changed hidden-axis edit is refused as `UnsupportedDependency` before publication. It has no user-hidden axes, so no positive hidden-axis E2 or E3/E4 mutation evidence is claimed. The former workspace lint diagnostics were resolved under the unchanged strict policy; current workspace verification remains tracked separately. The older generated-candidate Pages 14.4 attempt remains historical negative evidence; current fuzz verification is tracked in ADR 0008. The raw-ID Pages host route remains until creation parity and native mutation gates pass, with no focused-route fallback; attached Numbers/Keynote compatibility remains host-owned and global topology/deletion gates are unchanged. |
| Keynote physical `Sort Now` package tests | ✅ 26/26 focused source-level tests passed, covering selector/range semantics, stable scalar ordering, row/header/UID movement, exact no-op/inverse/conflict, unknown/member preservation, aggregate-only metadata compatibility, optional format-route validation, noncanonical IWA framing refusal, strict dependency refusal, and budget paths. This does not provide a public row/value reader or native-app evidence. |
| Keynote physical `Sort Now` native probe | 🟡 A disposable pre-hardening candidate opened and round-tripped in Keynote 14.4, but the current strict owner rejects the app-authored source because model field 39 identifies an unowned conditional-style CalculationEngine dependency graph. The run and hashes are recorded in ADR 0008 as exploratory evidence only; no current-owner E3/E4 promotion follows. The canonical `basic.key` fixture has no table; the separate `table-discovery.key` fixture is reserved for the bounded migration-host name/dimension discovery regression and does not certify physical sorting. |
| Keynote soundtrack-item lifecycle | ✅ Focused owner tests cover bounded read/add/insert/replace/remove, shared occurrences, malformed references, aggregate-only preservation, atomic errors, exact apply/inverse, and stale-source conflicts; the focused fuzz target and corpus are present. The duplicate legacy item API is retired behind a boundary ratchet. One genuine replacement candidate passed Keynote save/close/reopen without repair; operation-wide native certification, soundtrack creation, general asset CRUD, and all whole-monolith deletion gates remain open. |
| Keynote existing slide-media content/poster owner | 🟡 Selector-first `Package::{slide_media_data, edit_slide_media_data, apply_slide_media_data}` reads borrowed content/poster bytes and replaces one selected existing record after rooted graph, owner, digest/length, media-family, and finite-limit checks. Shared records remain shared; exact no-op/inverse/stale-patch, candidate reread, locality, unrelated-member preservation, and preview invalidation are bounded source-level contracts. The native [`media-replacement-native.key`](../test-data/iwork/keynote/media-replacement-native.key) source is 752,060 bytes with SHA-256 `f5763984974612f078486cb2310f408cbcb7cd06ef6adca494aae19eaf62609d`; it was saved, actually closed, and reopened in Keynote 14.4 with Audio/Audio/File/File source order and shared movie geometry/caption/playback intact. The combined content/poster candidate opened without repair and survived native save/actual-close/exact-reopen with all four objects and exact replacement payloads intact. Strict reread of the native-resaved artifact passed six shared-record reads and six exact no-op writes, establishing operation-specific E4; the poster result is byte-preserved storage evidence, not a visual-preview match. The three narrow host wrappers are removed, while generic `replace_media` and broad asset lifecycle remain host-owned. |
| Keynote media lifecycle native-oracle pass | 🟡 Permanent Keynote 14.4 audio duplicate/removal fixtures record three-audio/two-movie retention, final WAV removal, and movie/poster retention; both permanent artifacts passed native save, actual close, and exact-path reopen. The earlier movie duplicate/removal fixtures remain recorded above | These artifacts are native-authored expected-behavior evidence only. No Litchi-mutated lifecycle candidate, focused lifecycle E3/E4, candidate test count, or host-route retirement follows; the first-audio-removal intermediate snapshot is temporary and ungated |
| Keynote selector-first media lifecycle owner and native E4 | 🟡 `Package::{duplicate_slide_media, remove_slide_media}` plus typed movie/audio aliases now keep selectors, exact-source patches, graph closure, shared-data retention/final-owner GC, and one operation-wide budget behind the focused owner. The lifecycle integration target passes 21 cases (18 existing plus 3 header-reference refusal regressions); four native-focused tests cover six temporary native-saved outputs, and strict direct readback passes without rewriting all six outputs; sizes and hashes are recorded in the Sep 7 ADR/fixture receipt | Operation-specific native E4 only. Neutral identity passes 45 units including versioned accounting, the new lazy codec passes 10, and strict protobuf and Keynote Clippy pass. The private lifecycle unit slice passes 11 and existing replacement validation passes 28; regenerated candidate bytes equal the native-verified originals. Bounded lifecycle ASAN passed 256 runs and boundary units passed 942; implementation commit `182f31566` passed all normal workspace hooks; no host lifecycle route was retired and no generic media CRUD claim follows |
| Numbers persisted-sort/relocation physical compatibility | 🟡 Exact package snapshots stay on semantic focused transactions with terminal refusals; only historical source-built snapshots may use the private physical bridges for field 44 or table relocation. No public physical owner or ADR 0028 gate impact follows. |
| Broad Numbers baseline snapshot (2026-09-01; stale) | ❌ 405 unit tests passed/4 ignored; all integration targets passed except `table_data_list_reader_integration`, where 19 passed and 6 pre-existing comment/reply cases returned `InvalidSource` while validating synthetic comment metadata. This is a historical baseline snapshot, not a current-worktree result. |
| Numbers 14.4 Percentage probe | ✅ Rust candidate opened without repair; native save/close/reopen retained `4,200.000%`, three decimals, red-parentheses negatives, thousands separator, and Actual value `42`; strict Rust reread emitted an exact semantic no-op with zero touched components and the native-resaved SHA-256 unchanged |
| Historical broad Keynote baseline (2026-08-31) | ❌ 153 unit tests passed/1 failed; `soundtrack_order` integration was 3 passed/5 failed; rerun reproduced `selected_zip_suffix_and_central_records_allow_only_reassembly_fields` central-record failure |
| Documentation checks | ✅ `git diff --check` passed and every relative link in the audit and app-matrix documents resolves locally |
| Repository connectivity | ❌ `git fsck --connectivity-only --no-dangling` reports one missing commit and seven missing trees |
| Numbers existing-cell Custom owner | ✅ The selector-first, source-bound owner retains strict Buffa and deterministic source-built E1 coverage. Its admitted registry routes are `TN.DocumentArchive` field 9/message type 222 or `TN.super` field 8 → `TSA.custom_format_list` field 12. Focused integration validation passes 37 tests (19 owner, 7 native Custom Number, 6 native Custom Text, and 5 native Custom Date & Time); Numbers 14.4 opened the focused replacement and clear candidates without repair. Checked-in replacement and clear fixtures are documented in [`test-data/iwork`](../test-data/iwork/README.md#numbers-custom-format-raw-id-retirement-2026-09-06). The three dedicated raw-ID `NumbersEditor` Custom conveniences are retired; generic source-built/cross-format `DataFormat::Custom` compatibility and private attached Pages/Keynote table adapters remain host-owned. Boundary verification passes 927 policy tests. |
| Numbers existing-cell Duration owner | 🟡 Selector-first archive-free `Package::{table_cell_duration_format, edit_table_cell_duration_format, apply_table_cell_duration_format}` owns one rooted existing cell. The strict type-268 route admits fields 1/7/15/16/40, styles `0/1/2`, unit bits `1/2/4/8/16/32`, BNC value type 7/kind 4, and marker `0x0004` primary-only or `0x0005` with a valid generic Number secondary; only a true marker/kind/reference-free absence reads as `None`; marker-zero tuples that retain Duration kind/reference metadata are ambiguous inherited state and fail closed. Deterministic source-built focused package/codec tests provide E1 synthetic/self-round-trip evidence, including set/clear/reset, COW/refcounts, secondary retention/removal, scalar/formula-cache and opaque-byte preservation, inverse, candidate readback, locality, and typed refusals. An automated AppleScript-driven Numbers 14.4 open/save/close/reopen probe successfully round-tripped Litchi candidates for both admitted marker forms (`0x0004` primary-only and `0x0005` with retained generic Number secondary), with no reported error or repair/conversion indication; no GUI repair-dialog inspection was performed. Strict post-native rereads were semantic no-ops, exact inverse restoration held, and recorded candidate/native-resaved and Duration Tile/DataList member hashes matched across each native resave ([ADR 0008](adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-existing-cell-duration-owner-and-raw-id-host-route-retirement)). This is disposable, operation-specific E3/E4 evidence only. Apple-authored source/probe artifacts are provenance, not a checked-in E2 fixture, and do not establish broad native acceptance, arbitrary-producer parity, native byte parity, or package-wide performance. Dedicated raw-ID `NumbersEditor` Duration routes are retired; generic source-built/cross-format `DataFormat::Duration` and private Pages/Keynote attached-table adapters remain host-owned, focused refusals are terminal, and topology/debt/deletion-gate counts are unchanged. |

The four ignored Numbers native-oracle tests were not run because their private external fixture is unavailable.
The disposable Percentage probe establishes operation-specific native evidence, but formal E3/E4
promotion remains pending until its artifact and ledger are frozen and reproducible. The
Scientific record establishes operation-specific E3/E4 evidence for one existing-cell precision
change only: it does not certify arbitrary producers, other display families, package-wide
resource behavior, or a monolith deletion gate. The Fraction record establishes focused
source/build/test/fuzz plus operation-specific native E3/E4 evidence for the representative
Eighths-to-Hundredths scenario; field 20 true remains refused, and no all-variant native,
arbitrary-producer, native-byte-parity, package-wide performance, or monolith deletion-gate claim
is made. The Date & Time record establishes operation-specific E3/E4 evidence for the one
existing type-9 cell/pattern, including native save/close/exact-path reopen and a strict normalized
reread that was an exact byte-identical no-op; it does not generalize to other DateTime shapes,
producers, package bytes, or broad DateTime ownership. The dedicated raw-ID
`NumbersEditor` Date & Time route is retired as a production API boundary,
while generic source-built/cross-format `DataFormat::DateTime`, broad
smart-field lifecycle, and attached Pages/Keynote table compatibility remain
host-owned. The Custom record establishes a bounded existing-cell document-registry owner only:
its E1 source coverage is supplemented by operation-specific native Number, Text, and Date & Time
replacement/clear evidence. The admitted registry routes are `TN.DocumentArchive` field 9/message
type 222 or `TN.super` field 8 → `TSA.custom_format_list` field 12. Focused Custom integration
validation passes 37 tests, and boundary verification passes 927 policy tests. The dedicated raw-ID
`NumbersEditor` Custom conveniences are retired; generic source-built/cross-format
`DataFormat::Custom` compatibility and private attached Pages/Keynote table adapters remain
host-owned. This does not establish generic Custom-format authoring or any topology/debt/deletion-
gate change. The Duration record establishes a bounded
existing-cell native type-268/kind-4 owner only. A true marker/kind/reference-free absence reads as
`None`; marker-zero tuples that retain Duration kind/reference metadata are ambiguous inherited
state and fail closed. Its focused package and codec work is E1 synthetic/self-round-trip evidence.
An automated AppleScript-driven Numbers 14.4 open/save/close/reopen probe successfully round-tripped
Litchi candidates for both admitted marker forms (`0x0004` primary-only and `0x0005` with retained
generic Number secondary), with no reported error or repair/conversion indication; no GUI
repair-dialog inspection was performed. Strict post-native rereads were semantic no-ops, exact
inverse restoration held, and the recorded candidate/native-resaved and Duration Tile/DataList
member hashes matched across each native resave. This is disposable, operation-specific E3/E4
evidence only. The Apple-authored source/probe artifacts are provenance, not a checked-in E2 fixture,
and do not establish broad native acceptance, arbitrary-producer parity, native byte parity, or
package-wide performance. The dedicated raw-ID
`NumbersEditor` Duration route is retired, while generic source-built/cross-format
`DataFormat::Duration` and private attached Pages/Keynote table adapters remain host-owned; focused
refusals are terminal. The Keynote physical-sort run was performed against a pre-hardening candidate; current
strict admission rejects the app-authored source's unproven model field 39. Its artifacts are
external exploratory evidence only, not current-owner E3/E4 evidence; native acceptance remains
pending until a current-admitted native source is rerun.

## Safety and integrity findings

Strengths include hard ceilings for Snappy, wire fields/depth/work, archive objects/messages, ZIP entries and sizes, exact source authorization, candidate reopen/readback, and inert handling of external resources.

Material follow-ups:

1. The legacy `litchi-iwa` directory and media paths use check-then-open flows without the pinned descriptor/no-follow design already available in `FrozenDirectoryBundle`; this leaves a symlink/TOCTOU window ([bundle path](../crates/litchi-iwa/src/bundle.rs#L745), [safe snapshot](../crates/litchi-iwa-archive/src/directory.rs#L443)).
2. `litchi-iwa-graph` insertion/traversal and `litchi-iwa-index` construction lack unified aggregate node/edge/work ceilings. Index buffers use fallible reservation, but graph materialization/traversal still grows collections infallibly.
3. Archive accounting has no single retained-memory or outer-plus-nested budget; the generic IWA path charges aggregate bytes after decoding, and replacement Deflate output is not checked against the configured compressed-member ceiling.
4. Encryption/password-protected packages are rejected. Cryptographic signature verification/re-signing is absent.
5. Focused/archive `write_to` APIs remain caller-owned streaming sinks and do not themselves flush, sync, rename, or provide durable publication. Pages, Numbers, and Keynote `Package::save` now delegate to the archive-owned atomic publication boundary; durable patch serialization/history remains absent and platform-specific durability limits stay explicit.
6. Public `Patch::is_noop()` semantics are not uniform: a few owners expose semantic equality while apply/commit restores exact-artifact authorization. The API contract should distinguish those predicates.
7. The shared `roundtrip_iwork` example can feed opaque compressed bytes into the Store-only logical writer. Unsupported compression must use preserve-mode reassembly or be rejected.

These are bounded-surface findings, not evidence of active content execution: no URL fetching, process execution, macro execution, or external-media retrieval was found.

## Worktree and repository constraints

At the first snapshot, 12 tracked paths were modified or deleted and additional files were untracked. The active work included Keynote table-title alias removal and sort-order hardening, a Pages table-lock seam, a Numbers table-info projection, protobuf/fuzz changes, boundary-checker changes, corpus seeds, and temporary artifacts. The workspace changed concurrently during the audit; the final pre-handoff check showed four tracked modifications (`protobuf.rs`, Keynote sort source/test, and a Numbers fuzz target) plus the untracked seams/artifacts. The verification results above are dirty-state diagnostics between those observations and must be rerun on a frozen revision. All WIP must be classified, wired, tested, and committed or removed before it enters a support claim.

Repository history is also incomplete locally. `git rev-list HEAD...@{upstream}` fails because commit `32bfe7da90d6d7133b4b4cfaa33bdbbc48971ece` references missing parent `1f61a7a6f478bc112e09a1e5ac8899ef2e665f97`; `git fsck --connectivity-only --no-dangling` reports that commit plus seven missing trees. Exact `HEAD` inspection works, but ancestry, ahead/behind, merge-base, and historical completeness claims are unreliable until the objects are restored.

## Recommended sequence

### P0 — make progress auditable

1. Restore the missing Git objects and establish a clean, reproducible audit revision.
2. Maintain the new authoritative app matrices against a frozen revision and keep focused-owner support, legacy-host-only support, opaque preservation, and native evidence distinct.
3. Keep the completed Percentage, Currency, Scientific, Fraction, Text, Date & Time, Duration, and Custom slices isolated from unresolved sort, lock, projection, and other untracked work before promoting any additional support claim.
4. Restore the boundary ratchet and fix the reproducible Numbers storage/comment and Keynote soundtrack/ZIP-preservation failures.

### P1 — close product and release gates

1. Migrate fresh-package, lifecycle, rich text/table, chart, shape, and media ownership from `litchi-iwa` into the concrete app crates, retiring debts by capability rather than by raw API count.
2. Add richer checked-in native fixtures and an operation-indexed native accept/save/close/reopen/readback ledger.
3. Add leaf-feature isolation and iWork macOS/Windows CI; include fixture and fuzz-corpus paths in triggers.
4. Install/pin `protoc` in publish CI and test each unpacked published crate.
5. Close the directory path race and add bounded graph/index plus aggregate archive-memory/work accounting.

### P2 — normalize the public story

1. Align README/rustdoc claims with the immutable root facade versus writable focused owners versus legacy host.
2. Add runnable examples for the focused table/chart/movie/comment/footnote APIs that currently have tests but no examples.
3. Document encryption, signatures, durable save, preview invalidation, formulas/external links, and Windows directory behavior as explicit unsupported or inert rows.

## Keynote lifecycle cache-hardening follow-up (2026-09-07)

This extends the verified lifecycle baseline `ca9b9f5cd`; no host route is
retired. Typed movie/audio calls resolve complete source-order media positions,
including live-video and placeholder siblings, and report `KindMismatch`
before metadata rewriting. Selected direct comments return
`UnsupportedComment` until reply/author ownership has a dedicated transaction.
Comments on unselected media remain preserved.

A changed build topology invalidates only the selected `SlideNode` caches:
fields 15, 20, 22, and 23 are removed; fields 26 and 27 become `u32::MAX`.
Already-invalidated nodes keep their exact component bytes. Other node fields
and ZIP members remain preserved, and inverse patches restore the original
cache bytes. Five Keynote-authored source/duplicate/removal fixtures establish
the invalidated baseline state. Keynote may recompute a cache during save;
that native saved state is checked separately from the transaction output.
The sixth scalar-only Buffa projection keeps five generated files totaling
191,918 bytes under a 192 KiB bound, with no generated repeated views.

Source-built interoperability covers mixed movie/audio order, multiple builds,
titles/captions, duplicate/shared/final-owner removal, host reopen, data
retention/culling, exact inverse, and unselected-comment preservation. A narrow
Movie → CaptionInfo → ShapeStyle witness admits the producer's transitive
style references without ignoring unexplained header edges. Playback UUIDs
and registry UUIDs are separate identity domains. Clone identities use a
deterministic hash with collision checks instead of XOR, which collided for
source-built label graphs. Chunk UUID occurrences must agree within a build.

The focused allocation watermark remains monotonic: removals retain
`last_object_identifier`, and duplicates allocate above it. The host's
trailing-suffix release remains separate legacy behavior. Source-built
interoperability is E1 evidence; comment ownership and the remaining host
compatibility decisions still gate ADR 0028 deletion.

Targeted verification passed: 27 focused lifecycle integration tests, three
source-built interoperability tests, 16 neutral lifecycle codec tests, strict
Keynote/protos Clippy, 942 boundary unit tests, and the full scanner (64
packages, 238 internal edges, 11 explicit debts). The final cache candidate
also passed the environment-enabled native-saved strict readback test. Implementation commit `fe4592c68` passed the normal workspace hooks:
formatting, lint policy, all-feature library/integration tests, and
documentation tests. The codec source-tracking guard also covers the
nested node-cache module path. This follow-up retires no host routes.

## Keynote comment duplication follow-up (2026-09-07)

The focused owner implements selected direct comment/reply duplication through
the neutral lazy batch codec and a private comment-graph/author dependency
witness. Before the admission hardening, the focused lifecycle library slice
passed 28 cases and the integration target passed its 27 existing cases plus
11 comment-duplication cases. The default integration run does not enable
native saved-candidate readback, while the explicit native environment run
passed one strict reread. These counts are pre-hardening; the current final
rerun remains pending.

Native Keynote authored the baseline comment root on movie `2653286`, with
storage `2653723`, author `2653721`, and
`externalAnnotationAuthorStorage` `2652381`. Native Cmd-D produced movie
`2653814` and storage `2653826`, kept the storage UUID byte-exact, and shared
the author records. Native save, actual close, and exact-path reopen showed
comments on both copies in the three-movie/two-audio duplicate, while the
baseline remained two movies/two audio with its original comment.

The temporary native removal oracle culled the comment root while retaining
the author and author-storage records. Only component external edge
`(2652150, 2652381, 2653721)` was removed, taking the external set from 700 to
699; the duplicate external set remained 700. Permanent fixtures are
`test-data/iwork/keynote/media-comments-baseline-native.key` (SHA-256
`69d493b183308b6a0f336b6a944b2d78fe8439a40648a0dcb72ff1e171380fff`) and
`test-data/iwork/keynote/media-comments-duplicate-native.key` (SHA-256
`8c04282f5877ae671c808cb6a35022a3c41432459fb2bc0e43a22378d3c59cb6`). The
temporary removal oracle hash is
`1d0897ebd1f79b5e5ee3a33ce7f034f69baec9bac05be6cddad73d3dd828d509`.

The exported focused candidate
`focused-media-comment-duplicate-movie.key` changed from SHA-256
`d3c839320da017b48e79627d50f48eace995978c2344b3be8fb7739711ae1869` to
`8c3ed9b005a5bc280d9d9859d4c4b0dae71c958ec04dbc2de486e6d77a0472a8` after
Keynote Cmd-S, actual close to the theme chooser, and exact-path reopen. It
opened with three movies and two audio objects, retained two comments, and
showed no alert. The explicit strict reread verified the new comment IDs and
preserved storage UUIDs and authors. This focused native duplicate leaf remains
verified; native reply and comment removal remain unverified.

No native reply was authored because the prototype Keynote reply popover was
not operable; synthetic and source-built replies remain separate evidence. The
five host cases pass, covering positive selected movie/audio reply cloning plus
the legacy host reply-duplicate/remove refusal oracle. Comment removal remains
`UnsupportedComment`, and no host route has been retired.

The current extension rejects unknown raw-header references and unknown or
deprecated direct-comment references before publication. New selected-comment
admission is fail-closed when opaque unknown fields in the comment root, date,
or UUID could hide a reference back to the source graph. Admission requires a
strict known comment envelope and an exact header-edge census. The neutral
codec still preserves unknown bytes, and untouched or unselected comments
remain preserved. Full-inspection budgets are shared across graph, author,
header, codec, and metadata walks. Boundary verification previously passed
943 unit cases and `litchi-iwa-protos --lib` previously passed 846 cases,
including corrected mixed-UUID shrink/reply-varint-growth exact raw/Buffa
parity; those are pre-hardening counts and the current final rerun is pending.
The full scanner, strict Clippy, and normal hooks remain pending. No cleanup
result is claimed here.

The next goal is author culling plus metadata removal, with receipts required
before promotion.

Final selected-comment admission verification passes: the complete Keynote
library has 258 passing tests, with 27 existing media lifecycle and 12 comment
integration tests also passing. The explicit native-saved readback ran with its
path supplied. The final generated duplicate has the same SHA-256
`d3c839320da017b48e79627d50f48eace995978c2344b3be8fb7739711ae1869` as the
artifact verified in Keynote before native save. Strict Clippy for both changed
production libraries and workspace formatting pass. Boundary unit verification
passes 943 cases. Full boundary scanning and normal commit hooks are pending
the final receipt.

Final commit and cleanup receipt: `ce6ab2e43` passed the normal pre-commit
hooks: workspace formatting, all-feature production lint, all-feature workspace
library/integration tests, and workspace documentation tests. The full boundary
scanner passed for 64 packages and 238 internal dependency declarations, with
11 explicit migration-debt items. This supersedes the pending final-verification
notes above. No further host route was retired.

After verification, the owned temporary directory
`/private/tmp/litchi-media-comments-20260907t` was removed. `cargo clean` removed
14,129 files and reclaimed 9.8 GiB; free disk space was approximately 62 GiB.
The two checked-in native comment fixtures remain. Native reply authoring and
selected commented-media removal remain outside this completed duplication slice.

## Keynote selected commented-media removal follow-up (2026-09-07)

This current-turn record is additive to the completed comment-duplication
receipt. The focused owner implements selected commented movie/audio
removal while preserving a shared comment root and reply closure through a
global incoming-header census. Metadata author usage is limited to the source
component: a global `authorStorage` reference does not count as usage of that
source component.

The native policy retains physical author and author-storage records. The
owner removes only the exact strong current external edge, and only when the
source component is its final owner. It reuses the neutral
`ExternalReferenceRemoval` path and adds no protobuf encoding. This corrects
the previous next-slice shorthand: required cleanup is the final
source-component edge, not physical author or author-storage removal.

Two permanent Keynote fixtures were authored with Cmd-S, actually closed to the
theme chooser, reopened from their exact paths without alerts, and closed:

| Fixture | Native result | SHA-256 |
| --- | --- | --- |
| `test-data/iwork/keynote/media-comments-shared-removal-native.key` | 2 movies, 2 audio, 1 comment; removed clone `2653814` from the previous duplicate | `3dea8fded4216538b398942516f8fb9316d10fe40e1a0e1be54e020c59394c11` |
| `test-data/iwork/keynote/media-comments-final-removal-native.key` | 1 movie, 2 audio, 0 comments; removed original `2653286` | `c3c7c8952914263bb299df95f36a67be5fd46e7add6f37e90ab03e12810db5b7` |

The neutral metadata slice has 45 passing cases. Host-retirement gates remain
under review and the existing APIs are preserved. Final verification and
cleanup status are recorded below; no current-turn cleanup receipt is claimed.

### Current verification update (2026-09-07)

The root lifecycle library now passes 36 cases; strict Keynote all-feature
library Clippy, neutral metadata 45 cases, the existing focused lifecycle
integration 27-case slice, the host 5-case slice, and the boundary guard 943
case slice pass. The current comment-removal integration slice passes all 21
cases, including both explicit native-saved focused readbacks and the corrected
exact native-fixture oracle. The full current totals are 36 lifecycle-library,
21 comment-integration, 27 lifecycle-integration, and 5 host cases, plus
neutral metadata 45. Boundary scanning, normal hooks, and cleanup remain
pending.

Both focused native-saved candidates were Cmd-S saved, actually closed to the
theme selector, reopened from their exact paths without alerts, and closed.
Their byte transitions are:

| Candidate | Native result | Before SHA-256 | After SHA-256 |
| --- | --- | --- | --- |
| Shared removal | 2 movies, 2 audio, 1 comment | `5bd6aa9ad0921339bc851dcab5ebc350ddf64df32de020d38928cb4c83b70a36` | `164cdf00d8b4d25e0ff8c927ece58da8b446ebe19e074825c598df637519b32d` |
| Final removal | 1 movie, 2 audio, 0 comments | `0394b3a87c7cd957bae5795f6fbc6e970c303e6da4ad88b690242f7fbe5428ee` | `a7e5e735f94da1c2ba84320e9664e15ebe44067c1aab581c1553cbe8eb0f5c2f` |

Both explicit native-saved strict readbacks pass. The native final-removal
artifact rotates the unrelated locator `ViewState` to `ViewState-2654075`.
The corrected oracle compares the exact raw delta after excluding that
documented two-edge `ViewState` rotation and the author removal; no broad
semantic normalization is permitted. Raw edge counts remain 700 → 700 → 699.
The focused candidate comparison remains exact-raw-edge based.

Known `Movie` and `CommentStorage` survivor payload/header validation rejects
stale references before removal, preserves opaque survivor-root extensions, and
keeps selected cloning strict. The real Keynote selected-Audio Comment toolbar
was disabled and the authoring probe closed unchanged, so there is no native
commented-audio E4 claim. Host retirement remains gated by caption/title/
stand-in comment graphs, native audio coverage, unknown outer-reference
completeness, and identifier-watermark compatibility; the host APIs remain
retained and full host retirement is not claimed.

The two environment-path strict readbacks and corrected oracle are complete.
The boundary scanner, normal hooks, and cleanup remain pending. The current
disposable target is approximately 1 GiB with about 61 GiB free, and the owned
temporary payload at `/private/tmp/litchi-media-comment-removal-20260907u`
remains pending removal. The previous full clean belongs to the preceding turn.

Final source-target validation also resolves known survivor comment roots and
replies to same-component type-3056 objects, and authors to type-212 objects.
Missing or wrongly typed targets are rejected atomically even when payload and
header IDs agree. The full Keynote library passes 265 tests (including 36
lifecycle cases), alongside 21 comment and 27 lifecycle integration tests.
Both explicit native-saved readbacks pass after this hardening, and regenerated
shared/final candidates retain the before-save hashes recorded above.

### Final removal verification and cleanup receipt (2026-09-07)

Commit `b278aef8f` completed the focused direct-media comment removal slice.
Normal hooks passed Rust formatting, workspace all-feature library lint,
workspace all-feature library/integration tests, and workspace documentation
tests. The final boundary suite passes 943 tests; its full scanner reports
64 packages, 238 internal dependency declarations, and 11 explicit migration
debt items. These results supersede the pending hook/scanner notes above.

The owned `/private/tmp/litchi-media-comment-removal-20260907u` directory was
removed after verification: 31 temporary files totaling 11,979,409 bytes. The
two permanent native removal fixtures remain tracked. The rebuilt 9.3 GiB
Cargo cache is retained for the next slice, with approximately 53 GiB free;
the preceding turn's 9.8 GiB `cargo clean` remains the latest full-clean receipt.
The caption, native commented-audio, unknown-reference, and identifier-watermark
compatibility gates remain open; no additional host API was retired.
