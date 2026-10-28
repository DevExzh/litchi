# 0379 — W1-G: the shared iWork table-family crate `litchi-iwa-table`

- Status: groups 1 (cell formats) and 6 (hidden axes) landed on branch
  `uo2/iwa-w1g`, cut from trunk `feat/unified-ops-api` at `ca202291ca`.
  Group 1: the crate and the Numbers move (`35908bbc0a`, merge-ready at
  `7ea1e295b1`), then the Pages and Keynote families (`73f6fa976a`).
  Group 6 (§2a): the engine lift with the Pages owner (`e6e946e72b`), the
  Keynote family (`c4a13cadfa`), then the engine changes, the host builder
  change and the Numbers family (`3a9ef1b692`, `c0ec8404f2`,
  `1544f10604`). Group 2 (table-cell comments) now has a root-only shared
  engine and three facade slices on the continuation branch; focused code,
  parity, fuzz-lock and documentation gates pass, while the current boundary
  checker passes after removing the newly introduced root-level comment
  exports. The resource-dimension repair is committed as `03f991fcf6`; its
  current-source read/create/clear instruction probes are recorded in §4.3.
  The pre-repair WP source/probe receipt is `e09f2269d0`; its twelve-route
  read/create/Set/clear matrix is retained as historical evidence in §4.4,
  but does not qualify the repaired public API. Final clean-source and
  integrated performance and native verification remain pending (§2b, §6).
  Groups 3–5 (rich text, topology, conditional highlights) are not started
  (§5).
- Date: 2026-10-28 continuation; group 6 completed 2026-10-27.
- Review follow-up (`828d02cad9`): owner-extents coordinates require
  canonical value varints, and the Numbers facade preserves typed resolver
  and candidate-ingress limits/refusals. Numbers and Keynote hidden-axis
  parity plus the focused engine and facade tests pass after this fix.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (W1 table families, §3.2
  "one family taking a format enum"), the W1-A Numbers table records
  [0370](0370-iwa-w1a-numbers-tables.md)–[0370c](0370c-iwa-w1a-row-column-removal.md),
  the W1-F formula engine [0376](0376-iwa-w1f-formula-cache-engine.md),
  the harness and ledger [0369](0369-iwa-exit-parity-harness-and-ledger.md).
- `performance_claim`: the Numbers format, control and Pop-Up Menu routes
  are unchanged under callgrind (+0.09%, −0.03%, +0.01%); the new Pages
  family runs 88.2% below its host route and the Keynote family 0.9% above
  its host route (§4). Hidden axes (§2a): the engine lift leaves the Pages
  route unchanged (+0.008%); the Keynote and Numbers families run 60.4%
  and 18.1% below their host routes. Group 2's focused full-facade probes
  on source commit `cfc334e754` measured Pages −96.517%, Keynote −36.624%
  and Numbers +0.956% (§4.2). The pre-repair source `e09f2269d0` twelve-route
  matrix measured Pages −96.349% to −97.201%, Keynote −34.398% to −52.520%,
  and Numbers −31.796% to +1.043% (§4.4); those captures do not qualify the
  repaired public API, and final clean-source/integrated-tree qualification
  remains with the coordinator.

## 1. Why a new crate

Numbers sheets, Pages body tables and Keynote slide tables store the same TST
graph: a `TableModelArchive`, its tiles of BNC cells and the `TableDataList`
sidecars. The leaf engines that edit cell formats, rich text, comments and
topology lived only in `litchi-numbers` (~62k lines of
`package/table_*.rs`), tied to its sheet selectors, its `Package` state and
its error types. W1-A had already moved payload pieces into
`litchi_numbers_wire`. The coordinator decided to put the format-neutral
engines into one focused crate that the three leaves consume privately and
no facade re-exports.

`litchi-iwa-table` engines take:

- a `TableSource` (lib.rs): the component catalog, the exact physical source,
  an object lookup, the source bytes and the ingress limits. Numbers
  implements it over its object index, Pages and Keynote over theirs;
- a resolved `cell_format::Target`: the table model's component, object and
  message route, its lock state, the cell, and two leaf-chosen coordinates
  (`scope`, `table`) that come back in errors;
- one `cell_format::TransactionBudget`.

They return private archive bytes per changed member and a neutral
`cell_format::Error { Refused, InvalidSource, Limit, Allocation,
Verification, PatchConflict }` with a content-free `Path`. The leaf keeps
selectors, reassembly policy, candidate reopen and its public types.

## 2. Group 1: cell formats

### 2.1 The move (`35908bbc0a`)

| From `litchi-numbers/src/package/` | To `litchi-iwa-table/src/cell_format/` |
|---|---|
| `table_cell_control_native.rs` | `control.rs` (git mv) |
| `table_cell_display_format_native.rs` | `display.rs` (git mv) |
| `table_cell_pop_up_menu_native.rs` | `pop_up_menu_native.rs` (git mv) |
| `table_cell_pop_up_menu_metadata.rs` | `pop_up_menu_metadata.rs` (git mv) |
| engine half of `table_cell_pop_up_menu.rs`: `TransactionBudget`, control and Pop-Up Menu reads, native rewrite, metadata transition | `budget.rs`, `pop_up_menu.rs` |
| `metadata.rs` route helpers used by the above | `metadata.rs` (copy; Numbers keeps its own for its other users) |

Numbers changes only at call sites: `package.rs` aliases the moved modules
under their old names (`pub(crate) use litchi_iwa_table::cell_format::control
as table_cell_control_native`, …); the Pop-Up Menu module converts the
neutral `Error`, `Path` and `LimitKind` to its public types with lossless
`From` impls (the neutral `scope` is the sheet position); the nine format
facades and the control facade map through the same conversion
(`native_path` now returns the engine path). The public API is unchanged.

### 2.2 Pages and Keynote

Plan §3.2 asks for one family taking a format enum instead of the host's
per-format families. Two engine modules make that possible:

- `cell_format::data_format::{read, rewrite, family, is_display_family,
  maximum_members}` dispatches the neutral `DataFormat` over the
  per-family readers and writers. `read` reads controls first, then the
  nine display families, and reports Automatic only when a family reader
  saw an automatic cell (a format no reader admits is
  `UnsupportedDependency`, never Automatic). `rewrite` changes one display
  family at a time;
- `cell_format::publish::reassemble` is the exact-source reassembly the
  Numbers display facades run after a rewrite (preflight bound, Snappy
  under the retained limits, ZIP splice), shared instead of copied a
  tenth time.

`TransactionBudget::with_list_keys(ListKeys)` selects how format and
control lists allocate keys. Numbers keeps `Reuse` (lowest free key, the
watermark stays on release), so its output is unchanged. Pages and Keynote
use `Append`: a new entry takes `next_list_id`, and a removed last key is
handed back, which is what the host's Pages/Keynote format writer and the
shared cell writer (`table_write::list_release`, record 0370a) do. The
parity harness found the difference (`next_list_id` 5 against 6 on a reset,
key 3 against 6 on a replacement); with `Append` both match the host.

New public families (glossary rows and registrations in `pages`,
`keynote-minimal`, `keynote-full`):

- `litchi_pages::table::format`: `Package::{body_table_cell_format,
  edit_body_table_cell_format, apply_body_table_cell_format}`,
  `Edit::{set(DataFormat), reset}`;
- `litchi_keynote::slide::table::format`: `Package::{slide_table_cell_format,
  edit_slide_table_cell_format, apply_slide_table_cell_format}`, the same
  verbs. The dedicated `edit_slide_table_cell_number_format` family stays.

Each resolves its selector (Pages `resolve_body_table`, Keynote
`slide_table_core::select_table`) into an engine `Target`, implements
`TableSource` (Pages scans its catalog, Keynote uses its object index),
reopens the candidate with its own `Package` and verifies the read. A
change between two families runs two engine rewrites on a private
intermediate candidate; the patch keeps the source and final artifacts,
so `inverse` restores the source byte for byte. Controls, Pop-Up Menus and
numeral systems are read, and refused for writing
(`Refusal::UnsupportedFormat`): their writes also rewrite the package
metadata registry, which needs the Numbers control facade's publication
(save tokens, preview policy) lifted too (§5).

Parity (`crates/litchi-iwa/tests/parity/{pages,keynote}_cell_format_cases.rs`),
on B2 of `body-table-cells-native.pages` and the first stored automatic cell
of `slide-table-read-native.key`:

| Test | Routes | Tier |
|---|---|---|
| `pages_body_table_cell_data_format_set_parity`, `keynote_slide_table_cell_data_format_set_parity` | Automatic to Number, host `set_table_cell_data_format` | B |
| `pages_body_table_cell_format_set_parity` | a replacement inside Number, Currency, Percentage, Scientific and Fraction on the host's formatted cell, host family setters | B |
| `keynote_slide_table_cell_format_set_parity` | the same for Currency, Percentage, Scientific and Fraction | B |
| `*_cell_format_reset_parity` | each of those families reset to Automatic, host family resets (Keynote Number: the generic setter) | B |
| `*_cell_format_read_parity` | host family and generic reads against the leaf read on the host's output | A |
| `*_cell_format_divergence_is_reported` | pins where the routes differ (§5) | — |

Ledger: 31 rows migrate, 17 Pages (`set_`, `reset_` and `table_cell_` for
the five decimal families, `set_table_cell_data_format`,
`table_cell_data_format`) and 14 Keynote (the same for the four families
without Number, plus the two generic rows). Six macOS queue rows cover the
routes; none has native evidence.

### 2.3 The structural working set

The Numbers structural workspace (`package/structure.rs` and
`structure/{budget,collect,formula,metadata,plan,references}.rs`) moves to
`litchi_iwa_table::structure` as `Working<'p, P: TableLeaf>`. `TableLeaf`
is `TableSource` plus the leaf's strict `reopen_candidate`. Numbers keeps
its sheet and document helpers, the patch `Artifacts` and the
`structural_family!` macro, and aliases `type Working<'p> =
litchi_iwa_table::structure::Working<'p, Package>` and `type Published =
Published<Package>`, so its call sites are unchanged. Every Numbers test
(1,044 with the crate's) and the 151 Numbers parity cases pass; callgrind,
same probes as 0370/0370a (A before, B after, A B B A):

| Route | Before | After | Δ |
|---|---:|---:|---:|
| leaf topology insertion (`cost_probes::numbers_topology_insert_cost_abba`) | 415,361,948 | 415,363,167 | +0.0003% |
| leaf whole-cell text style (`cost_probes::numbers_text_style_cost_abba`) | 323,242,772 | 323,187,077 | −0.02% |

## 2a. Group 6: hidden axes

Ordered before rich text because W1-C waits on the Keynote family.

### 2a.1 The engine lift and the Pages facade

`litchi-pages` `package/body_table_hidden_axes.rs` (graph proof, profile
classification, extent and filter validation) and
`body_table_hidden_axes/rewrite.rs` (owner creation, UID-map and metadata
rewrite, git mv) move to `litchi_iwa_table::hidden_axes`. The engine is
generic over a `Leaf` (its selector type, its wire `Budget`, its target
resolution and strict reopen); its `Error`, `Refusal` and `LimitKind` keep
the Pages shapes, so the Pages owner converts both ways losslessly. The
page-layout and table-lock helpers the engine called (exact source,
editable component, recompression, selected-message metadata checks,
fingerprint, compressed bounds) are copied into `hidden_axes::support`
with the same error mapping; the Pages originals stay for their other users.

Pages keeps `table::hidden_axes::transaction::{Edit, Patch, Commit,
Diagnostics, Error, LimitKind, Path, Refusal}` and the three `Package`
methods unchanged, and implements the engine over a private `Leaf(Package)`
wrapper (its `table_lock::WireBudget` stays crate-private), so resolution,
proof and rewrite still share one budget. Every `litchi-pages` test passes
unchanged (754 with the crate's). Callgrind,
`native_hidden_axes::hidden_axes_cost_probe` (set then clear on
`body-table-hidden-axes-native.pages`, A B B A): 34,109,338 before,
34,112,098 after, +0.008%.

### 2a.2 Keynote

`litchi_keynote::slide::table::hidden_axes` (values re-exported from
`litchi_iwa_common::table::axis`, transactions in `transaction`) with
`Package::{slide_table_hidden_axes, edit_slide_table_hidden_axes,
apply_slide_table_hidden_axes}`. The facade resolves the slide and table
through `slide_table_core::select_table` into the engine `Target` (its new
`scope` is the slide position), runs the engine over a private
`Leaf(Package)` wrapper with the engine's own `hidden_axes::WireBudget`
(the Pages table-lock ledger, now a crate type), and keeps root previews
(`Leaf::DELETES_ROOT_PREVIEWS = false`), as the host does.

Parity (`crates/litchi-iwa/tests/parity/keynote_hidden_axes_cases.rs`), on
the first slide table of `table-discovery.key` and
`slide-table-merges-native.key`:

| Test | Route | Tier |
|---|---|---|
| `keynote_slide_table_hidden_axes_set_parity` | hide row 1 and column 0 (owner creation) | B |
| `keynote_slide_table_hidden_axes_read_parity` | reads on the source and on the host's output | A |
| `keynote_slide_table_hidden_axes_clear_divergence_is_reported` | clear on the host's output: the engine keeps the cleared state records with no user-hidden flag, the host drops them; both outputs read as unhidden | pinned |

Ledger: `set_slide_table_hidden_axes` (B) and `slide_table_hidden_axes`
(A) migrate, with two macOS queue rows. The engine refuses
`slide-table-read-native.key`'s table as an invalid source: its formula
owner has dependencies (§2a.3, §5).
Callgrind (`keynote_hidden_axes_cost_abba`, host A B B A leaf): host
852,289,662, leaf 337,487,517, **−60.4%**.

### 2a.3 Numbers

`litchi_numbers::table::hidden_axes` with `Package::{table_hidden_axes,
edit_table_hidden_axes, apply_table_hidden_axes}`, the Keynote facade's
shape over the sheet and table selectors (`table_headers::resolve` gives
the model and drawable routes), keeping root previews as the Numbers
cell-level families do.

Two engine changes, both from the Numbers parity:

- **Owners that record the table's extents.** A table owner's spanning
  envelopes (`FormulaOwnerDependenciesArchive` fields 7 and 8) may hold
  `total_range_for_table` and `body_range_for_table` and nothing else; the
  engine used to refuse any non-empty envelope. Extents carry no
  dependency and hidden axes leave the geometry alone, so they are now
  admitted. A spanning reference (`coord_refers_to_spans`) or a repeated
  extent is still `UnsupportedDependency`, a malformed range coordinate
  `InvalidSource` (`owner_extent_tests`); every coordinate key and varint
  must use canonical wire framing. The Numbers facade also preserves
  resolver and candidate-ingress limit/refusal classifications at the engine
  boundary.
- **Identity allocation follows the host.** Owner creation allocates the
  two filter sets before the two formula owners, as the host does (it was
  the other way round); every `litchi-pages` test still passes.

**Host builder change.** The host builder's `without_owner_extents` (the
Pages bootstrap's workaround, record 0370i) is removed (`c0ec8404f2`): the
host Pages builder now writes the owner extents that native files and the
host Numbers builder store. Only the host Pages bootstrap of a first table
uses it (`Builder::body_table`, and `Draft::add_table` on a table-less
source-built document); the Keynote host builder has its own calculation
archive. Checked against the parity harness's builder cases on the merged
tree:

- The harness has no tier-A builder case: the Pages and Numbers builder
  cases are tier B and the Keynote builder comparison is an untiered
  differential (record 0372). The leaf Pages builder (W1-B, record 0371)
  has no `body_table` (a record 0371 gap), and the one Pages builder case,
  `pages_builder_language_locale_parity`, builds without a table, so the
  host route never bootstraps one.
- No other parity case reaches the bootstrap. The table lifecycle cases
  (`pages_body_table_add_parity`, tier B; `pages_body_table_remove_parity`,
  tier A) run on `body-table-catalog-native.pages`, where the host clones
  a native table, and the leaf table creation refuses a document without
  a body table (`Refusal::NoTemplate`).
- The Numbers builder cases (`numbers_builder_*_parity`) are unchanged:
  the host Numbers builder always wrote the extents (0370i).
- Every builder case and every `litchi-iwa` test passes on the merged
  tree (§6), including the host tests that bootstrap a first Pages table
  (`repeated_body_table_calls_append_independent_tables`,
  `inserts_and_removes_first_table_without_a_template`, and the
  `pages_drawable_comment_registry` and
  `pages_source_built_stylesheet_closure` suites).

When W1-B ports `body_table` or table creation on a table-less document,
the host oracle it compares against writes the owner extents, as native
files do.

Parity (`numbers_hidden_axes_cases.rs`) on `basic.numbers`,
`custom-number-native.numbers` and a host-built 6×4 workbook:

| Test | Route | Tier |
|---|---|---|
| `numbers_table_hidden_axes_set_parity` | hide row 1 and column 0 on the natives; then hide more on each host output | B |
| `numbers_table_hidden_axes_read_parity` | reads on the sources and the host's outputs | A |
| `numbers_table_hidden_axes_clear_divergence_is_reported` | as Keynote | pinned |
| `numbers_table_hidden_axes_creation_divergence_is_reported` | owner creation on the host-built workbook: the engine registers the new objects' UUIDs and writes empty filter offsets; the host does neither | pinned |

Ledger: `set_table_hidden_axes` (B) and `table_hidden_axes` (A), two macOS
queue rows. Callgrind (`numbers_hidden_axes_cost_abba`): host 62,615,742,
leaf 51,289,709, **−18.1%**.

**Why `slide-table-read-native.key` (and nine Numbers fixtures) stay
refused.** The graph profiles are not the cause: every refused table has
the same version tuples (`[1, 0, 5]` info, `[3, 2, 10]` model) as the
admitted ones. The engine refuses them in
`validate_native_formula_owner_metadata`: their formula owner references
objects other than the table drawable, because the table holds formulas
or feeds a chart (`slide-table-read-native.key`'s D2 = B2 + 1). The
engine admits only owners whose dependency graph it proves empty, so a
hidden-state change cannot leave a formula or chart dependency stale. Lifting
that needs the formula-dependency engine (W1-F, record 0376) to re-plan the
owner's spanning and cell dependencies under the new hidden state; it is a
follow-up, not a profile calibration.

## 2b. Group 2: table-cell comments (root-only first slice)

The continuation now lands one complete vertical slice: root `Set`/create and
`Clear` rewrites for body and slide table cells. The shared implementation is
`litchi_iwa_table::comments_native`, below the three leaf selector adapters.
It owns the native table-list, tile-row, comment-storage, author and metadata
transitions, including format profiles for Numbers (`ListKeys::Reuse`), Pages
(`Append`) and Keynote (generated author and its native storage shape). The
old Pages and Keynote writer copies are removed; Numbers also calls this
engine for the focused root path. No facade passes native identifiers, ZIP
paths or protobuf values through a supported signature.

Before mutation, `Working` admits bounded object/reference census maps and
per-edge growth, then performs one graph walk. Comment-specific validation
checks list ownership, author locality, reply counts, unknown references and
the selected root. Prepared comment/table codecs replay their exact resource
requirements before execution; row offsets, cell buffers, metadata fields and
candidate bytes use checked or fallible reservations. Core compressed-byte,
reference, work, allocation and retained-resource refusals remain typed at
the leaf boundary. A candidate is reopened and read back before publication;
exact-source patches, `inverse()`, byte-preserving no-ops and locality checks
remain part of each facade contract.

The public leaf surfaces retain their selector-first shape:

- Pages: `body_table_cell_comment`, `edit_body_table_cell_comment` and
  `apply_body_table_cell_comment`; replies are read from `Comment::replies()`;
- Keynote: `slide_table_cell_comment`, `edit_slide_table_cell_comment` and
  `apply_slide_table_cell_comment`; replies are read from
  `Comment::replies()`; and
- Numbers: `table_cell_comment_replies`, `set_table_cell_comment`,
  `clear_table_cell_comment`, and the existing reply methods.

Reply reads remain supported through the root `Comment::replies()` projection
for Pages and Keynote; Numbers retains its existing selector-first
`Package::table_cell_comment_replies` reader and the parity case compares that
reader directly with the host result. Reply mutations remain explicit terminal
`UnsupportedDependency` refusals in the Pages and Keynote focused route, and
Numbers preserves its existing reply-bearing replacement refusal. Clearing or
replacing a root that has replies is therefore not claimed as a complete host
migration in this slice; the host-only ledger rows stay host-only until the
reply-thread lifecycle is lifted. No refusal falls through to the legacy host.

The Pages create witness pins the format-specific root-only profile rather
than masking a broad object class: on a stored, previously un-commented cell,
the host allocates one fresh type-212 author archive with the exact generated
`litchi-iwa` profile and appends its reference to the source's existing empty
type-213 author storage and the selected type-3056 comment. The leaf keeps the
existing type-213 payload empty and omits the author reference. The proof
selects the root from the host and leaf selected-cell readers, requires both
roots to be fresh against every source component, checks the selected 6005
list entry and exact allocation-derived metadata watermarks, and compares
all untouched objects and wire bytes. Only the selected new comment's date and
storage UUID are creation-time differences; unknown fields, existing comment
metadata, unrelated references and author-storage bytes are strict. Wrong
root/list, selected-text/unknown-field, and author-storage corruption are
negative controls. This is a Tier-B differential proof, not native Apple
acceptance evidence.

The differential cases currently run are:

| Format | Cases | Tier |
|---|---|---|
| Pages | `pages_body_table_comment_read_parity`, `pages_body_table_comment_set_parity`, `pages_body_table_comment_create_parity`, `pages_body_table_comment_clear_parity`, `pages_body_table_comment_replies_read_parity` | B for root writes; A for populated root/reply read |
| Keynote | `keynote_slide_table_comment_set_parity`, `keynote_slide_table_comment_clear_parity`, `keynote_slide_table_comment_replace_preserves_source_metadata`, `keynote_slide_table_comment_patch_rejects_unrelated_metadata_change`, `keynote_slide_table_comment_read_parity`, `keynote_slide_table_comment_replies_read_parity` | B for root writes; A for reads |
| Numbers | `numbers_table_cell_comment_set_parity`, `numbers_table_cell_comment_create_parity`, `numbers_table_cell_comment_clear_parity`, `numbers_table_cell_comment_host_oracle_control`, `numbers_table_cell_comment_replies_read_parity` (the read case also performs the direct semantic comparison) | B for root writes; A for populated root/reply read |

The host remains the differential oracle. The original slice plan covered 17
candidate rows. The final handoff disposition is 21 rows: five migrated reads
(including the two pre-existing Pages comment reads), fourteen host-only
mutation/reply rows, and two retired parity-only seams. The explicit
clear-oracle helper used only to prepare the independent lifecycle probes in
§4.3 is test provenance, not a supported capability or an additional
migration row. The 0363 macOS queue retains supported root create/set/clear
and deferred reply operations as separate pending planning entries; no E3/E4
native claim is made. Native-authored fixtures and synthetic/self-round-trip
tests are not Apple open/save evidence.

The Numbers mutation oracle is deliberately feature-gated
(`parity-host-oracle`) and calls the existing private legacy writer through
doc-hidden parity seams; the focused leaf route is therefore measured and
tested against an independent host path. It is not a supported Numbers editor
API and does not change the default parity target's existing coverage.

## 3. Tools and docs

- `tools/crate_boundaries.json`: the package with six edges (archive, core,
  protos, numbers-wire, common, core vocabulary) and `litchi-numbers ->
  litchi-iwa-table`; runtime-neutral.
- `tools/check_crate_boundaries.py`: the Numbers display-format, Custom,
  control and Pop-Up Menu source audits read the moved engine files
  (`IWA_TABLE_CELL_FORMAT_ENGINE_SOURCES`) for their markers and duplicate
  checks.
- `tools/non_iwork_gate.py` (+ test): 18 iWork packages, 21 excluded.
- `tools/api_glossary.json`: profile `iwa-table` (bans `_with_limits`
  suffixes and `LimitExceeded`).
- ADR 0024: topology amendment.

## 4. Performance

Callgrind, release with debug information (`lto = true`), CPU 3 pinned,
instruction counts. Probes (committed, ignored by default):
`tests/cell_format_cost_probe.rs` (number, currency, percentage, scientific,
fraction, duration, date-time and text set + clear on the data-format
fixture), `table_cell_control::control_cost_probe` (five controls set +
clear), `table_cell_pop_up_menu::pop_up_menu_cost_probe` (create, replace,
clear). Binaries of `59df2ec089` (before) and `35908bbc0a` (after) run
A, B, B, A.

| Route | Before (mean of A1, A2) | After (mean of B1, B2) | Δ |
|---|---:|---:|---:|
| display formats, 8 families | 32,233,052 | 32,263,295 | +0.09% |
| controls, 5 kinds | 14,444,461 | 14,439,824 | −0.03% |
| Pop-Up Menu | 4,736,325 | 4,736,939 | +0.01% |

Pages and Keynote, the same release binary,
`{pages,keynote}_cell_format_cases::*_cell_format_cost_abba` (a Currency
replacement on the host-written EUR cell; host family setter A, leaf
family B, run A B B A, open + edit + commit + serialize), collected with
`--toggle-collect=*probe_host*` and `*probe_leaf*`:

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| Pages `set_table_cell_currency_format` → `edit_body_table_cell_format` | 3,361,490,406 | 396,996,392 | **−88.2%** |
| Keynote `set_slide_table_cell_currency_format` → `edit_slide_table_cell_format` | 785,765,067 | 792,722,687 | +0.9% |

The Keynote leaf route pays the slide-table selection proof and its
inventory; it stays inside the 5% bar.

### 4.1 Hidden-axis review rerun

The review fix was measured again on source commit `d82fc08989`, based on
merged candidate `4abb90f713`, after the canonical coordinate-value check and
the lossless Numbers error mappings. The profiling binary was built from the
same source bytes before the commit was created (the commit records the
rustfmt-normalized mapping arms and the regression test) with the
release-derived `profiling` profile
(`lto = true`, debug information retained), then each probe ran on CPU 3 under
`taskset`/callgrind with the host and leaf probes collected separately. The
Numbers and Keynote probes used the prescribed A, B, B, A order; the Pages
probe is a single focused facade route and therefore has no host-versus-leaf
comparison.

| Route | Host instructions | Leaf/focused instructions | Δ |
|---|---:|---:|---:|
| Numbers `numbers_hidden_axes_cost_abba` host/leaf probe | 62,529,923 | 51,213,830 | **−18.097%** |
| Keynote `keynote_hidden_axes_cost_abba` host/leaf probe | 850,468,093 | 336,595,982 | **−60.422%** |
| Pages `native_hidden_axes::hidden_axes_cost_probe` focused route | — | 34,078,005 | −0.100% vs the prior after probe |

The Numbers and Keynote results are within the established cost bar. The
Pages value is reported as an engine/facade probe only, not as host-vs-leaf
parity or a qualification of the whole `File` stack. Raw callgrind output and
the profiling binary remain outside the repository under
`.coord/iwa-w1g/cg/`; the Numbers/Keynote profiling binary SHA-256 is
`a13fceee820290d82e4f70a6fb0dd02f2a1a3164295874df55ca0fa711ffbbb5` and the
Pages profiling binary SHA-256 is
`31659bbaa0f9cdc44b427f3df59d4ad802ee7b0afbf7918807d51527a9786ebd`.

The external artifact manifest is:

| Artifact | Size | SHA-256 |
|---|---:|---|
| `/home/zhuhe/code/litchi-worktrees/uo2-iwa-w1g/target/profiling/deps/parity-15e248d1d90252bd` | 472,404,512 bytes | `a13fceee820290d82e4f70a6fb0dd02f2a1a3164295874df55ca0fa711ffbbb5` |
| `/home/zhuhe/code/litchi-worktrees/uo2-iwa-w1g/target/profiling/deps/native_hidden_axes-007c6be3300edad5` | 34,327,072 bytes | `31659bbaa0f9cdc44b427f3df59d4ad802ee7b0afbf7918807d51527a9786ebd` |
| `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g/cg/current-numbers-host.out` | 465,419 bytes | `9c906f4b5a77a91933f8b28f2ff182567c7cb4c97b15864b93c36d506cfc5224` |
| `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g/cg/current-numbers-leaf.out` | 618,093 bytes | `aa13f8db95fbb50eff795e25c021d8a4be68209a07da2fa15664181da05acd5a` |
| `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g/cg/current-keynote-host.out` | 687,987 bytes | `655bdb7a7f8f28dd649ba8d378eb614169d6396bdbd3053444c5f3c4d37b341e` |
| `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g/cg/current-keynote-leaf.out` | 705,093 bytes | `394ab2d4fcbfa7cdda94dd6ed16c6738fb882550ddab3de82d35706a5cde7ec3` |
| `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g/cg/current-pages.out` | 599,645 bytes | `5baad48d0c85f6a81311c03e82f5cbd551e8da588ada687807b8623b0c499d1c` |

The two binaries are in this worktree's untracked `target/`; the callgrind
captures are in the external coordinator directory. None is repository
evidence. To regenerate them, source `iwa-w1g/env.sh`, build with
`CARGO_BUILD_JOBS=4 cargo test --profile profiling -p litchi-iwa --test parity
--all-features --no-run --locked` and
`CARGO_BUILD_JOBS=4 cargo test --profile profiling -p litchi-pages --test
native_hidden_axes --no-run --locked`, then run each ignored test with
`taskset -c 3 valgrind --tool=callgrind --quiet
--callgrind-out-file=<external-output> --toggle-collect=<probe>
<profiling-binary> --ignored --exact <test> --test-threads=1` in A, B, B, A
order. The Numbers and Keynote test names are
`numbers_hidden_axes_cases::numbers_hidden_axes_cost_abba` and
`keynote_hidden_axes_cases::keynote_hidden_axes_cost_abba`; the Pages test is
`hidden_axes_cost_probe`.

### 4.2 Root-comment facade rerun

The root-comment probes run on clean source commit `cfc334e754cc2e2edaba2598d31ef21325aae256`,
using the release-derived `profiling` profile (`opt-level = 3`, `lto = true`,
debug information retained) and CPU 3 pinned with `taskset`. Each ignored test
uses the shared `pages_cost_probes::abba` helper (host, leaf, leaf, host); the
Callgrind collection toggles isolate `probe_host` and `probe_leaf`. These are
full facade routes, including package open, selector resolution, edit, commit
and serialization. The Numbers host closure is the explicit parity-only
legacy writer, while the Pages source is host-authored before ABBA and the
Keynote/Numbers sources are native fixtures. This is source/build evidence,
not Apple open/save evidence.

| Route | Host instructions | Leaf instructions | Δ |
|---|---:|---:|---:|
| Pages `pages_body_table_comment_cost_abba` | 2,591,437,381 | 90,265,537 | **−96.517%** |
| Keynote `keynote_slide_table_comment_cost_abba` | 656,778,207 | 416,237,421 | **−36.624%** |
| Numbers `numbers_table_cell_comment_cost_abba` | 73,190,152 | 73,889,963 | **+0.956%** |

The external artifact manifest is below. The profiling binary is intentionally
not committed (497,767,840 bytes); all captures are below the repository's
evidence threshold and remain outside the checkout:

| Artifact | Size | SHA-256 |
|---|---:|---|
| `/home/zhuhe/code/litchi-worktrees/uo2-iwa-w1g/target/profiling/deps/parity-a78c3cd1bd84cfcd` | 497,767,840 bytes | `e01dd99c4f027fdb33350b876436ed1e2f27ec101f5d1a59b9f5660952291fee` |
| `comments-cfc334e754/pages-host.out` | 564,259 bytes | `4ae0ee2a654f736c5c7c8b1d71d7bd6024e7382954fac9288d14c83c9cd1922b` |
| `comments-cfc334e754/pages-leaf.out` | 657,162 bytes | `6ea7bf272312b5b6e18414ae544bdf097f6e2b5eebe92cc36fae2d3707a2c744` |
| `comments-cfc334e754/keynote-host.out` | 760,833 bytes | `6041c0866f9128eae85dddf64c5f44b44875556eac2abb34928b8fbd674b8fc2` |
| `comments-cfc334e754/keynote-leaf.out` | 731,394 bytes | `c4e11693ead3484856cc0b4b9dfc187f54200f126b975ba9e84a927eaaf02497` |
| `comments-cfc334e754/numbers-host.out` | 501,435 bytes | `8c129af94453bac47f3c3d9a0ccf0de7a2525ac5acb3adc5b6cef7a98b95935f` |
| `comments-cfc334e754/numbers-leaf.out` | 530,672 bytes | `626d1af2f31bb97201fa47e789db8a7c45bdcb790be95bd26e8bbb8dc26683ce` |

The external paths above are relative to
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g/cg/`. Regenerate the
binary with:

```sh
PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:/home/zhuhe/.cargo/bin:/usr/bin:/bin \
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-iwa-w1g/target \
CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_BUILD_JOBS=4 \
cargo test --locked --profile profiling -p litchi-iwa \
  --features parity-host-oracle --test parity --no-run
```

Then run each route twice, changing `--toggle-collect` and the output path:

```sh
taskset -c 3 valgrind --tool=callgrind --quiet --error-exitcode=1 \
  --callgrind-out-file=/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g/cg/comments-cfc334e754/<route>.out \
  '--toggle-collect=*probe_host*' \
  /home/zhuhe/code/litchi-worktrees/uo2-iwa-w1g/target/profiling/deps/parity-a78c3cd1bd84cfcd \
  --ignored --exact <parity-test> --test-threads=1
```

Use `*probe_leaf*` for the paired leaf capture; `<parity-test>` is one of
`pages_table_cases::pages_body_table_comment_cost_abba`,
`keynote_table_cases::keynote_slide_table_comment_cost_abba` and
`numbers_table_comment_cases::numbers_table_cell_comment_cost_abba`.

### 4.3 Root-comment read/create/clear rerun

The remaining supported root lifecycle routes were measured on clean source
commit `03f991fcf6`, after the typed comment/table resource-limit repair and
before the proof-only Pages read witness/ledger correction in `a5795869a8`.
That later commit changes no route implementation; the coordinator must still
repeat the captures on the final integrated tree. The
release binary was built with `CARGO_INCREMENTAL=0`, release debug information
disabled, and four build jobs:

```sh
PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:/home/zhuhe/.cargo/bin:/usr/bin:/bin \
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-iwa-w1g/target/profiling \
CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_BUILD_JOBS=4 \
cargo test --locked --release -p litchi-iwa --test parity \
  --features parity-host-oracle --no-run
```

The profiling binary is
`/home/zhuhe/code/litchi-worktrees/uo2-iwa-w1g/target/profiling/release/deps/parity-de33078d71de6252`,
40,900,464 bytes, SHA-256
`bc2e5ff92bf69cb0e88365b18f3bee0bb036e4a19798a00c7432d21e421871e0`.
Each ignored test ran on CPU 3 with `taskset`, Callgrind collection toggled
to the host or leaf probe, and the shared ABBA helper's host/leaf/leaf/host
order. The Pages read fixture and Keynote/Numbers populated read fixtures are
host-authored before the ABBA loop; create starts from a stored no-comment
cell; clear starts from one host-authored populated root prepared outside the
loop. Numbers' host side uses only the explicit feature-gated legacy oracle
described in §2b. These are full facade instruction counts, not wall time and
not Apple open/save evidence.

| Route | Host instructions | Leaf instructions | Δ |
|---|---:|---:|---:|
| Pages read | 1,232,839,805 | 34,536,873 | **−97.199%** |
| Pages create | 2,736,619,111 | 99,911,833 | **−96.349%** |
| Pages clear | 2,739,510,163 | 117,166,862 | **−95.723%** |
| Keynote read | 20,343,949 | 13,296,054 | **−34.644%** |
| Keynote create | 905,136,065 | 429,732,092 | **−52.523%** |
| Keynote clear | 58,029,332 | 37,518,767 | **−35.345%** |
| Numbers read | 3,699,962 | 2,973,596 | **−19.632%** |
| Numbers create | 169,394,744 | 115,553,384 | **−31.785%** |
| Numbers clear | 128,331,491 | 107,325,718 | **−16.368%** |

All nine leaf routes are below their paired host route and therefore meet the
normal instruction-cost bar. The coordinator must repeat these probes on the
final integrated trunk candidate; this WP result does not qualify a later
merge.

The raw captures remain outside the repository under
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g/cg/comments-read-create-clear/`:

| Artifact | Size | SHA-256 |
|---|---:|---|
| `pages-read-host.callgrind.out` | 112,982 bytes | `fdcc5a78012e6bde66c938ae4e48a16dcd23a1df97fc070dec9e79d521e14844` |
| `pages-read-leaf.callgrind.out` | 103,085 bytes | `5b3506631e2e46dad82e42f8446fb21bf8956d1838608fb363f8886c1d42278b` |
| `pages-create-host.callgrind.out` | 188,117 bytes | `7ceb2eca367ee5d6ffa1476862e897ef6dfe9bf087f8204ba176190d51a3892d` |
| `pages-create-leaf.callgrind.out` | 172,982 bytes | `53710f9fc80cfeb95c3e385224f2f1a7ae4c8e96e6e015df4ee35853320f1e9a` |
| `pages-clear-host.callgrind.out` | 175,901 bytes | `9f36ec5c83315133772fc0e22dc9c3a2618405a4846dafe54c0b559a5b3f9eda` |
| `pages-clear-leaf.callgrind.out` | 176,429 bytes | `2347b3a5f8429bf34c6653ef68ca918258553ed0d0732614e6ff85275e4448bf` |
| `keynote-read-host.callgrind.out` | 121,434 bytes | `7d7edbc8219aa98e7e41f1d14b49b59f0e66f8b6958b89661357a333381824cb` |
| `keynote-read-leaf.callgrind.out` | 90,999 bytes | `57d74a521cbf9eb6c63527fa2ce2cc6b801fed0be169cb57ccf08171470db6cc` |
| `keynote-create-host.callgrind.out` | 227,311 bytes | `fa42dcbabd470f0d3a670d102adb857bc65a559f02cd64d3c2d1ff67468a64c8` |
| `keynote-create-leaf.callgrind.out` | 176,155 bytes | `c47577a5f79385878919b7e22661c3c5e925ac02dd235eebb80d1b3e01f0b686` |
| `keynote-clear-host.callgrind.out` | 174,022 bytes | `32d8a703b537e5fc9a4aa0c49d639f9ab4275d779bc664286e11e00ac49b2686` |
| `keynote-clear-leaf.callgrind.out` | 158,398 bytes | `e7742c629eeea91c113f429015861121af92ae20316768ec9e0e7028df0c64e8` |
| `numbers-read-host.callgrind.out` | 89,085 bytes | `7b0c7214436e917775d2d8c8f84c781337b3c717b494b2d7949054bbf4e33ce3` |
| `numbers-read-leaf.callgrind.out` | 90,650 bytes | `794786b03bfa316f7e925cd7ab5d36a20e214abf2e6304b5485e8c0072f8bbd0` |
| `numbers-create-host.callgrind.out` | 168,655 bytes | `1a346e2a4097f825d237ce8306c9bd98162f2b3a0e21e7662a8844163e27f199` |
| `numbers-create-leaf.callgrind.out` | 149,413 bytes | `c50dee29f05b368164e1592a5f5c8228320fb36bb198241ab681ad034c5b32f7` |
| `numbers-clear-host.callgrind.out` | 157,398 bytes | `7da66a7a7ece4160cc47865dfc38ada8172726f03a9a44dec6bd6f3570083d7e` |
| `numbers-clear-leaf.callgrind.out` | 140,080 bytes | `4da8e64c76118ddebaa7dbcf59b641d6b3e9f32e6652d9021a7a15f0cbd951ee` |

To regenerate the captures, build the binary with the command above and run
each test with `taskset -c 3 valgrind --tool=callgrind --quiet
--error-exitcode=1 --callgrind-out-file=<external-output> <binary> --ignored
--exact <test> --test-threads=1`. Use
`--toggle-collect='*probe_host*'` for the host capture and
`--toggle-collect='*probe_leaf*'` for the paired leaf capture, changing the
output and toggle for each route. The exact test names are the nine
`*_comment_{read,create,clear}_cost_abba` functions in the Pages, Keynote and
Numbers parity modules.

### 4.4 Canonical-API twelve-route matrix

The previous captures are historical: they were taken from pre-repair source
`e09f2269d0`, before the duplicate Pages/Keynote Package setters, clearers and
plural reply getters were removed. The historical route values remain in the
external capture manifest and are not a qualification of the current public
API. A fresh matrix must be captured after the repair and recorded against
one clean source commit. The final route set is Pages, Keynote and Numbers
read/create/Set/clear (twelve full facade routes), with Pages/Keynote writes
using `Package::edit_*().set/clear()?.commit()` and read replies projected
from the root `Comment::replies()` snapshot.

## 5. Gaps

| Gap | Finding | Next step |
|---|---|---|
| Automatic to a decimal family other than Number, and cross-family changes on decimal cells | the engine admits a family only on a cell whose format entry already belongs to it, or an empty one; the host's compatibility writer converts (the Numbers 0370a §5 gap) | lift a conversion path into `data_format::rewrite` with parity against the host's `cell_data_format` writer |
| Date and Time, Duration, Text on number cells | the host writes a number under a date or duration kind, which the engine reads as an invalid cell; the host refuses Text on a number cell (pinned by `*_divergence_is_reported`) | decide with the macOS run which shape native files store |
| Custom, numeral system, controls, Pop-Up Menu for Pages and Keynote | read only; writes need the metadata transition and the control facade's publication | lift the Numbers control/Pop-Up Menu publication (`rewrite_transaction`, save tokens, preview policy) into the crate |
| empty cells | the engine edits stored cells only (`CellNotFound` on an empty slot); the host first stores the cell | reuse `table_write` to store a style-only cell first |
| hidden axes on tables whose formula owner has dependencies (`slide-table-read-native.key`, nine Numbers fixtures) | refused (`InvalidSource`): the engine proves an empty dependency graph before changing hidden state (§2a.3); the host edits them | re-plan the owner's dependencies with the W1-F formula engine |
| hidden-axis owner creation | the engine registers the new objects' UUIDs and writes empty filter offsets; the host does neither (pinned on Numbers) | macOS run decides |
| clearing hidden axes | the engine keeps cleared state records, the host drops them (pinned) | macOS run decides |
| group 2 | root-only table-cell comment Set/create/Clear uses the shared native graph/rewrite engine in all three leaves; reply-thread mutation rows remain host-only and all native evidence is pending | coordinator repeats the cost probes on the final integrated tree; then lift replies and clear the queued native routes |
| groups 3–5 | cell rich text, topology orchestration and conditional highlights are not started | next W1-G slices, in that order |

## 6. Verification

Build env of the coordinator rules (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`), sequential gate scripts outside the repository
(`.coord/iwa-w1g/gates{1,2}.sh`). Both pre-commit size checks printed
nothing for every commit.

| Command | `35908bbc0a` | `73f6fa976a` |
|---|---|---|
| `cargo fmt --all --check` | 0 | 0 |
| `cargo clippy -p litchi-iwa-table -p litchi-numbers -p litchi-pages -p litchi-keynote -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 (without pages, keynote, iwa) | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features numbers`, `pages`, `keynote` | 0, 0, —, — | 0, 0, 0, 0 |
| `cargo test --all-features` for `litchi-iwa-table`, `litchi-numbers`, `litchi-pages`, `litchi-keynote`, `litchi-iwa` (every target, `parity` included) | 0, 0, —, —, — | 0 each |
| `cargo test -p litchi --all-features` | 0 | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, nine groups of `--test` targets | 0 each | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-table -p litchi-numbers -p litchi-pages -p litchi-keynote` | 0 | 101 (two module links), 0 after `b21879b779` |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (every profile) | 0 | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 (69 packages, 300 declarations, 11 debts) | 0 (302 declarations) |
| `python3 tools/non_iwork_gate.py verify` | 0 | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` | 0 | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` | 0, 0, 0 | 0, 0, 1 (matrix stale for this record), 0 after `bd243a541d` |
| `python3 tools/check_feature_matrix.py` | 0 | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 | 0 |
| `cargo check --locked --offline --manifest-path crates/<c>/fuzz/Cargo.toml --bins` for `litchi`, `litchi-numbers` (`7ea1e295b1`), `litchi-pages`, `litchi-keynote` | 0 | 0 |

The capability and registry goldens and the OfficeCLI corpus pass unchanged
inside the litchi-ops chunks; no `litchi-ops` dispatch changed (owner
decision D7), so they were not re-blessed.

The `parity` integration target remains enabled under the default feature set.
Only the two Numbers mutation cases that need the independent legacy host
oracle (`numbers_table_cell_comment_set_parity` and
`numbers_table_cell_comment_host_oracle_control`) are compiled with the
explicit `parity-host-oracle` feature. The root-read witness is a doc-hidden
selector/position seam used only by the differential test; it is not a
supported editor API. Re-run the independent oracle cases with
`CARGO_BUILD_JOBS=4 cargo test -p litchi-iwa --features parity-host-oracle
--test parity numbers_table_comment_cases:: -- --nocapture`.

The table above is historical evidence from the earlier group-1 and group-6
commits; it is not a complete gate for the current integrated tree.  The
review follow-up was rerun on final source commit `86bd7f5e3a`, based on
merged candidate `4abb90f713` (code fixes `828d02cad9` and `86bd7f5e3a`, with the current
`feat/unified-ops-api` parent), as follows:

- the owner-extents unit tests (including the non-canonical value-varint
  refusal), Numbers hidden-axis facade tests (including the low-limit
  candidate readback regression), and the active Numbers/Keynote parity
  cases passed on that source tree;
- `cargo fmt --all --check`, the focused `litchi-iwa-table`/`litchi-numbers`
  Clippy gate, the exit-ledger check, crate-boundary and API-glossary checks,
  the non-iWork gate, and the affected no-`3rdparty` tests were rerun on this
  source tree;
- the release callgrind ABBA run in §4.1 was built and measured on
  `d82fc08989`, immediately before `86bd7f5e3a`; the latter only adds the
  semantic workbook-limit mapper and is not exercised by those host/leaf
  probes. It measured −18.097% for Numbers and −60.422% for Keynote, while
  the Pages focused route was 34,078,005 instructions. These are slice
  evidence, not the coordinator's final integrated-tree qualification.

The current group-2 checks are separate from the historical rows above. On
source commit `47f72dc6a0`, the focused parity runs passed: Numbers default
read (1/1), Numbers with `parity-host-oracle` (4/4), Pages comment cases
(3/3), Keynote comment cases (6/6), and the shared duplicate-storage unit
regression (1/1). The Keynote cases
include native D2 root read/clear, a host-authored populated reply read,
source-metadata-preserving replacement, and an unrelated-properties
patch-conflict control. The Numbers set case uses the explicit hidden legacy
oracle seam and its control; its populated reply read source is host-authored
before serialization. Pages uses a host-authored no-reply root for mutation
and table 0 B2 from the populated native comment fixture for root/reply read.
The three affected fuzz-lock manifests (`litchi-numbers`, `litchi-pages`, and
`litchi-keynote`) also passed `cargo check --locked --offline --bins` on this
source. The focused all-target/all-feature compile and Clippy gates passed,
and the full-facade cost probes in §4.2 passed on clean source commit
`cfc334e754`. These are differential proofs, not Apple open/save evidence;
all three routes remain `Evidence::Queued`.

For the handoff tree, `feat/unified-ops-api` at `b2a2e4c8` was merged into this
WP as `281aa05899`. On that merged tree the focused parity cases were rerun:
the shared duplicate-storage regression (1/1), Pages (3/3), Keynote (6/6),
Numbers default (1/1), and Numbers with `parity-host-oracle` (4/4). The three
affected fuzz-lock manifests, the ledger, crate-boundary, glossary,
no-`3rdparty`, and focused formatting checks also pass on the merged tree.
Those post-merge checks do not requalify the §4.2 measurements: the
coordinator must repeat the full gate and final performance on one integrated
trunk commit, and native verification remains pending.

On the current source commit `ac9f1e8e2b` (a rustfmt-only follow-up to the
proof witness), the focused comment parity target passed 14 cases with zero
failures, including the dedicated Pages comments
read witness and the populated Numbers semantic reply comparison; the
low-limit transaction/readback regression passed 1/1
and preserved the exact source while returning `TextBytes`. The all-target,
all-feature compile passed, as did library Clippy and the directly affected
Numbers comment-reply and `litchi-iwa` parity Clippy targets. The three
affected fuzz-lock manifests, exit ledger, crate-boundary, no-`3rdparty`,
formatting and rustdoc/glossary checks are rerun for the final handoff. The
broad all-target Clippy command still reports the pre-existing
`items_after_test_module` lint in Pages'
`package/body_drawable_graph.rs`; that file is outside this slice and was not
changed. The nine current-source read/create/clear Callgrind ABBA pairs and
their exact external artifacts are recorded in §4.3; the coordinator must
repeat them on the final integrated trunk candidate.

Continuation proof checkpoint (source commit `0acc4231ee`, after the Numbers
parity checkpoint `63d96fcfbf` was handed to F): the Pages comment target ran
with `parity-host-oracle` and passed the six focused checks (read, populated
reply read, root replacement, root create, root clear, and the exact surviving
storage witness; 6 passed, 4 ignored cost probes). The create witness is the
stored B2 cell in
`body-table-cells-native.pages`; it uses the selected-cell host and leaf
readers, validates the exact generated type-212 author profile and existing
empty type-213 update, proves source-wide fresh identities and exact metadata
watermarks, and exercises wrong-root, unrelated 6005/header-reference,
metadata-watermark, extra-author-allocation, generated-author archive-shape,
selected-comment unknown/text, and author-storage corruption controls. Its
creation-only 3056 mask rewrites only date/UUID bytes after rejecting unknown,
duplicate and noncanonical nested date/reference/UUID/reply fields; raw
unknowns in selected 6005 and metadata reference paths are explicitly shown
to survive normalization. A separate proof-only read witness on
`body-table-comment-hidden-native-saved.pages` checks both source-owned roots
and replies (two type-3056 payloads) byte-for-byte on host and leaf and rejects
an altered surviving payload. The feature-enabled all-target parity Clippy
command passed after replacing the obfuscated reference normalizer; the
default-feature parity target also compiled with `--no-run`, and
`cargo fmt --all -- --check` passed. The queue correction in this checkpoint
records the actual D2 Keynote/native and generated host-authored reply
fixtures, names only the folded Numbers
`numbers_table_cell_comment_replies_read_parity` semantic case, and separates
supported root create/set/clear native acceptance from deferred host-only
reply operations. No new release performance capture was made here: the
historical captures in §§4.2–4.4 remain non-qualifying until F's reviewed
Numbers proof and the coordinator's final clean-source twelve-route matrix.

Combined handoff source `444c331122` is a real merge of the Pages proof
follow-up and F's accepted Numbers proof tip `4b1f3000bc`. On that merged
source, the Numbers comment target passed nine non-cost checks: root read with
replies, set, create, create-identity mask control, clear, populated-neighbor
clear, clear policy controls, clear storage/text controls, and the independent
host-oracle control. The Numbers create/clear controls retain source-owned
3056 metadata and reject identity, locality, duplicate, and low-limit
corruption; their cost probes use the forward-only leaf route
(open/edit/commit/serialize), while older captures that included inverse
application remain historical and do not qualify final performance. The
merged source also passed the six Pages comment checks above, formatting,
feature-enabled all-target parity Clippy, and default-feature parity
`--no-run`; no final release performance capture was made after this merge.
