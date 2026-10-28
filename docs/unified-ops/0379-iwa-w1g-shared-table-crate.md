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
  checker reports an unchanged pre-existing Keynote drawable-comment export
  debt; final integrated performance and native verification remain pending
  (§2b, §6). Groups 3–5 (rich text, topology, conditional highlights) are not
  started (§5).
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
  and 18.1% below their host routes. Group 2 has no performance claim yet;
  measurement is held for the coordinator's quiet window.

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

- Pages: `body_table_cell_comment`, `body_table_cell_comment_replies`,
  `edit_body_table_cell_comment`, `apply_body_table_cell_comment`,
  `set_body_table_cell_comment` and `clear_body_table_cell_comment`;
- Keynote: the corresponding `slide_table_cell_comment*` methods; and
- Numbers: `table_cell_comment_replies`, `set_table_cell_comment`,
  `clear_table_cell_comment`, and the existing reply methods.

Reply reads remain supported. Reply mutations remain explicit terminal
`UnsupportedDependency` refusals in the Pages and Keynote focused route, and
Numbers preserves its existing reply-bearing replacement refusal. Clearing or
replacing a root that has replies is therefore not claimed as a complete host
migration in this slice; the host-only ledger rows stay host-only until the
reply-thread lifecycle is lifted. No refusal falls through to the legacy host.

The differential cases currently run are:

| Format | Cases | Tier |
|---|---|---|
| Pages | `pages_body_table_comment_set_parity`, `pages_body_table_comment_clear_parity` | B, root-only |
| Keynote | `keynote_slide_table_comment_set_parity`, `keynote_slide_table_comment_clear_parity`, `keynote_slide_table_comment_read_parity`, `keynote_slide_table_comment_replies_read_parity` | B for root writes; A for reads |
| Numbers | `numbers_table_cell_comment_set_parity` (including reply projection readback) | B, root-only |

The host remains the differential oracle. The 17 table-comment ledger rows
were reviewed: five focused read rows are migrated with these tests, while
the twelve root/reply mutation rows remain `host-only` because reply-bearing
threads are outside this profile. The 0363 macOS queue retains the routes as
pending planning entries with that boundary called out; no E3/E4 native claim
is made. Native-authored fixtures and synthetic/self-round-trip tests are not
Apple open/save evidence.

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
| group 2 | root-only table-cell comment Set/create/Clear uses the shared native graph/rewrite engine in all three leaves; reply-thread mutation rows remain host-only and all native evidence is pending | finish scoped Clippy/docs and final integrated performance, then lift replies and clear the queued native routes |
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

The current group-2 checks are separate from the historical rows above. On the
continuation source tree, after the admission, codec-limit, typed-error and
reply-boundary repairs, the following passed: `cargo fmt --all --check`, a
four-crate library compile (`litchi-iwa-table`, `litchi-pages`,
`litchi-keynote`, `litchi-numbers`), the Numbers table-cell comment suite
(31/31), the host-vs-leaf table-comment parity target (7/7), both affected
fuzz-lock checks (`litchi-numbers/fuzz` and `litchi-pages/fuzz`), and
`python3 tools/iwa_exit_ledger.py --check`. The same source tree also passed
the four-crate all-targets/all-features Clippy gate (`-D warnings`), focused
rustdoc (`RUSTDOCFLAGS=-Dwarnings`), `check_api_glossary.py`, and the affected
no-`3rdparty` test. The current `check_crate_boundaries.py` run is not green:
it reports the unchanged pre-existing export at
`crates/litchi-keynote/src/lib.rs:506`
(`litchi_iwa_common::table::read` under a focused drawable-comment public
surface). This slice did not modify that file or export. The parity cases are
Tier B for root writes and Tier A for reads, all with `Evidence::Queued`;
they are differential proof, not Apple acceptance. No performance
qualification or native E3/E4 receipt is claimed. The coordinator must repeat
the full gate and final performance on one integrated trunk commit after this
branch is reviewed.
