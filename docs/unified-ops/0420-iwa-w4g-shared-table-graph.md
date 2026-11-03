# 0420 — W4-G: the shared from-scratch table graph, Pages `Builder::body_table`, Numbers table add without a template

- Status: slice 1 on branch `uo2/iwa-w4g` (cut from trunk `bbf1e3b812`;
  trunk `d4121f4af7` merged). The package continues (§6).
- Date: 2026-11-02.
- Scope (coordinator, wave W4-G): a from-scratch table graph in
  `litchi_numbers_wire::table_create` consumed by Pages
  `Builder::body_table`, Keynote `add_slide_table` and the Numbers builder
  and table-add paths; Pages chart duplicate and remove; the shared
  `document.rs charts` row; Numbers shape image fill; Keynote line-end
  resolution on the shared `graph::Workspace`.
- `performance_claim`: the Pages builder route costs 13.7% fewer
  instructions than the host route, Numbers table add 79.8–91.2% fewer;
  the Numbers builder (now on the shared graph) is within 0.8% of the host
  (§4).

## 1. What landed

| Commit | What |
|---|---|
| `204c40a95c` *(shared)* | `litchi_numbers_wire::table_create::fresh`: the host's table construction without a template. `fresh::table(Ids, Placement, Styles, Shape, uuid, uid_seed, Profile)` returns the drawable, model, tile, header buckets, data lists, UID map and stroke sidecar as `Object { identifier, message_type, payload, references }`, each payload in the host's Prost field order; `preset`, `style_network`, `default_styles` (the host Numbers builder's table stylesheet), `engine`, `preset_network` and `preset_styles` (a theme preset's stylesheet context). `Profile::Builder` is the host builders' table (four data lists, the builder's reference order); `Profile::Bootstrap` the host editor's table for a workbook without a table (`bootstrap_empty_table_graph`: a control-cell-spec list, the BNC format-table reference, sorted model references). Leaf-neutral like `structure::duplicate`: the caller supplies identifiers, the parent and frame, and the stylesheet context, and attaches the drawable. `TableUuid::from_halves` lets an owner keep its own seed policy |
| `0858c753d8` | Numbers on the shared graph. The builder's table rewrite (name, size) emits `fresh::table` with the template's identifiers, styles and identity; `the_template_table_is_the_shared_fresh_graph` proves the graph reproduces the template's twelve table, preset and network payloads byte for byte. `table::lifecycle::add` drops its private copies of the `table_create` payload builders (owned storage, model, emptied storage, owner, engine registration; about 400 lines) and, for a workbook without a table to clone, builds the shared graph from the theme's first table preset (previously `Refusal::NoTemplate`, which now means "no table and no theme table preset") |
| `43724160de` | Pages `Builder::body_table(name, rows, columns)` and `PackageError::InvalidTable`. The first table is the host builder's graph under the host's identifiers 9–41: the shared table with its preset, network and default stylesheet, the attachment anchored by `U+FFFC` after the body text, the document root's engine, function-browser-state and custom-format-list references, the theme's table preset, a new `CalculationEngine` component (between the annotation storage and the document, as the host lists it) with the table's formula owner, and the registry's UUIDs and references. No bundled table template and no Pages→Numbers edge. Later tables clone the first through `table::creation`, as the host builder adds them with its editor |
| `b63ffa6c95` | callgrind probes |
| `11da80653b` | the shared `document.rs charts` row retires to the leaf chart-metadata reads (`document_charts_read_parity`, Tier A on `chart-data-native.{pages,numbers,key}`); helpers the trunk merge left unused |

## 2. Ledger and parity

| Host route | Destination | Parity | Tier |
|---|---|---|---|
| `pages/creation.rs body_table` | `litchi_pages::builder::Builder::body_table` | `pages_builder_body_table_parity` (one table after body text, one in an empty body, two tables) | B |
| `numbers add_empty_table` on a table-less workbook (row already migrated) | `edit_table_lifecycle().add` | `numbers_table_add_without_template_parity` | B |

One ledger row flips to migrated (`pages/creation.rs body_table`) and one retires (`document.rs charts`: the host's format-neutral list is the leaf `body_chart_metadata`, `sheet_chart_metadata` and `slide_chart_metadata` over every chart selector; `document_charts_read_parity` compares the host list with the leaf reads of every chart, as multisets, on one native file per format); `python3
tools/iwa_exit_ledger.py --check` passes. Both routes have W4-G rows in
`0363-mac-queue.md` (no host native evidence).

Tier B: the host draws the table UUID, the owner UID, the UID-map
identities and the registry's object UUIDs at random; the leaf derives
them from the document identity (the documented policy), and the cases
compare them by form. Every other changed object agrees with the host on
every known field (32 changed objects in the Pages case; a deliberate
change of the owner's internal identifier fails the case, so the harness
sees the engine member). The bytes differ only in the UUID-bearing
payloads, the body storage (the attachment table is appended, the host
writes it in tag order) and the registry.

Numbers parity is unchanged: `numbers_lifecycle_cases` (add, remove,
duplicate and resize on `basic.numbers` and the corpus) and
`numbers_builder_cases` (names, dimensions, language and locale) pass
before and after the move.

## 3. Findings

- The host Numbers builder's table is exactly reproducible from typed
  values: the template the leaf builder embeds (`create/blank.numbers`, the
  host builder's output) matches the shared graph byte for byte apart from
  its random UID map.
- The host Pages builder imports that table with a stylesheet of its own
  (object 40, unparented), separate from the Pages document stylesheet;
  the leaf writes the same structure.
- Pages table insertion (`table::creation`) still refuses a document
  without a body table (`Refusal::NoTemplate`), while the host's
  `Draft::add_table` bootstraps the first table on a scratch document.
  The builder module can serve that route (follow-up, §6).

## 4. Performance

Callgrind on a release build with debug info, CPU 3, ABBA, two calls of
each route; raw outputs in `.coord/iwa-w4g/cg/`.

| Route | Probe | Host Ir | Leaf Ir | Δ |
|---|---|---:|---:|---:|
| Pages builder, body text and a 5×4 table | `pages_builder_cases::builder_body_table_cost_abba` | 11,717,813 | 10,115,217 | −13.7% |
| Numbers table add (template clone, `basic.numbers`) | `cost_probes::numbers_table_add_cost_abba` | 1,186,367,946 | 104,739,401 | −91.2% |
| Numbers table add without a template | `cost_probes::numbers_table_add_without_template_cost_abba` | 33,729,174 | 6,803,155 | −79.8% |
| Numbers builder, 20×8, `fr`/`fr_FR` | `cost_probes::numbers_builder_cost_abba` | 2,502,506 | 2,523,095 | +0.8% |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --all-features --no-run
taskset -c 3 valgrind --tool=callgrind target/release/deps/parity-<hash> --ignored --exact <probe> --test-threads=1
callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep -E 'probe_(host|leaf)|(host|leaf)_(table_add|builder)'
```

## 5. Verification

Full gate (`.coord/iwa-w4g/gate.sh`, log `gate-s1b.log`) on `11da80653b`, crates
litchi-numbers-wire, litchi-numbers, litchi-pages and litchi-iwa (parity included), ops
features `numbers` and `pages`: 32 of 33 rows exit 0. Clippy (all and no-feature ops),
the crate tests, the litchi-ops lib/doc/test groups, rustdoc, boundaries, non-iWork,
glossary, pyo3, perf-baseline, the tool unit tests, ledger, delta, corpus, bench matrix
and feature matrix all pass. `cargo fmt --check` failed on one test file and was fixed
in `233cafee09`. After merging trunk `c599ef9cb8` (`5b2ecc75be`), fmt, strict clippy on
the changed crates, boundaries, the ledger, the bench matrix and the slice's parity
cases (23) exit 0. No litchi-ops dispatch arm changed, so the goldens were not re-blessed.

## 6. Left

| Item | Rows | Next |
|---|---:|---|
| Keynote `add_slide_table` | 1 | The leaf `edit_slide_drawables().add_table` (record 0329) clones the deck's first slide table on Keynote's own drawable workspace; the host clones the deck's detached creation scaffold (a `TST.TableInfoArchive` with no parent, which builder decks carry), keeps the scaffold's frame, then sets the frame, flags 3 and angle 0 and sizes every row and column uniformly in the header buckets (`set_uniform_table_dimensions`). `keynote_slide_table_add_divergence_is_reported` pins that no deck admits both. Next: when a deck has exactly one detached scaffold, clone it instead of the first slide table, with the model's default sizes kept (`Shape::row_height`/`column_width` `None`), then run the frame and uniform-dimension staging of `litchi_iwa_table::topology::{stage_frame, stage_uniform_dimensions}` (record 0412g) on the clone; turn the divergence test into a Tier B parity case on a builder deck. A lift of Numbers' template clone and owner creation into `litchi_iwa_table::structure::create` (draft in `.coord/iwa-w4g/create.rs.draft`) serves it if Keynote moves the add onto the shared `Working` |
| Pages table insertion without a template | — | route `table::creation` on a scratch document through `builder_table::install` at the anchor, as the host's `bootstrap_first_table_graph` |
| Pages chart duplicate and remove | 2 | not started; 0418 §6 has the plan (a chart variant on Pages `DrawableSelector` or `chart::lifecycle::Edit::{duplicate, remove}` mirroring Numbers `sheet_drawable_chart.rs`, with the attachment and U+FFFC handling of `body_media_lifecycle`) |
| Numbers shape image fill | 1 | not started: `sheet_drawable_style` publishes through `numbers_style_graph::publish`, which carries no data-member insertion and no registry data transitions; it needs the Pages 0413b/0415 sequence (`stage_image`, `own_image`, `released_image`, `ReleasingFills`) on `litchi_iwa_table::structure::metadata::Transition::{data_additions, data_owner_additions, data_record_removals}` |
| Keynote line endpoints on the shared workspace | — | not started (0418 §1 notes the workspace mismatch) |
