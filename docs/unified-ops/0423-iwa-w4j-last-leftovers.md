# 0423 — W4-J: Keynote slide-table add, Numbers shape image fill

- Status: slices 1 and 2 merge-ready on branch `uo2/iwa-w4j` (trunk
  `41ee75653b` merged); the remaining items are in §5.
- Date: 2026-11-02.
- Scope (coordinator, wave W4-J): the last leftovers of record 0420 §6 —
  Keynote `add_slide_table`, Pages chart duplicate and remove, Numbers
  shape image fill, Keynote line endpoints on the shared style workspace,
  and (optional) Pages table insertion into a document without a table.
- `performance_claim`: every migrated route uses fewer instructions than
  its host route (§3).

## 1. What landed

| Commit | What |
|---|---|
| `2bb9fd3987` *(shared)* | `litchi_iwa_table::structure::create`, lifted from Numbers' `table::lifecycle::add` (draft of record 0420): `clone_template` (the host's `create_empty_table_graph`: the template's drawable and model and an emptied copy of each owned storage object, styles shared), `fresh` (the host's `bootstrap_empty_table_graph` on `table_create::fresh`), `theme_styles` (a theme's first table preset and its style network) and `add_owner` (the table's CalculationEngine formula owner). Numbers keeps its template choice, sheet attachment, sidebar node and registry order on top; its parity cases pass unchanged |
| `5eed0680de` | Keynote `slide::table::lifecycle::Edit::add(slide, name, rows, columns, frame)`: an emptied copy of the deck's detached creation table (a `TST.TableInfoArchive` with no parent; builder decks carry one), framed with flags 3 and angle 0 (or the frame's rotation) and sized uniformly through `litchi_iwa_table::topology::{stage_frame, stage_uniform_dimensions}` (record 0412g), placed in front of the slide's owned drawables and z-order, with its own formula owner and the clones' UUIDs and cross-component edges. New refusals `InvalidName`, `InvalidSize`, `InvalidGeometry`, `NoTemplate` |
| `4992428ea4`, then folded onto trunk's `litchi_iwa_style::data_owner` (record 0422) after the merge | Numbers `sheet::drawable::style::Edit::set_image_fill(bytes, file_name, technique, tint)` over the new shared `litchi_iwa_style::shape::set_image_fill`: the embedded fill is planned on the style graph like any fill, the bytes are registered as the next `DataInfo` owned by the variation the shape then shows (`data_owner::own_new`), and the variation's message info records the data reference; `style_graph::publish` adds the `Data/` member. Reading a fill now returns an image fill instead of refusing. New refusal `InvalidImage` |

The first version composed the graph's edits on the structural working
set with `Transition::{data_additions, data_owner_additions}`; the
coordinator asked for trunk's shared `data_owner` transitions instead, so
no parallel path remains.

## 2. Ledger and parity

| Host route | Destination | Parity test | Tier |
|---|---|---|---|
| Keynote `add_slide_table` | `litchi_keynote::Package::edit_slide_table_lifecycle` (`add`) | `keynote_slide_table_add_parity` (builder deck) | B |
| Numbers `set_sheet_shape_image_fill` | `litchi_numbers::Package::edit_sheet_drawable_style` (`set_image_fill`) | `numbers_shape_image_fill_parity` (scale-to-fill with tint; tile) | B |

- **Keynote add.** The output equals the host's on every known field. The
  table UUID is the shared seed-derived identity on the leaf and the
  model-derived `00000000-0000-4000-8000-…` form on the host, so the model's
  `table_id`, the owner UID and the engine's owner map are compared by form
  (the documented policy of the Numbers add, record 0370f). The drawable's
  field order differs (Tier B). The host registers no edge from the new
  drawable to its slide, and neither does the leaf.
- **Keynote decks without the creation table** (every native deck) are
  `Refusal::NoTemplate`, as the host refuses them
  (`keynote_slide_table_add_without_template_is_refused`); the drawable
  family's `add_table` still copies the first slide table there (record
  0329). The pinned divergence test is replaced by these two cases.
- Both routes are macOS-queued (`0363-mac-queue.md`, rows W4-J). The
  ledger check passes: Keynote 3, Numbers 33, Pages 14 host-only rows.

## 3. Performance

Callgrind, release parity binary with debug info (SHA-256
`9e1d1109032b76445887642cab7d74cab06ed50d72817b8fbd1020572e17e617`, after
the trunk merge; the table-add row is from the previous build
`ccb993bd…`),
CPU 3, ABBA (two calls of each route), inclusive instructions of the
probe functions:

| Route | Probe | Host Ir | Leaf Ir | Leaf/host |
|---|---|---:|---:|---:|
| Keynote slide-table add (builder deck, 3×4) | `cost_probes::keynote_slide_table_add_cost_abba` | 146,844,352 | 22,314,779 | 0.152 |
| Numbers shape image fill (tile) | `cost_probes::numbers_shape_image_fill_cost_abba` | 373,909,767 | 155,392,006 | 0.416 |
| Numbers table add after the lift (`basic.numbers`) | `cost_probes::numbers_table_add_cost_abba` | 1,186,221,466 | 104,850,970 | 0.088 (0420: 104,739,401, +0.1%) |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --all-features --no-run
taskset -c 3 valgrind --tool=callgrind target/release/deps/parity-<hash> --ignored --exact cost_probes::<probe> --test-threads=1
callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep -E 'cost_probes::(host|leaf)_'
```

## 4. Verification

The gate script `.coord/iwa-w4j/gate.sh` (crates litchi-iwa-table,
litchi-numbers, litchi-keynote and litchi-iwa with parity; ops features
`numbers` and `keynote`) runs on the merge of trunk; its result is in the
hand-off.

## 5. Left

| Item | Rows | Next |
|---|---:|---|
| Pages chart duplicate and remove | 2 | not started here; plan of record 0418 §6 |
| Keynote line endpoints on the shared workspace | — | not started: Keynote's shape styles are staged inside `edit_slide_drawables` on Keynote's own workspace, so moving them onto `litchi_iwa_style::graph::Workspace` is a refactor of `slide_drawable/shape_style.rs` (1,280 lines) and its staging |
| Keynote slide-table add without a creation table | — | the shared `fresh` graph needs the theme's first table style; native Keynote themes archive their `TST.TableStyleNetworkArchive` in two messages, which `table_create::fresh::preset_styles` refuses (`Invalid`) |
| Numbers image-fill replacement and release | — | `Refusal::ImageFill` when the shape already shows an image fill; use `data_owner::release` with `ReleasingFills` as the Pages 0415 route |
| Pages table insertion without a template (optional) | — | not started |
