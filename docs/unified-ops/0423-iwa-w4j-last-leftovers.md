# 0423 — W4-J: Keynote slide-table add, Numbers shape image fill

- Status: slices 1 and 2 merge-ready on branch `uo2/iwa-w4j` (trunk
  `ba2d1dbd82` merged); the remaining items are in §5.
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
| `4992428ea4` | Numbers `sheet::drawable::style::Edit::set_image_fill(bytes, file_name, technique, tint)`: the embedded fill is planned on the shared style graph, then the graph's edits are composed on the structural working set, where the bytes become a new `Data/` member with its `DataInfo` record, the drawable's style variation takes the data reference and becomes its registered owner (the shared `Transition::{data_additions, data_owner_additions, entry_insertions}`, the Pages 0413b sequence). Reading a fill now returns an image fill instead of refusing. New refusal `InvalidImage` |

`Transition` already had the insert side (`data_additions`,
`data_owner_additions`, `entry_insertions`; the Numbers media add uses
them), so no shared metadata field was added.

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
`ccb993bdff1c578e5078b99d53b0b5a4e667d250a03b87d5904377b3c07f7f35`),
CPU 3, ABBA (two calls of each route), inclusive instructions of the
probe functions:

| Route | Probe | Host Ir | Leaf Ir | Leaf/host |
|---|---|---:|---:|---:|
| Keynote slide-table add (builder deck, 3×4) | `cost_probes::keynote_slide_table_add_cost_abba` | 146,537,720 | 22,305,824 | 0.152 |
| Numbers shape image fill (tile) | `cost_probes::numbers_shape_image_fill_cost_abba` | 374,042,992 | 218,763,122 | 0.585 |
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
| Numbers image-fill replacement and release | — | `Refusal::ImageFill` when the shape already shows an image fill; port the Pages 0415 release (`ReleasingFills`, `stored_fill`, `data_owner_removals`, `data_removals`) |
| Pages table insertion without a template (optional) | — | not started |
