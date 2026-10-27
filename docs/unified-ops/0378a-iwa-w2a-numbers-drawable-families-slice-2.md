# 0378a — W2-A slice 2: Numbers text, comments, media and chart series out of `litchi-iwa`

- Status: second merge-ready slice of wave W2-A on branch `uo2/iwa-w2a`
  (after [0378](0378-iwa-w2a-numbers-sheet-drawables.md); trunk
  `feat/unified-ops-api` merged at `2c9c985b00`, `c22bc02334`, `ea94815c6e`
  and the head merge of this slice). The wave continues (§6).
- Date: 2026-10-27.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave row W2-A; harness and
  ledger [0369](0369-iwa-exit-parity-harness-and-ledger.md); shared engines
  W1-D ([0374](0374-iwa-w1d-shared-text-engine.md): `list`, `drop_cap`,
  `ranged`, `smart_field`, `annotation`, `paragraph_style::named`), W1-A's
  storage-list publisher (0370b), W1-E
  ([0375a](0375a-iwa-w1e-style-variation-write-engine.md):
  `litchi-iwa-style` — `graph`, `series`, `drawing`, `media`; the
  `chart_style_codec::property` series-style rules).
- Donor: `feat/numbers-drawable-comments-focused-rereview-20261024` (the
  pre-0351 drawable-comment lifecycle) — re-implemented on the structure
  working set rather than copied (its 6k-line admission ledger duplicates
  what `structure::Working` and the metadata transition already enforce).
- `performance_claim`: every new route is 20–84% cheaper than its host
  route under callgrind ABBA (§5).

## 1. What landed

All code is in `litchi-numbers` (owned paths `src/package/sheet_{drawable,
chart,media}*`, `src/package/style_graph.rs`, `src/sheet/{drawable,
chart}.rs`) plus separated hunks in `package.rs` and the shared files below.

| Commits | Family | What |
|---|---|---|
| `8ec8178062`, `4eb975448e` | `sheet::drawable::list` | text-box paragraph lists, levels, numbering and list properties on W1-A's storage-list publisher (30 rows) |
| `367e072a55`, `8a8d80bb29` | `sheet::drawable::drop_cap` | Drop Caps (Rewrite/Derive/Detach plans; the theme's Drop Cap base preset) (4) |
| `90abe395c8`, `be6baa6c2c`, `eecb17ab4c` | `sheet::drawable::hyperlink` | text-box hyperlinks on the ranged engine (4) |
| `d472ce1639`, `c379cc6bc5` | `sheet::drawable::date_time` | Date & Time fields on the smart-field engine (5) |
| `58672971c9`, `d1ad1b59f8` | `sheet::drawable::annotation` | highlights, comments and replies; the generated annotation author (12) |
| `cfa8c8b873`, `8e5230df96`, `733cdbeece` | `sheet::drawable::named_style` | apply, create, `rename_style`, delete, delete-applied, redefine, catalog and applied reads (8) |
| `12a452476f`, `607e020403` | `sheet::drawable::comment` | direct drawable comments and replies (copy-on-write comment graph, author lifecycle, save-token stamps); `sheet_drawables` read (8) |
| `eac8b9bca0` *(shared structure)*, `353c22904b`, `12f8a1a090` | `sheet::drawable::Edit::remove` | images, movies and audio clips: data-reference owners and unreferenced data leave too; every removal clears the drawable's comment first (3) |
| `451c5b28a4` *(shared)*, `17d62d9904`, `9aacc7850c` | `sheet::drawable::media` | `media_assets`, `sheet_media_assets`, `extract_media`; `Edit::{replace_asset, replace_data, replace_poster}` (8) |
| `dfe476f31c` *(shared)*, `28d96c5412`, `f3a4c08777` | `sheet::drawable::Edit::{duplicate, duplicate_text_box}` | image/movie/audio and text-box duplication (4) |
| `8165b4cbdc` *(shared)*, `195bab4975`, `394d86ebef` | `sheet::drawable::Edit::{add_image, add_movie, add_audio}` | new media drawables with new embedded assets (3) |
| `21de340ac8`, `b6da31414f` | `sheet::chart::series` | per-series connection lines, wedge explosions, label distances, symbols, trendlines, error bars, pie labels, value-label visibility on `litchi_iwa_style::series` (32) |
| `b1a59e777e`, `136b8ffc57` | `sheet::chart::series` | series strokes, fills, symbol fills, symbol outlines and their resets (20) |
| `efa56edae5`, `bcc47a7ae3` | `sheet::chart::style`, new `sheet::chart::axis` | chart options (legend visible, border visible, hidden data, 3D bar shape, radar grid shape, gap spacing); axis-style switches (line, gridlines, minor ticks, minimum label) and tick-mark location (24) |
| `19fbd0abcc`, `3ab94c863d` | `sheet::chart::axis` | axis labels and category-axis series names (axis non-style) (4) |
| `b1359ee999` | infra | media payloads written by the leaf's canonical writer (no normal `prost` dependency, as the boundary check requires) |
| `e5627c4d74`, `7a830ad315`, `0dc8c6e3fe` | shared / infra | structure watermark-only patch; media inventory through `litchi_iwa_style::media` after the trunk merge (edge `litchi-numbers -> litchi-iwa-style`, fuzz locks); cost probe |

**Shared files changed (small hunks, for W1-A and the coordinator):**

- `src/package/structure/metadata.rs` (W1-A): `Registry::data_references`
  (`DataOwner`), `Transition::{data_owner_removals, data_owner_additions,
  data_additions, data_removals, entry_writes, entry_insertions}`,
  `Transition::release_unreferenced_data`, `next_data_identifier`, and the
  media-codec pass that runs **before** the registry pass. `remove_objects`
  no longer returns `Unsupported` for data owners — this is the data-owner
  removal the coordinator asked for (`eac8b9bca0`). Two codec facts shaped
  it: the registry codec refuses to drop the UUID of an object that still
  owns data, and it cannot remove one object's owners of two data records
  in one batch (`scan_data_reference_removals` matches removals by object
  only), so owners leave through `package_metadata_media_codec` first.
- `src/package/structure.rs` (W1-A): `Working::publish` writes
  `entry_writes`, inserts `entry_insertions` and deletes the released data
  members; `verify_locality` skips inserted members.
- `src/package/sheet_drawable_annotation.rs` helpers moved to the new
  `sheet_drawable_comment_graph.rs` (shared by the annotation and drawable
  comment families; its census now sees objects the transaction inserted).
- `Cargo.toml` of `litchi-numbers`: `litchi-iwa-style`, `sha1`, `prost`;
  `tools/crate_boundaries.json` edge `litchi-numbers -> litchi-iwa-style`.

## 2. Ledger

169 rows flipped to `migrated` in this slice (303 for W2-A so far);
`python3 tools/iwa_exit_ledger.py --check` passes. Every Tier B route has a
`docs/unified-ops/0363-mac-queue.md` row (no host native evidence; leaf
UUIDs derived deterministically, host random — the documented policy).

## 3. Parity findings

- **Named styles**: the crate-boundary checker flags any bare `rename(`
  call, so the verb is `Edit::rename_style` (as trunk's Keynote fix
  `ade498d1ae`).
- **Drawable comments**: the host clones the root (and a rewritten reply)
  under fresh identifiers for every reply change, then removes what the
  change leaves unreferenced (`comment_object_is_referenced`: archive
  references, drawable routes, reply lists, foreign UUID bindings,
  externals, data owners). The leaf census iterates the working archives,
  not the source, so clones the transaction inserted count (an earlier draft
  removed a live author).
- **Removal with a comment**: clearing the comment in a separate publish
  advanced the save token twice; it now runs inside the removal's working
  set (`comment::clear_in`), and the removal closure drops objects the clear
  already removed (the census reads the source graph).
- **Media removal** in the authored fixture meets a `DataMetadataMap`; data
  the map lists stays (the host's `has_data_metadata`), the rest is removed
  with the map passed to the media codec.
- **Native image replacement** is Tier B only by the ZIP local layout of the
  two rewritten members (content equal), as for Pages.
- **Series counts** follow the host topology: a scatter (X+Y) or bubble
  (X+Y+size) chart shows one series.
- **Series-style resets** of a series without a local value are host
  no-ops; the cases reset a stored value.

## 4. API notes

New public families (all in `numbers-full`): `sheet::drawable::{list,
drop_cap, hyperlink, date_time, annotation, named_style, comment, media}`,
`sheet::chart::series`; `sheet::drawable::Edit` gains `duplicate`,
`duplicate_text_box`, `add_image`, `add_movie`, `add_audio` and removes
media; `sheet::drawable::{ImageOptions, MovieOptions, AudioOptions}`;
`Refusal::InvalidMedia`. Host all-series and one-series `set_*`/`reset_*`
runs collapse into one verb per family (`Edit::set`, `Edit::reset`).
`docs/API_GLOSSARY.md` and `tools/api_glossary.json` list every family.

## 5. Performance

Callgrind, release with debug info (`CARGO_PROFILE_RELEASE_DEBUG=1`), CPU
3, Ir of `w2a_host_route_2` / `w2a_leaf_route_2` over one ABBA pass (two
calls each; open + operation + commit + serialize):

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| text-box list (bullet) | 264,399,926 | 163,992,950 | −38% |
| text-box highlight add | 233,059,820 | 93,756,202 | −60% |
| named style apply | 267,432,365 | 128,564,506 | −52% |
| drawable comment set | 215,822,874 | 108,908,842 | −50% |
| image remove (data released) | 577,294,564 | 131,259,448 | −77% |
| image data replace | 163,927,745 | 131,358,180 | −20% |
| image duplicate | 302,145,345 | 116,403,249 | −61% |
| image add | 320,048,159 | 111,588,495 | −65% |
| chart error bars | 288,100,479 | 152,154,432 | −47% |
| chart series fills | 266,225,618 | 43,777,457 | −84% |
| chart border visible | 181,184,544 | 43,278,713 | −76% |
| chart category axis line | 136,098,179 | 34,536,619 | −75% |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --all-features --no-run
W2A_ROUTE=<route> taskset -c 3 valgrind --tool=callgrind target/release/deps/parity-<hash> \
  --ignored --exact numbers_w2a_cost::numbers_w2a_route_cost_abba_2 --test-threads=1
callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep w2a_
```

The raw outputs (`s2_<route>.out`, 16 MB together) stay in the session
scratch (`.coord/iwa-w2a/cg/`), not in the repository.

## 6. Gaps and follow-ups

| Capability | Rows | Next |
|---|---:|---|
| shape styles (`sheet_shapes.rs`: fill, stroke, shadow, effects, line ends, presets) | 29 | the shape-style variation graph (0375a §5) |
| chart legend fill/font/shadow/stroke, axis title, bounds, scale, steps, number format, label affixes and angle, 3D label position, gridline stroke, background fill, border stroke, category labels, fonts, value-label affixes/formats/locations/auto-fit, error-bar auto-fit, pie leader lines, reference lines, lighting, radar series style, shadow | ~55 | per-property patches on the chart style codec (as `sheet::chart::{style, axis}` do) and the graph engine |
| titles and captions of shapes, images, movies, charts | 18 | the caption graph (CaptionInfo, storage, style) |
| chart lifecycle (add, duplicate, remove, data, direction, kind, list) | 7 | chart authoring |
| text-box columns and text layout | 6 | the shape-style variation graph |
| series and symbol image fills | 2 | image fills through the media transition |
| `semantic/media.rs` package I/O rows (`into_package`, `save`, `to_bytes`) | 3 | retire with the host draft (W3) |

## 7. Verification

Run on the branch with trunk merged (`8c2b44e22e` merged at the slice head),
build env of the coordinator rules (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`), test executables pruned after each test chunk. No
litchi-ops dispatch arm changed, so the capability/registry goldens and the
OfficeCLI corpus were not re-blessed.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features numbers` | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `cargo test -p litchi-numbers --all-features`; `-p litchi-iwa --all-features` (parity: 568 passed, 53 ignored probes); `-p litchi --all-features` | 0; 0; 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, nine groups of ≤20 `--test` targets (167 targets) | 0 each |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 (after `b1359ee999`: the first run refused a normal `prost` dependency of `litchi-numbers`; the media payloads now use the leaf's canonical writer) |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` | 0 |
| `python3 tools/iwa_exit_ledger.py --check`; `officecli_capability_delta.py --check`; `officecli_corpus.py --check`; `unified_bench_matrix.py --check`; `check_feature_matrix.py` | 0 each |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

After `b1359ee999` the affected gates were re-run: clippy (`litchi-numbers`,
`litchi-iwa`, `litchi`), `cargo test -p litchi-numbers`, the parity target,
the crate-boundary check and the perf-baseline check, all 0. Both pre-commit
size checks printed nothing for every commit. Target directory: 14 GB after
pruning.
