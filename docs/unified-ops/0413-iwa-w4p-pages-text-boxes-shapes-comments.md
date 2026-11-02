# 0413 — W4-P: Pages text-box text, annotations, fields, comments and shape rows out of `litchi-iwa`

- Status: wave W4-P on branch `uo2/iwa-w4p` (cut from trunk
  `feat/unified-ops-api` at `4b26bb3379`; trunk merged at `0ea960dc4b`,
  `fc7a879e7c`, `1bc0d4c996` and after the coordinator's slice-3 merge).
  Slices 1–3 are merged; slice 4 (removal, drawable comments, the drawable
  list, line segments) is ready. The remaining shape-lifecycle rows are
  listed in §6.
- Date: 2026-11-02.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (§0 W2-B scope, §3.2 API
  shape, §4 parity tiers); harness and ledger
  [0369](0369-iwa-exit-parity-harness-and-ledger.md); shared engines
  [0374](0374-iwa-w1d-shared-text-engine.md) (W1-D) and `litchi_iwa_style`
  (W1-E); the Numbers families this wave ports,
  [0378](0378-iwa-w2a-numbers-sheet-drawables.md) and
  [0378a](0378a-iwa-w2a-numbers-drawable-families-slice-2.md); the Keynote
  shape work in [0372](0372-iwa-w1c-keynote-slides-drawables-media.md); and
  Pages so far, [0371](0371-iwa-w1b-pages-body-sections-tables.md).
- `performance_claim`: every migrated route is 70–90% cheaper than its host
  route under callgrind ABBA (§5).

## 1. What landed

Pages text boxes are body-anchored `TSWP.ShapeInfoArchive` drawables
flagged as text boxes. `DrawableSelector::TextBox(Position)` selects one by
its text-flow position. The body-drawable ownership proof
(`body_drawable_graph::select_drawable`) accepts the host's text-box
ownership: an owned storage, and a deprecated storage slot that is absent or
names the same storage.

The text families are ports of the Numbers sheet-drawable publishers. They
run on the shared structural working set (`litchi_iwa_table::structure`,
record 0379) through a small Pages toolkit, `package::text_box_structure`.
The toolkit provides the working-set alias, the exact artifact pair and the
family macro, and the text-target resolver.

| Commit | Family | What |
|---|---|---|
| `71983118cb` | `DrawableSelector::TextBox`; `drawable::text_box_text` | `Package::{body_text_box_text, edit_body_text_box_text, apply_body_text_box_text}`. `Edit::{set, reset}` over the W1-D paragraph-style engine for all 24 properties, the following style included. `drawable::{properties, text, style}` accept text boxes, and `drawable::style` gains `TextColumns` |
| `50ed0f4c04` *(shared)* | `litchi_iwa_text_wire::list::publish` | The record 0370b storage-list publisher lifted into the text crate. The package side is a `Workspace` trait, so the text crate takes no table or core edge |
| `fc14233b1a` | `drawable::{list, named_style}` | Text-box lists, levels, numbering and per-level properties; named-style apply, create, `rename_style`, delete, delete-applied and redefine |
| `97bed34d25` *(shared)* | `litchi_iwa_table::structure::list_workspace` | One `Workspace` adapter over the shared working set. The Numbers cell-list and text-box list families move onto it, and Numbers' 637-line copy of the publisher is deleted |
| `22e12323cf` | `drawable::{language, drop_cap, hyperlink, date_time, annotation}`; `number_attachment::Target::Drawable` | Languages, Drop Caps (base from the Pages `TP.ThemeArchive`), hyperlinks, Date & Time fields, highlights, comments and replies with the generated annotation author, and number attachments on a text box's storage |
| `54150ddaa1` *(host fix)* | host `Draft::remove_text_box` | §3 |
| `d8dadf1e0b` | `drawable::lifecycle::Edit::remove` | Removes text boxes and shapes. The private graph includes the owned storage |
| `1cd2d20802` | `drawable::comment` | Direct drawable comments and replies on body text boxes, shapes, images, movies and audio |
| `82d4fa6523` | `drawable::properties::{Edit::set_line_segment, Snapshot::line_segment, Segment}` | Straight-line segments through the shared `shape_style_codec::{line, rewrite_line}` |

Every family is registered in the `pages` glossary profile, with its
`apply_*` result contract, and listed under this record in
`docs/API_GLOSSARY.md`. `Publisher::new` is pinned in the `iwa-text-wire`
`required_signatures`. No native identifier appears in a public signature,
and no litchi-ops dispatch file changed (D7).

## 2. Ledger

179 rows migrate and one is retired:

| Rows | Host routes | Destination | Tier |
|---:|---|---|---|
| 72 | `{set_,reset_,}text_box_{text,paragraph}_*` (23 properties) and the following-style triple | `drawable::text_box_text` | set/reset B, reads A |
| 7 | text-box geometry, properties, text layout and columns | `drawable::{properties, style}` | set B, rewrite/reset/read A |
| 4 | `{set,replace,clear}_drawable_text`, `drawable_text_storages` | `drawable::text`, `Package::body_drawable_text` | A |
| 30 | text-box paragraph lists | `drawable::list` | presets and properties B, level/numbering and reads A |
| 8 | named paragraph styles | `drawable::named_style` | B (rename and reads A) |
| 34 | comments (8), highlights (4), hyperlinks (4), Date & Time (5), languages (5), Drop Caps (4), number attachments (4) | `drawable::{annotation, hyperlink, date_time, language, drop_cap}`, `number_attachment` | creations and Drop Caps B, the rest A |
| 2 | `remove_text_box`, `remove_body_shape` | `drawable::lifecycle::Edit::remove` | B |
| 16 | Pages `*_drawable_comment*`, `drawables`; shared `comments.rs` comment/reply routes and `drawables` | `drawable::comment` | edits B, reads A |
| 3 | `{set_,}body_line_segment`, `Draft::builder` | `drawable::properties`, `Builder::new` | A, A, B |
| 1 retired | shared `comments.rs application` | — | the host accessor of its own application classification; leaf packages are format-typed |

`python3 tools/iwa_exit_ledger.py --check` passes. None of these routes has
host native evidence, so each destination has a W4-P row in
`0363-mac-queue.md` (79 rows, one of them for the host fix).

## 3. Parity findings

- **Following style.** The host's in-place style rewrite keeps every header
  reference. When "same style" replaces a named following style, the
  header therefore still references the named style. The Numbers
  publisher, which never exposed the following style, dropped such
  references. The Pages port keeps them, as the host does, so the rewrite
  cases are byte-identical.
- **Host defect, text-box removal.** `Draft::remove_text_box` left two
  things behind: the document's external references to the removed graph
  (the text box's style edge among them), and the body storage header's
  attachment reference. Its own "remains referenced after deletion" check
  then refused every host-added text box. It now prunes both, as the
  shape and media removals do (`54150ddaa1`). This is macOS-queued.
- **Line segments.** A segment's endpoints are derived back from the
  stored frame and path width, so the candidate check accepts `f32`
  rounding (1e-3 relative) on the endpoints only. The bytes are
  identical to the host's in both directions.
- **Tier B shapes.** The Tier B cases differ from the host only in the
  documented leaf policies, accepted by the harness: fresh deterministic
  UUIDs for new variations, presets, styles, comment storages and smart
  fields, and the save-token advance on structural publications.

## 4. API notes

- Module and method names mirror Numbers, with `sheet_` → `body_` and no
  sheet selector. Pages already used `drawable::text` for plain text, so
  the property family is `drawable::text_box_text`. The body's own
  Date & Time fields stay in `litchi_pages::date_time`; the text-box ones
  are `drawable::date_time`.
- `drawable::comment` refuses tables (no route). The media removal still
  refuses a commented drawable (`Refusal::Commented`). The host clears the
  comment first; a combined clear-and-remove is a follow-up.

## 5. Performance

The probes run under callgrind on a release build with debug info, pinned
to CPU 3, in host, leaf, leaf, host (ABBA) order. They measure inclusive Ir
of the probe functions of `pages_cost_probes::abba`, two calls per side.
Each call opens the package, performs one edit, commits and serializes.

| Route (probe) | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| text-box alignment set (`text_box_alignment_set_cost_abba`) | 602,007,823 | 151,702,468 | −74.8% |
| text-box columns set | 571,907,401 | 135,131,998 | −76.4% |
| text-box geometry | 433,766,962 | 41,418,237 | −90.5% |
| text-box bullet list | 533,922,952 | 156,684,871 | −70.7% |
| named style apply (over an override) | 659,570,213 | 199,406,490 | −69.8% |
| hyperlink add | 506,621,449 | 76,308,643 | −84.9% |
| text comment add | 669,105,389 | 85,641,041 | −87.2% |
| number attachment insert | 428,052,544 | 65,343,595 | −84.7% |
| drawable comment set | 434,049,289 | 99,952,051 | −77.0% |
| text-box remove | 524,505,391 | 66,098,371 | −87.4% |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --all-features --test parity --no-run
taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*probe_host*' \
  target/release/deps/parity-<hash> --ignored --exact <module>::<probe> --test-threads=1
# and '--toggle-collect=*probe_leaf*'
```

The raw outputs stay outside the repository, in
`.coord/iwa-w4p/cg/` (with `cg_slice{1..4}.txt` summaries). No wall-clock
or unified-`File` claim is made.

## 6. Gaps and follow-ups

| Rows | Capability | Next |
|---:|---|---|
| 7 | `add_text_box`, `add_body_{rectangle,shape,shape_with_fill,line,line_with_endpoints,line_with_style}` | Construct the body text-shape graph: the Numbers `sheet_drawable_author` builder, anchored with the media family's `insert_anchor`/`append_zorder` |
| 2 | `duplicate_text_box`, `duplicate_body_shape` | Shape-info and storage reference paths in the media graph clone |
| 2 | `{set_,}body_shape_preset` | Lift the preset path-source builder and classifier (duplicated in Numbers' `sheet_drawable_author` and Keynote's `slide_drawable::{author, preset}`) into `litchi-iwa-style` |
| 1 | `set_body_shape_image_fill` | Data-reference registry bookkeeping for a style variation (the 0371 image-fill gap) |
| 1 | `body_shapes` | A Pages shape listing (preset, line segment, text) once presets read |
| 1 | `pages/creation.rs body_table` | Unchanged 0371 gap (table bootstrap) |
| — | commented-drawable removal | clear the comment inside the removal's working set, as Numbers' `comment::clear_in` does |
| — | native evidence | every W4-P route waits on its macOS queue row |

## 7. Verification

Full gate on `7f16e7871b` with trunk `1bc0d4c996` and the coordinator's
slice-3 merge brought in. The build env followed the coordinator rules
(`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`), and test executables were
pruned between litchi-ops chunks.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pages -p litchi-numbers -p litchi-iwa-text-wire -p litchi-iwa-table -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features pages` and `--features numbers` | 0, 0, 0 |
| `cargo test -p litchi-pages --all-features` (765) | 0 (the first run failed one assertion that pinned the old shape-removal refusal; fixed in `7f16e7871b` and rerun) |
| `cargo test -p litchi-numbers --all-features`; `-p litchi-iwa-text-wire -p litchi-iwa-table` | 0; 0 |
| `cargo test -p litchi-iwa --all-features` (2,713, parity target included) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, nine groups of ≤20 `--test` targets (168) | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pages -p litchi-numbers -p litchi-iwa-text-wire -p litchi-iwa-table` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (73 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` | 0 |
| `python3 tools/iwa_exit_ledger.py --check`; `officecli_capability_delta.py --check`; `officecli_corpus.py --check`; `check_feature_matrix.py` | 0 each |
| `python3 tools/unified_bench_matrix.py --check` | 0 after regenerating `results/0289/bench-matrix.md` to list this record (1 before) |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| every `crates/*/fuzz/Cargo.toml`: `cargo metadata --offline --locked` | 0 |

No litchi-ops dispatch, capability or corpus-visible behaviour changed, so
the capability/registry goldens and the OfficeCLI corpus were not
re-blessed. Both pre-commit size checks printed nothing for every commit.
