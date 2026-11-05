# 0435 — a new caption uses the application's own paragraph style

- Status: merge-ready on branch `uo2/iwa-w8` (cut from trunk `feat/unified-ops-api` at
  `3001da641b`). **Five packages and four queue rows wait on Mac run 4.**
- Date: 2026-11-05.
- Decision: the owner, 2026-11-05: new captions use Numbers' own style (0433 §3).
- `performance_claim`: the caption route stays within the +5% budget against the previous
  leaf; the largest change is Keynote's new image caption, +0.08% (§6).

## 1. Result

| Format | New caption | New title | Native caption evidence | Changed? |
|---|---|---|---|---|
| Numbers | template's caption style, as native | title's style, byte-identical | chart, `chart-caption-native.numbers` | **Yes**, all kinds |
| Keynote | template's caption style, as native | title's style, byte-identical | chart and movie, `chart-caption-native.key` and the movie fixtures | **Yes**, the image and shape routes |
| Pages | title's style, as before | title's style | **none** | **No.** Pinned by a test |

No file in `test-data` holds a native title (`child_info_kind` 2), so a title's style has no
native evidence in any format. Titles keep the theme's first preset. The owner's rule is that
they keep "Caption Title", and the template names that preset so.

## 2. What native files show

A scan of the 119 single-file `.numbers`, `.key` and `.pages` directly under `test-data/iwork/*/` for
`TSA.CaptionInfoArchive` found captions only in Numbers (one file) and Keynote (33 files), all
of kind 1, all on the same style:

| Item | Numbers `904466` | Keynote `2651688` |
|---|---|---|
| Name stored in the style | `Object Caption` | `Object Caption` |
| `style_identifier` | `captions-0-paragraphstyle-Caption Text` | same |
| Font | HelveticaNeue-Light | not read |
| Place in the theme's caption styles | second | second |

**0433 §3 and the brief call this style "Caption Text". That is the end of its identifier;
the name it stores is "Object Caption".** The other preset, `904465`, stores "Object Title"
with an identifier ending "Caption Title". No bytes say what the Format panel prints, so
`RUN.md` asks for the words and compares them with Numbers' own file.

`TSS.ThemeArchive.application_presets.caption_styles` lists `[Title, Caption]` in all three
formats. The stylesheet maps these styles too, in its field 8, whose meaning we have not
established; the theme is the route the existing code already reads.

## 3. The change

`litchi_iwa_style::caption::theme_label_paragraph_style(theme, label, object)`:

- a title takes the first preset;
- a caption takes the preset whose object is a paragraph style with that identifier, whatever
  its position. No object number is written anywhere;
- otherwise, the first preset: see below.

Numbers' `theme()` and Keynote's `caption_theme()` and `caption_dependencies()` take the label.
Everything downstream follows: the shape style's text properties, the storage's paragraph
table, both reference lists, and the external reference (Numbers' componentized chart, Keynote's
`register_dependencies`). Pages still calls `theme_paragraph_style`.

**A template without the style: the existing fallback, not a refusal.** I first wrote a typed
`Refused`. A census of the theme in those 119 tracked files and the 797 frozen inputs then showed
real documents would lose a working route:

| Caption styles in the theme | Numbers | Keynote | Pages |
|---|---:|---:|---:|
| `[…Caption Title, …Caption Text]` | 301 | 268 | 251 |
| `[litchi-paragraph-caption]`, one style listed twice: authored by the host, some then saved by the application | 17 | 17 | 0 |
| `[…Caption Title]` alone: source-built | 0 | 0 | 62 |

Numbers opened and saved such files (`chart-grid-authored-native.numbers` has a receipt). Their one
style is a caption style, and the first preset is the only answer they offer. Tests hold no Numbers or Keynote file
with one preset. A refusal would have
broken them for a shape we hold no sample of: two presets and neither the native one. In that
shape the fallback gives the title's style, which 0431 to 0434 showed opens and saves. It writes
nothing different from before. Tests pin both cases.

### 3.1 Which drawables have evidence

Numbers' native caption is on a chart. Keynote's are on a chart and a movie. The image and shape
routes get the same style by inference: Keynote shows the choice does not vary between the two
kinds we can see, and the theme has one caption style. No native file holds an image or shape
caption. Run 4 shows those bytes open; it does not show what Numbers would have chosen.

### 3.2 Pages

No Pages file holds a caption Pages authored. Its templates list the same style, so the fix is
probably right, but that is a guess, and the instruction was not to guess. A caption typed in
Pages, saved and added to `test-data`, would settle it; the code change would be one call.

## 4. What stays different from native

Compared on `W8-native-parent-new-caption` against `chart-caption-native.numbers`:

| Part | Native | Leaf | Decision |
|---|---|---|---|
| Paragraph style | `904466` | `904466` | **Same now** |
| Paragraph font | HelveticaNeue-Light | same style object | **Same now** |
| Style 2025: identity | template's `904468`, unedited | new object, same identifier string, in the stylesheet | **Stays** |
| Style 2025: overrides, fill | 14, empty | 7, transparent | **Stays** |
| Storage field 7, list style `text-0-liststyle-None` | present, listed | absent | **Stays** |
| Caption info, placement, storage's other fields | | same fields | |

**Style 2025.** The template's preset is clearly what native uses. Using it means a caption with
no object of its own: the inline layouts would cross a component they do not cross today, under
the rule that an unlisted cross-component reference is fatal (0431), and Keynote's
`lists_graph_style`, witness and budget proofs all assume a private style. That is not small. The
empty and transparent fills draw nothing. Numbers keeps the leaf's object on re-save (0433 §3).
The difference is structural and harmless so far as we can see.

**Field 7.** Both native formats hold it, and Numbers adds and lists it when it re-saves the
leaf's (0433 §3). So it is correct, but Numbers repairs it and it shows nothing. Adding it needs
a third theme lookup, a codec change Pages shares, another listed reference in each layout and a
change to Keynote's dependency proof. Not small, for no visible effect.

Keynote's chart and movie routes still refuse a new caption on a native slide: the slide does not
list the document stylesheet, and `prove_caption_dependencies` requires it, unlike the image route,
which registers it. That was so before; `a_new_caption_on_a_native_chart_stays_refused` pins it.
Their new-caption code is exercised only on synthetic files.

The chart lifecycle's `paragraph_style` (new chart's presets) also takes the first preset. It is not
a caption, so it is unchanged and not examined.

## 5. Tests

`crates/litchi-numbers/tests/caption_native_style.rs`, `crates/litchi-keynote/tests/caption_native_style.rs`,
`crates/litchi-pages/tests/body_chart_caption.rs`, and a unit test in `caption.rs`:

- a new caption's style equals the native file's: object, identifier and name;
- a new title takes the title's style, and adding one beside a caption leaves the caption's;
- replacing a new caption and Numbers' own keeps the style, and removal gives a stand-in;
- a new caption after a removal takes the style again;
- each step's inverse gives the exact bytes before it;
- a template lacking the style, and a document Litchi built, fall back;
- Pages' new caption stays on the first preset.

**Rebuild as a regression test.** The examples `numbers_mac_queue_caption_routes` and
`keynote_mac_queue_caption_routes` write all 28 queue packages of the four affected rows from the
frozen inputs. **23 hash identical to the manifest**: every title, replacement and removal. Exactly
five differ, and field by field they differ only in the paragraph style:

| Package | Objects changed | External reference |
|---|---|---|
| `P0290`, Keynote image caption | style, storage: `2651687` to `2651688` | slide: title's out, caption's in |
| `P1129`, movie caption; `P1130`, image; `P1145`, shape | style, storage: `904465` to `904466` | none: same component's list holds it |
| `P1146`, chart caption | style, storage: `904465` to `904466` | engine: title's out, caption's in |

`control_kit.py audit`'s unlisted-reference check finds none in any of the 29.

## 6. Performance

Callgrind, release with debug lines, CPU 5, ABBA, inclusive Ir of a one-route probe over the same input.
Base is the four changed source files at `3001da641b` in this tree. Probes and binaries are in
`.coord/iwa-w8/perf/`, not committed; SHA-256 of the binaries: Numbers base `3afebca6…`, new `a8f3bd69…`;
Keynote base `f80d1cc6…`, new `219d5a9a…`. The comparison is the previous leaf: the host is gone.

| Route | Base Ir (A, A) | Branch Ir (B, B) | Change |
|---|---:|---:|---:|
| Numbers, new chart caption | 52,721,683; 52,730,903 | 52,747,876; 52,740,203 | +0.03% |
| Numbers, new image caption | 29,291,181; 29,294,637 | 29,307,586; 29,305,568 | +0.05% |
| Numbers, new image title | 29,288,627; 29,290,133 | 29,291,397; 29,287,557 | +0.00% |
| Keynote, new image caption | 73,475,740; 73,476,150 | 73,533,606; 73,538,798 | +0.08% |
| Keynote, new image title | 73,474,121; 73,470,295 | 73,498,754; 73,494,270 | +0.03% |

Budget +5%. No unified route changed, so there is no unified-against-leaf measurement.

## 7. Run 4 and the queue

`.coord/iwa-w8/out4/`, tar `.coord/iwa-w8/out4.tgz` (1,283,567 bytes, SHA-256
`98fa06d96819ed055b5049307bcd8359d4aebc1d33738ac0dba12a05ba6e7d6b`), not committed. `results/0435/run4-packages.tsv`
lists every package, its hash, what it replaces, and its rows. Seven packages:
`W8-P0290`, `W8-P1129`, `W8-P1130`, `W8-P1145`, `W8-P1146`, plus `W8-native-parent-new-caption` and
Numbers' file as `W8-native-caption-control`. **Every affected row's bytes rebuilt from frozen inputs;
none had to be left out.** `run_on_mac.sh --dry-run` on Linux verifies all seven hashes from the extracted tar,
exit 0; every path in `RUN.md` exists in it. No application ran.

**Queue edits, made by hand** in the Result and Artifact cells of six rows. `apply_results.py --evidence` was
not run, as it would drop the notes of 0434 §2. Row ids are unchanged (checked across all 767),
and the manifest of 0434 stays as it was:

| Row | Line | Result now | Artifact cell |
|---|---:|---|---|
| `q-fbf3a944f61d` W1-C `edit_slide_image_text` | 225 | pending run 4: `P0290` | `P0290`'s old re-save removed |
| `q-e8bd7cd5d7a2` W4-L `edit_sheet_drawable_caption` | 770 | pending run 4: `P1129`, `P1130`, `P1145` | those three removed |
| `q-077e77954c55` W4-L `sheet_drawable_caption` | 771 | pending run 4: `P1146` | `P1146` removed |
| `q-91327706dd7b` W4-N `edit_sheet_drawable_caption` | 775 | pending run 4: `P1146` | `P1146` removed |
| `q-c8ea770c712d` W4-H `pivot_categories` | 756 | `OPEN` replaced: discarded by the owner, 0429 §5 | unchanged |
| `q-3080a227051a` W4-K formulas | 773 | `OPEN` replaced: discarded by the owner, 0429 §5 | unchanged |

Each pending cell keeps the earlier pass for the row's other packages. The Pages rows are untouched:
their bytes did not change.

**The frozen manifest does not know the `W8-*` hashes.** A later `apply_results.py --evidence` against it
would put the old hashes back, mark these rows `pass` on bytes no longer written, and drop the notes. So a run-4
pass is entered by hand, from `run4-packages.tsv`. No tool option was added: this is one fold.

## 8. Gaps

- Pages' new caption stays on the title's style until a native file exists (§3.2).
- The image and shape routes' style is inferred from chart and movie evidence (§3.1).
- A title has no native evidence in any format (§1).
- Style 2025 and field 7 stay different (§4).
- Keynote's chart and movie routes refuse a new caption on native slides (§4).
- 0433 §10's caption gap is closed for Numbers and Keynote, once run 4 passes.

## 9. Not verified

- **No Mac ran.** Nothing here shows these bytes open, nor what the panel prints, nor how the caption looks.
- Whether a caption-styled Litchi caption re-saves with the style kept: run 4's re-saves show it.
- The font of Keynote's style was not read; it is the same identifier, not a checked font.
- The stylesheet's field 8 was read only as far as §2 says.
- The owner's discarding of the by-reference shape is taken from the brief; 0429 §5 itself says the parity cases were dropped by owner decision, and the discard commit is `e6a76cce0f`.

## 10. Gates

Logs are in `.coord/iwa-w8/gates/`. All exit 0 on the final run. No golden was re-blessed: no capability, corpus or
registry output changed, and the tests passed against the goldens as they were.

| Gate | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy` of `litchi-iwa-style`, `-numbers`, `-keynote`, `-pages`, `litchi-ops`, `litchi`, `litchi-py`, all targets and features (first run: 1 `needless_lifetimes` in the new test; rerun after the fix) | 0 |
| `cargo clippy -p litchi-ops --no-default-features`: none, numbers, pages, keynote | 0 |
| `cargo test --all-features` of the four changed crates | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, every `--test` in chunks of 28 | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc` of the changed crates | 0 |
| `check_crate_boundaries`, `non_iwork_gate verify`, `check_iwork_leaf_public_api`, `iwa_exit_ledger --check`, `generate.py --check` (767 rows, 0 errors), `check_example_targets` | 0 |
| Glossary, profiles `iwa-style`, `numbers-minimal`, `numbers-full`, `pages`, `keynote-minimal`, `keynote-full` | 0 |
| Tool unit tests: the rules' modules and `test_mac_queue`, `test_mac_control_kit`, `test_gate_runner`, `test_iwa_exit_ledger`, `test_check_example_targets`, `test_unified_bench_matrix` | 0 |
| `officecli_capability_delta --check`, `officecli_corpus --check`, `check_feature_matrix` | 0 |
| `tools/unified_bench_matrix.py --check`, after adding this record's exemption and regenerating (first run: 1) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `tools/perf-baseline`: `cargo check --locked --all-targets` | 0 |

The lint fix touches only a helper in the two new test files, whose tests were rerun; the full suites ran before it. No library
code changed after the kit was built.
