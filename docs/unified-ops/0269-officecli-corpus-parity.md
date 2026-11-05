# 0269 — Closing the largest OfficeCLI corpus gaps

Date: 2026-10-13. Branch `uo2/corpus-parity`, cut from trunk `8fd16dd11e`.
Follows record 0265 (the example-corpus harness) and its delta table
[`officecli-corpus-deltas.md`](officecli-corpus-deltas.md).

## 1. Result

Corpus outcomes (`crates/litchi-cli/tests/officecli_corpus/outcomes.tsv`),
script calls only (resident no-ops and read-backs excluded):

| Format | Lines | Passed before (0265) | Passed after | Warned after | Gap before → after | Cascade before → after | Bug |
|---|---:|---:|---:|---:|---:|---:|---:|
| docx | 949 | 466 | 473 | 172 | 379 → 231 | 98 → 67 | 0 |
| pptx | 2,959 | 418 | 1,097 | 438 | 1,853 → 863 | 564 → 418 | 0 |
| xlsx | 2,523 | 1,119 | 2,035 | 18 | 1,384 → 461 | 11 → 0 | 0 |
| **total** | **6,431** | **2,003** | **3,605** | **628** | **3,616 → 1,555** | **673 → 485** | **0** |

Passed counts grew in every script family (the test enforces this) and no
call that passed before fails now. By step: the cell `text` alias alone took
the total to 2,909 (XLSX 2,025); the warn policy and the PPTX writes took it
to 3,605. "Warned" calls applied with some properties skipped (§2); they are
not counted as passed. Every created file still reads back (200/200).

## 2. "Apply the rest and warn" (F3)

Record 0015 (Q2/Δ1) decided that unknown or unsupported properties are
refused by the library verbs and that OfficeCLI's warn-and-continue is a
command-channel switch only. This record implements that switch.

**Ops API** (`crates/litchi-ops/src/batch.rs`, `batch/execute.rs`):

- `batch::Unsupported::{Refuse, Warn}` (default `Refuse`).
- `batch::Options` (consuming setters `mode`, `unsupported`, `limits`),
  taken by `File::apply_with(item, &Options)` and
  `File::batch_with(items, &Options)`. `apply` and `batch_with_limits`
  delegate to them; `set`, `add` and `edit` closures stay strict.
- Under `Warn`, a `set` or `add` item whose target refuses a property with
  `unsupported_property` is retried without that property (every assignment
  of the key) until it succeeds or fails for another reason. Each attempt is
  isolated (`Edit::isolated`), so only the final attempt leaves changes: the
  valid subset applies atomically and is byte-identical to applying that
  subset alone (test `warn_applies_the_supported_subset_exactly`). A `set`
  whose every property is skipped is a no-op. Invalid values, path errors,
  limits and protection still fail the item (and an atomic batch).
- Per-property outcomes: `Applied::warnings: Box<[Error]>`, each the
  refusing `unsupported_property` error with its property and item index.
  Boxed so `ItemResult` stays four words; empty and unallocated under
  `Refuse`. `transaction::Diagnostics` is unchanged (it is `Copy` and
  content-free).
- Limitation: a refusal is matched to the caller's key by name. A refusal
  that names a canonical property the caller spelled with an alias is not
  guessed; it fails the item (record F3 row "still refused" in the table:
  0 corpus lines).
- Under `Warn`, atomic `set` items are not coalesced, because a coalesced
  run reports its refusal only when the run closes. See §6 for the cost.

**CLI**: `set`, `add` and `batch` take `--unsupported refuse|warn` (default
`refuse`). With `warn`, each skipped property prints
`warning[unsupported_property]: ...` with its `item`, `property` and
`suggestion` lines on stderr, the text report adds `(skipped: ...)`, the
`--json` item carries `warnings` (OfficeCLI `ErrorResult` objects), and the
run exits 9 (partial), the status a best-effort batch with failures already
uses. OfficeCLI exits 2 there; Litchi's 2 is a usage error (record 0256 D2).

**Corpus decision**: the stand-in passes `--unsupported warn` to every `set`,
`add` and `batch`, because that is OfficeCLI's own default behaviour. Calls
that skipped properties are recorded as `warned` with the skipped
properties as detail; `tools/officecli_corpus.py` lists each skipped
property in the per-format gap tables with refusal `warned`, cross-checked
against the registry like a refusal (a skipped property the registry claims
would be a bug; there are none). Recorded in record 0265 §3 and in the
generated table's "Kinds".

## 3. Workbook cell `text` (G7)

`text` is a registry alias of the cell `value` in XLSX (OfficeCLI's cell
handler, `case "value" or "text"`), XLSB and XLS; XLSX also gains
OfficeCLI's documented `bold` and `italic` aliases of `font.bold` and
`font.italic`. ODS has no unified cell writer, so it takes nothing yet. The
schema export now lists XLSB and XLS cell aliases (it listed only XLSX
tables). Tests: `crates/litchi-ops/tests/cell_text_alias.rs` (both
spellings write identical bytes; read-back after a byte round trip; exact
no-op; schema alias). Because the alias resolves to the same registry entry
and setter, it inherits the value writer's signature and encryption policy.

## 4. PPTX writes

- **Text-box formatting on `add`** (461 lines): the shape registry's
  formatting setters (`size`, `bold`, `italic`, `color`, `font`,
  `font.latin`, `align`) are now add-able. `add` stages `add_text_box` and
  one `set_text_properties` on the new shape in the same native edit (both
  owners), so the shape and its formatting are one commit.
- **OfficeCLI shape, paragraph and run aliases**: shape `left`, `top`, `w`,
  `h`, `fontSize`, `font.size`, `preset`, `shape`; paragraph `alignment`,
  `halign`; run `fontname`, `fontFamily`, `font.name`; exported in the
  schema.
- **Shape `fill` and `geometry`** (255 lines; add and set): leaf
  `opened::ShapeFill::{Solid, None}` and
  `TextProperties::{with_shape_fill, with_geometry}` splice the shape's
  `p:spPr` fill choice and `a:prstGeom` in schema order on a
  `TextTarget::Shape` edit, so both owners reuse `set_text_properties` and
  its read-back. `fill` takes `none` or a colour (opaque RGB only; scheme
  colours and alpha refuse with `unsupported_capability`), and `null`
  removes the local fill. `geometry` takes any `ST_ShapeType` token (exact,
  then ASCII case-insensitive, as OfficeCLI); assigning the current preset
  keeps its adjust values. Both are write-only: the read gap L-PPTX-R6b
  (the shape record keeps no `spPr` span) remains.
- **Slide `transition`** (161 lines; set): OfficeCLI's shorthand
  `TYPE[-DIR...][-SPEED|-MS]` (`format/pptx/write/transition.rs`, all 63
  type tokens of `transition.json` (`none` included) except `cube`, which the native model
  lacks) parses to `litchi_pptx::transition::Transition`, staged by the new
  leaf `opened::Edit::set_slide_transition` and source primary
  `Edit::set_slide_transition` (wrapping the existing
  `SourceBackedSlideEdit::{set_transition, clear_transition}` splice).
  `none`/`false`/`null` clear. Set only: a source-backed slide insertion
  refuses markup compatibility in the new slide, which Morph and the
  PowerPoint 2010+ effects need, so `add slide --prop transition=` is
  `unsupported_property` (6 corpus lines, warned).
- **Scene reads after an extension transition** (leaf
  `shape/reader.rs`): markup compatibility that occurs only after the
  slide's `cSld` (the transition owner) can no longer mark the shape scene
  rewritten, so byte-splicing shape edits stay available on slides with
  Morph or p14 transitions. The rule is conservative: the source prefix
  through `</cSld>` must not name the markup-compatibility namespace at all
  (directly or by character references). Markup compatibility inside the
  shape tree still rewrites (test
  `shape_edits_work_after_an_extension_transition`).

Tests: leaf `opened/text_format_tests.rs` (3 new: round trip, refusals,
source owner equals eager owner with inverse), `opened/transition_edit_tests.rs`
(4); ops `tests/pptx_corpus_parity.rs` (5: add with formatting, fill,
geometry and aliases on both owners with read-back after a byte round trip;
set fill/geometry with no-op repeat and clear; transitions with no-op repeat
and clear; signed source refusal; encrypted source refusal) and
`format/pptx/write/transition.rs` unit tests. Capability and registry
goldens re-blessed (`pptx.json`, `xlsx*.json`, `xlsb.json`, `xls.json`).
OfficeCLI write-capability delta for pptx: 25 → 43 supported, 38 → 25 gap.

## 5. Findings X1 and X2

One rule for refused additions, documented in `litchi_ops::error`:

| Refusal | Code |
|---|---|
| The element kind cannot be added in this format (yet), whatever the parent | `unsupported_element` |
| The kind is addable, but not under this parent | `unsupported_type` |
| The kind is addable here, but a declared gap is pending | `unsupported_capability` with the follow-up |

PPTX and XLSX now check the element kind before the parent, so
`add /Sheet1 --type chart` answers "this element kind cannot be added to an
XLSX workbook yet; only sheets can" (`unsupported_element`, was
`unsupported_capability` with the sheet message: X1), and PPTX `add
/slide[1] --type chart` is `unsupported_element` (was `unsupported_type`
with the slide-parent message). XLSX sheet-under-sheet is `unsupported_type`
like PPTX and XLSB. DOCX already answered `unsupported_element`; its
`unsupported_capability` rows are declared registry gaps (third row). The
CLI exit status is 5 for all three codes, as before.

## 6. Measurements

Paired ABBA medians, 15 rounds (30 samples per leg) × 3 trials, release
build, `taskset -c 29`, shared 32-core host (load 15–21 from other agents).
Binary `target/release/examples/corpus_parity_bench`, SHA-256
`6ca2a846e13d1c6647fac70336f7c2b2cba6f030880360e34ca59c3168144827`, built from
this record's commit with
`cargo build --release -p litchi-ops --features pptx,xlsx --example corpus_parity_bench`.
Each pair first checks that both legs write identical bytes. Raw:
[`results/0269-corpus-parity/trial-{1,2,3}.jsonl`](results/0269-corpus-parity/).

| Scenario (PPTX eager owner; deck of N text boxes) | N | leaf µs (t1/t2/t3) | unified µs | unified/leaf |
|---|---:|---|---|---|
| add text box with size, bold, colour, fill, geometry | 1 | 328.2 / 329.1 / 330.3 | 457.1 / 457.5 / 458.3 | 1.39 |
| same | 100 | 2,525.0 / 2,519.6 / 2,519.1 | 2,801.5 / 2,794.4 / 2,787.1 | 1.11 |
| same | 1,000 | 23,517 / 24,167 / 24,469 | 25,000 / 25,662 / 25,952 | 1.06 |
| add plain text box (pre-existing path, reference) | 1 | 304.1 / 303.2 / 305.0 | 429.9 / 429.1 / 434.8 | 1.41–1.43 |
| same | 100 | 2,064.6 / 2,108.3 / 2,099.3 | 2,321.7 / 2,375.3 / 2,358.5 | 1.12–1.13 |
| same | 1,000 | 21,227 / 19,695 / 21,745 | 22,858 / 21,182 / 23,207 | 1.07–1.08 |
| set shape fill + geometry | 1 / 100 / 1,000 | 277.4 / 1,287.2 / 10,431 (t1) | 281.1 / 1,289.3 / 10,414 (t1) | 1.00–1.02 |
| set slide transition | 1 / 100 / 1,000 | 281.0 / 1,306.2 / 10,483 (t1) | 280.7 / 1,303.2 / 10,422 (t1) | 0.99–1.01 |

| Scenario (XLSX, created workbook) | A median µs | B median µs | B/A |
|---|---|---|---|
| one cell `apply_with`, Refuse (A) vs Warn (B), nothing refused | 23.6 / 23.4 / 23.6 | 23.5 / 23.6 / 23.5 | 1.00–1.01 |
| same cell with one refused property: subset under Refuse (A) vs Warn (B) | 22.9 / 23.0 / 23.5 | 31.8 / 31.6 / 32.1 | 1.37–1.39 |
| 100-cell batch, Refuse (A, coalesced) vs Warn (B, per-item) | 355.0 / 363.2 / 363.1 | 17,423 / 17,600 / 17,562 | 48.4–49.1 |

Reading:

- The new set paths (fill/geometry, transition) are within 2% of the leaf.
- Formatted text-box add is over the 10% single-op budget at 1 and 100
  shapes, but the overhead is the add path's own (plain add: 1.42 / 1.13 /
  1.07); the formatting adds no unified overhead beyond it. Not fixed here
  (follow-up P1).
- The warn policy costs nothing when nothing is refused (single command),
  and one retry (+9 µs, 1.37–1.39×) per refused property.
- **Miss**: a Warn batch does not coalesce, so a 100-cell XLSX batch is about 49×
  the coalesced Refuse batch (about 176 µs per item, the best-effort cost),
  and each item retains its own snapshot: about 850 XLSX items reach the
  default 256 MiB snapshot budget, exactly as `--best-effort` does today.
  Corpus batches are small. Follow-up P2.

## 7. Gaps left and follow-ups

- **P1** PPTX `add` overhead (1.4× at one shape): the unified add reads the
  slide scene for the direct-shape ordinal and adopts the new owner; the
  leaf leg does neither.
- **P2** Coalescing under `Unsupported::Warn`: pre-admit each item's
  properties against the registry so atomic warn batches keep the native
  runs.
- **P3** Python: `File.apply`/`File.batch` should take
  `unsupported="warn"` and surface `warnings` (`LitchiWarning`, record 0015
  §2.2). Not done here (owned by the bindings stream).
- **P4** Shape `fill` with alpha (`00000000`, 25 lines of
  `ppt/3d-model.sh`) and scheme colours need an alpha/scheme-aware leaf
  colour; `opacity`, `line`, `lineWidth`, `lineColor` (L-PPTX-R6b) remain.
  Shape `name` on add (111 warned lines) and `autoFit` (117, L-PPTX-R7b).
- **P5** Reads: shape `fill`/`geometry` and slide `transition` are
  write-only (read gaps L-PPTX-R6b, M1-extended).
- **P6** `add slide --prop transition`: the source-backed insertion's
  markup-compatibility refusal (`opened/copy_plan.rs`, `reject_mce`) blocks
  inserting a slide with an extension transition in one edit.
- **P7** Clearing a Morph/p14 transition keeps the owner's `mc:Fallback`
  (`<p:fade/>`) with an empty `mc:Choice` (leaf
  `clear_alternate_fragment`); OfficeCLI removes the whole owner. Readers
  that select the Choice see no transition; older readers still fade.
- Chart `add` (764 lines) belongs to `uo2/s19-charts`.
- New delta **D4b**: a batch item adding a slide without `layout` (or with
  `layout=blank`) is not translated by the harness (6 lines).

## 8. Gates

Run with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pptx -p litchi-cli -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-pptx --all-features` (1,770 passed) | 0 |
| `cargo test -p litchi-ops --all-features` (1,896 passed) | 0 |
| `cargo test -p litchi --all-features` (260 passed) | 0 |
| `cargo test -p litchi-cli --all-features` (24 passed, corpus included) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pptx -p litchi-ops -p litchi-cli` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (21 profiles, 1,043 bans, 33 families) | 0 |
| `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI> --check` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_no_3rdparty_references tools.test_officecli_capability_delta` | 0 |

Python bindings are unchanged (they compile under the clippy gate; P3).

## 9. Not claimed

- "Passed" still means `litchi` accepted the call and the file reads back;
  content is not compared with OfficeCLI's output.
- Warned calls are partial by definition; the table keeps every skipped
  property visible as a gap.
