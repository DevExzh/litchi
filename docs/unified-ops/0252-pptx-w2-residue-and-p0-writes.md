# 0252 — PPTX ADR 0034 W2 residue and P0 writes

Status: complete for this wave on branch `uo2/p7c-pptx` (plan 0248 §6 steps 5-pptx and 7c); gaps listed in §8.

## 1. Ingress options and patch conflict (W2 residue, part 1)

- `litchi_pptx::OpenOptions` (`source_limits`, `cache_limits`,
  `execution_context`, feature `encryption`: `encryption_limits`) is the one
  configured-open envelope. `OpenPolicy` names the lazy-only policies.
- Owning `Package`: `open_with_options`, `from_reader_with_options`,
  `from_bytes_with_options`, `open_with_password_and_options`,
  `from_reader_with_password_and_options`,
  `from_bytes_with_password_and_options`. Cache/execution policies are refused
  before input is read with `Error::UnsupportedOpenPolicy` (ops code
  `invalid_argument`).
- Deleted without aliases: `Package::{open_with_limits,
  from_reader_with_limits, from_bytes_with_limits, from_vec,
  from_vec_with_limits, from_vec_with_password,
  from_vec_with_password_and_limits, open_with_password_and_limits,
  from_reader_with_password_and_limits, from_bytes_with_password_and_limits}`
  and the six `open_with_*` combinations on `SourceBackedPresentation` and
  `SourceBackedPresentationEditor` (now `open_with_options`).
- Consumers formerly calling `from_vec(bytes)` now call
  `from_reader(Cursor::new(bytes))`, which keeps physical-source provenance
  (cross-slide copy authorization) at the cost of one bounded copy.
  `from_bytes` keeps its borrowed, provenance-free ingress. Affected
  perf-baseline legs (`pptx_*` routes in `tools/perf-baseline`) therefore
  include that copy; no performance claim about them is made here.
- `Error::StaleSource` is renamed `Error::PatchConflict` (ADR 0034 §6); ops
  still classifies it as `io_error` (the source changed under the handle).
- Glossary: both PPTX profiles ban the removed names; pptx-full also gains the
  primary-transaction bans pptx-minimal already had (plan 0248 §2 hole).

Not done here: positional `from_path_*`/`from_read_at_*` constructor
combinations on the source owners still exist (pending, as for DOCX in 0038).

## 2. Primary family completion (W2 residue, part 2)

- `litchi_pptx::opened` exports `Edit`, `Patch`, `Commit`, `Diagnostics`,
  `Error` and `LimitKind`; the family is registered in both PPTX glossary
  profiles, so the gate now checks one PPTX family.
- `opened::LimitKind` names the seven dimensions of `opened::Limits`
  (`Parts`, `PatchBytes`, `TextBytes`, `HistoryEntries`, `HistoryBytes`,
  `RetainedCandidateBytes`, `RetainedMceBytes`); `Limits::maximum(kind)`
  reads one bound.
- The 48 refusals that reported an `opened::Limits` bound as
  `Error::Limit { resource: "...", limit }` now report
  `Error::EditLimit { kind, maximum }` (ops: `limit_exceeded`). The
  LPTX0003 durable-patch golden transcript changed only in its 13
  `*.limits.*.refused` rows, which print that error.
- `opened::Error` is the crate `Error` re-exported. The ADR's per-module
  error split of the PPTX god-enum, and `observed` on edit limits, are not
  done (follow-up).

## 3. Rich text writes (plan 7c item 1) and table cells (item 2)

Leaf (`litchi-pptx`):
- `opened::{TextEdit, TextTarget, TextProperties, Rgb}` and one verb,
  `set_text_properties`, on the eager `opened::Edit`, the source
  `source::Edit` (composes through `stage_slide_edit`, so source-local
  replay and inverse patches cover it) and `SourceBackedSlideEdit`.
- Targets: a shape's whole body, a paragraph, a run (`a:r`/`a:br`/`a:fld`,
  counted as the text-body reader counts them), or a table cell.
  Properties: `bold`, `italic`, `size` (hundredths of a point,
  100–400000), `color` (`a:solidFill/a:srgbClr`), `font` (latin+ea+cs),
  `latin`, `align` (`a:pPr@algn`), paragraph/run/cell `text`, and cell
  `fill` (`a:tcPr` solid fill). Every property can be removed.
- Byte splices only: untouched attributes, siblings and extension lists keep
  their bytes; inserted children follow the `a:rPr`/`a:tcPr` schema order
  and the receiving element's prefix. The candidate slide is read back as a
  scene with identical shape count and identities. `table::Cell::fill()`
  reads a cell's local RGB fill.

Unified (`litchi-ops`), registry entries flipped to `set`:
- run: `bold`, `italic`, `size`, `color`, `font.latin`, plus write-only
  `font` and `text`;
- paragraph: `align`, `size`, plus write-only `text`, `bold`, `italic`,
  `color`, `font` (applied to every run);
- shape/placeholder: `bold`, `italic`, `size`, `color`, `font`,
  `font.latin`, `align` (every paragraph and run; shape `text`/geometry and
  formatting stage in one native edit);
- `tc`: readable/settable `fill`, plus write-only `text`, `bold`,
  `italic`, `size`, `color`, `font`, `align`. The `tc fill` gap
  (L-PPTX-R8) is closed.
- Colours: only opaque RGB without transforms is writable
  (`unsupported_capability` otherwise, with a suggestion). Run text with a
  line break is `invalid_value`.
- Path and query targets (`/slide[N]/shape[M]/paragraph[K]/run[L]`,
  `/slide[N]/table[M]/tr[R]/tc[C]`, scoped `slide[1] run` queries) fan out
  through the same native verb, 256 edits per call; batch and `edit`
  closures reuse `set` (the shape coalescer declines formatting payloads).
- Tests: `crates/litchi-pptx/src/opened/text_format_tests.rs` (6, both owners
  and inverse replay) and `crates/litchi-ops/tests/pptx_text_format.rs` (6:
  read-back via `get` on both owners and after a byte round trip, clears,
  query fan-out, atomic `edit` closures, typed refusals). Capability golden
  `pptx.json` re-blessed; OfficeCLI delta regenerated (pptx: 153→164
  supported, 280→271 missing).

## 4. Scoped family vocabulary (W2 residue, part 3)

- 12 snapshot-scoped `Transaction` types are renamed `Edit`;
  `AuthorPresenceTransaction`/`CommentThreadingTransaction` become
  `AuthorPresenceEdit`/`CommentThreadingEdit`. Three `rollback()` methods
  are deleted (no caller used them).
- Public `is_changed()` (and `changed()` aliases) on 40 edit, commit and
  patch types are removed. Edits and commits answer through
  `diagnostics() -> litchi_pptx::diagnostics::Diagnostics` (`changed()`);
  families with their own diagnostics type keep it. Patches answer through
  `is_noop()`; patch `is_empty()` is renamed `is_noop()`.
- Callers in the crate, its tests, `litchi-ops` and `tools/perf-baseline`
  are migrated. Both PPTX glossary profiles ban `Transaction`,
  `AuthorPresenceTransaction`, `CommentThreadingTransaction`, `rollback` and
  `is_changed` crate-wide; the gate passes.
- Not done: the scoped families keep their `Snapshot::edit()` entry and the
  crate error; the `Package::{x, edit_x, apply_x}` reshape and per-family
  `Error`/`LimitKind` are a follow-up.

## 5. Paired measurement: run formatting

Route family `pptx_{eager,source}_{native_primary,unified_file_set}_run_format_v1`
(new harness `tools/perf-baseline/unified-pptx-text-format`). One sample sets
`bold`, `size` 24pt and `color` 1F4E79 on `/slide[1]/shape[1]/paragraph[1]/run[1]`;
the native leg is one `Edit::set_text_properties` call plus one commit, the
unified leg one prepared `File::set`. Generated one-slide decks of 1–10,000
text boxes, both owners, 9 ABBA rounds (18 samples per leg), release build
(`debug = 1`), `taskset -c 30`, shared 32-core host (load ≈ 8–12 from other
agents). Binary SHA-256
`1735a2c30eacb6a2da38fe4ac246ab1902ea32b9dbfc8105a3a0d98b4b8da43f`, built at
commit 0de12f04d5 with
`CARGO_TARGET_DIR=target/bench CARGO_NET_OFFLINE=true cargo build --release --manifest-path tools/perf-baseline/unified-pptx-text-format/Cargo.toml`.
Warm-up outputs were byte-identical between legs in every case. Raw results:
`docs/unified-ops/results/0252-pptx-p0-writes/run-format-trial-{1,2,3}.json`.

| owner | shapes | native median µs (t1/t2/t3) | unified median µs | unified/native |
|---|---|---|---|---|
| eager | 1 | 305.5 / 302.2 / 301.9 | 318.8 / 321.1 / 314.4 | 1.044 / 1.062 / 1.042 |
| eager | 10 | 418.3 / 418.0 / 416.7 | 454.8 / 453.3 / 450.4 | 1.087 / 1.085 / 1.081 |
| eager | 100 | 1496.5 / 1498.4 / 1501.9 | 1762.0 / 1765.6 / 1758.7 | 1.177 / 1.178 / 1.171 |
| eager | 1000 | 12030.1 / 12088.5 / 12170.6 | 14607.9 / 14670.1 / 14579.1 | 1.214 / 1.214 / 1.198 |
| eager | 10000 | 116708.6 / 117570.0 / 118087.9 | 142280.8 / 143775.8 / 141843.3 | 1.219 / 1.223 / 1.201 |
| source | 1 | 56.7 / 56.4 / 56.6 | 61.3 / 60.9 / 61.0 | 1.082 / 1.080 / 1.077 |
| source | 10 | 271.6 / 270.6 / 270.7 | 279.4 / 279.2 / 276.5 | 1.028 / 1.032 / 1.021 |
| source | 100 | 2360.1 / 2399.9 / 2386.9 | 2398.0 / 2429.0 / 2401.9 | 1.016 / 1.012 / 1.006 |
| source | 1000 | 23652.8 / 23588.4 / 23407.5 | 23922.1 / 23918.8 / 23492.3 | 1.011 / 1.014 / 1.004 |
| source | 10000 | 234679.3 / 236586.0 / 234624.8 | 237656.6 / 238103.3 / 235161.8 | 1.013 / 1.006 / 1.002 |

Verdict against the single-set budget (≤ 10 %):
- source owner (the default open route): **met** at every size (0.2–8.2 %);
- eager owner: met at 1 and 10 shapes (4–9 %), **missed** at 100–10,000
  shapes (17–22 %). The overhead grows with slide size: the unified path
  resolves `/slide/shape/paragraph/run` by reading the eager slide's shape
  scene and text body, which the eager owner does not cache between calls,
  while the native leg is handed the positions. Follow-up: an eager scene
  cache in the PPTX node layer (plan step 8 owns `format/pptx/node.rs`), or
  a resolve-free fast path for direct positional run paths.

Not measured: paragraph, shape-level and cell formatting, query fan-out, and
the 1k-op batch for text formatting (batches reuse per-item `set`; the
shape coalescer declines formatting payloads).

## 6. Notes, comments, pictures and layout placeholders (plan 7c items 3–5)

Leaf (`litchi-pptx`):
- `opened::Edit::set_notes_text` creates a notes slide when the selected
  slide has none: `/ppt/notesSlides/notesSlideN.xml` related to the slide
  and the single notes master, slide → notes relationship, read back before
  the edit changes; failures restore the edit. Strict decks and decks
  without exactly one notes master are refused (`UnsafeEdit`).
- `opened::Edit::{add_comment_text, set_comment_at}`: positional legacy
  comment append/replace over the existing comment codec; authors are
  matched by exact name and created (next id) when absent; indexes follow
  the author's `lastIdx`.
- Slide insertion (`opened::Edit::insert_slide` and the source-backed
  `source::Edit::insert_slide`) instantiates the selected layout's
  placeholders as PowerPoint does (port of reference 9474e2b6ce
  `add_slide.rs`, adapted: the trunk insert APIs and their exact-name layout
  selection, durable patches and source topology are kept; only the slide
  body changes). Repeated placeholder identities get their layout bounds.
  The new slide is read back against the plan.
- `TextTarget::Shape` text replaces a whole body, so an empty placeholder
  (no run for the run-slot writer) can be filled.

Unified (`litchi-ops`):
- notes: `set(/slide[N], {notes})` creates or replaces the page;
  `set(/slide[N]/notes, {text})` replaces it (registry: slide `notes`
  write-only set, notes `text` set).
- comments: `add(/slide[N], comment, {text, author, initials, date, x, y})`
  returns `/slide[N]/comment[K]`; `set(/slide[N]/comment[K], {text, x, y})`.
- pictures: `add(/slide[N], picture, {src|path, name, x, y, width, height})`
  where `src` is a base64 `data:` URI or a local file path, bounded by the
  staging limit before decoding/reading, classified by signature
  (PNG/JPEG/GIF/BMP/TIFF); default extent 2in × 1.5in. `base64` becomes a
  dependency of the ops `pptx` feature.
- shape `text` on an empty body (a new placeholder) goes through the body
  edit; the shape coalescer declines such targets.
- These verbs edit separate parts the source-backed slide edit cannot stage,
  so a source-backed handle is promoted to an eager package (one bounded
  copy of the current bytes) and the result is an eager handle.
- Tests: `crates/litchi-pptx/tests/new_slide_content.rs` (4) and
  `crates/litchi-ops/tests/pptx_slide_content.rs` (6), both owners, byte
  round trips. Capability golden re-blessed; OfficeCLI delta: pptx 164→166
  supported.

Gaps left (follow-ups): picture `set` (geometry for `p:pic` needs the
source geometry proof extended to pictures), comment remove through the
unified `remove`, modern (threaded) comment writes, notes on strict decks
or decks without a notes master, p14 section maintenance for inserted and
moved slides (reference `sections.rs` not ported; trunk still refuses
section decks in source-backed order edits), and no paired measurement for
notes/comment/picture/slide-add (they are not single-property hot paths;
the source→eager promotion costs one copy and reparse of the package).

## 7. p14 section maintenance (eager owner)

Port of reference 9474e2b6ce `opened/sections.rs` (refusals re-typed as
`SlideOrderPlan { kind: Sections }`; no `SlideAdd` error family added):
- eager `insert_slide`: the new slide joins the section of the slide before
  it, or at the start the section of the slide after it, or the first
  section; eager slide-insertion admission now accepts the p14 section list
  and its `p:extLst/p:ext` wrapper (removal admission is unchanged);
- eager `move_slide`: the moved slide leaves its section and joins its new
  neighbour's; `swap_slides`/`reorder_slides`: membership is kept and each
  section is re-sorted into the new order;
- edits splice `p14:sldId` elements only and read the entry back.

Tests: `crates/litchi-pptx/tests/slide_sections.rs` (3) and the six ported
unit tests in `opened/sections.rs`. Not done: the source-backed owner (the
default unified open route) still refuses insertion and order edits on
section decks, because its root validators (`source_removal`,
`source_cross_copy`) do not model the extension; unified slide add/move on a
section deck therefore needs `SourceMode::Eager` until that follow-up.

## 8. Fixtures, merge, gates and remaining work

- Owner rule (no `3rdparty/` references): the three LibreOffice
  `sd/qa/unit/data/pptx` fixtures used by
  `crates/litchi-pptx/tests/source_backed_svg_lifecycle.rs` are vendored into
  `test-data/libreoffice-core/sd/qa/unit/data/pptx/` with `PROVENANCE.md`;
  drawingml doc comments name ECMA-376 / MS-ODRAWXML without the path. The
  9474e2b6ce port adds no such reference.
- Trunk (0249, 0250, 0254–0256) is merged. The Python write test that
  added slides from the first (title) layout now uses the `Blank` layout,
  because inserted slides now hold their layout's placeholders.
- Final gates (after the merge, all exit 0): `cargo fmt --all --check`;
  `cargo clippy -p litchi-pptx -p litchi-drawingml -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`;
  `cargo test` for litchi-pptx, litchi-drawingml, litchi-ops, litchi and
  litchi-cli (all features); `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps
  --all-features -p litchi-pptx -p litchi-drawingml -p litchi-ops`;
  `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify`;
  `tools/check_api_glossary.py` (21 profiles, 890 bans, 16 families);
  `tools/officecli_capability_delta.py --check`;
  `crates/pyo3-litchi/tests/run_tests.sh` (242 tests).

Remaining (follow-ups):
1. Eager-owner single-set overhead for run formatting at ≥100 shapes
   (17–22 %, §5): needs an eager scene cache or a positional fast path.
2. `opened::Error` is the crate god-enum re-exported; `EditLimit` has no
   `observed`; scoped families keep `Snapshot::edit()` and the crate error.
3. Positional `from_path_*`/`from_read_at_*` constructor combinations on the
   source owners.
4. Picture `set` (geometry/replace), comment `remove`, modern comment writes,
   notes on strict decks or decks without one notes master.
5. Source-backed owner: section decks are still refused for insertion and
   order edits; notes/comment/picture writes promote to an eager package.
6. `docx` still exposes `StaleSource` (outside this branch's ownership).
