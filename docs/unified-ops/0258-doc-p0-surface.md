# 0258 — DOC P0 read+write surface through the unified API

Date: 2026-10-13. Status: implements record 0248 §6 Phase C step 11.
Branch `uo2/p11-doc`, cut from trunk `e467bcf4e8`.

DOC (Word 97–2003) now reaches the P0 read and write surface through
`litchi_ops::File`. Reads cover runs, top-level tables, header and footer
stories, notes and comments. Writes cover run and paragraph character
formatting, paragraph text, and paragraph add, remove, move and swap. Every
write goes through the primary `litchi_doc::Package::edit` transaction.

## 1. Reader port (stale `uo/m1`)

Six `uo/m1` commits were ported onto the immutable-owner trunk. Trunk had none
of them. The old hashes are those of the stale branch.

| Old commit | Content |
|---|---|
| 59e3003ca2 | Word 97 tables that store no table depth |
| 4ad07d24a4 | compressed 0x93/0x94 decoded as curly double quotes |
| 63e55e1c21 | per-story reads (`story_blocks`, `story_paragraphs`, `body_blocks`, `body_paragraph`, `body_table`), `CpRange` spans on paragraphs, runs, tables, rows and cells, and the weighted story-index memo (ADR 0005) |
| d08a33bce6 | font table and run font names (`font_table`, `font_name`) |
| 46f79e0d74 + 7a606ea812 | header and footer slots by section and kind, with inheritance; notes and comments by index |

How the port was adapted:
- State fields are reached through `Document(Arc<DocumentState>)`.
- Tests use `raw::Package` and `writer::Builder`.
- Positional readers take `Position` instead of `usize` (ADR 0034 §8).

Two later additions:
- `StoryBlocks::paragraph_text` returns body paragraph text through the index without building anything.
- Corpus tests check that the indexed text equals the body scan. 492 paragraphs were compared and all matched.

The `crates/litchi` Markdown corpus golden changes on three DOC fixtures. Each change was reviewed against its cause:
- `3endnotes.doc` gains 14 curly quotes and `saved-by-table.doc` gains 6. Each quote adds 2 bytes.
- `HeaderFooterProblematic.doc` is now read; it rendered 6 bytes before. Its Word 97 tables store no depth.

## 2. Leaf write engine (`litchi-doc`)

The primary `Edit` gains four operations:

- **`TextTarget::BodySpan { paragraph, start, end }`**: a UTF-16 range of one ordinary main-story paragraph, that is, one run. Text replacement and character formatting accept it, and durable patches encode it.
- **`set_character_value(target, CharacterValue::{Size, Color, Font})`**:
  - Size is written with sprmCHps.
  - Color is written with sprmCCv. sprmCIco is added when the color is one of the 16 Word 97 palette colors. `Color(None)` clears the direct color.
  - Font is written with sprmCRgFtc0. The font must already be in the font table. An unknown font is `Refusal::FontNotFound`, because adding a font is a resource change.
  - Setting the value the target already has is an exact no-op.
- **`insert_paragraph`, `remove_paragraph`, `move_paragraph`**:
  - An inserted paragraph copies its neighbour's PAPX and the CHPX of the neighbour's paragraph mark.
  - A moved paragraph keeps its exact units, CHPX runs and PAPX.
  - The editor appends one Unicode piece, shifts the modeled CP tables, rebuilds the CHPX/PAPX FKPs and the CLX, and splices Selsf.
- **Refusals**:
  - `SoleParagraph`: removing the only body paragraph.
  - `AnchoredContent`: a field, bookmark, note, comment, drawing or proofing anchor strictly inside the paragraph, or a section the edit would empty.
  - `TableBoundary`: the edit would join two tables or end the story in a table.
  - Tracked text, structural text and unmodeled CP tables are also refused.
- **Encrypted packages**: `Package::edit` now refuses them with `Refusal::Encrypted`. Before this change, the editor reported encrypted ciphertext as a corrupt file.

Patches invert exactly, because they hold byte snapshots. The durable envelope refuses the new structural and valued-format changes with a typed `PatchError`.

A corpus sweep tried an insert, a removal and a move on every corpus file. 48 edits were admitted and read back as expected; 18 were refused with typed refusals. Tests: `crates/litchi-doc/tests/doc_body_structure.rs`.

## 3. Unified adapter (`litchi-ops/src/format/doc`)

### Reads

Paths use the OfficeCLI DOCX vocabulary:

| Path | Properties |
|---|---|
| `/body/p[N]` | `text`, `align`, `spaceBefore`, `spaceAfter`, `indent`, `rightIndent`, `firstLineIndent`, `keepNext`, `keepLines`, `pageBreakBefore` |
| `/body/p[N]/r[M]` | `text`, `font`, `size`, `bold`, `italic`, `color`, `underline`, `strike`, `caps`, `smallcaps`, `superscript`, `subscript` |
| `/body/tbl[N]` | `rows`, `cols` |
| `/body/tbl[N]/tr[M]` | `header` |
| `/body/tbl[N]/tr[M]/tc[K]` | `text`, `fill`; cell paragraphs and runs below it |
| `/header[N]`, `/footer[N]` (also `[@type=default\|even\|first]`) | `text`, `type` |
| `/footnote[N]`, `/endnote[N]` | `text`, `number` |
| `/comments/comment[N]` | `text`, `author`, `initials` |

Headers, footers, notes and comments also expose their paragraphs and runs.

The body lists paragraphs and top-level tables in document order, through the story index. A body without tables keeps the single-scan borrowed walk. Main-story paragraphs are lazy: the native paragraph is built only when its format or runs are read.

Capabilities:
- The 22 element gaps of record 0249 shrink to 13: hyperlink, equation, section, style, numbering, bookmark, field, formfield, picture, ole, revision, textbox and shape.
- The capability golden `golden/capabilities/doc.json` was re-blessed.
- The OfficeCLI delta table is unchanged, because DOC is not yet in it.

### Writes

Each `set`, `add`, `remove`, `move_to` and `swap` call is one native commit:
- `set` on `/body/p[N]` takes `text`, `bold`, `italic`, `underline` (`none`/`single`), `size`, `color` (RGB or `auto`) and `font`. Format keys apply to the whole paragraph.
- `set` on `/body/p[N]/r[M]` takes the same keys for one run.
- A scoped query sets every match in one commit. Run targets are staged from the last span to the first.
- `add(p)` under `/body` accepts any placement.
- `remove` takes paragraphs, or a run, whose text it clears.
- `move_to` and `swap` act on paragraphs, which keep their runs and PAPX.

Transactions and batches dispatch the same calls.

Behaviour:
- **Exact no-ops**: a format key is not staged when every selected run already reads as that value. Unchanged text is skipped. A call that changes nothing returns the source bytes.
- **Signed documents**: every verb refuses with `document_protected` before staging. After `Signed::unsign`, edits are allowed.
- **Encryption**: an unlocked encrypted DOC refuses every write with `document_protected` (`Refusal::Encrypted`). Nothing is written as plaintext.

Error codes: `FontNotFound` → `not_found`, `InvalidFontSize` → `invalid_value`, `SoleParagraph` → `invalid_argument`, `Encrypted` → `document_protected`. Other refusals remain `unsupported_capability`.

Tests: `crates/litchi-ops/tests/doc_unified_p0.rs` (9 tests: reads, round trips after reopen, no-ops, structure, refusals, signature, encryption, transactions, and a corpus check that runs spell their paragraph over 304 paragraphs in 30 files). The narrow-capability test in `doc_unified_ops.rs` and the Python capability-agreement test were updated for the new verbs.

## 4. Performance

Setup:
- Release build of `rustc 1.95.0`, pinned to CPU 26 (idle).
- ABBA rounds: 2 warm-up and 9 measured, giving 18 samples per leg.
- Both legs' output bytes were compared before timing.
- Source commit `d5ff71feaf`.
- Build: `CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=<target> cargo build --release --manifest-path tools/perf-baseline/<package>/Cargo.toml`.
- Binary SHA-256 values are in `results/0258-doc-p0/binaries.sha256`. The binaries are not committed.

**Unified against the primary leaf** (`tools/perf-baseline/unified-doc-p0`, new route; receipts `results/0258-doc-p0/p0-trial-{1,2,3}.json`). Medians are from trial 1. The pre-fix ratio is a single diagnostic run before the indexed lookup (`p0-before-lookup.json`).

| Scenario | N | native µs | unified µs | unified/native, trials 1–3 | pre-fix |
|---|---:|---:|---:|---|---:|
| `paragraph_text_set` | 1 | 45.7 | 47.5 | 1.040 / 1.030 / 1.043 | 1.036 |
| `paragraph_text_set` | 100 | 233.3 | 242.9 | 1.041 / 1.034 / 1.034 | 0.996 |
| `paragraph_text_set` | 1000 | 2719.3 | 2766.5 | 1.017 / 1.022 / 1.019 | 1.010 |
| `run_format_set` | 1 | 63.8 | 67.8 | 1.063 / 1.085 / 1.055 | 1.081 |
| `run_format_set` | 100 | 363.9 | 374.5 | 1.029 / 1.024 / 1.024 | 1.008 |
| `run_format_set` | 1000 | 3943.1 | 3938.1 | 0.999 / 1.007 / 1.000 | 1.018 |
| `paragraph_insert` | 1 | 46.1 | 48.0 | 1.040 / 1.041 / 1.038 | 0.990 |
| `paragraph_insert` | 100 | 282.7 | 285.8 | 1.011 / 1.038 / 1.044 | 1.038 |
| `paragraph_insert` | 1000 | 3462.9 | 3493.8 | 1.009 / 1.010 / 1.011 | 1.021 |
| `paragraph_text_get` | 10 | 0.2 | 0.3 | 1.261 / 1.261 / 1.280 | 3.040 |
| `paragraph_text_get` | 100 | 2.1 | 0.3 | 0.143 / 0.139 / 0.150 | 2.134 |
| `paragraph_text_get` | 1000 | 24.8 | 0.3 | 0.014 / 0.013 / 0.013 | 2.025 |
| `run_bold_get` | 100 | 1.0 | 1.6 | 1.684 / 1.671 / 1.698 | 4.041 |
| `run_bold_get` | 1000 | 1.0 | 1.7 | 1.677 / 1.727 / 1.714 | 27.671 |

Rows for N=10 are in the receipts; all of them stay below 1.09 for writes. Findings:
- **Writes** (text set, run format set, insert) stay within the 10% single-op budget in every trial. The worst is 8.5%, for `run_format_set` at N=1.
- **Fixed in reads**: before the fix, the adapter resolved `/body/p[N]` by scanning the main story, and then scanned it again to read the text. It now resolves through the memoized story index and keeps the text on the node. `paragraph_text_get` therefore costs a constant 0.3 µs, where the leaf's scan costs 25 µs at N=1000. `run_bold_get` falls from 27.7× to 1.7× the leaf.
- **Not fixed in reads**: about 0.6 µs of fixed node, path and format cost remains on point reads. This shows as ratios of 1.7× on a 1 µs leaf, and 5× on a 0.1 µs leaf at N=1.

**Frozen 0188/0190 route** (`tools/perf-baseline/unified-doc-paragraph`: native body-text `Snapshot` against `File::set`, with a fresh `File` per sample). Trunk `e467bcf4e8` and this branch were built side by side and run alternately (T M M T T M). Receipts are `results/0258-doc-p0/frozen-paragraph-{T,M}-*.json`. Unified overhead:

| N | trunk (3 trials) | branch (3 trials) |
|---:|---|---|
| 1 | +1.2 / +0.4 / +2.4% | +1.2 / +0.3 / +0.3% |
| 10 | −2.9 / −1.0 / −0.7% | +2.7 / +1.2 / +1.6% |
| 100 | +9.6 / +7.9 / +10.1% | +8.5 / +8.7 / +7.4% |
| 1000 | −0.4 / −0.1 / +0.8% | +1.5 / +1.8 / +0.9% |

What this shows:
- The 13.6% of record 0190 does not reproduce on this machine today. Trunk measures 7.9–10.1% at N=100, and this branch measures 7.4–8.7%, within the 10% budget.
- At N=100, absolute time drops by about 12% on both legs: native 213 → 188 µs, unified 233 → 204 µs. The likely cause is the ported reader, which scans each range from its own offset; this is not separately proven.
- A profile of the N=100 route puts 18% of samples in `memmove`, copying whole-file buffers in the native validation, the package commit and source projection paths. Those copies are leaf work shared by both owners. The unified layer adds path resolution and handle construction, which the primary-leaf pairing above bounds at ≤4% for this scenario.

No claim of improved DOC open or write speed beyond these receipts is made.

## 5. ADR 0034 W4 residue

- **Part-level ingress**: the six CFB part editors (annotation bookmarks, captions, command bars, envelope, route slip, subdocuments) replace `open_with_limits` and `open_with_limits_and_policy` with `open_with_options(bytes, &EditOptions)`. The options' revision limits bound the parse, and their protection policy governs changed protected content.
- **Other configured opens**:
  - `SourceSnapshot` keeps `open_with_options(source, SourceBackedOptions)`.
  - `raw::Package::open_with_limits` becomes `open_with_options(path, OpenOptions)`.
  - Smart tags use `Snapshot::parse_with_options_and_policy` with `Editor::from_snapshot`.
- **Rollback**: all ten DOC `rollback` methods are deleted. Dropping a transaction is the rollback.
- **Glossary**:
  - `tools/api_glossary.json` bans `open_with_limits`, `open_with_limits_and_policy` and `rollback` crate-wide in doc-minimal and doc-full.
  - `docs/API_GLOSSARY.md` flips the Configured-open and Rollback rows.
  - The new DOC vocabulary is recorded under "DOC P0 surface and W4 residue".

## 6. Verification

Environment: `CARGO_TARGET_DIR=<worktree>/target`, `CARGO_INCREMENTAL=0`, debug info off, toolchain 1.95.0.

| Command | Exit | Result |
|---|---|---|
| `cargo fmt --all --check` | 0 | |
| `cargo clippy -p litchi-doc -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 | |
| `cargo test -p litchi-doc --all-features` | 0 | 1,478 passed |
| `cargo test -p litchi-ops --all-features` | 0 | 1,780 passed |
| `cargo test -p litchi --all-features` | 0 | 260 passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-doc -p litchi-ops` | 0 | |
| `python3 tools/check_crate_boundaries.py` | 0 | |
| `python3 tools/non_iwork_gate.py verify` | 0 | 35 facade-safe trees, 1 combined tree |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir <wt>/target/glossary` | 0 | 21 profiles, 845 banned names, 14 families |
| `python3 -m unittest tools.test_check_api_glossary` | 0 | OK |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 | 242 tests OK |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI> --check` | 0 | unchanged |

The `soapberry-zip` and `litchi-opc` test targets read corpus files that the
2026-10-05 history rewrite purged, so a workspace-wide `--all-targets` check
fails there. This happens on trunk too and is unrelated to this change.

## 7. Gaps and follow-ups

- **Element gaps (13)**: hyperlink, equation, section, style, numbering, bookmark, field, formfield, picture, ole, revision, textbox and shape. Follow-up: a record after 0258.
- **Writes not admitted, all typed refusals**:
  - writes to paragraphs of cells, headers, footers, notes and comments;
  - run `add`;
  - underline styles other than `none`/`single`;
  - theme and transformed colors;
  - fonts the font table does not name, because adding one is a resource edit;
  - paragraph-property writes (`align`, spacing, indents);
  - text replacement over runs with byte-different CHPX, such as rsid-only differences. The leaf's uniform-format proof refuses these.
- **Structural edit limits**:
  - Removal refuses paragraphs with an anchor strictly inside, which includes proofing marks. On real documents with spelling state, this refuses often.
  - An inserted paragraph that ends at a bookmark-end CP joins that bookmark.
- **Durable patches** refuse the new change kinds.
- **Unscoped queries**: `query("p")` now also returns paragraphs of tables, headers, notes and comments, and writes refuse non-body targets. Use `/body > p` for body-only selection.
- **Views** (`text`, `annotated`, `outline`, `stats`) still list body paragraphs only.
- **DOC in the delta table**: not yet added (record 0249 follow-up).
- **Not done (W4)**: scoped `Transaction` type renames (doc 11), and `open(bytes)` on part editors, which is not the canonical byte entry.
