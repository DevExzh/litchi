# 0275 — RTF to parity (plan step 21)

Plan step 21 of [record 0248](0248-coordinator-restart-plan.md) extends the
unified RTF adapter, which records 0232–0247 left at the main-story
paragraph/run loop, to tables, header/footer stories, footnotes/endnotes,
comments, hyperlinks, pictures, styles and lists, and attempts the tiny-input
overhead misses of records 0236–0246 (owner rule D5).

## Leaf changes (`litchi-rtf`)

**Exact-source destination publication.** Table-cell, header/footer
paragraph, footnote/endnote and comment text edits used to publish through the
canonical writer. That writer reserializes the whole document and emits every
body table after the body story, so editing one cell moved the table to the
end of the document and dropped source formatting such as the `\super` around
note references. The parser now records, in a crate-private provenance table,
the source byte range of every top-level cell, header/footer paragraph, note
body and comment body, and the offset of the control or brace that ends it.
`edit/destination_splice.rs` splices the replacement into that range:

- the range may contain only text-producing tokens at its own depth and
  balanced groups of character-formatting controls below it, so the splice
  cannot change parser state for later content; anything else (a `\pard`
  inside a cell, a destination, binary data) is a typed `UnsupportedSource`;
- an empty story is filled at its terminator, with a control-word delimiter
  only when the preceding control word would absorb the replacement;
- replacement text is escaped, tabs become `\tab`, non-ASCII becomes a scoped
  `{\uc1\uN?}` group, and note/comment paragraph breaks become `\par`;
- the candidate is reparsed under the source's open policy and compared with
  the source on every retained story (body text, every cell, header/footer
  paragraph, note, comment, field, picture and resource count) before
  publication.

Every byte outside the selected stories is copied verbatim, unknown
destinations included, and inverse patches restore the exact source bytes.
The canonical writer remains for the other destination edits (shape text,
SmartTag, move bookmarks) and as the fallback when a splice is refused in a
document without body tables; with body tables it now refuses instead of
relocating them.
Two refusal tests that asserted the canonical path's opaque-syntax refusal
now assert exact splices.

**Linear destination batches.** Staging checked destination uniqueness by
rebuilding every staged operation's effect strings, and the candidate check
searched the operation list once per story: both quadratic. A per-edit effect
set and keyed override maps make a native 1,000-cell batch about 12× the
100-cell cost for a source that is also 10× larger (previously 65×).

**Parser fixes.** `\tab` inside a table cell now belongs to the cell story
(it was appended to the body text). `\footnote\ftnalt` is recognized as an
endnote (it was a footnote plus an opaque control), and the canonical writer
emits endnotes as `\footnote\ftnalt` instead of the non-standard `\endnote`.

**Read accessors.** `text::Paragraph::text_range`,
`ParagraphFormat::{style, list_override, list_level}`,
`Format::character_style` and `Package`/`Document::picture_body_offset` (the
body offset of pictures opened directly in the main story).

## Unified surface (`litchi-ops`)

| Path | Reads | Writes |
|---|---|---|
| `/body/tbl[N]`, `/tr[M]`, `/tc[K]` | table `rows`, `cols`; row `header`, `cantSplit`, `height`; cell `text`, `width`, `valign`, `vmerge`, `hmerge` | cell `text` |
| `/header[N]`, `/footer[N]` (`[@type=default\|first\|even\|odd]`), `/p[M]` | story `text`, `type`; paragraph `text` | paragraph `text`; story `text` when it has one paragraph |
| `/footnote[N]`, `/endnote[N]` | `text`, `number` | `text` |
| `/comments/comment[N]` | `text`, `author`, `initials`, `date` | `text` |
| `/body/p[N]/hyperlink[M]` | `text` (cached result), `url`, `anchor`, `tooltip`, `tgtFrame` | — |
| `/body/p[N]/picture[M]` | `width`, `height`, `type` | `remove` |
| `/styles/<name>` | `name`, `type`, `id`, `basedOn`, `next` | — |
| `/numbering/abstractNum[N]`, `/level[N]`, `/numbering/num[N]` | list `id`, `name`, `levels`; level `ilvl`, `numFmt`, `lvlText`, `start`; override `id`, `abstractNumId` | — |
| `/body/p[N]` | adds `style`, `numId`, `numLevel` | — |
| `/body/p[N]/r[M]` | adds `rStyle` | — |

Destination writes resolve against one snapshot and publish as one native
transaction for a path, a query fan-out (`tc`, `footnote`, `comment`,
`picture`) or an edit closure: consecutive direct destination sets coalesce
into one native publication, closing the window only on a family change, a
full window or a second write to the same story. Mixing body paragraphs,
runs, destination stories and pictures in one selector is a typed
`unsupported_type`; setting a read-only property of a projected element is
`unsupported_capability`.

The native editor addresses a header/footer by section and kind, so a story
that repeats an earlier kind in its section reads normally but refuses writes
with `unsupported_capability`.

Native identities (`commentId`, list `id`, `ilvl`, `footnoteId`) are declared
for the Word vocabulary but stay raw-path only (ADR 0034 §8); `level[N]` is
the Nth level. Header/footer paragraphs are addressed by path and listed by
views; unscoped `p` queries stay in the main story, preserving the records
0232–0247 query semantics. Body children list paragraphs before tables,
because the native model keeps tables apart from the paragraph flow.

The RTF settable/addable/removable element lists and write gaps moved from
`capabilities.rs` into `format/rtf/spec.rs`, so later RTF changes stay in
owned files. `node::view::Source` has one `Rtf` variant delegating to the
adapter's `Props`.

## Gaps left (typed in `capabilities()`)

Element read gaps: shape, textbox, field, formfield, bookmark, ole,
equation, revision, section. Operation gaps (follow-up named per gap):

- tables: add/remove table, row and cell; table/row/cell properties other
  than cell `text`; nested-table cells; per-paragraph cell nodes (`tc/p`);
  body child order;
- header/footer, note and comment lifecycle (add/remove); header/footer
  paragraph runs; whole multi-paragraph story `text`; a second story of the
  same kind in one section; stories with drawings, fields or page events
  (native refusal); `query` does not reach header/footer paragraphs;
- hyperlink set/add/remove (needs a native field-instruction editor);
- picture add (native `\pict` insertion) and property set; removal is limited
  to standalone hexadecimal PNG/JPEG body groups, at most 64 per transaction;
- stylesheet set/add/remove and style application (`style`, `rStyle`):
  applying a style needs the referenced definition's properties emitted with
  the handle, which the paragraph-layout writer refuses today;
- list editing and `numId`/`numLevel` writes;
- performance: direct `remove` commands are not coalesced (the transaction
  coordinator coalesces `set` only), and single-target query fan-out and
  single edit-closure sets keep a fixed overhead (see below).

## Verification

Native tests: `destination_splice.rs` (exact-byte splices around a table
between paragraphs, formatted and empty cells, escapes and Unicode, header and
note splices in documents with tables, a 1,000-cell batch, exact inverse
bytes, refusal of a `\pard` inside a cell, header paragraph breaks, cell
tabs) and `story_references.rs` (picture anchors, style/list references,
`\ftnalt`). Unified tests: `rtf_unified_stories.rs` covers every new path and
property, exact-byte destination sets, empty cells and Unicode through save,
query fan-out as one publication with exact inverse, atomic mixed-story
batches, typed refusals, picture removal, capability listings, coalesced
direct destination sets and repeated header kinds.

Intended golden changes: the RTF registry snapshot (re-blessed with
`LITCHI_BLESS=1`), the OfficeCLI delta table (regenerated; RTF elements
supported 7 → 22, property reads supported 25 → 68), and three lines of the
facade Markdown golden: two fixtures whose table-cell tabs no longer leak into
the body text, and one `\ftnalt` endnote fixture that Markdown no longer
refuses. `rtf_unified_read.rs` now asserts that header stories read while
`shape` stays a pending capability.

Gates at `9abcf6b51a` (each exit status checked):

| Command | Exit | Result |
|---|---:|---|
| `cargo fmt --all --check` | 0 | |
| `cargo clippy -p litchi-rtf -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 | |
| `cargo clippy -p litchi-ops --no-default-features --features rtf --all-targets -- -D warnings` | 0 | RTF-only facade |
| `cargo test -p litchi-rtf --all-features` | 0 | 1,537 tests/doctests |
| `cargo test -p litchi-ops --all-features` | 0 | 1,980 tests/doctests |
| `cargo test -p litchi --all-features` | 0 | 260 tests/doctests |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-rtf -p litchi-ops` | 0 | |
| `python3 tools/check_crate_boundaries.py` | 0 | 67 packages, 271 dependencies, 11 debts |
| `python3 tools/non_iwork_gate.py verify` | 0 | |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile rtf-minimal` / `ops-rtf` | 0 / 0 | 86 bans + 1 family / 1 ban |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 | 252 tests |
| `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI`, then `--check` | 0 / 0 | identical to the vendored-schema output |



## Performance

### Destination writes

The new [standalone harness](../../tools/perf-baseline/unified-rtf-destination/README.md)
(`rtf_destination_benchmark.v1`) compares the native destination editor with
three unified routes: one direct path `set` per target inside an edit closure
(coalesced), one query fan-out writing a shared value (against a native leg
writing the same value), and one direct `File::set`/`remove` for single
targets. Paths, properties and native operations are prepared outside timing;
every sample's output is compared byte-for-byte with the native oracle.
Three trials pinned to CPU 30 at `9abcf6b51a` (two warm-up and nine
`ABCDEEDCBA` rounds, 18 samples per route) give these ranges:

| Family | Targets | Native median µs | Batch/native | Fan-out/native | Direct/native |
|---|---:|---:|---:|---:|---:|
| table_cell_text | 1 | 5.1–5.2 | 1.114–1.129 | 1.358–1.383 | 1.028–1.041 |
| table_cell_text | 100 | 182.6–187.0 | 1.156–1.165 | 1.137–1.140 | — |
| table_cell_text | 1000 | 2185.8–2231.3 | 1.138–1.141 | 1.101–1.115 | — |
| footnote_text | 1 | 5.0–5.2 | 1.103–1.140 | 1.325–1.358 | 1.010–1.039 |
| footnote_text | 100 | 186.6–192.4 | 1.105–1.112 | 1.054–1.066 | — |
| footnote_text | 1000 | 2490.4–2522.5 | 1.090–1.101 | 1.033–1.041 | — |
| header_paragraph_text | 1 | 4.4–4.5 | 1.116–1.122 | — | 1.027–1.046 |
| header_paragraph_text | 100 | 73.0–73.3 | 1.304–1.313 | — | — |
| header_paragraph_text | 1000 | 840.1–848.0 | 1.281–1.295 | — | — |
| comment_text | 1 | 5.4–5.6 | 1.104–1.126 | 1.348–1.353 | 1.027–1.033 |
| comment_text | 100 | 221.5–224.5 | 1.087–1.120 | 1.063–1.075 | — |
| comment_text | 1000 | 2536.8–2713.1 | 1.076–1.158 | 1.028–1.101 | — |
| picture_removal | 1 | 3.6–3.7 | 1.088–1.106 | 1.527–1.531 | 1.046–1.056 |
| picture_removal | 64 | 45.8–48.7 | 52.383–55.583 | 1.346–1.407 | — |

- **1,000-target batches** stay within the 2× budget: 1.08–1.16 for cells,
  notes and comments, 1.28–1.30 for header/footer paragraphs (whose native
  per-target cost is lowest, so the per-command transaction cost shows most).
- **Single writes** through `File::set`/`remove` are 1.01–1.06×, inside the
  10% budget. The same single write inside an edit closure is 1.09–1.14×, and
  a one-target query fan-out 1.33–1.53×: the query plan and walker cost is not
  amortized over one target. These are misses.
- **Picture removal** is capped at 64 pictures per native transaction. One
  query `remove` of 64 pictures is 1.35–1.41×; 64 direct `remove` commands are
  52–56× because each publishes separately (removal is not coalesced; the
  transaction coordinator only coalesces `set`).
- **Native scaling.** Before the linear-batch fix a probe run of the same
  harness measured native 1,000-target batches at 28.0 ms (cells),
  16.7 ms (notes), 25.8 ms (header paragraphs) and 16.7 ms (comments), 46–83×
  their 100-target cost. After it they are 2.19–2.23, 2.49–2.52, 0.84–0.85
  and 2.54–2.71 ms, 11.4–13.5× the 100-target cost for a source that also
  grows 10×. The probe is not a receipt.

Receipts: [trial 1](results/0275-rtf-parity/destination-trial-1.json),
[trial 2](results/0275-rtf-parity/destination-trial-2.json),
[trial 3](results/0275-rtf-parity/destination-trial-3.json), with raw samples,
commit/tree, compiler, CPU and binary size/SHA-256.

### Tiny-input overhead (D5)

Profiling the one-paragraph `set` showed generic path resolution costing about
190 ns of a 195 ns native no-op edit: the resolver validates the path by
formatting it, folds names through the vocabulary and walks the selector
machinery for `/body/p[1]`. Exact positional main-story paths (`/body`,
`/body/p[N]`, `/body/p[N]/r[M]`, `/body/p[N]/picture[M]`) and exact destination
paths now resolve directly against the native owner; any other path, and any
miss, takes the generic resolver so errors are unchanged. Single targets stage
through an inline one-or-many collection instead of heap vectors, the node and
child cursor no longer carry duplicate paragraph/run copies, and the hot
paragraph/run child axes keep only the cursor they read.

The four existing harnesses were rebuilt from the base commit (`git archive
395d0ffa9f`) and from `9abcf6b51a`, and run in process-level
base/new/new/base order pinned to CPU 30. Ratios are unified/native medians
within each process, as each harness reports them:

| Case | Base `395d0ffa9f` | This branch |
|---|---:|---:|
| `add` paragraph, 0 paragraphs | 1.115–1.134 | 1.077–1.090 |
| `add` paragraph, 1 paragraph | 1.103–1.111 | 1.058–1.077 |
| `add` run, 0–1 groups (12 cases) | 1.083–1.131 | 1.043–1.091 |
| layout, 1 paragraph, path | 1.260 | 1.144–1.157 |
| layout, 10 paragraphs, path | 1.092 | 1.048 |
| layout, 1 paragraph, query | 1.440–1.456 | 1.449–1.468 |
| layout, 10 paragraphs, query | 1.149–1.150 | 1.165–1.179 |
| edit-closure set, 1 paragraph text | 1.171–1.173 | 1.175–1.184 |
| edit-closure set, 1 layout+alignment | 1.349–1.353 | 1.275–1.304 |
| edit-closure set, 1 run text / compound format | 1.082–1.120 | 1.079–1.126 |
| edit-closure set, 1,000 layout+alignment | 1.684–1.721 | 1.590–1.602 |

Paragraph and run insertion on empty and one-group inputs are now inside the
10% single-write reference; the layout path improves but stays a miss
(1.14–1.16): the remaining unified cost is coercing nine string-valued
properties and the paragraph-property scan, which the native leg receives
pre-typed. Single-target query fan-out (1.45×) and single sets inside an edit
closure (1.08–1.30×) are unchanged misses: query planning and the transaction
coordinator cost a fixed few hundred nanoseconds that one target cannot
amortize. These need work in the shared query and transaction layers
(follow-up), outside this step's owned files. All other cases with at least
ten targets are within their budgets. Receipt:
[tiny-input ABBA](results/0275-rtf-parity/tiny-input-abba.json) (per-case
ratios and native medians of all four processes, binary SHA-256 of both
builds).

## Commits and cleanup

`1de948afe7` (exact-source destination publication, cell tabs),
`401c45fca4` (read accessors), `9b807de843` (linear destination batches,
`\ftnalt`, Markdown golden), `9abcf6b51a` (unified surface, overhead work,
registry golden, delta table, harness) and this record. Harness builds used
the root `Cargo.lock` copied next to each standalone manifest (not committed)
and an external target; no binary or evidence file of 1 MiB or more is
committed.

## Not claimed

No full-corpus round-trip claim, no allocation qualification, and no claim
for formats other than RTF. Shared-host timings support only the stated
cases.
