# 0251a — DOCX element-path writes: nested paragraphs/runs, hyperlink text, table structure

Plan step 7a items 1–2 of [0248](0248-coordinator-restart-plan.md) §6 Phase B.
Branch `uo2/p7a-docx`; follows [0251](0251-docx-metadata-builder.md).

## Leaf: path-addressed main-document edits

`litchi_docx::document` gains these public types:

- `ElementPath`: the direct-child steps from `w:body`. The empty path is the
  body.
- `ElementStep`: `Paragraph`, `Run`, `Hyperlink`, `Table`, `Row`, `Cell` and
  `ContentControl`, each with the position among the container's direct
  WordprocessingML children of that name.
- `ElementAnchor`: `Start`, `Before(step)`, `After(step)` or `End`.
- `NewTable`.

Both the snapshot `Edit` and the package `Edit<P>` take these new verbs:

| Verb | Targets |
|---|---|
| `format_paragraph_at` / `format_run_at` | paragraph / run anywhere a path reaches |
| `replace_text_at` | paragraph, run or hyperlink text |
| `replace_cell_text_at` | a cell's blocks become one paragraph; `tcPr` and the first paragraph's `pPr` are kept |
| `insert_run_at` | runs in a paragraph or hyperlink, including `<w:p/>` |
| `insert_paragraph_at` / `insert_table_at` | body, cell or block content control, at an anchor |
| `insert_row_at` | a row with the table grid's widths, or a given cell count |
| `remove_element_at` | paragraph, run, hyperlink, table, row, cell, content control |
| `move_element_at` | any of the above into a container at an anchor |
| `element_xml` | read the staged bytes |

A path naming a direct-body paragraph, or a direct run of one, is routed to
the existing body-position operation, so those edits keep their incremental
layout path.

Three exact-markup operations record the edits:

- `Operation::ReplaceElement { path, before, after }`;
- `InsertElement { container, child, xml }`;
- `RemoveElement { container, child, xml }`.

Replay checks the exact bytes at the path or index, and each inverse is
exact. They have durable encodings, `element.replace`, `element.insert` and
`element.remove`, with targets `element:<path>` and
`container:<path>/child:<n>`. They take part in the composition effect keys.
They are supported by the source-backed main-document overlay. They retain
exact markup, so a commit holding one publishes without compaction.

**Refusals.**

- A removal is refused when the element holds one end of a range whose other
  end is outside it, a story reference, or a section break; a run is also
  refused when it holds a field character.
- A cell's last paragraph, a row's last cell and a table's last row cannot be
  removed or moved out.
- A row or cell position is refused when a row- or cell-level `w:sdt` or
  `w:customXml` wrapper precedes it, because readers count wrapped rows and
  cells and the raw position would differ.

`NewTable` builds the table's markup:

- the default single 4-eighths-point borders (as Word and OfficeCLI do),
  `tblW auto` and an optional `tblStyle`;
- a `tblGrid` with 9,360 twips divided evenly, or explicit widths;
- one paragraph per cell, holding its text.

Generated markup uses the container's prefix. An unprefixed container gets a
declared `w` prefix.

`LimitKind` gains `TableRows` and `TableColumns`; a new table's dimension is
capped at 4,096.

## Unified surface

**Path forms.** All are 1-based in path text.

- `/body/tbl[N]/tr[R]/tc[C]/p[K]` and `/body/tbl[N]/tr[R]/tc[C]/p[K]/r[M]`
- `/body/sdt[N]/p[K]`
- `/body/p[N]/r[M]`, where M may be a run inside a hyperlink
- `/body/p[N]/hyperlink[M]`
- `/body/tbl[N]`, `/body/tbl[N]/tr[R]` and `/body/tbl[N]/tr[R]/tc[C]`
- nested tables `…/tc[C]/tbl[J]/…`

**Proof.** A node's visible address becomes a leaf `ElementPath`
(`format/docx/element.rs`) only when each step is proven:

- body blocks use `IndexedBlock::transaction_position`;
- blocks inside cells and content controls use `direct_position`, which is
  valid under the body's `raw_positions_proven`;
- rows and cells use their visible ordinals, and the leaf refuses the
  wrappers that would make them differ;
- a run maps to a direct run, or to a direct run of a resolved direct
  hyperlink, only when exactly those runs are the paragraph's addressable
  set;
- a hyperlink requires that the paragraph holds no unresolved
  `w:hyperlink`.

Anything else is `unsupported_capability`.

**Verbs.**

- **set**:
  - Paragraphs and runs accept their full property registries at any proven
    path.
  - A hyperlink accepts `text`, and a cell accepts `text`. Both are new
    write-only registry rows (`Setter::HyperlinkText`, `Setter::CellText`).
  - Direct-body targets keep the existing route.
- **add**:
  - `paragraph` under the body, a cell or a block content control.
  - `run` under any proven paragraph, beside its direct runs. A body
    paragraph that has hyperlink runs now takes this route, so the returned
    `r[N]` is the addressable ordinal. Before this change the direct-run
    ordinal was returned there.
  - `table` (`tbl`) under the body, a cell or a content control. It takes
    the add-only `rows`, `cols` and `data` (`"a,b;c,d"`), plus `style` and
    `colWidths` (twips, comma-separated).
  - `row` (`tr`) under a table, with an optional `cols`.
  - Returned paths are visible ordinals, computed from the container index
    before the insert.
- **remove**:
  - Any of `p`, `r`, `hyperlink`, `tbl`, `tr`, `tc` and `sdt`.
  - Query fan-outs drop descendants of removed ancestors and remove in
    reverse raw order.
- **move**:
  - Tables, rows and nested paragraphs.
  - `Start`, `End`, `Position(n)` (counted after the detach, as for body
    paragraphs), `Before`, `After` and `Into`.
- **swap**: unchanged; direct-body paragraphs and runs only.

**Capabilities.** The DOCX element operations flip:

| Element | Operations |
|---|---|
| `hyperlink`, `cell` | set and remove |
| `table`, `row` | add and remove |
| `sdt` | remove |

Table `cols`, `rows`, `style` and `colWidths` are now add properties, and
`data` is add-only.

The six `nested-or-unproven-*` operation gaps become `wrapped-or-unproven-*`,
and their reasons are rewritten. Two new gaps are listed: `add cell` and
`set table/row` properties.

Both DOCX capability goldens are re-blessed (`LITCHI_BLESS=1`, all-features
and docx-only). The OfficeCLI delta table is regenerated
(`python3 tools/officecli_capability_delta.py --officecli …`, `--check` exit
0).

## OfficeCLI parity

| OfficeCLI | Litchi now | Delta |
|---|---|---|
| `table` add: `rows`, `cols`, `data`, `style`, `colWidths` | same keys | OfficeCLI's `style` vocabulary (`medium1`, …) is mapped to Word style ids; Litchi writes the value as a `tblStyle` id. `width`, `align`, `border.*`, `caption` and `layout` at add are not integrated. |
| `table-row` add: `cols`, `height` | `cols` | `height`, `header`, `cantSplit` and grid before/after at add are absent. |
| `table-cell` add: `text`, `width` | refused (`add cell` gap) | The cell grid is not edited. |
| `table-cell` set: `text` | `text` | Run formatting across all cell runs, `fill`, `valign`, merge and borders are not integrated. |
| `hyperlink` set: `url`, `text` | `text` | Changing `url` needs a relationship write (item 3). |
| paragraph and run set at nested paths | full paragraph and run registries | none for proven paths |

## Measurements

Command:
`taskset -c 5 target/release/examples/docx_element_set_bench`, release
build of `crates/litchi-ops/examples/docx_element_set_bench.rs` at this
record's commit.

| Setting | Value |
|---|---|
| Fixture | 200 body paragraphs (one with a hyperlink) and a 20×5 table; 3,553 bytes |
| Owners | eager, opened once |
| Sampling | 2 warm-up rounds; 9 ABBA rounds of 64 operations per sample, giving 18 samples per leg |
| Machine | AMD EPYC 9R45, pinned to CPU 5 |
| Trials | 3: [1](results/0251a-docx-element-writes/trial-1.json), [2](results/0251a-docx-element-writes/trial-2.json), [3](results/0251a-docx-element-writes/trial-3.json) |

Median unified/leaf ratios:

| Case | Leaf µs | Unified µs | Ratio (trials) |
|---|---:|---:|---|
| cell paragraph `align` | 163 | 202 | 1.238 / 1.244 / 1.238 |
| cell `text` | 161 | 198 | 1.227 / 1.230 / 1.239 |
| hyperlink run `bold` | 155 | 166 | 1.071 / 1.078 / 1.057 |
| table add 3×3 | 147 | 155 | 1.046 / 1.052 / 1.051 |

**Budget.** The single-op budget is 10%.

- Table add and the hyperlink-run set are inside it.
- Writes to paragraphs inside table cells miss it by about 13–14 points.
  Profiling attributes about 36 µs of the 39 µs overhead to read-side path
  resolution of `tr[R]/tc[C]`. Each resolution rebuilds the table's row set
  (`Table::rows` over the whole table) and the cell's block index.
- Two changes in this record already cut the gap:
  - the element-path walk reuses the resolved node's own block entry;
  - it no longer re-materializes the table.

  This took the cell cases from 1.48 to 1.24.
- **Follow-up:** a row/cell seek, or a per-generation table cache, in the
  DOCX read adapter.

## Verification

| Command | Exit |
|---|---|
| `cargo test -p litchi-docx --all-features` | 0 (9 new element-edit unit tests: exact inverses and durable replay both ways) |
| `cargo test -p litchi-ops --all-features` | 0 (7 new `docx_unified_nested` cases on both owners, with read-back and reopen) |
| `cargo clippy -p litchi-docx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo fmt --all --check` | 0 |

Rustdoc, the glossary gate and the facade suite are run at the step's final
commit; see 0251c.

## Not claimed

- No relationship-bearing writes (hyperlink `url`, new hyperlinks, comments,
  pictures, numbering). Those are items 3–6.
- Table and row property set, cell insertion and cell merges are not
  claimed.
- Swap of nested elements is not claimed.
- No change to direct-body paragraph and run hot paths. Their routing is
  unchanged except for the run-add redirect for hyperlink-bearing
  paragraphs.
