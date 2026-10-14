# 0284 — Typed SmartArt CRUD through the unified API

Date: 2026-10-14. Status: implements record 0248 §6 Phase E step 19 (the
SmartArt half; the chart half is records 0270, 0271 and 0283). Branch
`uo2/smartart`, cut from trunk `feat/unified-ops-api` at `80e24f8745`.

PPTX SmartArt graphics can now be read, authored from a text outline,
edited node by node and removed through `File::{get, query, add, set,
remove}` (and therefore `edit`, `batch`, Python and the CLI). DOCX and XLSX
remain gaps (§7).

## 1. Inventory before this change

| Owner | Could author | Could edit | Could read |
|---|---|---|---|
| `litchi-drawingml::diagram` | `SmartArtBuilder` and `generate_smartart_*` (a homegrown layout, used by the DOCX writer for new documents) | nothing | `DiagramDataModel` (a lossy semantic codec: drops `spPr`, prSet formatting and rich text) |
| `litchi-pptx` | `shape::diagram::graphic_frame` (frame XML only, never wired) | nothing | frames counted in the scene (`Shape::Diagram`) |
| `litchi-docx` | writer `smartart` (new documents only) | nothing | `Package::smart_arts` inventory |
| `litchi-xlsx` | nothing | nothing | nothing |
| `litchi-ops` | nothing | remove the frame | `/slide[N]/smartart[M]` frame properties only |
| OfficeCLI | no typed authoring: SmartArt is a read-only `smartart` node and round-trips through `dump` + `add-part`/`raw-set`; `--type diagram` is a mermaid renderer, a different feature | — | `SmartArtToNode` (name, id, geometry, z-order) |

OfficeCLI therefore sets no vocabulary for authoring; the v1 scope below is
Litchi's own, with OfficeCLI's read projection kept as the frame part.

## 2. What landed

### Leaf: `litchi-drawingml`

- `diagram::authoring` (new): `Spec` is a built-in `Layout` (`list`,
  `process`, `cycle`, `hierarchy`, `pyramid`, `matrix`, `venn`, with
  PowerPoint's unique identifiers `layout/default`, `process1`, `cycle2`,
  `hierarchy1`, `pyramid1`, `matrix3`, `venn1`) and a tree of `Node`s, built
  directly or parsed from a text outline (one node per line, one level per
  leading tab or two spaces; OfficeCLI's `\n`/`\t` escapes accepted).
  It writes the data model (content points with parent and sibling
  transitions, presentation points and `presOf`/`presParOf` connections
  matching its layout definition, `dsp:dataModelExt`), Litchi's layout
  definition, quick style and colors, and **Litchi's own cached drawing**:
  a deterministic integer-EMU layout per family (grid, row with arrows,
  ring, layered tree with links, stacked slices, two-column grid,
  overlapping circles), descendants shown as bullet lines where a layout
  draws only the top level. Every drawn shape's `modelId` is the
  presentation point that shows it. `Spec::{from_parts, recognize}` answer a
  spec only when regenerating the data model and the layout definition from
  what they say reproduces them byte for byte.
- `diagram::edit` (new): `Outline::read` projects any data model (both
  namespaces, GUID or numeric ids) onto its content nodes in document order
  with layout, style and color ids and the drawing reference; `set_text`
  replaces one node's text in place, keeping body/list properties and the
  first paragraph's and run's properties; `drawing_text` makes the same
  change in the cached drawing when every shape presenting the node shows
  exactly that node's current text, and answers `None` otherwise;
  `set_drawing_reference` names, renames or removes the drawing reference.
  Edits re-read their output and refuse anything that does not read back.

### Leaf: `litchi-pptx`

- `opened::Edit::add_smartart(slide, &Spec, name, bounds)` stages the five
  diagram parts (`/ppt/diagrams/{data,layout,quickStyle,colors,drawing}N.xml`,
  first free `N`), their slide relationships and the graphic frame in one
  candidate.
- `opened::Edit::{smartart, set_smartart}` read and replace a SmartArt
  frame's parts (`SmartArtParts`); `set_smartart` creates, replaces or
  removes the cached drawing and its slide relationship and rewrites the
  data model's drawing reference to match, so the two can never disagree.
  Identical parts are an exact no-op.
- `remove_shape` now also removes a SmartArt's cached-drawing relationship
  and part with the frame (it was left behind: the drawing is named only by
  the data model, not by the frame).
- `Slide::diagram_xml`, `SourceSlide::diagram_xml`,
  `source::Slide::diagram_xml` read one data part by the frame's `r:dm`.

### Unified layer: `litchi-ops`

| Path | Read | Write |
|---|---|---|
| `/slide[N]/smartart[M]` | frame properties, `layout` (Litchi's token, else the last segment of the layout id, e.g. `process3`), `layoutId`, `nodeCount`; `text` is the outline | `add` under a slide: `layout` (default `list`), `text` (required), `name`, `x`, `y`, `width`, `height`; `set`: `layout`, `text`; `remove` (frame and all parts) |
| `/slide[N]/smartart[M]/node[K]`, nested `/node[K]/node[J]` | `text`, `level` (1-based) | `add` under a smartart or node with `text` (placement start, end or position); `set text`; `remove` (with descendants) |

`query("node")`, children listings and `capabilities()` include the new
element (`node`, a litchi-only element of the PPTX vocabulary). The
command channel no longer refuses line breaks in a SmartArt `add`'s `text`
(an outline, not text controls).

## 3. Keeping the cached drawing consistent (decision)

The brief left a choice: regenerate the drawing (`dsp:`) deterministically,
or remove it so consumers lay the graphic out again. The answer differs by
who wrote the graphic:

1. **Graphics Litchi authored** (recognized exactly by `Spec::recognize`)
   are regenerated whole — data model and drawing, layout definition on a
   layout change — after every edit. Deterministic, and nothing is lost,
   because nothing beyond the spec was there.
2. **Node text edits of any other graphic** are made in place in the data
   model, and in the drawing shapes that present exactly that node. When no
   shape maps unambiguously (no presentation point, one shared with another
   node, a drawing that already disagrees with the data model), the drawing
   is **removed** with its relationship and reference, so consumers lay the
   graphic out from the data model.
3. **Adding or removing nodes of any other graphic is refused**
   (`unsupported_capability`). Measured with LibreOffice: removing the
   drawing of a foreign `process3` graphic is re-laid out correctly, but a
   node added without presentation points is **not drawn** (LibreOffice
   lays out from the data model's presentation points, which only the
   original layout engine can regenerate), so "remove the drawing" would
   publish a graphic that silently hides the new node. The refusal's
   suggestion names the way out: `set layout=...` rebuilds the graphic as a
   Litchi graphic from its outline (node text and hierarchy kept, per-node
   formatting and the original layout and styles not kept), after which
   nodes can be added and removed.

## 4. Evidence

| Check | Result |
|---|---|
| Unit, `litchi-drawingml::diagram::authoring` | every layout: outline round trip, parts read back (also through the legacy `DiagramDataModel`/`DiagramDefinition` readers), determinism, recognition only of exact parts; bounds; empty spec; foreign in-place text with formatting kept; drawing patch and refusal; drawing reference add/rename/remove |
| Unit, `litchi-pptx::opened` | add stages five parts deterministically; read back; no-op; drawing removal and re-creation keep relationship and reference consistent; frame removal removes every diagram part; patch replay and inverse |
| Integration, `crates/litchi-ops/tests/smartart.rs` | add/get/set/remove for source and eager owners after a byte round trip; nested node paths; query; placement; no-op set equal bytes; layout and outline changes; all seven layouts; the foreign fixture `test-data/ooxml/pptx/smartart-accent-process.pptx` (LibreOffice `sd/qa/unit/data/pptx/smartart-accent-process.pptx`, MPL-2.0, already tracked) reads, takes a text edit in place, refuses structural edits with the input unchanged, and is rebuilt by a layout set; refusals (missing text, unknown layout, bad indentation, empty outline, unknown key, out-of-range position) leave the input unchanged; capabilities |
| Python | `test_write_api.PptxVerbTests.test_smartart` (add, node add, set, remove, reopen) |
| Goldens | `tests/golden/{capabilities,registry}/pptx.json` re-blessed (both feature sets): `smartart` becomes addable and settable with `layout`/`layoutId`/`nodeCount`/`text`; new `node` element; no other element changed; delta table regenerated (one new litchi-only element) |
| OfficeCLI corpus | re-recorded: no outcome changed (the corpus has no SmartArt writes) |
| LibreOffice 26.2.6.3 headless | one deck with all seven layouts authored through the CLI, then node text set, node add, node remove and a layout change, plus a text edit of the foreign fixture; converted to PDF, every expected node text is rendered and removed ones are not ([`results/0284/lo-check.json`](results/0284/lo-check.json), script [`results/0284/lo_check.sh`](results/0284/lo_check.sh)). The foreign edit removed the drawing (no unambiguous shape) and LibreOffice re-laid it out with the new text |

## 5. Measurements

Harness `tools/perf-baseline/unified-smartart` (README there): release
build (`CARGO_NET_OFFLINE=true cargo build --release --locked
--manifest-path tools/perf-baseline/unified-smartart/Cargo.toml`, source
commit `5caf561857`, binary SHA-256
`4baa02f145fd410e6cb93882d52e9ff11e73666419f250ef86747133a413ac7f`, not
committed), pinned with `taskset -c 7`, 9 ABBA rounds (18 samples per leg),
eager owners, two runs on a shared machine (load average 11–15)
([`smartart-writes.json`](results/0284/smartart-writes.json),
[`smartart-writes-repeat.json`](results/0284/smartart-writes-repeat.json),
raw samples included). The add legs publish equal decoded members (the
eager owner's writer orders members differently); the node legs publish
identical bytes.

| Route | N | Native median | Unified median | Ratio (run 1 / run 2) |
|---|---:|---:|---:|---|
| `pptx_smartart_add_v1` | 0 | 0.911 / 0.903 ms | 0.920 / 0.915 ms | 1.010 / 1.014 |
| `pptx_smartart_add_v1` | 10 | 2.635 / 2.582 ms | 2.649 / 2.606 ms | 1.005 / 1.009 |
| `pptx_smartart_add_v1` | 50 | 8.872 / 8.816 ms | 8.905 / 8.828 ms | 1.004 / 1.001 |
| `pptx_smartart_node_text_v1` | 1 | 0.836 / 0.835 ms | 0.850 / 0.848 ms | 1.017 / 1.015 |
| `pptx_smartart_node_text_v1` | 10 | 1.158 / 1.156 ms | 1.173 / 1.167 ms | 1.013 / 1.009 |
| `pptx_smartart_node_text_v1` | 50 | 2.540 / 2.533 ms | 2.560 / 2.573 ms | 1.008 / 1.016 |
| `pptx_smartart_node_add_v1` | 1 | 0.859 / 0.861 ms | 0.869 / 0.869 ms | 1.013 / 1.009 |
| `pptx_smartart_node_add_v1` | 10 | 1.176 / 1.183 ms | 1.188 / 1.199 ms | 1.010 / 1.013 |
| `pptx_smartart_node_add_v1` | 50 | 2.567 / 2.567 ms | 2.573 / 2.573 ms | 1.002 / 1.002 |

Every route is within the single-op write budget (≤ 10%); the unified
overhead is 1–2% (path resolution reads the frame's data part once more to
find the node). A first version read the data model twice inside the
unified edit and measured 1.05–1.07 on the node routes; it now reads it
once (`Spec::recognize`). Source-backed owners are not measured: the
unified write promotes a source owner to an eager package first, as chart
writes do, which the native API has no counterpart for. The foreign
in-place text route is not measured (no generated corpus of foreign
graphics at scale).

## 6. OfficeCLI deltas

OfficeCLI has no SmartArt authoring, so every write here is litchi-only.
Its read projection is kept: `smartart` nodes keep `name`, `id`, `x`, `y`,
`width`, `height` and `zorder` as before. The new `node` element and the
`layout`, `layoutId`, `nodeCount` and outline `text` reads are additions.

## 7. Gaps and follow-ups

| Gap | Follow-up |
|---|---|
| DOCX SmartArt: `litchi-docx` has a read inventory (`Package::smart_arts`) and a writer for new documents, but no edit-side inline drawing insertion and no ops element (`/body/.../smartart[N]`) | 0284a: a WORD-vocabulary `smartart` element over `Package::smart_arts`, then inline `w:drawing` authoring next to `insert_picture_at` reusing `authoring::Spec` (coordinate with the DOCX chart insertion of record 0283) |
| XLSX SmartArt: no model in `litchi-xlsx` at all (drawing parts only) | needs a worksheet-drawing diagram anchor in the leaf first |
| Node add/remove on foreign graphics refused (§3) | generate presentation points from a layout definition (a layout engine), or clone a sibling's presentation subtree for list-like layouts and verify each against LibreOffice |
| A layout set on a foreign graphic replaces its layout, style and colors and drops per-node formatting | keep the content points' `spPr`/`t` formatting through `Spec` (a richer node model) |
| Node formatting (fill, font, color) is not read or written; no node `id` identity in paths | add node properties once a consumer needs them |
| PowerPoint rendering was not run; the embedded layout definitions are Litchi's own and simpler than PowerPoint's (PowerPoint may lay out from its built-in definition of the same id) | verify in PowerPoint; consider Litchi-specific layout ids if PowerPoint's re-layout disagrees with the cached drawing |
| SmartArt frame geometry `set` (x/y/width/height) | the graphic-frame geometry write gap shared with charts (record 0270) |
| Master and layout SmartArt, and SmartArt inside groups, are read-only | as for charts |
| The legacy `generate_smartart_*`/`SmartArtBuilder` generators remain (used by the DOCX writer) | move the DOCX writer onto `authoring::Spec` and delete them (ADR 0004) |
| A source-backed `remove` of a SmartArt frame keeps its parts (exact source replay policy) | the dependency sweep on publication shared with charts |

## 8. Not claimed

- Rendering fidelity beyond LibreOffice drawing every node text; PowerPoint
  was not run.
- Any OfficeCLI SmartArt write behavior (OfficeCLI has none).
