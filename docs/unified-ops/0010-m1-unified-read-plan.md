# 0010 — M1 unified read: executable plan

- Milestone: M1 (design §15; goal "EXECUTION PLAN": node view, `get`/`query`,
  `view(outline|text|annotated|stats)`, capability skeleton, paired
  leaf-versus-unified read benchmarks)
- Status: plan. No production code is changed by this record.
- Date: 2026-10-05
- Base: `uo/0010-m1-plan`, branched from `uo/m0a` at `2d8971feac`. M0a steps
  S1–S5e are done there; S6 (`File` enum, family views, `compat`), S6b, S7,
  S8a–d, S9 (deletion of the old facade) and S10 are not. Every `file:line`
  below was read in that tree. Record 0008 is not on this branch; it is cited
  from `feat/unified-ops-api` at `e9c3868109`.
- OfficeCLI was read from the untracked clone at
  `/home/zhuhe/code/litchi/3rdparty/OfficeCLI/` (commit `dced0d74`, v1.0.151,
  2026-09-17).
- `performance_claim: none`. No Litchi code was built, run or measured for
  this record; the surveys behind it ran only read-only Python scans of
  `test-data` (see "What is not claimed"). "(inferred)" marks a statement read
  from code but not executed.

Path prefixes (all relative to the repository root unless they start with
`OC/`):

| Prefix | Directory |
|---|---|
| `O/` | `crates/litchi-ops/src/` |
| `SEL/` | `crates/litchi-core/src/selector/` |
| `DX/` | `crates/litchi-docx/src/` |
| `PX/` | `crates/litchi-pptx/src/` |
| `XX/` | `crates/litchi-xlsx/src/` |
| `XB/` | `crates/litchi-xlsb/src/` |
| `OPC/` | `crates/litchi-opc/src/` |
| `L/` | `crates/litchi/src/` (the facade) |
| `OC/` | `/home/zhuhe/code/litchi/3rdparty/OfficeCLI/src/officecli/` |

Inside `OC/`: `W.Nav`, `W.Q`, `W.Sel`, `W.View` are
`Handlers/Word/WordHandler.{Navigation,Query,Selector,View}.cs`; `P.Q`,
`P.NB`, `P.Path`, `P.Sel`, `P.Res`, `P.View` are
`Handlers/Pptx/PowerPointHandler.{Query,NodeBuilder,Helpers.Path,Selector,Resolve,View}.cs`;
`X.Q`, `X.RW`, `X.Sel`, `X.Sheet`, `X.Node`, `X.Cell`, `X.View` are
`Handlers/Excel/ExcelHandler.{Query,Query.RowWhere,Selector,Helpers.Sheet,Helpers.Node,Helpers.Cell,View}.cs`;
`AF` is `Core/AttributeFilter.cs`, `OF` is `Core/OutputFormatter.cs`, `GXQ`
is `Core/GenericXmlQuery.cs`, `CB.GQ`/`CB.View` are
`CommandBuilder.{GetQuery,View}.cs`.

## 0. Decisions (provisional; open questions in §11.3)

| # | Decision | Section |
|---|---|---|
| D1 | The generic node view is a borrowed projection (`View<'n, 'f>`) over a typed node, serialized by a streaming `serde::Serialize` into OfficeCLI's eight-member `DocumentNode` JSON. `format` is produced from static per-kind `PropSpec` tables with `&'static str` keys and typed values; no per-node map, `DocumentNode` struct or path `String` is built. Python dicts are built from the same accessors. | §2 |
| D2 | Typed nodes are one flat `Node<'f>` enum, one variant per `File` variant, each wrapping the leaf's own ADR 0007 values. Nodes never retain a PPTX `SourceSlide`, and the unified layer keeps no parsed-value cache; every index is a leaf-owned weighted clean-value cache. | §3 |
| D3 | M1 step R8 deletes `compat::{word, slide}` and the document/presentation family views after S9, porting the facade's Markdown to typed nodes and re-backing the Python legacy classes; `WorkbookView` stays as the formula evaluator's seam. | §3.4 |
| D4 | `/body/p[N]` is OfficeCLI's direct-body paragraph (customXml flattened, `oMathPara` wrappers excluded), counted on the markup-compatibility-visible view, for `get`, `query` and the node view. M2 maps it to the transaction's raw direct-body ordinal, stored in the body index, under record 0011's admission proof, and refuses otherwise. | §4.5.1 |
| D5 | PPTX shape paths use OfficeCLI's per-kind counters over direct spTree children; canonical paths use `@id`. | §4.6 |
| D6 | Workbook sheet names are caseless in every format; `/sheet[N]` is the Nth tab; `Sheet1!A1`, `3:3`, `B:B` are accepted; any in-grid coordinate resolves, `Missing` included. | §4.7 |
| D7 | Errors: syntax → `invalid_path`/`invalid_selector`/`max_depth_exceeded`; unknown name → `invalid_path`; concept absent from the format → `unsupported_path`; leaf lacks a model → `unsupported_capability` (listed); nothing there → `path_not_found` (`get`) or `None` (`find`). Suggestions are computed on the error path only. | §4.4 |
| D8 | Cost: the first touch of an axis scans the part that holds it once per payload generation; each step is then O(1) or O(log n). Only parts on the path are decoded. | §4.10 |
| D9 | `get` → `Result<Node>`, `find` → `Result<Option<Node>>`; `IntoPath` accepts text or a parsed `Path`; `query::Plan` is the reusable compiled query; iterators are fallible per item. | §8.1 |
| D10 | Query compiles once and streams one iterative depth-first walk in document order with pruning, incremental paths, `limit`, general `:has`/`:not`/combinators, and OfficeCLI's operator and header-row semantics; unknown keys are warnings; regex values are refused in M1. | §5 |
| D11 | `view` in the text, annotated, outline and stats modes, per format with OfficeCLI's content and JSON shapes, streamed with early stop; `html`, `svg`, `screenshot`, `pdf`, `issues`, `forms` are `unsupported_capability`. | §6 |
| D12 | `capabilities()` is generated from the same `ElementSpec` tables the resolver runs on, guarded by const checks, exhaustive matches, liveness tests and golden snapshots; `PropSpec` is the read-only subset of an M2 registry entry, and `litchi_ops::value` holds the coercions M2 extends. | §7 |
| D13 | Tier 1 gets the full core axes, tier 2 the same paths after its correctness fixes, tier 3 a stated lower-fidelity subset; the ten formats with no `File` variant stay outside M1. | §4.8, §4.9 |
| D14 | R1–R3 and the leaf requests start now from `uo/m0a`; R5 waits for S6; R8 for S9; M1 merges after M0a. | §10 |
| D15 | Every scenario has a paired `unified_`/`leaf_` route with provisional triggers, plus untimed decode-count, allocation and path round-trip gates. | §9 |

## 1. Starting point

### 1.1 What M1 builds on

- **Handles exist; `File` does not.** Every M0a format handle is one `Arc`
  with a typed read surface limited to whole-document text, counts, metadata
  and tolerance reports: `Docx` (`O/format/docx.rs:241`, `:270`, `:310`,
  `:337`), `Pptx` (`O/format/pptx.rs:290-424`: text, slide count, slide text
  and name, catalog, size), `Xlsx` (`O/format/xlsx.rs:307-399`, including the
  lifetime-free `SelectedWorksheet::{cell, cells}` at
  `O/format/xlsx/selection.rs:107-162`), `Xlsb` (`O/format/xlsb.rs:273-335`),
  `Doc` (`O/format/doc.rs:112-130`), `Xls` (`O/format/xls.rs:130-189`), `Ppt`
  (`O/format/ppt.rs:160-190`), and the tier-3 `Rtf`, `Odt`, `Ods`, `Odp`
  handles. The `File` enum, `Locked::unlock` and `Signed::file` arrive in S6
  (`O/file.rs:4-7`). The modules M1 fills are placeholders: `O/node.rs:1-3`,
  `O/path.rs:1-4`, `O/handler.rs:1-4`, `O/capabilities.rs:1-3`,
  `O/registry.rs:1-4`.
- **The grammar is done (M0b).** `Path`/`Step` (`SEL/path.rs:41-44`,
  `:219-222`), `Predicate::{Index, Last, Filter}` (`SEL/predicate.rs:33-43`),
  `Query`/`Complex`/`Compound`/`Pseudo` (`SEL/query.rs:41-44`, `:192-195`,
  `:280-284`, `:400-412`), the story `Axis` table (`SEL/axis.rs:22-90`), and a
  `Copy` `ParseError` whose `code()` is `invalid_path`, `invalid_selector` or
  `max_depth_exceeded` (`SEL/error.rs:317-327`). `ErrorCode::PathNotFound` is
  reserved for resolution (`SEL/error.rs:33-35`). Record 0005 defers to M1:
  element aliases and case folding (delta 6), bare-key meaning such as
  `col[B]` (delta 5), the `col.` prefix (delta 27), `:has`/`:not`/combinator
  semantics (deltas 15–17), the `Sheet1!A1` form as a target (delta 23), and
  numeric/unit comparison (record 0005, "What is not claimed").
- **The unified error exists.** `Code` has the 36 OfficeCLI codes plus
  `unsupported_capability`, `limit_exceeded` and `cancelled`
  (`O/error.rs:56-139`); `Error` is one pointer (`O/error.rs:338-349`). There
  is no `From<ParseError>` yet and no `Source` variant for the grammar
  (`O/error.rs:245-293`).

### 1.2 What the leaves can answer today

The decisive fact for M1 is that the **default** owner of each tier-1
handle, the source-backed one, reaches very little below whole-document text:

| Format | Source-backed owner (default route) | Eager owner |
|---|---|---|
| DOCX | Main story only: `blocks()`, `paragraphs()`, `paragraph(i)`, `tables()`, text (`DX/source_backed.rs:3284-3436`). Headers, footers, notes, comments, styles, numbering, settings and theme are eager-only; the source routes that exist for them (`story_text_snapshot`, section inventory, document variables) refuse markup-compatibility (MCE) markup, which every Word-authored fixture declares (record 0008 §1: all 19 `test-data/ooxml/docx` fixtures; `DX/source_backed.rs:2667-2828`). | Everything above, plus stories and parts, each re-parsed per call. |
| PPTX | Slide position, id, name, text, pictures (`PX/presentation/source.rs:1306-1460`, `:3378-3473`, `:3485-3825`). No shapes, tables, charts, notes, layouts or comments. | `Slide::shapes()` → `Scene` (`PX/shape/reader.rs:176-238`) and the ADR 0004 `Shape` enum (`PX/shape/model.rs:426-440`), tables, chart info, notes, masters, layouts, comments — mostly by decoding every slide or every related part. |
| XLSX | Catalog, `cell(at)`, `cells(area)`, `visit_cells` (`XX/workbook/source.rs:684-1017`). No rows, columns, tables, names, comments or sheet objects; the first cell read of a typical sheet parses the whole worksheet (`XX/workbook/source.rs:1219-1253`). | Rows, columns, merges, extents, tables, names, validation, conditional formats, views, drawings (`XX/workbook/model.rs:1408-1745`). Cell formatting is opaque on both (`XX/style.rs:205-249`). |
| XLSB | Catalog and `materialize()`, which refuses workbooks with pivot tables, slicers or timelines (`XB/workbook/source.rs:1264-1298`). | A `BTreeMap` cell store (`XB/sheet.rs:128`) re-decoded on every `worksheet_by_*` call (`XB/workbook/model.rs:346-356`). |

So M1 is as much a leaf-read programme as a facade programme. §4.11 lists
the leaf requests; each one follows the goal's rule that a missing API is
added in the leaf crate rather than worked around in `litchi-ops`
(`docs/UNIFIED_OPS_GOAL.md:105-107`).

No cross-format read trait or content-event stream exists despite ADR 0007's
wording (`docs/adr/0007-office-object-models.md:9-11`): `litchi-word` and
`litchi-slide` are vocabulary crates with no users, and the only facility all
leaves share is `litchi_core::text_output` (`crates/litchi-core/src/text_output.rs:15-48`).
The `litchi_core::sheet` traits box every worksheet and cell
(`crates/litchi-core/src/sheet/traits.rs:40`, `:56`, `:63`, `:76`, `:102`,
`:108`, `:124`, `:138`, `:145`, `:157`) and allocate a `String` per
`coordinate()` (`:13`); M1 does not use them on any path (goal constraint 2).

### 1.3 What OfficeCLI does

OfficeCLI has no shared walker: each handler resolves paths, builds nodes and
evaluates selectors its own way, and the three disagree on depth, counting,
case and error codes (details in §2, §4, §5). Where M1 must pick one rule,
this record picks the one that keeps cost proportional to the request and
records the OfficeCLI behaviour it departs from as a numbered delta (M1-Δn,
collected in §11.4).

## 2. The generic node view

### 2.1 Contract

The node view is design §5's `{path, type, text, format, children}`,
serialized exactly the way OfficeCLI serializes its `DocumentNode`
(`OC/Core/DocumentNode.cs:14-29`), which carries three more members than the
design lists. M1 emits all eight, with OfficeCLI's names, order and omission
rules (`OF:101-134`):

| Member | JSON | Emitted |
|---|---|---|
| `path` | string | always; canonical form, §2.6 |
| `type` | string | always; OfficeCLI's type string for the element (§4.5–§4.9) |
| `text` | string | when the element has text; omitted (not `null`) otherwise |
| `preview` | string | per the OfficeCLI rule for the kind (§2.4); omitted otherwise |
| `style` | string | Word paragraphs only, as in OfficeCLI (`W.Nav:4004`) |
| `childCount` | int | when known without decoding a part the request did not touch (§2.4; M1-Δ1) |
| `format` | object | always, possibly `{}` |
| `children` | array | always, possibly `[]`; filled to the requested depth (§2.5) |

The view is a projection, never a model:

- it is not stored: a `View` borrows one typed node (§3) and is serialized,
  walked or dropped;
- it builds no per-node map: `format` is produced at serialization time from
  a static per-kind property table (§2.3);
- typed `get`, `find` and `query` never build it (goal, "typed `get` returns
  borrowed/typed nodes without building `format` maps").

The OfficeCLI CLI wraps nodes in `{"success":true,"data":{"matches":N,"results":[...]}}`
(`CB.GQ:97-103`; `OF:25-31`). That envelope belongs to the CLI and Python
command channel (M3); the Rust view is the inner node.

### 2.2 Types

```rust
// litchi_ops::node (re-exported by litchi)
pub struct View<'n, 'f> { node: &'n Node<'f>, options: ViewOptions }

#[derive(Debug, Clone, Copy)]
#[non_exhaustive]                // built with ViewOptions::default().depth(2) etc.
pub struct ViewOptions {
    pub depth: u32,      // default 1; values above MAX_VIEW_DEPTH (64) are clamped
    pub max_nodes: u32,  // default 100_000; exceeding it is `output_too_large`
    pub dense: Option<u32>, // ranges: emit empty-cell stubs up to this many cells
}

impl<'f> View<'_, 'f> {
    pub fn path(&self) -> NodePath<'_>;          // Display; written straight to the sink
    pub fn kind(&self) -> &'static str;          // `type`
    pub fn text(&self) -> Result<Option<Cow<'_, str>>>;
    pub fn preview(&self) -> Result<Option<Cow<'_, str>>>;
    pub fn style(&self) -> Result<Option<Cow<'_, str>>>;
    pub fn child_count(&self) -> Result<Option<u32>>;
    pub fn format(&self) -> Result<Format<'_>>;  // Iterator<Item = (Key, Value<'_>)>
    pub fn children(&self) -> Children<'_, 'f>;  // child views, one level down
}

impl serde::Serialize for View<'_, '_> { /* streams; §2.7 */ }

pub type Key = &'static str;                     // M2: the registry key, same text

#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Value<'a> {
    Bool(bool),
    Int(i64),
    Number(f64),            // finite only
    Text(Cow<'a, str>),
    Token(&'static str),    // closed vocabularies: "single", "center", ...
    Length(Length),         // EMU; rendered as OfficeCLI does (§2.3)
    Color(Color),           // RGB or a theme reference with transforms
}
```

`Format` reads the node's property set once (one parse of `w:rPr`, one
`cNvPr`/`spPr` scan, one cell-style lookup) and then yields entries by table
index; it allocates nothing beyond what the leaf accessor returns.

### 2.3 The format map

- **Keys** are the OfficeCLI property names for the element (examples:
  Word run `bold`, `size`, `font.latin`, `color`, `underline`,
  `superscript` at `W.Nav:2437-2527`; paragraph `align`, `spaceBefore`,
  `keepNext`, `numId`, `outlineLvl`, `paraId` at `W.Nav:4009-4366`; PPT shape
  `name`, `id`, `isTitle`, `x`, `y`, `width`, `height`, `fill`, `phIndex` at
  `P.NB:911-1026`, `:920`; Excel cell `type`, `formula`, `cachedValue`,
  `font.bold`, `font.size`, `numFmtId` at `X.Node:303-739`). Each key a node
  can emit is declared once in its element's `PropSpec` table (§7.2), which
  the query compiler (§5) and `capabilities()` (§7) read too. M1 declares
  only the keys a leaf can answer; M2's registry adds the rest under the same
  names.
- **Values** keep OfficeCLI's JSON types: booleans as booleans, counts and
  indices as numbers, positions and extents as unit-qualified strings chosen
  the way OfficeCLI chooses them (cm, then pt, then in, then EMU, whichever
  round trips exactly; `OC/Core/EmuConverter.cs:181-198`), font sizes in
  points (`"11pt"`, `W.Nav:2485`; `X.Node:563`), colours as `#RRGGBB`
  (`OC/Core/ParseHelpers.cs:241-266`). The key's `PropSpec` value type picks
  the rendering, so the same `Length` prints the way OfficeCLI prints that
  key. Theme colours stay references with
  their transforms, never collapsed to RGB (ADR 0004,
  `docs/adr/0004-semantic-api-design.md:812-813`). This departs from design
  §5's "flat `key → string`" (finding F-1, §11.2): the canonical string form
  of every value is still available as `Value::to_string()`, and it is the
  form `set` will accept in M2.
- **Local, not effective.** The map reports direct (local) properties, which
  is what `set` writes, as ADR 0004 requires layers to stay separate
  (`docs/adr/0004-semantic-api-design.md:807-810`). An explicitly-off flag
  is emitted as `false` and an unset one is absent, as OfficeCLI does for Word
  runs (`W.Nav:2486`). OfficeCLI's read-only `effective.*` keys
  (`OC/Handlers/Word/WordHandler.StyleList.cs:389-470`) are deferred
  (M1-Δ2).
- **No private keys.** OfficeCLI leaks `_`-prefixed round-trip keys and raw
  XML blobs into `format` (`W.Nav:5592`, `:5654`, `:6274-6297`;
  `P.NB:2113`, `:2139`). Litchi emits only declared keys (M1-Δ3).

### 2.4 `text`, `preview`, `style`, `childCount`

M1 follows OfficeCLI's per-kind rules where the leaf can answer them without
touching another part:

| Kind | `text` | `preview` | Other |
|---|---|---|---|
| Word paragraph | full visible text | first 50 characters + `...` (`W.Nav:4002-4005`) | `style` = the paragraph's style **id**; OfficeCLI prints the display name, which needs the styles part (M1-Δ4) |
| PPT shape | full text | text up to 50 characters, else the shape name (`P.NB:907-908`) | — |
| PPT slide | — | title text, else `(untitled)` (`P.Q:1298-1304`) | — |
| Excel cell | stored value as written; formatted display text arrives with L-XLSX-R3 (M1-Δ5) | the A1 reference (`X.Node:360-366`) | missing cell: `text` = `(empty)` (`X.Q:965`) |

`childCount` is the number of children the view would list at depth ≥ 1 for
that node (§4.5–§4.9 define the child axes). OfficeCLI always emits it, with
its own per-handler notion of children (`W.Nav:3911`, `:4006`; `P.Q:129-135`;
`X.Q:51`, `:621`). M1 emits it whenever it is known from the parts the
request already decodes or from catalog metadata, and omits it when counting
would decode a part nothing else in the request needs (for example each
sheet's rows under `get /` of a workbook, or each slide's shapes under
`get /` of a deck). That keeps "decode only the parts the path touches" true
for root views (M1-Δ1).

### 2.5 Depth

One rule for every format: `depth` is the number of child levels below the
node; 0 lists no children; the default is 1, as in OfficeCLI's CLI
(`CB.GQ:14-26`). OfficeCLI applies depth differently per handler (Word's root
ignores it, `W.Nav:118-478`; a PPT slide lists its shapes even at depth 0,
`P.Q:1328`; an Excel `get /` at depth 1 lists every row and cell of every
sheet, `X.Q:53-59`; a range ignores depth and lists every coordinate,
`X.Cell:358-413`). Each of those is M1-Δ6. Ranges list stored cells only;
`ViewOptions::dense` asks for empty-cell stubs up to a cell budget, in line
with ADR 0007's sparse `cells(range)` versus budgeted `grid(range)`
(`docs/adr/0007-office-object-models.md:56-58`).

OfficeCLI clamps depth to 256 silently and has no node or output cap
(`CB.GQ:37-38`; `OC/Core/DocumentLimits.cs:41`). M1 clamps to 64, which
bounds the serializer's recursion, and stops with `output_too_large` after
`max_nodes` nodes (M1-Δ7; ADR 0006 bounded output).

### 2.6 The `path` member

The emitted path is canonical and re-resolves to the same node. M1 uses
OfficeCLI's canonical forms, which prefer stable identity where the element
has one (`W.Nav:1748-1778`; `P.Path:633-639`; `W.Q:444`; `X.Q:234-242`):

| Element | Canonical path |
|---|---|
| Word paragraph with `w14:paraId` | `/body/p[@paraId=1A2B3C4D]` |
| Word comment | `/comments/comment[@commentId=3]` |
| Word block content control with an id | `/body/sdt[@sdtId=7]` |
| Word footnote / endnote | `/footnote[@footnoteId=2]` |
| PPT shape with a `cNvPr` id | `/slide[1]/shape[@id=5]` |
| Workbook sheet addressed as `/sheet[N]` | `/<real sheet name>/...` |
| Anything else | positional, 1-based (`/body/tbl[2]/tr[1]/tc[3]`) |

An identity form is used only when the identity is readable without a
scan (`paraId` needs L-DOCX-R7); until then the positional form is emitted.
`Node::positional_path()` always returns the positional form (a Litchi
extension). Paths are rendered on demand from the node's compact address (a
story or sheet anchor plus `(kind, ordinal or identity)` steps) by building a
borrowed `litchi_core::selector::Path` and using its `Display`, so the text is
the grammar's canonical spelling and parses back to an equal `Path`
(`SEL/path.rs:197-210`; record 0005's round-trip property). A query keeps sibling counters in its walk frames,
so rendering a match's path never rescans its siblings (§5.3).

### 2.7 Serialization and Python

- `View` implements `serde::Serialize` with `serialize_struct` in the member
  order of §2.1; `format` is a streamed `serialize_map`, `children` a
  streamed `serialize_seq` of child views, and `path` goes through
  `collect_str`. Writing JSON is `serde_json::to_writer(sink, &view)`: no
  intermediate `DocumentNode`, `HashMap` or `String` per node. `serde` is
  already a workspace dependency (`Cargo.toml` `[workspace.dependencies]`).
- `View` also exposes the members as accessors (§2.2). The Python bindings
  build `dict`s from those accessors directly, with no JSON round trip (goal,
  "Build result dicts directly from the projection").
- Text output (`OF:786-819`'s one-line-per-node form) is a CLI concern (M3).

## 3. Typed nodes per format

### 3.1 Shape

`get`, `find`, `query` and `children` yield `Node<'f>`, a flat data-bearing
enum with one variant per `File` variant (the flat-per-format rule of record
0003 D1, applied to nodes). Each variant wraps that format's own node enum,
whose variants carry the leaf's ADR 0007 values, never a copy into a second
model:

```rust
#[non_exhaustive]
pub enum Node<'f> {
    #[cfg(feature = "docx")] Docx(format::docx::Node<'f>),
    #[cfg(feature = "xlsx")] Xlsx(format::xlsx::Node<'f>),
    #[cfg(feature = "pptx")] Pptx(format::pptx::Node<'f>),
    #[cfg(feature = "xlsb")] Xlsb(format::xlsb::Node<'f>),
    #[cfg(feature = "doc")]  Doc(format::doc::Node<'f>),
    #[cfg(feature = "xls")]  Xls(format::xls::Node<'f>),
    #[cfg(feature = "ppt")]  Ppt(format::ppt::Node<'f>),
    #[cfg(feature = "rtf")]  Rtf(format::rtf::Node<'f>),
    #[cfg(feature = "odt")]  Odt(format::odt::Node<'f>),
    #[cfg(feature = "ods")]  Ods(format::ods::Node<'f>),
    #[cfg(feature = "odp")]  Odp(format::odp::Node<'f>),
}
```

| Format | Node kinds (M1) | Leaf value carried |
|---|---|---|
| DOCX | root, body, paragraph, run, hyperlink, table, row, cell, sdt, section, header, footer, footnote, endnote, comment, styles, style, numbering, abstractNum, num, level, bookmark, field, picture, equation, unknown | `litchi_docx::Block`/`Paragraph`/`Run`/`Table`/`Row`/`Cell`/`OpaqueBlock` (`DX/document/model.rs:132-141`, `:69-73`; `DX/paragraph/model.rs:453-467`); these hold an `Arc<Vec<u8>>` plus a range (`DX/paragraph/model.rs:295-318`), so they outlive the package borrow; story and part values from L-DOCX-R5/R6 |
| XLSX | root, sheet, cell, range, row, column, table, namedrange, comment, chart, validation, conditionalFormatting, autofilter | `SelectedWorksheet` (`O/format/xlsx/selection.rs:47-55`), `SourceCellView`/`cell::View` (`XX/cell.rs:149-163`), row/column records, `table::Table` |
| PPTX | root, slide, shape (with its OfficeCLI type: title, placeholder, textbox, shape, picture, table, chart, connector, group, ...), paragraph, run, tr, tc, notes, slidemaster, slidelayout, comment | the ADR 0004 `Shape` (`PX/shape/model.rs:426-440`) through an owned scene handle (L-PPTX-R2), never a `SourceSlide` (§3.3) |
| XLSB | root, sheet, cell, range, row, column | the leaf cell record, normalized to zero-based coordinates (§4.7) |
| DOC, XLS, PPT | per §4.8 | the reader values the handles already hold (`O/format/doc.rs:137`, `O/format/ppt.rs:196`) |
| RTF, ODT, ODS, ODP | per §4.9 | the existing handle payloads (`O/format/rtf.rs:89`, `O/format/odt.rs:150`, `O/format/ods.rs:143`, `O/format/odp.rs:159`) |

Every per-format node also stores its compact address (§2.6) and exposes the
leaf value: `Node::as_docx()` → `&format::docx::Node`, whose
`as_block()`/`as_paragraph()`/`as_run()` return the `litchi_docx` types;
`Node::as_xlsx()` → … `as_cell()` → the `cell::View`; and so on. Generic
accessors on `Node` (`kind`, `path`, `text`, `children`, `view`) dispatch by
one `match`, statically.

### 3.2 Reaching a typed node from a `File`

```rust
impl File {
    pub fn root(&self) -> Result<Node<'_>>;                         // "/"
    pub fn get<'p>(&self, path: impl IntoPath<'p>) -> Result<Node<'_>>;
    pub fn find<'p>(&self, path: impl IntoPath<'p>) -> Result<Option<Node<'_>>>;
    pub fn query<'q>(&self, selector: impl IntoQuery<'q>) -> Result<Matches<'_, 'q>>;
}
```

Each method is one `match` on the `File` variant that calls the format's
resolver (§4) with the handle. `File::Signed(s)` delegates to `s.file()`
(S6), because reads are allowed on signed input (record 0003 §2.6).
`File::Locked(l)` returns `document_protected` with the unlock suggestion
(`O/error.rs:392`); `capabilities()` still answers (§7).

### 3.3 Lifetimes and what a node may retain

- `Node<'f>` borrows the `File`. Leaf values that already own an `Arc` range
  (DOCX paragraphs, runs, tables; XLSX `SelectedWorksheet`) are moved into
  the node; values that borrow per-call parsed bytes (PPTX `Scene<'a>` and
  `Shape<'a>`, `PX/shape/reader.rs:170-182`) need the owned scene handle of
  L-PPTX-R2 before they can sit in a node.
- **A node must not retain a `SourceSlide`.** The PPTX editor builder refuses
  ("shared") while any `SourceSlide` or presentation clone is alive
  (`PX/presentation/source.rs:1205-1208`, `:1216-1257`). A PPTX node
  therefore keeps the slide position and an owned scene, so M2's sole-owner
  edit path (record 0009) is not defeated by a live read node.
- **The unified layer adds no parsed-value cache of its own.** Every index or
  memo M1 needs (body-block offsets, a slide's scene, a sheet's cell store)
  is a leaf-owned, weighted clean-value cache keyed by payload allocation, as
  ADR 0005 requires (`docs/adr/0005-io-memory-and-performance.md:14-17`) and
  as the DOCX paragraph memo already does (`DX/parts/document_part.rs:234-375`;
  design `docs/performance/0680-docx-paragraph-memo-design.md`, which rules
  out "a bare package-level `OnceLock<Option<Arc<ParagraphIndex>>>`"). ADR
  0032 does not cover these, because their size grows with the payload
  (`docs/adr/0032-snapshot-derived-value-memos.md`, "Consequences"). A live
  node pins the values it holds; that is ADR 0005's "active handles pin
  them".

### 3.4 Retiring the M0a compatibility surface

M0a step S6 moves the wrapper enums unchanged into
`litchi_ops::compat::{word, slide}` and adds the family views
`File::{document, presentation, workbook}` (record 0003 §2.4); S7 points the
facade's Markdown at `compat` and S8a points the Python bindings at it; S9
then deletes `L/document/`, `L/presentation/` and the old sheet types (record
0003 §4). After S9, `compat` is the only home of the wrapper enums.

M1 retires that surface in one step (R8, §10), after S9 has merged and after
the typed nodes it needs exist:

1. **Markdown** moves from `compat` to the per-format typed nodes. It needs
   ordered blocks; per-paragraph heading level, quote/code block kind and
   resolved list data; per-run bold, italic, strike, direct underline,
   vertical position and inline formulas; images, hyperlinks with targets,
   note references and bodies; cell spans; per-slide number and text
   (`L/markdown/writer.rs:341-1888`, `L/markdown/document.rs:128-226`,
   `L/markdown/docx.rs:76-434`). Those are leaf reads (DOCX `Block`/`Inline`/
   `Run`, DOC paragraph properties, ODT blocks, RTF `Inline`), so Markdown
   matches on `Node` variants and leaf types; it does not go through the
   generic view, whose `format` map is lossy by design.
2. **Python's legacy classes** (`Document.paragraphs()`, `Paragraph.runs()`,
   `Run.bold`/`italic`, `Table.rows()`, `TableRow.cells()`, `TableCell.text`,
   `Presentation.slides()`, `Slide.text`; `crates/pyo3-litchi/src/document.rs:64-286`,
   `presentation.rs:63-118`) are re-backed on typed nodes with unchanged
   semantics. In particular `Document.paragraphs()` keeps today's meaning
   (every outermost paragraph, table cells included, which is the leaf's
   `paragraphs()`), even though `/body/p[N]` counts body paragraphs only
   (§4.5). The two are different APIs and are documented as such.
3. **Facade tests** that use the wrappers (`crates/litchi/tests/pages_facade.rs:52-75`,
   `markdown_docx_semantics.rs:41-50`, `:100-111`) move to typed nodes or to
   `litchi::iwork` for Pages.
4. `compat::{word, slide}`, `DocumentView` and `PresentationView` are then
   deleted, not deprecated (`-D deprecated`, ADR 0004).
   `WorkbookView`'s static `WorkbookTrait` implementation stays: it is the
   seam `litchi_eval::FormulaEvaluator` needs (record 0003 §2.4), not a
   compatibility wrapper. Its boxed-trait internals are a separate follow-up
   for the eval owner.

Retiring `compat` inside M0a instead would make M0a wait for M1's typed
nodes and leaf requests; keeping the interim surface for one milestone is the
cheaper order.

## 4. Path resolution

### 4.1 Engine

```rust
// crate-internal; static dispatch only
pub(crate) trait Navigate<'f> {
    type Kind: Copy + Eq + 'static;      // per-format element kinds
    type Node;                           // the per-format node enum
    const SPEC: &'static FormatSpec<Self::Kind>;   // §7.2
    fn root(&self) -> Result<Self::Node>;
    fn child(&self, parent: &Self::Node, step: &Resolved<'_, Self::Kind>)
        -> Result<Lookup<Self::Node>>;
    fn children(&self, parent: &Self::Node, kinds: KindSet<Self::Kind>)
        -> Result<impl Iterator<Item = Result<Self::Node>> + '_>;
}

pub(crate) enum Lookup<N> {
    Found(N),
    Missing(Available),          // → path_not_found, with what exists
    NotApplicable,               // → unsupported_path
    Gap(&'static GapSpec),       // → unsupported_capability, listed in capabilities()
}
```

`File::get` matches the variant and calls `resolve::<H>(handle, path)`, which:

1. Starts at `root()` for an absolute path, or at the node for
   `Node::get(relative)` (`SEL/path.rs:96-99`, `parse_relative`).
2. For each step, normalizes the name to a `Kind` with the format's static
   alias table (§4.2) — a binary search over `&'static [(&str, Kind)]`, the
   same allocation-free pattern as `Axis::classify` (`SEL/axis.rs:97-111`) —
   and classifies the predicates (§4.3) into a `Resolved` step.
3. Asks `child()` for the next node. Only the leaf calls listed for that
   axis in §4.5–§4.9 run; no other part is decoded.
4. On a miss, builds the error of §4.4 with the prefix that did resolve and
   the failing step's byte offset in the original text. Suggestions (did you
   mean, available children) are computed only on this error path.

Resolution never mutates the document. OfficeCLI's PowerPoint `get` can
clone a layout placeholder into the slide while reading
(`P.Res:302-401`; `P.Q:931-961`); Litchi resolves such a path to the layout's
placeholder and marks the node as inherited (M1-Δ8; ADR 0007 requires edits
of inherited content to name the source or a local override,
`docs/adr/0007-office-object-models.md:93-96`).

### 4.2 Names: aliases and case folding

Element names are matched ASCII case-insensitively through per-format alias
tables. The tables below are the M1 set, harvested from OfficeCLI's path and
query handlers; an alias is valid in paths and selectors alike.

| Format | Input names (any case) | Canonical kind | OfficeCLI source |
|---|---|---|---|
| Word | `paragraph`, `p` | `p` (type `paragraph`) | `OC/Core/PathAliases.cs:15`; `W.Q:2129-2190` |
| Word | `run`, `r` | `r` (type `run`) | `PathAliases.cs:16` |
| Word | `table`, `tbl` / `row`, `tr` / `cell`, `tc` | `tbl` / `tr` / `tc` | `PathAliases.cs:17-19` |
| Word | `hyperlink` | `hyperlink` | `PathAliases.cs:20` |
| Word | `sdt`, `contentcontrol` | `sdt` | `W.Q:2129-2190` |
| Word | `picture`, `image`, `img` | `picture` | `W.Q:2129-2190` |
| Word | `equation`, `math`, `formula`, `oMath`, `oMathPara` | `equation` | `W.Q:2129-2190`; `W.Nav:1482-1501` |
| Word | `toc`, `tableofcontents` | `toc` | `W.Q:461-503` |
| Word | `ole`, `oleobject`, `object`, `embed` | `ole` | `W.Q:113-127` |
| Word | `level`, `lvl` | `level` | `W.Q:306-401` |
| Word | `break`, `br`, `pagebreak` | `break` | `W.Q:2129-2190` |
| PPT | `shape`, `textbox`, `title` (selectors) | the shape family | `PathAliases.cs:23-24`; `P.Q:1658-1680` |
| PPT | `picture`, `pic` / `connector`, `connection` / `hyperlink`, `hlink` | `picture` / `connector` / `hyperlink` | `P.Path:87-120` |
| PPT | `paragraph`, `p` / `run`, `r` | `paragraph` / `run` | `P.Path:87-120` |
| PPT | `tr`, `row` / `tc`, `cell` / `col`, `column`; `cell[R,C]` → `tr[R]/tc[C]` | table parts | `P.Path:87-120` |
| PPT | `moderncomment`, `modern-comment`, `thread`, `threadedcomment` | `moderncomment` | `P.Q:1658-1680` |
| PPT | `video`, `audio`, `media` | media pictures | `P.Q:1658-1680` |
| PPT | `slidemaster`; `slidelayout`, `layout` | axes | `SEL/axis.rs:76`, `:85-86` |
| Excel | `col`, `column` / `comment`, `note` / `table`, `listobject` | `column` / `comment` / `table` | `X.Q:993-999` |
| Excel | `namedrange`, `definedname` / `pivottable`, `pivot` / `validation`, `dataValidation` / `cf`, `conditionalformatting` | as named | `X.Q:152-217`, `:511-526`, `:644-661` |
| Excel | `merge`, `mergedrange` (selectors) | `mergeCell` | `X.Q:1261` |
| Excel | `ole`, `oleobject`, `object`, `embed`; `media`, `image` | `ole`; `picture` | `X.Q:787-797` |
| All | story and part axes (`body`, `header`, `footer`, `footnote(s)`, `endnote(s)`, `comment(s)`, `notes`, `slide`, `chart`, `sheet`, `styles`, `numbering`, `theme`, `settings`, ...) | `Axis` | `SEL/axis.rs:64-90` |

Excel selector property aliases map `bold`, `italic`, `font`, `size`,
`color` to `font.bold`, `font.italic`, `font.name`, `font.size`, `font.color`
(`X.Sel:201-213`); row predicates honour the `col.`/`column.` prefix (forces
a header column) and `@key` (forces a row property) (`X.RW:28-31`, `:47-54`),
which closes record 0005 delta 27.

Case rules M1 adopts:

| Item | M1 rule | OfficeCLI | Delta |
|---|---|---|---|
| Element and alias names | ASCII case-insensitive | case-insensitive in selectors and PPT paths; the Word generic path match is case-sensitive (`W.Nav:1611`) | M1-Δ9 for that Word fallback |
| Sheet names | Unicode caseless for **every** workbook format, with XLSX's key function (`XX/sheet.rs:178-188`; ADR 0004 `docs/adr/0004-semantic-api-design.md:440-447`) | case-insensitive (`X.Sheet:414-423`) | changes XLSB and ODS, which match exactly today (`XB/workbook/model.rs:349-356`; record 0003 F15) |
| Predicate keys | ASCII case-insensitive | same (`AF:584-585`) | — |
| `=`, `!=` text values | case-insensitive; colours compared after normalizing `#` | same (`AF:646-656`) | — |
| `~=` and `:contains` | case-insensitive substring | `~=` case-insensitive (`AF:439-471`); `:contains` case-sensitive on Word body paragraphs and runs (`W.Sel:154`, `:371`) | M1-Δ10 |
| Identity values | per key, as OfficeCLI: `@paraId` and PPT `@name` case-insensitive, Word bookmark `@name` and `/styles/<id>` exact | `W.Nav:1027-1039`, `:1186-1211`; `P.Sel:412-424`; `W.Q:640-648` | — |

At the top level of a workbook, a step **without** predicates is first a
sheet name (caseless); only if no sheet has that name is it tried as a
workbook-level element (`namedrange`, `chart`, ...). A step **with**
predicates is an element (`sheet[2]`, `namedrange[@name=X]`). So a sheet named
`Chart`, `Styles` or `Body` stays reachable as `/Chart/A1`, which
`Step::axis()` alone would misclassify (`SEL/path.rs:261-265`). Unknown bare
names fail with `path_not_found` and the list of sheet names
(`X.Sheet:425-434` does the same).

### 4.3 Predicates

| Predicate (`SEL/predicate.rs:33-43`) | In a path step | In a selector compound (§5) |
|---|---|---|
| `Index(N)` (`[N]`) | the Nth child of that kind under the parent, 1-based in text, zero-based `Position` in Rust. Exceptions copied from OfficeCLI: Excel `row[N]` is row number N (`X.Q:596-642`); Word `footnote[N]`/`endnote[N]` is note id N, while `/footnotes/footnote[N]` is positional (`W.Q:430-459`; `W.Nav:1596-1599`); numbering `level[L]` is `w:ilvl` = L, zero-based (`W.Q:306-401`) | on the subject compound, the Nth match (`OC/Core/SelectorPositionalIndex.cs:43-51`); on an ancestor compound, the Nth sibling of that kind |
| `Last` | the last child of that kind | the last match |
| identity `[@key=value]` (`SEL/predicate.rs:78-87`) | the child whose identity matches; identity keys per kind: `paraId`, `textId` (paragraph), `commentId`, `sdtId`, `footnoteId`/`endnoteId`, `name` (bookmark, shape, named range, table), `id` (shape, abstractNum, num), `type` (header/footer, placeholder), `idx` (placeholder); any other `@key` is `invalid_path` listing the kind's identity keys | an attribute comparison |
| bare key `[B]` (`SEL/predicate.rs:66-73`) | kind-specific: `col[B]` column letter, `placeholder[ctrTitle]` placeholder type, `namedrange[NAME]`, `formfield[name]`; elsewhere `invalid_path` | "key present and non-empty" (`AF:510-579`) |
| other `Filter(expr)` | refused with `invalid_path` and the suggestion to use `query`, as OfficeCLI refuses content filters in `get` (`GXQ:254-261`) | compiled predicate (§5.2) |

Stacked predicates (`placeholder[@type=body][@idx=1]`) must all hold; an
ambiguous result (more than one child) is `invalid_path` naming the
candidates, as OfficeCLI does for placeholders (`P.Path:528-579`).

### 4.4 Error mapping

| Condition | Code | Notes |
|---|---|---|
| The text is not path syntax | `invalid_path` | from `ParseError::code()` (`SEL/error.rs:317-327`); message carries the byte offset, suggestion from `ParseError::suggestion()` (`:329-347`) |
| The text is not selector syntax | `invalid_selector` | same |
| Step or nesting limit (`SEL/limits.rs:24`, `:32`) | `max_depth_exceeded` | same |
| A step name the format does not know (no kind, alias, axis or sheet) | `invalid_path` | suggestion: nearest names by edit distance, computed on the error path only |
| A kind the grammar knows but the format family lacks (`/slide[1]` on a DOCX, `/body` on a workbook, `/Sheet1/A1` on a deck) | `unsupported_path` | suggestion: the family's root forms |
| A kind the format has but Litchi's leaf cannot answer yet | `unsupported_capability` | the `GapSpec` names the follow-up (§4.11); the same entry appears in `capabilities()` |
| A valid kind and predicate with nothing at that position or identity | `path_not_found` from `get`; `Ok(None)` from `find` | suggestion: what exists (`Available at /body: p[1]..p[12], tbl[1]..tbl[3]`), OfficeCLI's message context (`W.Nav:1789-1815`) moved into the structured `suggestion` |
| An unsupported `@key` or a content filter in a path | `invalid_path` | §4.3 |
| The file is `Locked` | `document_protected` | `O/error.rs:392` |
| A leaf failure while decoding (malformed part, limit, changed source) | the leaf's classification | through the existing `From` conversions (`O/error.rs:918-969`); never flattened |

OfficeCLI reports most misses as `not_found`, inferred from message text, and
some malformed paths fall through to `internal_error` (`OF:317-715`;
`P.Q:1263-1270`). M1 uses `path_not_found`, as design §2 requires, and
`invalid_path` for every syntax or vocabulary failure (M1-Δ11). Missing
spreadsheet coordinates inside the grid are not misses: `/Sheet1/Z99` on an
empty cell is a cell node in the `Missing` state, and `row[N]` for a row with
no record is an empty row node, as in OfficeCLI (`X.Q:958-986`, `:596-642`)
and as `cell::View::Missing` models it (`XX/cell.rs:149-163`). Coordinates
outside the grid fail `CellRef::parse` (`SEL/axis.rs:159-181`) and are
`invalid_path`.

`ParseError` gains a `From` conversion into `Error` with a new
`Source::Selector(ParseError)` variant (it is `Copy`, 24 bytes), so the typed
offset survives (`O/error.rs:245-293`).

### 4.5 DOCX

#### 4.5.1 What `/body/p[N]` means, and how M2 addresses the same paragraph

Three paragraph index spaces exist in `litchi-docx` today (record 0008 §6
A4):

| Space | Counts | Where |
|---|---|---|
| (a) read ordinal | every outermost `w:p` anywhere: body, table cells, body-level `sdt` content | `Document::paragraph(i)` and `paragraphs()` (`DX/parts/document_part.rs:1072-1095`; test `:1176-1201`); the source-backed `paragraph_text(i)` too, although its rustdoc says "direct body paragraph" (`DX/source_backed.rs:3296-3321`; finding F-6) |
| (b) transaction ordinal | direct `w:p` children of `w:body`, in the **raw** part | the edit `Snapshot` (`DX/document/transaction.rs:937-951`, scan `:6395-6420`, `paragraph(Position)` `:1389-1426`) |
| (c) OfficeCLI path ordinal | direct `w:p` children of `w:body`, with `w:customXml` wrappers flattened, and paragraphs that only wrap an `m:oMathPara` excluded (those are `/body/oMathPara[N]`) | `OC` `W.Nav:1044-1069`, `:1459-1501`, `:1506` (`Elements<Paragraph>()`) |

**Decision (D4).** `/body/p[N]` means (c), evaluated on the
**markup-compatibility-processed (visible) view** of the main part, for `get`,
`find`, `query`, and for the children of `/body` in the node view. `/body`'s
children are its direct block children in visible order, typed: `paragraph`,
`table`, `sdt`, `equation` (display math), `section` (a body-level `sectPr`),
and `unknown` for anything else, which keeps OfficeCLI's block kinds. Table
cells (`/body/tbl[i]/tr[j]/tc[k]/p[m]`), content controls
(`/body/sdt[i]/p[m]`) and stories each have their own direct-children spaces
built the same way. Space (a) remains the meaning of the legacy Python
`Document.paragraphs()` only (§3.4).

Why the visible view: readers already see the projection
(`DX/source_backed.rs:726-737`; `DX/parts/document_part.rs:927-942`), and
record 0008's invariant I2 says a selector resolved by the reader must land on
the object the editor changes. Where they differ is body-level
`mc:AlternateContent`: OfficeCLI navigates raw markup through the Open XML
SDK and would count the raw children, while Litchi counts the selected branch
(M1-Δ12). Record 0008 found no body-level `AlternateContent` anywhere in the
corpus; its one occurrence is inside a run (record 0008 §2.3, §5).

**How M2 addresses the same paragraph.** The M1 body index (L-DOCX-R1) stores,
for each visible body paragraph, its ordinal in space (b) when it is a direct
`w:body` child, and "none" when it sits inside `w:customXml`. M2's handler
then:

1. resolves the unified path on the visible view to that index entry;
2. requires record 0011's admission proof (stages 1–3 of record 0008: the
   MCE projection drops only directives and ignorable content without
   WordprocessingML, so raw and visible agree on the direct body sequence);
3. maps the entry to `Position(b)` and edits through the transaction;
4. refuses with `unsupported_capability` when the entry has no (b) ordinal
   (inside `customXml`) or the admission fails (body-level
   `AlternateContent`, record 0008 W2), never guessing;
5. in test and debug builds, re-projects the committed candidate and checks
   that the visible paragraph at (c) carries the edit (record 0008 §5's
   projection oracle).

The same rule holds for tables (`tbl[N]` ↔ the snapshot's direct table
ranges) and block content controls.

#### 4.5.2 Axes

Owner columns: S = source-backed (default route), E = eager. Costs are
"(inferred)" from the code; none was measured.

| Path form | Predicates | Leaf API today | Decoded / cost per step | M1 action |
|---|---|---|---|---|
| `/` (type `document`) | — | handle; core properties `Docx::metadata` (`O/format/docx.rs:270-289`) | children `body`, `styles`, `header[N]`, `footer[N]`, `numbering`, non-empty `footnotes`/`endnotes`/`comments`, as OfficeCLI (`W.Nav:118-478`); listing headers needs relationship metadata only | L-DOCX-R5 gives part listings without decoding story payloads |
| `/body` | — | S `Package::document()` (`DX/source_backed.rs:674`); E `Package::document()` (`DX/package/model.rs:573-595`) | decodes the main part (ADR 0030) and **re-runs the MCE projection on every call** (`DX/parts/document_part.rs:498-517`; `DX/source_backed.rs:730`), an owned rewrite whenever the MCE namespace appears (`crates/litchi-ooxml-common/src/mce/codec.rs:751-759`), which is every Word-authored fixture (record 0008 §1) | L-DOCX-R2: memoize the visible payload per part generation |
| `/body/p[N]`, `tbl[N]`, `sdt[N]`, `oMathPara[N]`, `[last()]` | index, last | only `blocks()` (`DX/document/package/model.rs:443-445`): two full passes and one view per child on every call (`DX/parts/document_part.rs:843-884`); `sdt`, `customXml`, `sectPr` come back as `Unknown` (`:689-700`) | one body-index build per payload generation, then O(1) | **L-DOCX-R1** (body index), **L-DOCX-R3** (typed `sdt`) |
| `/body/p[@paraId=X]` (also `@textId`) | identity | `Paragraph::extensions()` builds a self-contained copy, O(offset) per call (`DX/paragraph/model.rs:176-261`, `:717-720`) — O(n²) over a document | O(start tag) per paragraph; optional memo | **L-DOCX-R7** |
| `/body/p[N]/r[M]` | index, last | `Paragraph::runs()` (`DX/paragraph/codec/runs.rs:32-149`): outermost runs flattened through hyperlinks, `ins`/`del`, `smartTag`, inline `sdt`, `fldSimple`, `customXml`; textbox runs excluded (`:74`) | one pass over the paragraph | **L-DOCX-R8**: OfficeCLI's run set also drops comment-reference runs and runs inside inline `sdt`, `fldSimple` and `mc:AlternateContent` (`W.Nav:970-1004`); until then M1-Δ13 |
| `/body/p[N]/hyperlink[M]` | index | `inlines_with_relationships` (`DX/paragraph/package.rs:30-42`); relationships unreachable on S | paragraph pass | L-DOCX-R8 (S relationships) |
| `/body/tbl[N]/tr[M]/tc[K]` | index, last | `Table::rows()`, `Row::cells()` (`DX/table.rs:310-318`, `:484-492`); `Row::clone` drops the cell cache (`:375-383`) | table range only | adapter keeps rows by reference; `grid_span`/`v_merge` read a nested cell's `tcPr` (`DX/table.rs:93-149`; finding F-7) → **L-DOCX-R4** |
| `…/tc[K]/p[J]`, `…/tc[K]/tbl[J]` | index | `Cell::paragraphs()` returns nested-table paragraphs too (`DX/table.rs:718-736`) | cell range | L-DOCX-R1's container scanner gives direct cell children |
| `/header[N]`, `/footer[N]` (+ `@type`) | index (package part order, `W.Q:410-426`), `@type` | E `headers()`/`footers()` re-scan every `sectPr` per call and carry no part name (`DX/document/package/package.rs:88`, `:159`; `DX/header_footer/package.rs:95-202`); S only text through `story_text_snapshot`, which refuses MCE (`DX/source_backed/story_text.rs:1184-1195`, `:2419-2491`) | one story part | **L-DOCX-R5** (story documents). OfficeCLI ignores `@type` and returns header 1 (`W.Q:418`); M1 honours it (M1-Δ14) |
| `/footnote[N]`, `/footnote[@footnoteId=N]`, `/footnotes/footnote[N]`, `/endnote…` | id; ordinal under `footnotes` | E `footnotes()`/`endnotes()` re-run MCE per call (`DX/footnote.rs:262-343`) | one notes part | L-DOCX-R5 |
| `/comments/comment[N]`, `[@commentId=X]` | index, identity | E `comments()` has no paragraphs (`DX/comment.rs:127-224`) | one comments part | L-DOCX-R5 |
| `/styles/<id or name>`, `/numbering/abstractNum[@id=N]/level[L]`, `/section[N]` | name (exact), id, ilvl, index | E `styles()` linear lookup (`DX/document/package/codec.rs:110-125`; `DX/styles.rs:252-306`), `numbering()` (`DX/document/package/package.rs:531-551`); section inventory on both, but S refuses MCE (`DX/source_backed.rs:2667-2828`) | one part each | **L-DOCX-R6**; the S section inventory waits for record 0011 |
| `/bookmark[...]`, `/field[N]`, `/formfield[...]`, `/toc[N]`, `/chart[N]`, `/watermark`, `/revision[...]`, pictures, equations | per OfficeCLI (§4.3) | E-only flat lists without locations (`DX/document/package/package.rs:480-483`; `DX/document/package/codec.rs:135-1648`) | full scans | M1-extended (§10, R4x); gaps listed in `capabilities()` until then |

#### 4.5.3 DOCX format keys in M1

Paragraph: `style` (id), `numId`, `numLevel`, `spaceBefore`/`spaceAfter`/
`lineSpacing` (`DX/paragraph/codec/paragraph_properties.rs:41-264`), `paraId`.
Run: `bold`, `italic`, `underline`, `strike`, `superscript`/`subscript`,
`size`, `font.latin` (`DX/paragraph/codec/run.rs:145-578`). The boolean
readers scan past `w:rPrChange` and nested textbox runs
(`DX/paragraph/codec/run.rs:450-488`, `:598-633`; finding F-8), so they are
used only after L-DOCX-R4's checked `direct_properties()` lands; until then
the keys are listed as gaps rather than emitted wrong. Alignment,
indentation, colour, highlight and keep flags have no reader yet (L-DOCX-R4).

### 4.6 PPTX

**Counting rule (D5).** As in OfficeCLI, each shape kind has its own 1-based
counter over the **direct** children of the slide's `p:spTree` (or of a
group's): `shape` counts `p:sp` (title, placeholder, text box and equation
shapes included), `picture` counts non-media `p:pic`, `table` and `chart`
count the graphic frames of that kind, `connector` counts `p:cxnSp`, `group`
counts `p:grpSp`, `video`/`audio` count media pictures, and `placeholder[K]`
is the Kth `p:sp` with a `p:ph` (`P.Q:1370-1497`; `P.Res:403-428`). The
slide's node-view children are its shapes in spTree (z) order, each with its
kind-specific path, as OfficeCLI's node builder emits them (`P.NB:134-228`).
The leaf `Scene` already yields top-level shapes in spTree order
(`Scene::roots()`, `PX/shape/reader.rs:309`) and classifies them into the ADR
0004 `Shape` variants (`PX/shape/model.rs:426-440`; mapping
`PX/shape/reader.rs:929-946`, `:1561-1589`), so the per-kind counters are a
filter over `roots()`. OfficeCLI gives SmartArt frames a `shape` index in the
node builder but not in `shape[N]` resolution (`P.NB:198-213`;
`P.Q:1372`); Litchi counts them as `smartart` in both (M1-Δ15).

| Path form | Predicates | Leaf API today | Decoded / cost per step | M1 action |
|---|---|---|---|---|
| `/` (type `presentation`) | — | S `slide_count()` and the slide list, no I/O (`PX/presentation/source.rs:1306-1460`); slide size on both (`:1366-1372`) | catalog only; per-slide `childCount` omitted (§2.4) | — |
| `/slide[N]`, `[last()]` | index, last | S `slide(position)` reads nothing (`PX/presentation/source.rs:1454-1460`); `name`/`text` load that slide only (`:3398-3421`). E `Presentation::slide(i)` decodes slide N plus an MCE root check (`PX/presentation/package.rs:267-278`; `PX/parts/slide.rs:1178-1184`), but `Package::presentation()` is rebuilt per call (`PX/package/model.rs:176-182`) and E `slide_count()` decodes every slide to check content types (`PX/presentation/package.rs:33-47`) | S: none until a child is asked for; E: slide N | ops reads the E catalog once through `Presentation::part().slide_references()` (`PX/presentation/model.rs:89`), which parses only `presentation.xml`; **L-PPTX-R5** removes the content-type decodes |
| `/slide[N]/shape[K]`, `picture[K]`, `table[K]`, `chart[K]`, `connector[K]`, `group[K]`, `placeholder[K or type]`, nested under `group[M]` | index, last, `@id`, `@name` (case-insensitive, `P.Sel:412-424`), `@type`/`@idx` (placeholders) | E only: `Slide::shapes()` → `Scene::read` (`PX/slide/model.rs:103`; `PX/shape/reader.rs:190-238`), one full index of the slide per call, borrowed, nothing cached; S: none | the slide part, once per payload generation after L-PPTX-R2 | **L-PPTX-R1** (scene on S), **L-PPTX-R2** (owned, memoized scene with index handles); `@id` lookup scans the records (`Scene::get` handles names only, `PX/shape/reader.rs:334-360`) |
| `…/shape[K]/paragraph[P]/run[R]`, `…/shape[K]/run[R]` | index | only flattened text: `Common::text()` (`PX/shape/model.rs:340-347`; `PX/shape/reader.rs:948-979`) | — | **L-PPTX-R7** (text-body reader); until then `text` only, paragraph/run paths are `unsupported_capability` |
| `…/table[K]/tr[R]/tc[C]`, `col[C]`, `cell[R,C]` | index | `table::Table::from_graphic_frame` over the frame XML (`PX/table/shape.rs:52-150`); `cell(r,c)` rescans rows (`:135-140`); no merge spans | the frame range | adapter bridges `Shape::Table` → `table::Table`; **L-PPTX-R8** adds spans |
| `…/chart[K]` (+ `/series[S]`, `/axis[@role=R]`) | index | `Slide::charts()` in relationship order, decoding every chart part (`PX/chart/package.rs:203-218`); `chart_info()` only (`PX/chart/codec.rs:55`); frames do not record their chart relationship | the chart part | **L-PPTX-R8** (frame → part, `litchi_drawingml::chart::reader::read`, as DOCX does at `DX/chart/codec.rs:243`) |
| `/slide[N]/notes` | — | E `notes::load` reads every slide, the notes master, its theme and every notes slide (`PX/notes/package.rs:290-739`) | everything | **L-PPTX-R9** (per-slide notes) |
| `/slidemaster[N]`, `/slidemaster[N]/slidelayout[M]`, `/slidelayout[M]` (+ `/shape[K]`) | index | E `slide_masters()`/`layouts()` decode every master and every layout of master N (`PX/presentation/package.rs:319-360`; `PX/parts/slide.rs:1566-1592`); S none | masters and layouts are few, so E's cost is acceptable for M1 | **L-PPTX-R4** for S |
| `/slide[N]/comment[M]`, `/slide[N]/moderncomment[K]/reply[R]` | index | E legacy per slide (`PX/parts/slide.rs:1355-1367`); modern comments load every part (`PX/modern_comments/package.rs:61-125`); text only as raw XML | one part (legacy); all parts (modern) | L-PPTX-R9 |
| `/theme`, transitions, animations, sections | — | E only (`PX/shape/theme/package.rs:79-83`; `PX/transition/reader.rs:111-118`) | — | M1-extended |

PPTX format keys in M1: `name`, `id`, `isTitle`, `placeholder` type and
`phIndex` (raw token today, `PX/shape/reader.rs:828-871`), `x`, `y`,
`width`, `height` from local `a:off`/`a:ext` (`:872-916`; rotation, flips and
group coordinate spaces are not captured), and `alt` once L-PPTX-R6 keeps
`descr` (not captured today; `:794-827`). Fill and line have no reader.

A PPTX node never holds a `SourceSlide` (§3.3).

### 4.7 XLSX and XLSB

**Workbook rules (D6).**

- **Sheet names** resolve caselessly in every workbook format (§4.2). For
  XLSX the leaf already does it (`XX/workbook/model.rs:880-905`, binary search
  over `sheet::key`); for XLSB, XLS and ODS the adapter folds the catalog's
  names once per handle and maps to the leaf's position, because their leaves
  compare exactly or ASCII-only (`XB/workbook/model.rs:349-356`;
  `crates/litchi-xls/src/workbook/source.rs:1320-1330`;
  `crates/litchi-ods/src/facade/mod.rs:951-953`).
- **`/sheet[N]`** is the Nth tab in workbook order, all sheet kinds
  included, as XLSX lists them (`XX/workbook/model.rs:873`). The returned path
  uses the real sheet name (`X.Q:234-242`). XLSB lists worksheets only and
  keeps its sheet kinds private (`XB/workbook/package.rs:1320-1364`;
  `XB/workbook/source.rs:54-61`); until L-XLSB-R1, XLSB's `sheet[N]` counts
  worksheets (M1-Δ16).
- **Text forms.** `get("Sheet1!A1")`, `'My Data'!A1`, `Sheet1!3:3` and
  `Sheet1!B:B` are accepted and normalized to `/Sheet1/A1`, `row[3]` and
  `col[B]`, as OfficeCLI does (`X.Sheet:101-184`). This closes record 0005
  deltas 23 and 28 for workbooks; the grammar is unchanged (the `&str`
  `IntoPath` implementation does the rewrite before `Path::parse`).
- **Missing is not absent.** A coordinate inside the grid always resolves:
  `/Sheet1/Z99` with no record is a cell node in `cell::View::Missing`, with
  `text` = `(empty)`, and `row[N]` (N is the row number, not an ordinal) with
  no record is an empty row node (`X.Q:596-642`, `:958-986`;
  `XX/cell.rs:149-163`).
- **Coordinates are zero-based in Rust** everywhere. The XLSB adapter
  normalizes the mixed bases of today's code (`XB/sheet.rs:411-420` zero-based
  eager cells versus the one-based ops source sheet,
  `O/format/xlsb/source.rs:270-319`; pinned by
  `O/format/xlsb/tests.rs:256-279`; `docs/unified-ops/m0a-log.md:67`).

| Path form | Predicates | Leaf API today | Decoded / cost per step | M1 action |
|---|---|---|---|---|
| `/` (type `workbook`) | — | XLSX `Workbook::sheets()`/`SourceBackedWorkbook::sheets()` (`XX/workbook/model.rs:873`; `XX/workbook/source.rs:670`) | catalog only; each sheet's `childCount` omitted (§2.4) | — |
| `/<Sheet>`, `/sheet[N]` | name, index, last | `Xlsx::sheet(selector)` (`O/format/xlsx.rs:383-398`) → `SelectedWorksheet` | catalog only (S defers payloads, `XX/workbook/source.rs:107-115`) | expose kind, visibility, active (delegates; `XX/workbook/source.rs:748-811`) |
| `/<Sheet>/A1` | — | `SelectedWorksheet::cell(at)` (`O/format/xlsx/selection.rs:107-124`) → `Missing`/`Covered`/`Stored` | E: first read parses the whole sheet into a store, then O(log n) (`XX/workbook/model.rs:1767-1786`; `XX/cell.rs:974-994`). S: a sheet with only `dimension`/`sheetData`/`mergeCells` is re-streamed on **every** call; any other sheet is decompressed twice and fully parsed, then O(log n) (`XX/workbook/source.rs:829-846`, `:1219-1253`; `XX/raw/worksheet/selected.rs:1388-1425`) | **L-XLSX-R1**: keep the parsed store as a weighted clean value after the first read, and admit inert children (`sheetViews`, `sheetFormatPr`, `pageMargins`) to the streaming scan |
| `/<Sheet>/A1:C5` (type `range`) | — | E `Worksheet::cells(area)` lazy (`XX/workbook/model.rs:1658-1661`); S `cells(area)` returns a `Vec` after a full stream, `visit_cells` cannot stop early (`XX/workbook/source.rs:935-1042`) | children: stored cells in the rectangle; dense stubs only with `ViewOptions::dense` (§2.5) | **L-XLSX-R6**: `visit_cells` with `ControlFlow` |
| `/<Sheet>/row[N]` | row number, last | E `row(at)` and its cells, O(log R + k) (`XX/workbook/model.rs:1669-1683`); S none | sheet store | **L-XLSX-R2** (S delegates to the store) |
| `/<Sheet>/col[B]`, `col[2]` | bare letter, index | E `column("B")` properties (`XX/workbook/model.rs:1703-1727`); a column's cells cost O(rows × log c) (`XX/cell.rs:1148-1185`) | sheet store | L-XLSX-R2 for S |
| `/<Sheet>/table[N]`, `table[@name=X]`, `/table` | index, name (case-insensitive) | E `Worksheet::tables()` re-parses the table parts each call (`XX/workbook/model.rs:1587`; `XX/workbook/worksheet.rs:223-256`); names via `column_names()` (`XX/table/model.rs:283-285`); no ranges, no lookup by name; S none | the table parts | **L-XLSX-R4** |
| `/namedrange[NAME]`, `[@name=X]` | name (case-insensitive), index | E `defined_names()` (`XX/workbook/model.rs:923`); S parses names at open and drops them (`XX/workbook/source.rs:77-104`) | workbook part (already decoded) | **L-XLSX-R2** keeps them |
| `/<Sheet>/comment[N]` | index | no accessor; package-only `comments::load_from_worksheet` (`XX/workbook/comments/package.rs:21-42`); threaded comments have a codec and no loader | the comments part | **L-XLSX-R5** |
| `/<Sheet>/chart[N]` (+ `series`, `axis`) | index | E `drawing(ordinal)`: chart presence only (`XX/workbook/drawing.rs:572`; `XX/drawing/model.rs:127-137`) | drawing part | **L-XLSX-R7** |
| `validation[N]`, `cf[N]`, `autofilter`, hyperlinks, panes | index | E-only full re-scans (`XX/workbook/model.rs:1408-1582`), with 1,000,000-event caps that also count `sheetData` events (`XX/auto_filter/package.rs:20`, `:238-250`; `XX/data_validation/mod.rs:41-44`), so a large sheet can fail even without the object (inferred) | whole sheet XML | M1-extended; **L-XLSX-R2** adds one memoized scan of the blocks after `sheetData` |

Cell format keys in M1: `type`, `formula`, `cachedValue` (from the stored
`Cell` and `formula::Formula`, `XX/cell.rs:80-202`; `XX/formula.rs:63-145`).
Font, fill, border, alignment and number format are unreadable through any
public API today: `Style` is opaque (`XX/style.rs:205-249`) and the public
`stylesheet::Styles::parse` cannot reach the package's `styles.xml`
(`XX/style/stylesheet/mod.rs:100`) — **L-XLSX-R3**. Until then the cell's
`text` is the stored value, not the formatted display text (M1-Δ5).

**XLSB specifics.** The eager `Workbook` re-decodes a worksheet on every
`worksheet_by_*` call (`XB/workbook/model.rs:346-356`); the source owner's
`materialize()` refuses any workbook with a pivot table, slicer or timeline
(`XB/workbook/source.rs:1264-1298`) and `write_text_to` has no fallback
(`:1071-1103`), so today's ops source-mode cell reads fail on such workbooks
(inferred). M1 serves XLSB cells from the materialized sheet (weighted cache
in the leaf, **L-XLSB-R2**) and, where materialization is refused, from the
leaf's cached-value index (`indexed_values().cached_value`,
`XB/workbook/source.rs:174`, `:1173`) with values only and the formula keys
listed as gaps. XLSB styles exist only on the eager owner
(`XB/workbook/model.rs:285`).

### 4.8 Legacy OLE (tier 2): DOC, XLS, PPT

Tier 2 uses the same paths and kinds as its OOXML sibling; what the binary
format lacks is `unsupported_path`, and what the leaf does not model yet is
`unsupported_capability` with a follow-up.

**DOC.** The reader decodes the FIB, the text of all seven stories, fields,
notes, comments, bookmarks, sections, lists, styles and the piece and
property tables at open (`crates/litchi-doc/src/document/model/state.rs:67-255`),
but paragraphs, runs and tables are rebuilt on every call by scanning each
story from text index 0 (`crates/litchi-doc/src/parts/paragraph_extractor.rs:148-156`).
Three paragraph spaces exist here too: all stories, split at `\r` and U+0007
(`crates/litchi-doc/src/document/model/semantic/layout.rs:8-99`); main-story
paragraphs outside tables (`crates/litchi-doc/src/body_text.rs:4276-4312`);
and every 0x0D of the main story, the L-DOC-1 editor's space
(`crates/litchi-doc/src/body_text/source.rs:2120-2165`). `/body/p[N]` is the
second, the same meaning as OOXML's, and M2 maps it to the third through
L-DOC-R2's character-position spans.

| Path form | Leaf API today | Cost | M1 action |
|---|---|---|---|
| `/body`, `/body/p[N]`, `/body/tbl[N]` | `Document::elements()` is the only ordered API; it covers all stories with no story tag and clones every paragraph (`…/semantic/layout.rs:242-297`) | full re-extraction per call; quadratic in table size per in-table paragraph (`paragraph_extractor.rs:414-477`) | **L-DOC-R1** (per-story reads), **R2** (CP spans), **R3** (story index memo) |
| `…/p[N]/r[M]` | `Paragraph::runs()` (`crates/litchi-doc/src/paragraph/model.rs:122-124`); merged CHPX runs, raw field codes inside (`paragraph_extractor.rs:489-604`) | paragraph | M counts merged runs (M1-Δ17); display text needs **L-DOC-R5** |
| run keys | `Run::{bold, italic, underline, strikethrough, font_size, color, …}` (`crates/litchi-doc/src/paragraph/model.rs:397-480`); font indices only, no font-name table reader | cascaded at open | **L-DOC-R4** (font names) |
| paragraph keys, heading | `ParagraphProperties` incl. `outline_level` (`crates/litchi-doc/src/parts/pap/model.rs:21-181`) | O(1) | — |
| `/body/tbl[N]/tr/tc` | `Table`/`Row`/`Cell` (`crates/litchi-doc/src/table.rs:66-313`); nested tables skipped (`…/model/validation.rs:77-79`); in-table paragraphs without a nesting level never form a table (`crates/litchi-doc/src/parts/pap/model.rs:20`, `:91`; `…/semantic/layout.rs:289-293`) | as above | **L-DOC-R6** (nested), **L-DOC-R7** (level-0 tables; correctness, finding F-9) |
| `/header[N]`, `/footer[N]` | `headers()`/`footers()`, section index lost and empty slots skipped (`…/semantic/subdocuments/headers.rs:21-59`; `crates/litchi-doc/src/parts/headers.rs:29-31`) | per call | **L-DOC-R8** (section slots) |
| `/footnote[N]`, `/endnote[N]`, `/comments/comment[N]` | whole lists per call; one malformed comment fails the call (`…/subdocuments/comments.rs:27-55`) | per call | **L-DOC-R9** |
| metadata | handle at open (`O/format/doc.rs:90-94`); the core projection drops custom properties and edit time (`crates/litchi-ole-common/src/property_set/model.rs:1065-1102`) | — | **L-OLE-R1** |
| unsupported in the format | modern comments, Word 2007+ content controls (`sdt`), custom XML parts, OMML equations, DrawingML charts | — | `unsupported_path` |

**XLS.** The two owner states disagree today: the byte route's eager reader
decodes everything and silently drops any worksheet that fails
(`crates/litchi-xls/src/workbook/package.rs:221-226`), while the path route's
source owner decodes only the globals, refuses BIFF5 and encrypted files, and
on sheets with embedded charts reads phantom cells and stops at the chart's
EOF (`crates/litchi-xls/src/workbook/source.rs:2103-2107`, `:3371-3384`,
`:2787-2797`; the eight-file divergence of `docs/unified-ops/m0a-log.md:93-98`). The causes are the
nested chart `BOF` and an out-of-range `WINDOW2` zoom
(`crates/litchi-xls/src/view/codec.rs:28-38`;
`crates/litchi-xls/src/workbook/codec/semantic/worksheet.rs:520-522`). M1 resolves XLS through the
eager model where it exists and through `materialize_eager()`
(`crates/litchi-xls/src/workbook/source.rs:1058-1157`) on the source owner
when a rich axis is asked for, **after** L-XLS-R1 and R2 make the two readers
agree. Cell lookups on the source owner scan the sheet until a second lookup
builds an index within a 2 MiB budget (`…/workbook/query_cache.rs:158-176`;
`…/workbook/source.rs:60`); **L-XLS-R5** builds it on first use. Formula text,
cell formatting, rows, columns, merges, names, comments, validation and
conditional formats exist on the eager owner only (§4.11). Threaded comments, sparklines, slicers and dynamic arrays do
not exist in BIFF8 and are `unsupported_path`. The `Xls` handle has no reader
accessor today (`O/format/xls.rs:124-219`); R6b adds crate-internal ones.

**PPT.** Slides are cheap to select by position (`slide_count`, the slide
directory; `crates/litchi-ppt/src/presentation/model.rs:249-273`), but each
`slide_at` builds a new factory whose first parse rebuilds a notes index by
re-parsing the whole document container
(`crates/litchi-ppt/src/slide/factory.rs:163-168`;
`crates/litchi-ppt/src/slide/notes.rs:71-80`). Shapes are parsed per `Slide`
value; `shape[K]` counts only the shapes the reader keeps
(`crates/litchi-ppt/src/slide/types/codec.rs:140-158`, `:347`, `:444`);
names and alt text are never decoded; and placeholders whose text is stored
in the slide list come back **empty**, because `OutlineTextRefAtom` is parsed
but never resolved (`crates/litchi-ppt/src/escher_textbox.rs:299`;
`crates/litchi-ppt/src/slide/types/semantic.rs:190`, `:413-421`; finding
F-10). The leaf's slide and shape values are not `Sync`
(`crates/litchi-ppt/src/slide/types/mod.rs:24`;
`crates/litchi-ppt/src/shapes/shape.rs:201`, `:309`), so the PPT adapter
projects each requested shape into an owned node value inside the call; it
caches nothing. Layouts do not exist as objects in the binary format (only a
`geom` value, `crates/litchi-ppt/src/records/record.rs:709`), nor do modern
comments, SmartArt or DrawingML charts: `unsupported_path`. Masters exist but
expose only raw records (`crates/litchi-ppt/src/master/mod.rs:32-60`).

| PPT path form | M1 action |
|---|---|
| `/slide[N]` | **L-PPT-R2** (one notes index) |
| `/slide[N]/shape[K]`, `placeholder[...]`, `picture[K]`, `table[K]/tr/tc` (text only) | **L-PPT-R1** (slide-list text; correctness), **R3** (name, alt), **R4** (stable shape index) |
| `…/paragraph[P]/run[R]` | text boxes only today (`crates/litchi-ppt/src/shapes/textbox.rs:382-388`); **L-PPT-R5** for placeholders and auto shapes |
| `/slide[N]/notes` | `speaker_notes()` (`crates/litchi-ppt/src/slide/types/semantic.rs:337-349`) |
| `/slidemaster[N]` | **L-PPT-R7** (master shapes) |
| `/slide[N]/comment[M]` | legacy comments only (`…/slide/types/semantic.rs:450`) |

### 4.9 Tier 3: RTF, ODT, ODS, ODP (lower fidelity, stated)

The owner override lets tier-3 formats read "through the existing handles"
until M6 (`docs/UNIFIED_OPS_GOAL.md:35-43`). M1 gives them the same paths and
node view at **lower fidelity**, and lists every missing axis in
`capabilities()`. Small handle or leaf fixes are included where they are
one-liners; everything else is M6.

| Format | Supported in M1 | `unsupported_capability` (leaf has no model yet) | `unsupported_path` | Fixes in M1 |
|---|---|---|---|---|
| RTF | `/body/p[N]`, `/body/p[N]/r[M]` with run keys (`bold`, `italic`, `underline`, `strike`, `size`, `font`, colours, `hidden`, `caps`, baseline), paragraph style and outline level, `/body/tbl[N]/tr/tc` (text, merges), `/header[k]`, `/footer[k]`, `/footnote[N]`, `/endnote[N]`, `/comments/comment[N]` (flat text) | runs inside cells, notes, comments and headers; the interleaving of tables with paragraphs (tables carry no body anchor: `crates/litchi-rtf/src/codec/parser/codec/tables.rs:569-581`), so `/body`'s children list paragraphs, then tables (M1-Δ18) | slides, sheets | the handle holds `litchi_rtf::Document` instead of the raw `RtfDocument`, because only `Document::body()` builds the lazy `Story` (`crates/litchi-rtf/src/api/facade.rs:179`, `:260-268`; `crates/litchi-rtf/src/api/story.rs:53`); this changes `paragraph_count` to count `\par` paragraphs (`crates/litchi-rtf/src/model/document/model.rs:1024-1065` vs `api/facade.rs:254`), recorded as a behaviour change |
| ODT | `/body/p[N]` and headings from `elements()` (top-level paragraphs, headings, tables, lists; `crates/litchi-odt/src/elements/parser.rs:35-47`, `:89-218`), paragraph style and resolved paragraph properties, `/body/tbl[N]/tr/tc` (text, value, spans), notes, comments, header/footer regions, images | runs (`r[M]`) and character formatting: spans are dropped by every paragraph reader (`crates/litchi-odt/src/elements/text.rs:737-765`, `:1939`) | slides, sheets | one adapter path over the public parsers for both owners (the source owner lacks about 20 eager reads; `crates/litchi-odt/src/document/source.rs:324-332`) |
| ODS | sheet list, `/<Sheet>/A1` (value, displayed text, raw formula, direct style, merge role, links), `A1:C5`, `row[N]`, `col[B]` (values only), named definitions | row and column properties, hidden or active sheet, resolved cell formatting, rich text, charts and validations on the source owner | bodies, slides | cell lookups go through the owner's indexed `cell()` (`crates/litchi-ods/src/facade/mod.rs:958`; index after 64 queries, `crates/litchi-ods/src/facade/cell_locator.rs:6-7`), not the slice's linear `Sheet::cell` (`crates/litchi-ods/src/worksheet/model.rs:697`); caseless sheet names in the adapter |
| ODP | `/slide[N]` (title, text, index; name, master and layout from `litchi_odp::page::parse`, `crates/litchi-odp/src/model/page_metadata.rs:454`), `/slide[N]/notes` (flat), `/slidemaster[N]` and layouts (`crates/litchi-odp/src/model/page_layout.rs:354`) | `/slide[N]/shape[K]`: the slide parser drops title and text frames (`crates/litchi-odp/src/codec/parser/codec/xml/semantic.rs:301-323`), so indices would not match OfficeCLI's; shape paragraphs, runs and tables; comments on the source owner | bodies, sheets | the handle reports the owner's `metadata()` instead of a default value (`O/format/odp.rs:134-140`; m0a-log S5e) |

Tier-3 text views reuse `litchi_core::text_output`, which every tier-3 leaf
already writes through on at least one owner
(`crates/litchi-core/src/text_output.rs:15-48`; ODT
`crates/litchi-odt/src/document/semantic.rs:71`, ODS
`crates/litchi-ods/src/facade/mod.rs:331`, ODP
`crates/litchi-odp/src/presentation.rs:596`). The seven ODF minor formats and
the three iWork formats have no `File` variant until M6 (record 0003 D3,
record 0007 Q2), so M1 does not read them through `litchi::open`: iWork
stays readable through `litchi::iwork::Document`, unchanged (design §16.4),
and `litchi::open` on an iWork file keeps returning `unsupported_format` with
that suggestion (`O/error.rs:385-387`); the seven minor ODF formats have no
unified read until M6. Design §15's "after M1 agents can read anything" is
therefore not met for those ten formats (finding F-2).

### 4.10 Cost model

For a path of depth *d*:

```
cost(get) = Σ over steps s of [ index(s) + select(s) + value(s) ]
```

- `index(s)`: the first time a payload generation is asked for the axis of
  step *s* (a part's body blocks, a slide's scene, a sheet's cell store), the
  leaf builds that axis's index in one pass over **that part only**; later
  requests, from any clone of the handle, hit the leaf's weighted clean-value
  cache (§3.3). This is the leaf's existing scan, paid once instead of per
  call; it keeps today's error timing, because the build still scans the
  whole part (`DX/parts/document_part.rs:579-582` explains why the DOCX
  scanners do).
- `select(s)`: O(1) by ordinal from the index, O(log n) for a cell or a
  sheet name, O(siblings) for an `@name`/`@id` lookup without an identity
  memo.
- `value(s)`: O(size of the selected element) to build its leaf value (a
  paragraph's run list, a shape's record), never more.

Parts decoded are exactly the parts on the path, plus their relationship
parts. A node view adds the cost of each emitted node's `format` (one
property parse per node) and, for depth ≥ 1, the same per-child costs; the
`childCount` rule of §2.4 keeps it from decoding parts the request does not
touch.

| Format | `/…/x[N]` today (worst step) | After the M1 leaf requests |
|---|---|---|
| DOCX | O(main part) per call: MCE rewrite plus two full block passes | O(main part) once per generation, then O(1) per step |
| PPTX (S) | not answerable below slide text | O(slide part) once per generation, then O(1) |
| XLSX (S) | O(sheet part) per call on minimal sheets; two decompressions plus a full parse, then O(log n), on others | O(sheet part) once per generation, then O(log n) |
| XLSB | O(sheet) per `worksheet_by_*` on the eager owner | O(sheet) once, then O(log n) |
| DOC | O(all stories) per call, quadratic inside tables | O(story) once, then O(1) |
| XLS (S) | full sheet scan per lookup until the second | O(sheet) once, then O(log n) |
| PPT | O(document container) per `slide_at` | O(slide) per call (shapes cannot be cached, §4.8) |

"Proportional to path depth and the scanned axis, never to total document
size" (goal, PERFORMANCE BAR) holds after the first touch of each axis; the
first touch costs one pass over the part that holds the axis, not over the
document.

### 4.11 Gap list: leaf read requests

Every request below is a change in a non-iWork leaf crate, which the goal
prefers over a workaround in `litchi-ops` (`docs/UNIFIED_OPS_GOAL.md:105-107`).
"Core" requests gate an M1 axis listed in §4.5–§4.9; "extended" requests gate
an axis M1 may ship as a listed gap; "correctness" requests fix a wrong read
today and gate the format's adapter. Sizes are estimates (S: one focused
change; M: a few days; L: a week or more). Ids continue the M0a `L-*-1`
series with an `R` for read.

| Id | Crate | Change | Class | Size |
|---|---|---|---|---|
| L-DOCX-R1 | litchi-docx | Container block index: for `w:body` (and, through the same scanner, a cell, a block `sdt`, a header/footer/note/comment root) record each direct child's kind (paragraph, table, sdt, equation, sectPr, altChunk, unknown; `customXml` flattened), range, namespace bindings, and the paragraph's transaction ordinal (§4.5.1). Memoized beside the paragraph index under the same weighted clean-value cache rules (`DX/parts/document_part.rs:234-375`; design `docs/performance/0680-docx-paragraph-memo-design.md`). Accessors on both owners: `body_block_count`, `body_block(i)`, `body_paragraph(i)`, `body_table(i)`, `body_sdt(i)`. | core | M |
| L-DOCX-R2 | litchi-docx | Memoize the MCE-visible main payload per part generation (weighted clean value, same keying as `DX/parts/document_part.rs:166-175`), so `document()` stops rewriting the part on every call (`DX/parts/document_part.rs:498-517`; `DX/source_backed.rs:730`). | core | S–M |
| L-DOCX-R3 | litchi-docx | `Block::Sdt` (or `OpaqueBlock::{namespace, local_name}`) with id, tag and alias, and the content's blocks. | core | S |
| L-DOCX-R4 | litchi-docx | Checked `Run::direct_properties()` and `Paragraph::direct_properties()` (namespace- and depth-checked like `parse_run_underline`; colour, highlight, caps, `rStyle`, all `rFonts` slots, `jc`, `ind`, `outlineLvl`, keep flags), fixing the `rPrChange` and nested-run leak (finding F-8); `Table::properties()`; `Cell::properties()` reading the cell's own `tcPr` (finding F-7). | core | M |
| L-DOCX-R5 | litchi-docx | An owned, MCE-visible story document for headers, footers, footnotes, endnotes and comments on both owners, with part name, relationship id and header/footer type, listed in package part order without decoding payloads; built on R1's scanner. | core | M–L |
| L-DOCX-R6 | litchi-docx | Source-owner access to styles, numbering, settings and theme (related part by relationship type; public owned constructors for `StyleDefinitions`, `DocumentSettings`, `Theme`). | extended | M |
| L-DOCX-R7 | litchi-docx | `Paragraph::{para_id, text_id}` from the captured namespace bindings in O(start tag); optional `paraId → index` memo. | core | S |
| L-DOCX-R8 | litchi-docx | The OfficeCLI run set for `r[N]` (§4.5.2), `OpaqueInline::kind()`, and main-part relationships on the source owner for hyperlinks. | core | S–M |
| L-DOCX-R9 | litchi-docx | Word counts with paragraph separators (finding F-12); `table(i)` through R1 instead of collect-then-nth (`DX/document/package/model.rs:506-509`). | extended | S |
| L-PPTX-R1 | litchi-pptx | Scene access on the source owner (`SourceSlide` reads its part through the bounded cache and builds a `Scene`). | core | S–M |
| L-PPTX-R2 | litchi-pptx | An owned, `Arc`-shared scene (processed payload plus records) memoized per slide payload as a weighted clean value, with index-based shape handles exposing the ADR 0004 `Shape` accessors. | core | M |
| L-PPTX-R3 | litchi-pptx | Source-owner identity (`relationship_id`, `part_uri`, `is_hidden`, already held privately at `PX/presentation/source.rs:105`, `:339-343`) and related-part reads (layout, notes, comments, charts) through the bounded cache. | core | M |
| L-PPTX-R4 | litchi-pptx | Source-owner presentation-level reads: `presentation.xml` (held privately, `PX/presentation/source.rs:115`), masters, layouts, sections. | extended | M |
| L-PPTX-R5 | litchi-pptx | Content-type checks through `part_metadata` instead of decoding every slide (`PX/presentation/package.rs:42`; comments, structure and custom-show loaders likewise). | perf | S |
| L-PPTX-R6 | litchi-pptx | Keep `descr`, `title`, `hidden`, the payload relationship id and the `xfrm` span in each shape record (`PX/shape/model.rs:213-224`; `PX/shape/reader.rs:794-946`). | core | S |
| L-PPTX-R7 | litchi-drawingml, litchi-pptx | A text-body reader (paragraphs, runs, local run and paragraph properties), shared with SpreadsheetDrawing's existing reader (`crates/litchi-spreadsheet-drawing/src/shape/reader.rs:742-770`). | core | M |
| L-PPTX-R8 | litchi-pptx | Table cell spans and a `Shape::Table` → `table::Table` bridge; chart frame → part mapping and `litchi_drawingml::chart::reader::read` for series and axes. | extended | M |
| L-PPTX-R9 | litchi-pptx | Per-slide notes (`Slide::notes()`) and per-slide modern comments with a text accessor. | core | S–M |
| L-PPTX-R10 | litchi-pptx | `PlaceholderKind::from_token`, `SlideLayoutKind::from_token`. | core | S |
| L-XLSX-R1 | litchi-xlsx | Keep a sheet's parsed store as a weighted clean value after the first read on the source owner, and admit inert children to the streaming scan. | core | M |
| L-XLSX-R2 | litchi-xlsx | `SourceWorksheet` delegates for rows, columns, merges, extents and retained defined names; one memoized scan of the blocks after `sheetData` for validation, conditional formats, autofilter, hyperlinks and views (removing the 1M-event exposure). | core | M |
| L-XLSX-R3 | litchi-xlsx | Readable cell formatting: `Style::format()` (font, fill, border, alignment, number-format code) over a cached stylesheet; a style key on `SourceCell`. | core | M |
| L-XLSX-R4 | litchi-xlsx | Tables on the source owner; `Table::{range, header_range, data_range, totals_range}`; case-insensitive lookup by name. | core | S–M |
| L-XLSX-R5 | litchi-xlsx | Comments accessor on both owners; a threaded-comment loader. | extended | M |
| L-XLSX-R6 | litchi-xlsx | `visit_cells` that can stop early (`ControlFlow`), for streaming query. | core | S |
| L-XLSX-R7 | litchi-xlsx | Chart read through the drawing (`chart::read`); drawings on the source owner. | extended | M |
| L-XLSB-R1 | litchi-xlsb | Public `SheetKind`, `kind()` and `state()`; the full tab list. | core | S |
| L-XLSB-R2 | litchi-xlsb | Cache a decoded worksheet (weighted clean value) on the eager owner; `cells_in(rect)` over the `BTreeMap`; one coordinate base. | core | S–M |
| L-DOC-R1–R3 | litchi-doc | Per-story reads; character-position spans on paragraphs, runs and tables; a per-story element index memo. | core | M |
| L-DOC-R4 | litchi-doc | Font-name table (SttbfFfn) and `font_name(ftc)`. | core | S |
| L-DOC-R7 | litchi-doc | In-table paragraphs with nesting level 0 form level-1 tables (finding F-9). | correctness | S |
| L-DOC-R8, R9 | litchi-doc | Header/footer slots with section and kind; notes and comments by index, tolerating one bad comment. | core | S–M |
| L-DOC-R5, R6, R10–R14 | litchi-doc | Display text without field codes; nested tables; text-box paragraphs; drawing shapes on the reader; list labels; alt text and floating pictures; zero-copy story text. | extended | S–M each |
| L-XLS-R1, R2 | litchi-xls | Skip nested `BOF … EOF` substreams in both readers; record unreadable worksheets as defects instead of dropping them; tolerate an out-of-range `WINDOW2` zoom. Together they reconcile the eight-file divergence. | correctness | S–M |
| L-XLS-R3–R5 | litchi-xls | All tabs with kind and visibility on the source owner; one Unicode-caseless name matcher; build the cell index on first lookup. | core | S each |
| L-XLS-R6–R11 | litchi-xls | Source formatting and formula text; number-format codes; ranges and counts; drawings per sheet. | extended | S–M each |
| L-PPT-R1 | litchi-ppt | Resolve `OutlineTextRefAtom` slide-list text into the shapes that reference it (finding F-10). | correctness | M |
| L-PPT-R2–R4 | litchi-ppt | One notes index per presentation; shape name and alt text (properties 0x0380/0x0381); a stable shape index. | core | S each |
| L-PPT-R5–R10 | litchi-ppt | Rich text for placeholders and auto shapes; `Send + Sync` shapes; master shapes; typed slide-to-master link; counts on the path route; image-store cache. | extended | S–M each |
| L-OLE-R1 | litchi-ole-common | Project custom properties and edit time into `Metadata` (`crates/litchi-ole-common/src/property_set/model.rs:1065-1102`). | core | S |

Record 0011 (DOCX MCE admission, stages 1–3 of record 0008) is a
prerequisite for the DOCX source-owner section inventory and for M2's
paragraph mapping (§4.5.1), not for M1's reads.

## 5. Query

### 5.1 Contract

- `File::query(selector)` returns `Matches<'f, 'q>`, a lazy iterator of
  `Result<Node<'f>>` in document order. Nothing is collected: the walk stops
  when the caller stops pulling, after `limit` matches (`Matches::limit(n)`),
  or at the first error, after which the iterator is fused.
- The selector is compiled once, before the first node is visited, into a
  `query::Plan`; a caller can also prepare the plan once and reuse it on any
  file of the same format (§8.1). The walk never looks at selector text again.
- Zero matches is an empty iterator, as in OfficeCLI (`CB.GQ:269-273`).
  `zero_matches` is reserved for M2 verbs that need a target.
- Compile-time diagnostics are warnings, not errors, as OfficeCLI's
  zero-result diagnostics are (`AF:320-417`, `:1042-1055`): an unknown key
  (with a did-you-mean over the subject kinds' `PropSpec` names) and a
  relational comparison with a non-numeric literal. They are available from
  `Plan::diagnostics()` before the first node, and the condition evaluates as
  "absent".

OfficeCLI materializes every match into a full node (a Word paragraph's node
even resolves `effective.*`), sorts PowerPoint results by kind, and has no
limit (`W.Q:3168`; `P.Q:2050-2345`; `OC/Core/IDocumentHandler.cs:83`).
Litchi streams typed nodes and builds a view only when asked (M1-Δ19).

### 5.2 Compilation

1. **Scope.** No scope means the format's default scope: Word walks the body,
   then headers, footers, footnotes, endnotes and comments, in OfficeCLI's
   order (`W.Q:3183-3239`), entering a story part only if a subject kind can
   occur in it; presentations walk slides in order; workbooks walk sheets in
   order. `Scope::Path` (`SEL/query.rs:49-57`) resolves through §4 (a scope
   that does not resolve is `path_not_found`); `Scope::Sheet(name)` selects a
   sheet caselessly and is `unsupported_path` outside workbooks.
2. **Element test.** `Element::Named` goes through the §4.2 alias table and
   yields a `KindSet` (a bit set over the format's `Kind`); `Element::Any`
   (`*`) is the scope's direct child kinds (Word: body paragraphs and tables,
   as `W.Q:1758-1775`); `Element::Implicit` is every kind; a cell or range
   name (`B`, `A1:C5`) is the cell kind plus a column or range filter.
3. **Predicates.** `Index`/`Last` on the subject compound select the Nth or
   last match overall (`OC/Core/SelectorPositionalIndex.cs:43-51`); on an
   ancestor compound they select among that kind's siblings. Excel `row[N]`
   and `col[B]` keep their path meaning (row number, column letter;
   `X.Q:1096-1141`). Each `Condition` key is looked up once in the subject
   kinds' `PropSpec` tables (with the Excel selector aliases of §4.2); the
   literal is coerced once into the key's value type through
   `litchi_ops::value` (§7.3).
4. **Operators** (OfficeCLI semantics, `AF:510-579`, `:666-760`):

   | Operator | Meaning |
   |---|---|
   | `=` | case-insensitive text; colours compared after `#` normalization; lengths compared in EMU (with OfficeCLI's ±500 EMU tolerance when either side has a unit); numbers numerically |
   | `!=` | negation of `=`; an absent key matches |
   | `~=` | case-insensitive substring. A regex literal `r"..."` is refused with `unsupported_capability` in M1: the workspace has no regex engine (Q3) |
   | `>`, `<`, `>=`, `<=` | numeric only, after stripping `pt`, `px`, `cm`, `in`, `em`, `%` or converting through EMU; a non-numeric value never matches |
   | `[key]` | present and non-empty |

5. **Pseudo-classes.** `:contains(t)` is a case-insensitive substring over
   the node's text (M1-Δ10); `:empty` an early-exit text probe; `:no-alt` the
   picture/shape alt-text test (needs L-PPTX-R6 / L-DOCX-R4 data; a gap until
   then); `:has(list)` a nested plan evaluated in the candidate's subtree,
   stopping at the first hit; `:not(list)` a compound list matched against
   the candidate itself. OfficeCLI implements `:has`/`:not` only as literal
   Excel forms (`X.Sel:118-127`); M1 implements the grammar's general forms
   (record 0005 deltas 15–16).
6. **Combinators.** `>` and descendant, both general (record 0005 delta 17).
   OfficeCLI honours one `>` in Word and slide scoping in PowerPoint
   (`W.Sel:15-54`; `P.Sel:108-144`).

### 5.3 Execution

- **One iterative depth-first walk** over the scope, with an explicit stack:
  no recursion, so document nesting cannot grow the call stack (CRUD
  checklist, "Iterative hostile traversal").
- **Pruning.** The walk enters a container only if the static child table
  (`ElementSpec::children`, §7.2) can reach a subject kind; `slide` never
  descends into shapes, a Word `cell` query never enters a run.
- **Paths are incremental.** Each stack frame keeps per-kind sibling
  counters, so a match's canonical path is known when it is yielded
  (identity forms read the identity attribute from the element's start tag).
- **Cheapest test first:** kind, then positional filter, then structural
  attributes, then property predicates (one property parse per candidate),
  then text pseudo-classes. Ancestor compounds are checked against the frames
  on the stack: `>` against the parent frame, a descendant combinator against
  any ancestor frame, with backtracking bounded by depth × compounds.
- **Unions** (`a, b`) are evaluated together on each node, so each node is
  yielded once, in document order; OfficeCLI runs each alternative and
  de-duplicates by path (`AF:977-995`).
- **Workbooks** walk each sheet's cells through L-XLSX-R6's early-stopping
  visitor (row-major), never through the boxed `litchi_core::sheet` traits.

### 5.4 Header-row predicates

`row[Salary>5000 and Region=EMEA]` follows OfficeCLI's algorithm
(`X.Q:1716-1791`; `X.RW:60-327`; `X.Q:2124-2244`), which M1 implements in the
workbook adapters:

1. Each leaf condition is a **column** (default, or a `col.`/`column.`
   prefix) or a **row property** (`height`, `hidden`, `outlineLevel`,
   `collapsed`, `customHeight`, or any `@`-prefixed key). Mixing both, a bare
   numeric key, or a key that is both is `invalid_selector`.
2. The header comes from a **table (ListObject) that owns every referenced
   column**, matching column names case-insensitively, with a column letter
   inside the table as fallback; its data rows exclude the header and totals
   rows. With no such table, a **detected table** is sniffed: an anchor with
   nothing above or to its left, a header of at least two contiguous
   non-empty cells starting with text, at least one data row, ending at the
   first blank row. It is never "the first row of the used range". More than
   one candidate table is `invalid_selector`.
3. Comparisons use the raw stored value when the operator is relational or
   the literal is numeric, otherwise the displayed value; equality is
   case-insensitive; hidden rows match.
4. An unknown column is `not_found`, with the table's column names (all of
   them up to 20, else the five nearest) as the suggestion.

Binding happens once per sheet in scope: table metadata through L-XLSX-R4,
detection through one pass over the sheet's store. Results are row nodes.

### 5.5 Cost model

```
cost(query) = Σ first-touch index builds of the parts in scope
            + Σ over visited nodes (kind test + predicate cost)
            + Σ over candidates passing cheap tests (text pseudo-class cost)
```

The visited set is the pruned scope, cut short at `limit`; memory is
O(stack depth + plan size); no node view or `format` map is built unless the
caller asks for one. A `:has` costs at most its candidate's subtree and stops
at the first hit; header binding is once per sheet.

## 6. `view(mode)`

### 6.1 Contract

```rust
pub mod view {
    #[non_exhaustive]
    pub enum Mode { Text, Annotated, Outline, Stats, Issues, Forms, Html, Svg, Screenshot, Pdf }
    #[derive(Default)] #[non_exhaustive]   // built with Options::default().max_lines(50) etc.
    pub struct Options { pub start: Option<u32>, pub end: Option<u32>, pub max_lines: Option<u32>,
                         pub columns: Option<ColumnSet>, pub range: Option<CellRange> }
}
impl File {
    pub fn view(&self, mode: view::Mode) -> Result<String>;
    pub fn view_to(&self, mode: view::Mode, options: &view::Options, out: &mut impl std::fmt::Write) -> Result<()>;
    pub fn view_json(&self, mode: view::Mode, options: &view::Options) -> Result<view::Json<'_>>; // Serialize
}
```

- `Mode` parses OfficeCLI's names and one-letter aliases (`CB.View:143-693`).
- `Html`, `Svg`, `Screenshot` and `Pdf` return `unsupported_capability`
  (design §10; ADR 0033 §10). `Issues` and `Forms` also return
  `unsupported_capability` in M1: they need validators that do not exist yet
  (M1-Δ20; `issues` is a candidate for M2's `validate`).
- Views stream into the caller's `fmt::Write` through the §5.3 walker, so
  `end` and `max_lines` stop the walk early; `start` skips without building
  skipped output. `range` and `columns` apply to workbooks only; elsewhere
  they are `invalid_argument`, as in OfficeCLI
  (`OC/Core/IDocumentHandler.cs:47-56`).
- JSON shapes follow OfficeCLI: text, outline and stats have structured JSON
  per format (`W.View:897-1172`; `P.View:338-513`; `X.View:322-492`);
  annotated JSON is `{"view":"annotated","content":…}` (`OF:717-724`).

### 6.2 Content per format

| Mode | Word (DOCX, DOC, RTF, ODT) | Presentation (PPTX, PPT, ODP) | Workbook (XLSX, XLSB, XLS, ODS) |
|---|---|---|---|
| text | one line per body element, `[canonical path] text`, `[Table: N rows]`, `[Equation]`; `start`/`end` count body elements (`W.View:330-508`) | `=== /slide[N] ===`, then each shape's text, groups and table cells included; `start`/`end` count slides (`P.View:19-70`) | `=== Sheet: name ===`, then `[/Sheet/row[N]] A1=v⇥B1=v`, empty cells omitted; `start`/`end` count rows per sheet, `max_lines` is global (`X.View:16-117`) |
| annotated | per run: `[path] 「text」 ← style \| format` (`W.View:510-693`) | per slide, a tree of title, text box, table, picture (with "no alt text"), group (`P.View:72-163`) | per cell: `A1: [value] ← =formula \| type` (`X.View:119-181`) |
| outline | counts line, headers and footers, heading tree by outline level (`W.View:752-809`) | `File \| N slides`, then `Slide k: "title" - counts` (`P.View:165-196`) | per sheet: rows × columns, formulas, errors, tables, charts (`X.View:183-240`) |
| stats | paragraphs, words, characters, style and font distributions (`W.View:811-895`) | slides, shapes, text boxes, pictures, charts, words, slides without title, pictures without alt (`P.View:220-336`) | sheets, total, empty, formula and error cells, value-type distribution (`X.View:260-320`) |

### 6.3 What each format can deliver in M1

| Format | text | annotated | outline | stats | Missing pieces (follow-up) |
|---|---|---|---|---|---|
| DOCX | yes | after L-DOCX-R4 (run format) | yes; heading levels from `outlineLvl` (L-DOCX-R4) or the style's outline level (L-DOCX-R6) | yes | list prefixes need list-label rendering on the source owner (L-DOCX-R6); the eager `statistics()` undercounts words (finding F-12), so M1 counts words itself |
| PPTX | yes (S: one slide at a time, `PX/presentation/source.rs:1397-1450`) | after L-PPTX-R1/R2/R6 | titles from title placeholders (L-PPTX-R1/R2) | after L-PPTX-R1/R2 | font usage needs L-PPTX-R7 |
| XLSX | yes, through L-XLSX-R6 | values and formulas; type labels need L-XLSX-R3 | yes; per-sheet dimension from the store | yes | error cells come from `Value::Error` (`XX/cell.rs:190-202`); displayed values need L-XLSX-R3 |
| XLSB | yes | values; formulas where materialized | yes | yes | pivot workbooks fall back to cached values (§4.7) |
| DOC | yes (main story, after L-DOC-R1) | after L-DOC-R1/R2 | headings from `outline_level` | yes | font names (L-DOC-R4) |
| XLS | yes (after L-XLS-R1/R2) | eager owner | yes | yes | source owner via `materialize_eager()` |
| PPT | yes (built from slide and notes text, not `Presentation::text()`, which skips failing slides and returns a fixed sentence when empty: `crates/litchi-ppt/src/presentation/model.rs:1291-1318`) | text boxes only (L-PPT-R5) | slide titles | yes | slide-list text (L-PPT-R1) |
| RTF, ODT, ODS, ODP | text through `litchi_core::text_output` | paragraph and style level only | ODT headings, RTF outline levels, ODS sheets, ODP titles | counts | runs (ODT), shapes (ODP), formatting (ODS); §4.9 |

Text extraction follows each leaf's default projection, "visible accepted
content" (ADR 0007, `docs/adr/0007-office-object-models.md:13-14`); review
projections (deleted text, both MCE branches) are not offered in M1.

## 7. `capabilities()` in M1

### 7.1 What is reported

`File::capabilities(&self) -> Capabilities` is infallible and allocates
nothing: it is a view over `static` tables plus the handle's state. It
reports, for the file's format:

- **state**: `open`, `signed` (reads allowed, edits `document_protected`) or
  `locked` (everything `document_protected` until `unlock`); whether the
  source was retained for editing (`OpenOptions::retain_source`).
- **verbs**: `get`, `query`, and `view` with each mode, as `supported` or
  `unsupported` plus the code it returns (`unsupported_capability` for the
  renderer modes, `issues` and `forms`; `unsupported_capability` for every
  M2 verb in M1).
- **elements**: one entry per element kind the resolver knows for the
  format. Each entry uses OfficeCLI's schema vocabulary
  (`3rdparty/OfficeCLI/schemas/help/_schema.json:6-199`): `element` (the OfficeCLI type
  string), `elementAliases`, `parent`, `paths` (`stable` identity forms and
  `positional` forms), `children` (`element`, `pathSegment`), `operations`
  (`get`, `query` true; `add`, `set`, `remove` false in M1), and `properties`
  (name, value type, and the `get`/`set`/`add` flags; M1 sets `get` only).
  Litchi adds `identity` (the `@key`s §4.3 accepts) and `predicates` (the
  keys a selector may compare).
- **gaps**: every element, axis or property that OfficeCLI's schemas have for
  the format and Litchi does not answer, with the code
  (`unsupported_path` when the format lacks the concept,
  `unsupported_capability` when the leaf lacks a model), and the follow-up
  (a leaf request id from §4.11 or a milestone).

OfficeCLI's per-property flags cover only `add`, `set` and `get`; read-only
means `get` without `add`/`set` (352 such properties); its `help` command
reads the schemas, but `get` and `query` never consult them
(`OC/Help/SchemaHelpLoader.cs:115-343`). Litchi's report is generated from the
tables the resolver itself runs on, so it cannot claim what the code does not
do.

### 7.2 Generation: one table drives resolution, the view and the report

Per format, one `static FORMAT: FormatSpec<Kind>`:

```rust
pub(crate) struct ElementSpec<K: 'static> {
    pub kind: K,                          // FORMAT.elements[k as usize].kind == k
    pub name: &'static str,               // OfficeCLI `type`
    pub aliases: &'static [&'static str], // §4.2, sorted, ASCII case-insensitive
    pub parents: &'static [K],
    pub children: &'static [K],           // also the query walker's pruning table
    pub identities: &'static [&'static str],
    pub paths: &'static [&'static str],   // canonical templates, e.g. "/body/p[N]"
    pub props: &'static [PropSpec],       // M1: readable; M2: registry entries
    pub gap: Option<&'static GapSpec>,    // Some = unsupported_capability, with follow-up
}
pub(crate) struct PropSpec { pub name: &'static str, pub ty: ValueType }
```

The step normalizer (§4.2), the query compiler (§5.2), the node view's
`format` (§2.3) and `capabilities()` all read these entries. Drift is caught
by:

1. `const` assertions that each `ElementSpec` sits at its `Kind`'s index and
   that alias tables are sorted (the `Axis::TABLE` pattern,
   `SEL/axis.rs:64-95`).
2. An exhaustive `match` on `Kind` in each resolver: a kind without a
   resolver arm does not compile.
3. A path-template test: every `paths` template, filled with fixture values,
   resolves on a fixture to a node of that kind.
4. A property-liveness test over the fixture corpus: every key a node emits
   is declared in its `PropSpec` table, and every declared key is emitted by
   at least one fixture.
5. Golden JSON snapshots of `capabilities()` per format under
   `crates/litchi-ops/tests/golden/capabilities/`, regenerated only on
   request (`LITCHI_BLESS=1`).
6. The OfficeCLI delta table (goal, "VERIFICATION GATES") is generated by a
   script from those golden files and OfficeCLI's `schemas/help/**`. The
   schemas are not in git, so the script runs locally; the committed table
   records the SHA-256 of both inputs, and a CI test fails when a golden
   file's hash no longer matches the one the table was generated from.

### 7.3 The M2 seam

M1's `PropSpec` is the read-only subset of an M2 registry entry (name and
value type). M2 replaces each `props` slice with registry entries (adding
constraints, coercion and the typed setter) without changing `ElementSpec`,
the node view, or the shape of the `capabilities()` JSON; it only turns the
`set`/`add` flags on. The value parsing the query compiler needs (numbers,
lengths with units and the line-width exception, colours in OfficeCLI's
forms, booleans, tokens) lives in `litchi_ops::value` from M1 on, allocation
free on the success path (goal, "Allocation discipline"), so M2's coercions
extend it instead of duplicating it.

## 8. Ergonomics

### 8.1 Rust

```rust
// litchi_ops, re-exported by litchi
impl File {
    pub fn root(&self) -> Result<Node<'_>>;
    pub fn get<'p>(&self, path: impl IntoPath<'p>) -> Result<Node<'_>>;           // miss: path_not_found
    pub fn find<'p>(&self, path: impl IntoPath<'p>) -> Result<Option<Node<'_>>>;  // miss: Ok(None)
    pub fn query<'q>(&self, selector: impl IntoQuery<'q>) -> Result<Matches<'_, 'q>>;
    pub fn view(&self, mode: view::Mode) -> Result<String>;
    pub fn view_to(&self, mode: view::Mode, options: &view::Options, out: &mut impl fmt::Write) -> Result<()>;
    pub fn view_json(&self, mode: view::Mode, options: &view::Options) -> Result<view::Json<'_>>;
    pub fn capabilities(&self) -> Capabilities;
}

pub trait IntoPath<'p> { fn into_path(self) -> Result<PathArg<'p>>; }
// &'p str and &'p String parse (borrowing; zero allocations for short paths,
// record 0005); Path<'p> and &'p Path<'p> are used as is, without parsing.

pub trait IntoQuery<'q> { fn into_query(self) -> Result<QueryArg<'q>>; }
// &'q str parses and compiles; &'q Query<'q> compiles; &'q query::Plan<'q> is used as is.

pub mod query {
    pub struct Plan<'q> { /* compiled; Send + Sync; reusable */ }
    impl<'q> Plan<'q> {
        pub fn compile(query: &'q Query<'q>, format: FileFormat) -> Result<Self>;
        pub fn diagnostics(&self) -> &[Diagnostic];
    }
}

pub struct Matches<'f, 'q> { /* Iterator<Item = Result<Node<'f>>> */ }
impl<'f> Matches<'f, '_> {
    pub fn limit(self, n: usize) -> Self;
    pub fn first(self) -> Result<Option<Node<'f>>>;
}

impl<'f> Node<'f> {
    pub fn kind(&self) -> &'static str;
    pub fn format(&self) -> FileFormat;
    pub fn path(&self) -> NodePath<'_>;            // canonical, §2.6
    pub fn positional_path(&self) -> NodePath<'_>;
    pub fn text(&self) -> Result<String>;
    pub fn write_text(&self, out: &mut impl fmt::Write) -> Result<()>;
    pub fn children(&self) -> Children<'f>;        // Iterator<Item = Result<Node<'f>>>
    pub fn get<'p>(&self, relative: impl IntoPath<'p>) -> Result<Node<'f>>;
    pub fn find<'p>(&self, relative: impl IntoPath<'p>) -> Result<Option<Node<'f>>>;
    pub fn view(&self) -> View<'_, 'f>;            // depth 1
    pub fn view_with(&self, options: ViewOptions) -> View<'_, 'f>;
    pub fn as_docx(&self) -> Option<&format::docx::Node<'f>>; // one per format
}
```

Conventions:

- **`get` and `find` (D9).** ADR 0004 asks for `Result` or `Result<Option<_>>`
  selectors where a missing catalog object is `None` and bad syntax is an
  error (`docs/adr/0004-semantic-api-design.md:449-456`). `find` is that form;
  `get` is `find` plus `path_not_found`, the shape agents and the goal's
  example use (`docs/UNIFIED_OPS_GOAL.md:57-60`). They are two contracts, not
  an alias.
- **Parse once, reuse.** A `Path` is the reusable compiled path: resolution
  only adds an allocation-free alias lookup per step. A `query::Plan` binds a
  query to one format's kinds and property tables; using it on another
  format is `invalid_argument`. `Path` is 760 bytes (record 0005, "Performance
  evidence"), so `IntoPath` takes it by reference as well as by value.
- **Iterators are fallible per item**, because a lazily decoded part can fail
  mid-walk; `Matches` and `Children` are fused after the first error.
- **No string-typed hot loops**: typed callers pass `Path`/`Plan` values and
  read `node.as_docx()` etc.; strings are the edge layer.
- Every public item gets rustdoc with a runnable example (goal, "ERGONOMICS
  BAR").

### 8.2 Python (feeds M0a S8a and M3)

M0a S8a re-points the existing `Document`, `Presentation` and `Workbook`
classes at `File` plus `compat` (record 0003 §5.1). M1 step R9 adds the
unified read surface on top, without changing those classes' signatures:

```python
f = litchi_py.open("report.docx")          # File, immutable handle, context manager
f.format                                    # FileFormat
node = f.get("/body/p[3]", depth=1)         # dict: OfficeCLI node shape (§2.1)
nodes = f.query("paragraph[style=Heading1]", limit=None, depth=0)   # list[dict]
f.find("/body/p[99]")                       # None when absent
text = f.view("text", start=None, end=None, max_lines=None)         # str
outline = f.view("outline", json=True)      # dict
caps = f.capabilities()                     # dict
```

- **Dicts are built from `View` accessors** in Rust, with no JSON string in
  between (goal, PERFORMANCE BAR, Python). `format` values keep their JSON
  types (`bool`, `int`, `float`, `str`).
- **The GIL is released** around resolution, query walks and view writing,
  with the result collected as typed nodes or text and converted to Python
  objects after reacquiring it (the goal's ~100 µs rule).
- **Exceptions**: one base `litchi_py.LitchiError` with `code`, `message`,
  `suggestion`, and subclasses `PathError` (`invalid_path`,
  `invalid_selector`, `path_not_found`, `unsupported_path`,
  `max_depth_exceeded`), `UnsupportedError` (`unsupported_capability`,
  `unsupported_format`, `unsupported_type`), `IoError` (`file_not_found`,
  `io_error`), `CorruptError` (`corrupt_file`), and `ProtectedError`
  (`document_protected`); M2 adds `PropertyError`. Today's mapping to
  `IOError`/`ValueError`/`Exception` (`crates/pyo3-litchi/src/common.rs:19-31`)
  is replaced in the same step.
- **Legacy classes** (`Document.paragraphs()` and the rest) are re-backed on
  typed nodes in R8/R9 with unchanged semantics (§3.4).
- **Stubs and tests**: `.pyi` entries for every new item, a stub-versus-module
  test, and the first pytest suite (none exists, `crates/pyo3-litchi/BUILD.md:125-136`):
  open per format, `get`/`find`/`query`/`view` on fixtures, each exception
  class, and the `depth` and `limit` options.
- **Fixes carried in R9**: `__repr__` slices text by bytes and can panic on
  multibyte text (`crates/pyo3-litchi/src/document.rs:143`, `:200`, `:293`;
  `presentation.rs:125`; finding F-13); the crate does not enable the `ppt` and
  `xls` features although its documentation promises them
  (`crates/pyo3-litchi/Cargo.toml:13-23`; finding F-14).
- **M3** adds the write verbs to the same `File` class, the command-dict
  channel (`apply`, `batch`, `dump`), whose results wrap nodes in OfficeCLI's
  `{"success","data":{"matches","results"}}` envelope (`CB.GQ:97-103`), and
  a streaming `iter_query`.

## 9. Performance plan

### 9.1 Routes to add

All new routes are opt-in (outside `Case::DEFAULT`, `tools/perf-baseline/src/lib.rs:1644`)
and named with record 0003's `unified_` infix (record 0003 §5.3) for the
unified leg and a `leaf_` infix for its paired leg. Because "unified" already
appears in evidence strings for today's facade (`tools/perf-baseline/src/lib.rs:26155`,
`:38457`, `:42166`), each new route's help text names `litchi_ops::File`
explicitly. Both legs run in the same binary, on the same corpus, with the
same owner state the unified open picks (`SourceMode::Auto`).

| Scenario | Unified leg | Leaf leg | Corpus |
|---|---|---|---|
| get one paragraph, warm handle | `file.get("/body/p[N]")` + `text()` | L-DOCX-R1 `body_paragraph(N-1)` + `text()` | generated `docx-semantic-{medium,large}` (`tools/perf-baseline/src/lib.rs:11548`); real control `test-data/ooxml/docx/drawing.docx` (227 paragraphs, 12 tables) |
| get one paragraph, cold (open + get) | `open` + `get` | leaf open + `body_paragraph` | same |
| get one table cell | `/body/tbl[i]/tr[j]/tc[k]` | `body_table` → `rows` → `cells` | same |
| get one shape | `/slide[100]/shape[4]` | scene `roots()` filter (L-PPTX-R1/R2) | generated `pptx-source-backed-media` (200 slides × 8 text boxes, `tools/perf-baseline/src/filesystem.rs:1297`); control `test-data/poi/test-data/slideshow/bug62513.pptx` |
| get one cell | `/Sheet1/B2`, `/Sheet2/IV256` | `Workbook::sheet` → `cell` | generated `XlsxShape::DenseWide` (2 sheets × 256 × 256, `tools/perf-baseline/src/lib.rs:431-437`, `:70588-70590`); control `test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx` (75,770 cells) |
| get a range | `/Sheet1/A1:J100` | `cells(area)` | same |
| query, all matches | `paragraph[style=Heading1]`; `shape:contains("…")`; `row[Amount>5000]` | hand-written loop over the leaf with the same predicate | as above; the row query needs a generated shape with a table header (new `XlsxShape` variant) |
| query, first match | the same selectors with `limit(1)` | — | medium and large corpora, to show the cost does not grow with document size |
| view text, large document | `view_to(Text)` into a preallocated sink | leaf `write_text_to` (DOCX `DX/document/package/model.rs:70-81`, `DX/source_backed.rs:828`; PPTX `PX/presentation/source.rs:1397`) or the row visitor (XLSX) | large generated corpora |
| node view JSON, depth 1 | `serde_json::to_writer(sink, &node.view())` | none (no leaf twin; reported with allocation counts) | as above |
| tier 2: get and view text | `/body/p[N]`; `/slide[1]/shape[1]`; `/Sheet1/A1` | leaf reader calls | the five real files the OLE2 facade routes use (`tools/perf-baseline/src/facade_ole2.rs`; `FloatingPictures.doc`, `SampleShow.ppt`, …) plus `test-data/ole/xls/ConditionalFormattingSamples.xls` |

Registering a route updates `EXPECTED_SELECTABLE_CASE_COUNT`
(`tools/perf-baseline/src/lib.rs:62337`), the name/parse/help registry test
(`:62409-62466`), the pinned name lists (`:63371`, `:63397`) and the CRUD
coverage index, which is already one selector behind (`tools/test_crud_coverage_index.py:28`
pins 533; `lib.rs` has 534; finding F-11) and is fixed first.

### 9.2 Budgets (review triggers, provisional)

The goal fixes the read-path properties, not percentages, except for open
(`docs/UNIFIED_OPS_GOAL.md:122-131`). M1 adopts these triggers and tightens
them once the noise of each route is measured, as the goal requires:

| Scenario | Trigger |
|---|---|
| get, warm | unified ≤ leaf + max(5%, 1 µs) |
| get, cold (open included) | unified ≤ leaf + max(5%, 20 µs), the open budget |
| query, all matches | unified ≤ leaf loop + 10% |
| query, first match | p50 on the large corpus ≤ 2× p50 on the medium corpus (match at the same position) |
| view text | unified ≤ leaf text + 5% |
| node view | no timing trigger; allocations per emitted node ≤ the owned values the leaf returns for it (no map, no per-node `String` for paths) |
| peak RSS (cold get) | unified ≤ leaf + 5% |

Untimed gates, run as tests on every change:

- **Decode counts.** Each route's test asserts the parts it decodes, through
  `OpcPackage::deferred_decode_counters` (`OPC/package.rs:2027`) and
  `SourceBackedPackage::cache_diagnostics` (`OPC/source_backed.rs:7503`): for
  example, `get("/slide[3]/shape[1]")` decodes slide 3 and nothing else.
- **Allocation counts** for short-path parse plus resolve on a warm handle,
  taken with the harness's isolated allocator binaries
  (`tools/perf-baseline/src/allocation_metrics.rs:1-19`), so `litchi-ops`
  gains no `unsafe` test allocator.
- **Path round trip**: over the corpus, `get(node.path())` returns the same
  node for every node a depth-2 walk visits.

### 9.3 Method

The program's rules apply (`docs/GOAL.md:678-701`): before and after on the
same build and machine, ABBA where noise warrants, machine-readable results,
every individual result reported. M1 reuses the S1/S10 capture and analysis
scripts (`docs/unified-ops/results/0003/before/capture.py`, `analyze.py`):
alternating children, median of per-child p50s, trigger max(5%, route noise),
a 0.5 µs floor under 5 µs (`docs/unified-ops/m0a-log.md:34`). Results go under
`docs/unified-ops/results/<record>/`, each file under 1 MiB; raw captures stay
outside the repository with their name, size, SHA-256 and regeneration
command (AGENTS.md). A regression above 5% in an existing route caused by M1
blocks the merge until explained or fixed, as for M0a.

## 10. Step breakdown

### 10.1 Branching and ordering

- Steps R1–R3 and the leaf requests touch only new `litchi-ops` modules
  (`node`, `path`, `query`, `value`) or leaf crates, so they start now, each
  on its own branch from `uo/m0a` at `2d8971feac`.
- The format adapters (R4, R6, R7) need R1–R3 and their leaf requests, not
  the `File` enum: they resolve against the M0a handles that already exist.
- R5 (`File` dispatch and `capabilities()`) needs M0a S6.
- R8 (retiring `compat` and the family views) needs S9, so it runs after the
  M0a merge.
- An `uo/m1` integration branch is cut from `uo/m0a` once S6 is on it; M1
  merges into `feat/unified-ops-api` after M0a's single merge (record 0007
  Q5). Every step leaves the workspace buildable and green (ADR 0008):
  adapters land behind no public entry point until R5 exposes them.

### 10.2 Steps

Owner size: 1 = one implementer for one change; 2 = twice that, or two
implementers. Tier as in the goal's priority.

| Step | Content | Depends on | Parallel with | Size | Tier |
|---|---|---|---|---|---|
| **R1** | `litchi_ops::node`: `Node` skeleton (cfg-gated variants), `NodePath` (address + `Display` through `Path`), `View`, `ViewOptions`, `Value`, `PropSpec`/`ElementSpec`/`FormatSpec`/`GapSpec`, streaming `Serialize`, depth clamp and node budget (`output_too_large`). Unit tests against a test-only fake format. | — | R2, R3, L-* | 1 | — |
| **R2** | Resolution core: `IntoPath` (incl. `Sheet1!A1`, `B:B`, `3:3` normalization), `From<ParseError>` with `Source::Selector`, the `Navigate` trait, step normalization over static alias tables, predicate classification (§4.3), error construction with suggestions (§4.4). | — | R1, R3, L-* | 1 | — |
| **R3** | Query core and `litchi_ops::value`: `IntoQuery`, `query::Plan` compilation, operators and unit/colour/number coercions (allocation-free), the iterative walker with pruning and incremental paths, `:contains/:empty/:no-alt/:has/:not`, combinators, unions, `limit`, diagnostics. Fuzz target over plan compilation, seeded from record 0005's selector corpus. | R2 interfaces | R1, L-* | 2 | — |
| **L1** (wave 1) | Tier-1 core leaf requests (§4.11): L-DOCX-R1, R2, R3, R7 (one owner); L-DOCX-R4, R8; L-DOCX-R5; L-PPTX-R1, R2, R3, R10; L-PPTX-R6, R7; L-PPTX-R9; L-XLSX-R1, R2, R6; L-XLSX-R3, R4; L-XLSB-R1, R2. Each with its own tests in the leaf. | — | R1–R3 | about 9 owners, 1–2 each | 1 |
| **R4a** | DOCX adapter: kinds, alias table, `Navigate`, `PropSpec`s, view modes; decode-count and round-trip tests; §4.5.1's transaction ordinal stored for M2. | R1–R3; L-DOCX core | R4b–d, R6, R7 | 2 | 1 |
| **R4b** | XLSX adapter, including header-row binding (§5.4) and the workbook text forms. | R1–R3; L-XLSX core | R4a, c, d | 2 | 1 |
| **R4c** | PPTX adapter (per-kind counters, scene handles, no retained `SourceSlide`). | R1–R3; L-PPTX core | R4a, b, d | 2 | 1 |
| **R4d** | XLSB adapter (shared sheet code from R4b, coordinate normalization, cached-value fallback). | R4b; L-XLSB core | R4a, c | 1 | 1 |
| **R5** | `File::{root, get, find, query, view, view_to, view_json, capabilities}`; `Signed` delegation, `Locked` refusal; capability generation, golden snapshots, the generated OfficeCLI delta script. | M0a S6; R1–R3; at least R4a | R4*, R6, R7 | 1 | — |
| **L2** (wave 2) | Tier-2 leaf requests, correctness first: L-XLS-R1, R2; L-PPT-R1; L-DOC-R7; then L-DOC-R1–R4, R8, R9; L-XLS-R3–R5; L-PPT-R2–R4; L-OLE-R1. | — | wave 1, R4* | about 5 owners, 1–2 each | 2 |
| **R6a / R6b / R6c** | DOC, XLS, PPT adapters (crate-internal reader accessors on `Xls`). | R1–R3; L2 for the format | each other, R7 | 1 each | 2 |
| **R7** | Tier-3 adapters at the §4.9 fidelity: RTF handle holds `litchi_rtf::Document`; ODT, ODS, ODP adapters; ODP metadata fix. | R1–R3 | R6 | 2 | 3 |
| **R8** | Retire `compat::{word, slide}`, `DocumentView` and `PresentationView`: port the facade's Markdown to typed nodes, re-back the Python legacy classes, move the facade tests (§3.4). | M0a merged (S9); R4a, R4c, R6a, R6c, R7 | R9, R10 | 2 | — |
| **R9** | Python read surface (§8.2): `litchi_py.open` → `File`, `get`/`find`/`query`/`view`/`capabilities`, exception classes, stubs, the first pytest suite, the two Python fixes. | M0a S8a (merged); R5 | R8, R10 | 1 | — |
| **R10** | Perf routes and the paired measurement record (§9). | R4a–d, R5; R6 for tier-2 routes | R8, R9 | 1 | — |
| **R11** | OfficeCLI corpus test: each of the 950 harvested paths and 132 selectors (record 0005) either resolves on a fixture or fails with a listed gap or a recorded delta, never `internal_error` or a panic; the generated delta table; the M1 closure record. | all | — | 1 | — |
| **R4x** (extended) | Extended axes as their leaf requests land: DOCX styles, numbering, sections, fields, bookmarks, revisions, charts; PPTX masters, layouts, comments, charts, themes; XLSX validation, conditional formats, autofilter, comments, charts. Each axis is a listed gap until its step lands. | its L-* request | anything | 1 each | 1–2 |

Critical path: L-DOCX-R1/R2 → R4a → R5 (with S6) → R8 (with S9) → R11. Up to
about twelve implementers can work at once in the first wave (R1, R2, R3 and
about nine leaf owners).

### 10.3 M1 exit criteria

1. Every `File` variant answers `get`, `find`, `query` and
   `view(text|annotated|outline|stats)` for the axes of §4.5–§4.9; every other
   axis is a listed gap or an `unsupported_path` the format genuinely lacks.
2. `compat` and the document/presentation family views are deleted; the
   facade's Markdown and the Python legacy classes run on typed nodes with
   unchanged output on the existing tests.
3. The verification gates of the goal pass, plus the decode-count,
   allocation and round-trip gates of §9.2.
4. The paired routes of §9.1 are recorded, with every trigger met or
   explained in a record with a follow-up.
5. R11's corpus test is green and its delta table committed.

## 11. Risks, findings, open questions and deltas

### 11.1 Risks

- **Risk 1 (leaf volume).** M1's core axes depend on about 35 core leaf requests,
  plus four correctness fixes, across nine crates (§4.11). If they slip, M1 ships with many listed gaps.
  Mitigation: the core/extended split; gaps are explicit in `capabilities()`;
  wave 1 starts now in parallel with R1–R3; correctness fixes go first.
- **Risk 2 (retained memory).** Memoized visible payloads, block indexes, scenes
  and sheet stores raise resident memory for handles that are read deeply.
  Mitigation: they are the leaves' weighted clean-value caches with ADR 0005
  eviction, never handle-level `OnceLock`s (§3.3; the rule of performance change 0680);
  R10 measures peak RSS on every route.
- **Risk 3 (error timing).** Index builds still scan their whole part, so a
  malformed tail is reported on the first touch of that axis, as today; only
  the query walk stops early. A caller who expects `get("/body/p[1]")` to
  succeed on a document with a broken tail gets the same refusal the leaf
  gives today. This is deliberate (`DX/parts/document_part.rs:579-582`).
- **Risk 4 (MCE addressing).** Counting paragraphs on the visible view departs
  from OfficeCLI on body-level `mc:AlternateContent` (M1-Δ12), and M2's
  mapping depends on record 0011. If 0011 slips, M2 refuses DOCX body edits
  on MCE-declaring files; M1 reads are unaffected.
- **Risk 5 (two paragraph meanings).** The legacy Python `Document.paragraphs()`
  keeps space (a) while `/body/p[N]` is space (c) (§4.5.1). Both are
  documented; a test pins each.
- **Risk 6 (shared PPTX source).** A node that kept a `SourceSlide` would make
  M2's sole-owner editor path fail (`PX/presentation/source.rs:1205-1208`).
  A test asserts that after dropping every node, `try_into_editor` succeeds.
- **Risk 7 (legacy PPT cost).** PPT slide and shape values are not `Sync`
  (§4.8), so each `get` re-parses the slide (O(slide) per call) until
  L-PPT-R6.
- **Risk 8 (XLS divergence).** Until L-XLS-R1/R2, a path-opened and a
  byte-opened `.xls` can list different sheets and the source owner can
  report phantom cells (§4.8). R6b is gated on those two requests.
- **Risk 9 (OfficeCLI parity).** Uniform rules depart from OfficeCLI in about
  twenty places (§11.4). Agents trained on OfficeCLI may hit them. Mitigation:
  the delta table, `capabilities()`, and R11's corpus test.
- **Risk 10 (Markdown port).** R8 rewrites the facade's 2,161-line Markdown writer
  onto typed nodes (`L/markdown/writer.rs`). Mitigation: byte-compare the
  Markdown output over the corpus before and after; the existing Markdown
  tests stay green.
- **Risk 11 (noise).** The S1 capture found about a third of rows above 5%
  spread on the shared host (`docs/unified-ops/m0a-log.md:34`); M1's
  sub-microsecond `get` triggers need ABBA and per-route noise floors.
- **Risk 12 (serializer recursion).** `Serialize` recurses once per depth level;
  the depth clamp of 64 bounds it. A test serializes a depth-64 synthetic
  tree.

### 11.2 Findings (disagreements with the design or existing records)

- **F-1.** Design §5 calls `format` a flat `key → string` map; OfficeCLI emits
  typed JSON values (`OF:114-133`). M1 follows OfficeCLI; canonical strings
  remain available (§2.3).
- **F-2.** Design §15 says agents can read anything after M1. The three iWork
  and seven minor ODF formats have no `File` variant until M6 (record 0003
  D3), so that holds for 11 of the 21 claimed formats.
- **F-3.** ADR 0007's shared `Story` traversal and cross-format content-event
  streams (`docs/adr/0007-office-object-models.md:9-11`, `:30-33`) do not
  exist in code; M1 builds per-format story access (L-DOCX-R5) rather than a
  shared trait.

Code findings (inferred from reading unless stated):

- **F-4.** The `litchi_core::sheet` traits box every worksheet and cell
  (`crates/litchi-core/src/sheet/traits.rs:40-157`); unsuitable for M1 paths.
- **F-5.** OfficeCLI's PowerPoint `get` mutates the in-memory document when a
  placeholder exists only on the layout (`P.Res:302-401`).
- **F-6.** `source_backed::Document::paragraph_text` is documented as "direct
  body paragraph" but counts every outermost paragraph
  (`DX/source_backed.rs:3296-3321`).
- **F-7.** DOCX `Cell::grid_span`/`v_merge` can read a nested cell's `tcPr`
  (`DX/table.rs:93-149`).
- **F-8.** DOCX run boolean readers also see `w:rPrChange` and nested textbox
  runs (`DX/paragraph/codec/run.rs:450-488`, `:598-633`).
- **F-9.** DOC in-table paragraphs with nesting level 0 never form a table
  (`crates/litchi-doc/src/parts/pap/model.rs:20`, `:91`;
  `crates/litchi-doc/src/document/model/semantic/layout.rs:289-293`); the
  survey's byte-level probe counted 351 such paragraph-property records in
  `HeaderFooterProblematic.doc` and 10 in `tdf71749_with_footnote.doc`.
- **F-10.** PPT placeholders whose text lives in the slide list come back
  empty: `OutlineTextRefAtom` is parsed and never resolved
  (`crates/litchi-ppt/src/escher_textbox.rs:299`;
  `crates/litchi-ppt/src/slide/types/semantic.rs:413-421`). The survey's
  byte-level scan found that atom in 2 of the 6 files under
  `test-data/ole/ppt` and 7 of the 21 under
  `test-data/poi/test-data/slideshow`.
- **F-11.** `tools/test_crud_coverage_index.py:28` pins 533 selectors while
  `tools/perf-baseline/src/lib.rs:62337` registers 534.
- **F-12.** DOCX `statistics()` counts words over text with no paragraph
  separators, so words in adjacent paragraphs merge
  (`DX/document/package/model.rs:713-754`; `DX/parts/document_part.rs:1006`).
- **F-13.** The Python `__repr__`s slice text at byte 50 or 30 and panic on
  multibyte text (`crates/pyo3-litchi/src/document.rs:143`, `:200`,
  `:293`; `presentation.rs:125`) — read, not run.
- **F-14.** The binding's `litchi` dependency enables neither `ppt` nor
  `xls` (`crates/pyo3-litchi/Cargo.toml:13-21`), while its README promises
  both (`crates/pyo3-litchi/README.md:8`, `:74`, `:94`).
- **F-15.** M0a's `Docx::section_inventory` on the source route
  (`O/format/docx.rs:367-373`) calls a leaf path that refuses any MCE
  markup (`DX/source_backed.rs:2667-2828`), so it fails on typical Word
  files; record 0011's admission work should cover it.
- **F-16.** XLSB's source `materialize()` refuses workbooks with pivot
  tables, slicers or timelines, so today's ops source-mode cell and text
  reads fail on them (`XB/workbook/source.rs:1264-1298`, `:1071-1103`).
- **F-17.** XLSX `ArrayFormula.row`/`column` are documented one-based but
  filled zero-based (`XX/workbook/worksheet.rs:48-51` vs `:289-292`).

### 11.3 Open questions (provisional defaults adopted, for owner confirmation)

| Id | Question | Default |
|---|---|---|
| Q1 | `format` values: OfficeCLI's JSON types or design §5's strings? | JSON types (F-1) |
| Q2 | `childCount`: omit when counting would decode an untouched part, or always compute? | omit (M1-Δ1) |
| Q3 | Regex values (`~= r"…"`): add the `regex` crate (new dependency) or refuse? | refuse with `unsupported_capability` until the owner approves a dependency |
| Q4 | `:contains` case: uniform case-insensitive, or OfficeCLI's per-kind rules? | uniform (M1-Δ10) |
| Q5 | Depth: one rule for every format, clamp 64? | yes (M1-Δ6, Δ7) |
| Q6 | Sheet names caseless in every workbook format (changes XLSB and ODS; record 0003 F15)? | yes |
| Q7 | `/sheet[N]`: all tabs in workbook order? | yes; XLSB counts worksheets until L-XLSB-R1 |
| Q8 | Node `path`: OfficeCLI's identity forms (`p[@paraId=…]`, `shape[@id=…]`)? | yes, plus `positional_path()` |
| Q9 | Range children: stored cells only unless `dense` is asked for? | yes |
| Q10 | Both `get` (`path_not_found`) and `find` (`Option`)? | yes (D9) |
| Q11 | The RTF handle switch changes `paragraph_count` to count `\par` paragraphs. Accept? | accept, record the behaviour change |
| Q12 | Retire `compat` in M1 (R8) rather than keep it to M3? | M1 |
| Q13 | Python `query` default `depth`: 0, or OfficeCLI's one-level JSON hydration? | 0 in the Python method; the M3 command channel hydrates like OfficeCLI |

### 11.4 OfficeCLI deltas introduced by M1

The text cites these as M1-Δn.

| Id | OfficeCLI | M1 | Disposition |
|---|---|---|---|
| Δ1 | `childCount` always emitted (`OF:808`; `W.Nav:3911`) | omitted when it would decode an untouched part | improvement (cost) |
| Δ2 | `effective.*` read-only keys (`WordHandler.StyleList.cs:389-470`) | not emitted | deferred (needs style resolution) |
| Δ3 | `_`-prefixed keys and raw XML in `format` (`W.Nav:5592`, `:6274-6297`; `P.NB:2113`) | only declared keys | wontfix |
| Δ4 | paragraph `style` = display name (`W.Nav:4004`) | style id | deferred (L-DOCX-R6) |
| Δ5 | Excel cell `text` = formatted display value (`X.Cell:96-111`) | stored value | deferred (L-XLSX-R3) |
| Δ6 | depth per handler; ranges list every coordinate (`W.Nav:118-478`; `P.Q:1328`; `X.Q:53-59`; `X.Cell:358-413`) | one depth rule; sparse ranges | improvement (cost) |
| Δ7 | depth clamp 256, no output cap (`CB.GQ:37-38`) | clamp 64, `output_too_large` after `max_nodes` | improvement (bounded output) |
| Δ8 | `get` may clone a layout placeholder into the slide (`P.Res:302-401`) | reads never mutate; the layout node is returned as inherited | improvement |
| Δ9 | Word generic path segments case-sensitive (`W.Nav:1611`) | case-insensitive | improvement |
| Δ10 | Word body `:contains` case-sensitive (`W.Sel:154`, `:371`) | case-insensitive | improvement (Q4) |
| Δ11 | misses are `not_found` by message text; some malformed paths `internal_error` (`OF:317-715`) | `path_not_found`, `invalid_path` | improvement (design §2) |
| Δ12 | paths count raw markup | paths count the MCE-visible view | deliberate (record 0008 I2) |
| Δ13 | `r[N]` excludes comment-reference, inline-`sdt`, `fldSimple` and `AlternateContent` runs (`W.Nav:970-1004`) | the leaf's run set until L-DOCX-R8 | deferred |
| Δ14 | `/header[@type=…]` ignored, resolves header 1 (`W.Q:418`) | honoured | improvement |
| Δ15 | SmartArt takes a `shape` index in nodes but not in resolution (`P.NB:198-213`; `P.Q:1372`) | `smartart` in both | improvement |
| Δ16 | `sheet[N]` counts sheets (`X.Q:234-242`) | XLSB counts worksheets until L-XLSB-R1 | deferred |
| Δ17 | — (no DOC support) | DOC `r[N]` counts merged property runs | inherent |
| Δ18 | — (no RTF support) | RTF `/body` children: paragraphs, then tables | deferred (M6) |
| Δ19 | query materializes full nodes; PPT results grouped by kind (`P.Q:2050-2345`) | streamed typed nodes in document order; `limit` | improvement |
| Δ20 | `issues` and `forms` views (`CB.View:618-695`) | `unsupported_capability` | deferred (M2 `validate`) |
| Δ21 | `:has`/`:not` only as literal Excel forms; one `>` in Word (`X.Sel:118-127`; `W.Sel:15-54`) | general forms | superset (record 0005 deltas 15–17) |

## What is not claimed

- No Litchi code was written, built or run for this record; no number in it
  is a measurement. Costs marked "(inferred)" are read from code.
- The leaf facts come from seven read-only surveys of this tree and of the
  OfficeCLI clone, with the load-bearing citations re-checked by hand; a
  citation can still be off by a few lines.
- The counts in F-9 and F-10 come from throwaway byte-level Python scans
  the surveys ran over `test-data`, not from Litchi code; the scans were
  deleted.
- Sizes in §4.11 and §10 are estimates.
- The budgets in §9.2 are review triggers to be tightened after measuring
  noise, not targets or claims.
- OfficeCLI behaviour is described from its source at `dced0d74`; where its
  documentation and code disagree (README JSON shapes, SKILL.md's
  `.data.Results`), this record follows the code.

