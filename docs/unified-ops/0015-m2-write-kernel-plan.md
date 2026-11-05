# 0015 — M2 write kernel: executable plan

- Milestone: M2 (design §15: "`set/add/remove/move/swap` for OOXML ×3,
  single-op transaction sugar, atomic batch (meeting the ≥1k-operation
  performance budget), property registry + coercions, typed errors with
  suggestions — plus the P0 structural surface"), and its tier-1/tier-2
  continuation M2b (goal, "EXECUTION PLAN": "Write handlers for xlsb (tier 1)
  and legacy OLE doc/xls/ppt (tier 2)").
- Status: plan. No production code is changed by this record. A throwaway
  probe measured leaf write costs (§5.6) and was deleted.
- Date: 2026-10-05
- Base: `uo/0015-m2-plan` at `a8ba410287` (= `uo/m1` after the merge of M0a
  S8a–S8d and records 0014a/b). Every `path:line` below was read in that
  tree, except those prefixed `m1@056bc55b30:`, which were read at the R5
  merge that landed on `uo/m1` after this branch was cut (`File::get`/`find`/
  `query`/`view`/`capabilities`, the capability report and the OfficeCLI delta
  table).
- OfficeCLI was read from the untracked clone at
  `/home/zhuhe/code/litchi/3rdparty/OfficeCLI/` (commit `dced0d74`).
- `performance_claim: leaf costs only.` §5.6 reports leaf measurements from a
  throwaway probe on one shared host; every unified-overhead figure in this
  record is an estimate, and every budget a review trigger (§10.2).

Path prefixes (relative to the repository root unless they start with `OC/`):

| Prefix | Directory |
|---|---|
| `O/` | `crates/litchi-ops/src/` |
| `OPC/` | `crates/litchi-opc/src/` |
| `D/` | `crates/litchi-docx/src/` |
| `X/` | `crates/litchi-xlsx/src/` |
| `P/` | `crates/litchi-pptx/src/` |
| `DML/` | `crates/litchi-drawingml/src/` |
| `B/` | `crates/litchi-xlsb/src/` |
| `W/` | `crates/litchi-doc/src/` |
| `L/` | `crates/litchi-xls/src/` |
| `T/` | `crates/litchi-ppt/src/` |
| `OC/` | `/home/zhuhe/code/litchi/3rdparty/OfficeCLI/src/officecli/` |

## 0. Decisions (provisional; open questions in §12.3)

| # | Decision | Section |
|---|---|---|
| D1 | Verbs: `set`, `add`, `remove`, `move_to`, `swap` (sugar, one-item batches), `edit` (closure, atomic), `batch` (atomic or best effort), `apply` (OfficeCLI's command channel), `save`/`save_as`/`save_with`/`write_to`/`to_bytes`, `create`. Every write takes `&self` and returns a new handle (record 0009 D1). Sugar returns `File`; `edit` returns `Commit<T>` (file, patch, diagnostics, value); `batch` returns `BatchCommit` with one result per item. | §2 |
| D2 | Sequential semantics: item k sees items 1…k−1. Coalescing is an optimization that never changes a result; every batch test also runs one item per transaction and compares bytes. | §5.1 |
| D3 | The registry is M1's `PropSpec` tables grown in place: a `registry!` macro emits the readable `PropSpec` slice and a parallel `PropEntry` (write access, setter enum, constraint, aliases) table per element, with sorted permutations; lookup is ASCII case-insensitive, allocation-free; suggestions on the error path only. | §3 |
| D4 | Writes refuse invalid values (`invalid_value`, key named, suggestion) and unknown keys (`unsupported_property`, "did you mean"); nothing is clamped, auto-corrected or dropped. OfficeCLI's warn-and-continue is a switch on the command channel only. | §3.4–3.5, §12.4 |
| D5 | After a write, eager owners yield a new owner; a source-backed DOCX handle (and a PPTX handle on its source text editor) holds a pending commit published by `save` in one leaf publication, and rebases once, lazily, for the next read or write. | §4.1 |
| D6 | Runs: consecutive items of one family and engine, within the leaf's caps, with no addressing conflict (barriers); fusion into the leaf's batched operations; chaining between runs is native on eager engines and a rebase on OOXML source-backed engines until option D. | §5.3–5.4 |
| D7 | Engines: XLSX promotes to the eager `Workbook` at the first write (exact source); PPTX runs shape-text runs that fit the source-backed editor there and promotes for anything else; DOCX stays on its owner's engine and rebases. | §4.2–4.4, §5.6 |
| D8 | Addressing reuses M1's resolution and maps to leaf coordinates per format (DOCX transaction ordinals, XLSX zero-based cells, PPTX staged slide index and shape key); positions the leaf cannot prove are refused, never guessed. | §4 |
| D9 | Atomic batches stop at the first failure and leave the caller's handle untouched; best-effort batches skip failing items; commit-time failures are attributed by bisection on the error path only; `Error` gains `item()` and `property()`. | §5.5 |
| D10 | Leaf prerequisites in order: record 0009 S1–S3 (and S4–S5 for M2b), then addressing and admission fixes, then the parity engines (DOCX formatting, XLSX style allocation, PPTX shape properties), metadata families and legacy creators. | §6 |
| D11 | P0 structural surface: worksheet lifecycle is wiring (XLSX); metadata, style application and numbering basics need the leaf families of §6.4. | §9 |
| D12 | `capabilities()`, the help-schema JSON, Python introspection and the OfficeCLI delta table are generated from the registry and handler tables; goldens re-blessed; a write-liveness test round-trips every settable key. | §3, §11 W6 |
| D13 | Save is the leaf save. The facade writer is a policy layer over `litchi_opc::atomic` and CFB's `publish_staged` (exported by L-OLE-W1); it adds no rename or fsync code of its own. | §7.2 |
| D14 | A file `Signed::unsign` returned writes the unsigned copy when saved, edited or not; the removal happens at the first write or save. | §8.1 |
| D15 | Sensitive (unlocked) files save encrypted (`SaveOptions::reencrypt(password)`) or not at all; plaintext needs `File::into_plaintext()`; editing a sensitive file is P1. | §8.2 |
| D16 | `Patch` is the batch's ordered list of leaf patch steps; `File::apply_patch` works where no rebase separates the steps. | §4.1 |
| D17 | Bytes a rebase or promotion creates are charged to the open's execution context when it has one, and count against the open's input ceiling otherwise. | §4.1 |
| D18 | Branching: `uo/m2` from `uo/m1` after R5; steps W1–W8 plus leaf steps; merge order M0a → M1 → M2 into the trunk; M2b on the same branch after the core exit. | §11 |
| D19 | The raw layer is outside M2's core: `raw` and `add-part` are an M2-extended step, `raw-set` and `validate` need new engines (XPath, schema validation) and move to M4; the batch parser accepts their keys from the start. | §2.3 |

## 1. Starting point

### 1.1 What M2 builds on

- **`File` and handles.** `File` is a flat, `#[non_exhaustive]`, 16-byte enum,
  `Clone + Send + Sync`, one variant per format plus `Locked`/`Signed`
  (`O/file.rs:131-185`). Each tier 1–2 handle is `Arc<Opened<State>>`: one
  open's state shared by every clone, plus the handle's signature disposition
  (`O/format/signature.rs:22-93`). No handle records the path it was opened
  from (grep over `O/format/` and `O/open*`).
- **Editor builders, written in M0a and unused until M2** (each carries
  `expect(dead_code, reason = "… M2")`):

  | Handle | Builder | Receiver | Returns |
  |---|---|---|---|
  | `Docx` | `edit_document` (`O/format/docx.rs:472-499`) | `&self` | `DocxEditor::{Edit(Box<document::Edit>), Unsigned(Box<Package>)}` (`:127-138`) |
  | `Xlsx` | `into_editor` (`O/format/xlsx.rs:441-472`) | **`self`** | `XlsxEditor::{Source(Box<SourceBackedPackage>), Eager(Workbook)}` (`:141-156`); falls back to `materialize` (`:504-506`) when shared |
  | `Pptx` | `into_editor` (`O/format/pptx.rs:518-585`) | **`self`** | `PptxEditor::{Source, Eager(Box<Transaction>), Unsigned}` or `EditorRefused` with the handle (`:152-181`) when shared |
  | `Xlsb` | `editor` (`O/format/xlsb.rs:345-368`) | `&self` | `XlsbEditor::{Eager(&Workbook), Materialized(Box<Workbook>)}` (`:100-121`); `to_workbook()` on every call |
  | `Doc` | `body_text_editor` (`O/format/doc.rs:187-203`) | `&self` | `Option<SourceSnapshot>` |
  | `Xls` | `cell_values_editor` (`O/format/xls.rs:246-273`) | `&self` | `cell_values::Snapshot` (copies and re-indexes the input per call) |
  | `Ppt` | `text_editor` (`O/format/ppt.rs:241-272`) | `&self` | `Option<text_edit::SourceSnapshot>` |

  All seven refuse a still-signed handle with `document_protected` and build
  over an explicitly unsigned copy after `Signed::unsign` (record 0013 §1).
  The two consuming builders and the XLSX `materialize` fallback are what
  record 0009 S1 and S3 replace (§6).
- **M1 read side.** The DOCX (R4a) and XLSX (R4b) adapters resolve paths on
  the leaves' indexes, carry the leaf values a write needs (a DOCX block's
  `transaction_position()`, `O/format/docx/node.rs:427-455`; an XLSX sheet's
  position and handle, `O/format/xlsx/node.rs:97-104`, `:194-220`), and
  declare their properties in static spec tables (`O/format/docx/spec.rs:409-640`:
  170 OfficeCLI keys over 12 tables; `O/format/xlsx/spec.rs`). R5 merged
  into `uo/m1` after this branch was cut (`056bc55b30`): `File::{root, get,
  find, query, view, view_to, view_json, capabilities}`, one `match` per entry
  point (`m1@056bc55b30:crates/litchi-ops/src/read.rs:1-41`), and a capability
  report whose `Verb` enum already lists `Set`, `Add`, `Remove`, `Move`,
  `Swap`, `Batch` and `Save` as `unsupported_capability`
  (`m1@056bc55b30:crates/litchi-ops/src/capabilities.rs:221-261`), with the
  `add`/`set`/`remove` flags of `Element::operations` and of each property
  left false for M2 (`:1110-1124`, `:1243-1263`). The PPTX adapter (R4c) is
  committed on `uo/m1-r4c-pptx` (`5e9137372c`) and not yet merged.
- **Placeholders M2 fills.** `O/batch.rs:1-4` ("Batch executor: atomic and
  best-effort batches with Edit-family coalescing"), `O/registry.rs:1-4`,
  `O/handler.rs:1-4`.
- **Values and errors.** `litchi_ops::value` coerces lengths, colours,
  numbers, booleans and vocabularies without allocating
  (`O/value.rs:1-80`; `ValueType::coerce`, `O/value/ty.rs:70`); "did you
  mean" is `value::suggest::nearest` with OfficeCLI's threshold
  (`O/value/suggest.rs:44-56`). The unified `Error` is one pointer with 39
  codes (`O/error.rs:58-140`, `:435-452`); it has no item index, property
  key or diagnostics slot yet (m0a-log S2 follow-up "(M2) Add a diagnostics
  slot").

### 1.2 What the leaves can write today

Record 0002 surveyed this on 2026-10-04; the table below is re-read at this
base (§4 has the detail and the citations). "Existing content" means editing
what an opened file already holds.

| Format | Edit entry (receiver) | Commit shape | Chains? | Formatting of existing content | Structure | Atomic path save |
|---|---|---|---|---|---|---|
| DOCX | `edit_document(&self)` on both owners | snapshot, patch, diagnostics | eager yes (inferred); source-backed no (`StaleSource`) | none: readers only (`D/properties.rs:301`; `D/paragraph/codec/paragraph_properties.rs:282`) | plain paragraph insert; no public remove or move | eager `save*`; source-backed: streams only, except tail append's `write_to_path` |
| XLSX | `Workbook::edit(&self)` (eager); 16 per-feature source editors | workbook, patch; no diagnostics | yes | apply an existing `Style` only; no allocation | sheets add/remove/rename/move/hide; no row/column shift | `Workbook::save*` |
| PPTX | `opened_presentation_transaction(&self)` (eager); source editor via `try_into_editor(self)` | snapshot, patch; no diagnostics | yes, unapplied commits chain | none (fill only at creation) | slide move/remove (dependency-free); shapes add/remove; no slide add | eager `save*`; source-backed streams only |
| XLSB | `edit_cell_values(&self)`; structure edit | snapshot, patch; applied with `&mut self` | per worksheet | existing XF only | rename only | `Package::save*` |
| DOC | eager `body_text::Snapshot::edit(&self)`; narrow source editor | snapshot, patch (source commit: + diagnostics) | yes | bold, italic, underline | none | source commits only |
| XLS | `cell_values::Snapshot::edit(&self)` | snapshot, patch, diagnostics | yes | XF index or raw XF bytes | rows/columns insert/delete, rename | source-backed numeric commits only |
| PPT | `slide_order::Snapshot::edit(&self)`; narrow text source editor | snapshot, patch | yes | none | slide move/remove | text source commits only |

The consequence for M2: every tier-1 leaf can change text, structure and
workbook values, but none can change the *formatting* of content that
already exists, and the OfficeCLI corpus formats nearly everything it adds
(§6.4). That gap, not the facade, is M2's largest cost.

### 1.3 What OfficeCLI does on the write side

- **One handler per format, item by item.** `IDocumentHandler` declares
  `Set` (returns the keys it did not apply), `Add` (returns the new path),
  `Remove`, `Move`, `CopyFrom`, `Raw`, `RawSet`, `AddPart`, `Validate` and
  `Save` (`OC/Core/IDocumentHandler.cs:84-126`); `Swap` is per handler
  (`OC/Handlers/Word/WordHandler.Mutations.cs:1627`). A batch runs its items
  in order through one handler (`OC/CommandBuilder.Batch.cs:62-153`) with no
  coalescing; its speed comes from deferring serialization
  (`OC/Handlers/WordHandler.cs:183-212`).
- **Atomicity by file copy.** An atomic CLI batch runs on a hidden sibling
  copy of the file and replaces the original only if every item succeeded
  (`OC/CommandBuilder.Batch.cs:481-643`); without `--stop-on-error` every item
  still runs before the rollback (`:119-131`); the in-process executor has no
  rollback at all (`OC/Core/BatchExecutor.cs:31-61`).
- **Lenient properties.** Unknown keys are written as raw XML when they match
  an OOXML name, otherwise reported as `unsupported_property` warnings with a
  suggestion while the other keys apply; a key one edit away from a known one
  is auto-corrected (`OC/CommandBuilder.cs:764-806`, `:1704-1821`;
  `OC/CommandBuilder.Set.cs:329-351`). Invalid values throw and become
  `invalid_value` (`OC/Core/OutputFormatter.cs:546-554`). Error codes are
  inferred from message text, so many write failures surface as
  `internal_error` (`:317-715`).
- **Schemas.** 152 help schemas declare, per element, `operations` and, per
  property, `add`/`set`/`get` flags, `type`, `values`, `aliases` and more
  (`3rdparty/OfficeCLI/schemas/help/_schema.json:7-177`); after their `extends` bases are merged,
  DOCX declares 963 properties (659 settable, 641 addable), XLSX 693 (402,
  450) and PPTX 687 (510, 417). They drive help only and never gate a write
  (`OC/Core/TrackingPropertyDictionary.cs:10-15`).
- **Usage.** The 96 example scripts (14 Word, 26 Excel, 56 PowerPoint) call
  `add` about 2,430 times and `set` about 1,180 times, against `create` 99,
  `validate` 88, `raw-set` 42, `remove` 11 and `move`/`swap` never (a grep for
  the CLI's verbs over `3rdparty/OfficeCLI/examples/{word,excel,ppt}/**/*.sh`,
  comments included, so the counts are approximate).

## 2. The verb surface

### 2.1 Rust

All verbs are methods of `File` in a new module `O/write.rs`, the write twin
of R5's `read.rs`: one `match` on the variant per verb, statically
dispatched to the format's handler (§4.1). `File::Signed` and `File::Locked`
refuse through the existing builders (record 0013 §1) and `Locked`'s read
error; tier-3 variants refuse with `unsupported_capability`, follow-up M6.
The receivers are record 0009's (D1, §7): every write verb takes `&self`,
returns a new handle, never consumes its receiver and never reads a
reference count, so a file's bytes are a function of the snapshot and the
operations only.

```rust
// litchi_ops, re-exported by litchi (ergonomics bar: one entry point).
pub fn create(format: FileFormat) -> Result<File>;
pub fn create_with(format: FileFormat, options: &CreateOptions) -> Result<File>;

impl File {
    // Single-operation sugar: one leaf Edit and commit (one-item batch, §5).
    pub fn set<'p>(&self, target: impl IntoTarget<'p>, props: impl IntoProps) -> Result<File>;
    pub fn add<'p>(&self, parent: impl IntoPath<'p>, node: impl IntoNode) -> Result<Added>;
    pub fn remove<'p>(&self, target: impl IntoTarget<'p>) -> Result<File>;
    pub fn move_to<'p, 'q>(&self, path: impl IntoPath<'p>, to: impl IntoPlacement<'q>) -> Result<File>;
    pub fn swap<'p, 'q>(&self, a: impl IntoPath<'p>, b: impl IntoPath<'q>) -> Result<File>;

    // Several operations, atomic: the closure stages into a transaction proxy.
    pub fn edit<T>(&self, f: impl FnOnce(&mut Transaction<'_>) -> Result<T>) -> Result<Commit<T>>;
    // OfficeCLI's batch, atomic or best effort; its single-command channel.
    pub fn batch<'i, I>(&self, items: I, mode: BatchMode) -> Result<BatchCommit>
    where
        I: IntoIterator,
        I::Item: Into<BatchItem<'i>>;
    pub fn apply<'i>(&self, item: impl Into<BatchItem<'i>>) -> Result<Commit<Applied>>;

    // Publication of this snapshot, repeatable on any handle (§7).
    pub fn save(&self) -> Result<Saved>;
    pub fn save_as(&self, path: impl AsRef<std::path::Path>) -> Result<Saved>;
    pub fn save_with(&self, to: SaveTo<'_>, options: &SaveOptions) -> Result<Saved>;
    pub fn write_to<W: std::io::Write>(&self, sink: W) -> Result<u64>;
    pub fn to_bytes(&self) -> Result<Vec<u8>>;

    // Security states (§8).
    pub fn is_sensitive(&self) -> bool;
    pub fn into_plaintext(self) -> File;   // consuming, as `Signed::unsign`
}
```

`move` is a Rust keyword, hence `move_to` (Python keeps `move`). Arguments:

| Type | What it accepts | Notes |
|---|---|---|
| `IntoPath<'p>` | `&str` (parsed once, borrowing) or a parsed `Path` | M1's trait (`O/path/arg.rs:36`), unchanged |
| `IntoTarget<'p>` | a path, or a query (`&Query`/`query::Plan`) for `set` and `remove` | a query target applies to every match, as OfficeCLI fans selectors out (one `set` or `remove` per match, removes in reverse document order: `OC/Handlers/Word/WordHandler.Set.cs:132-153`, `OC/Handlers/Word/WordHandler.Mutations.cs:44-62`); M2 runs all matches as one batch, so they coalesce. A selector that is not scoped (no leading `/`, no top-level `/`, not `Sheet!Ref`) is refused with `bare_selector_rejected`, OfficeCLI's guard (`OC/Core/MutationSelectorGuard.cs:37-56`; record 0005 delta 24) |
| `IntoProps` | `props!{ "font.bold" => true, prop::SIZE => 12.pt() }` (an array on the stack), `&[(K, V)]`, `Vec<(K, V)>`, the OfficeCLI string forms (`&[&str]` of `"key=value"`) | `K: Into<Key>` (`&str` or a typed key); `V: Into<PropValue>` |
| `PropValue<'a>` | `bool`, integers, `f64`, `&str`/`String`, `Length` (`12.pt()`, `2.cm()` through the `Units` extension trait), `Color` (`Color::rgb(..)`, `"1F4E79"`, `"accent1+lumMod75"`), `PropValue::Null` (clear the property) | coerced through the key's `ValueType` (§3.3) |
| `IntoNode` | `NewNode::new("paragraph").prop("text", "Hi").prop("bold", true)`, `("paragraph", props)`, a bare type name | the element type is resolved through the format's alias table (`p`, `paragraph`), as M1 resolves path steps |
| `IntoPlacement<'q>` | `Placement::{Start, End, Index(usize), Before(path), After(path), Into(parent path, Option<index>)}`, from `usize` (index) | OfficeCLI's `--index` is zero-based and counted after the element is detached (`OC/CommandBuilder.Add.cs:419`); Litchi keeps zero-based indexes (record 0010 D6) |

**Return shapes.**

| Verb | Returns | Contents |
|---|---|---|
| `set`, `remove`, `move_to`, `swap` | `File` | the new handle; diagnostics are dropped (use `edit` to keep them) |
| `add` | `Added { file, path }` | the new handle and the new element's canonical path (OfficeCLI's `Add` returns the path, `OC/Core/IDocumentHandler.cs:88`); `Added: Deref<Target = File>` and `into_file()` |
| `edit` | `Commit<T> { file, patch, diagnostics, value }` | design §3's normalized shape (ADR 0003, `docs/adr/0003-snapshots-edits-and-patches.md:21-23`) plus the closure's value |
| `batch` (atomic) | `Ok(BatchCommit { file, patch, diagnostics, results })` or `Err(error)` with `error.item()` | every result `Applied`; the error names the first failing item and the caller's handle is unchanged |
| `batch` (best effort) | `Ok(BatchCommit)` | `results[i]` is `ItemResult::Applied(Applied { path, changed })` or `ItemResult::Failed(Error)`, in item order; the file reflects the applied items |
| `apply` | `Commit<Applied>` | one item through the batch path |
| `save*` | `Saved { bytes: u64, path }` | |
| `write_to` | the byte count | |

`Diagnostics` is a list of advisory entries (design: "Advisory/warning codes
(`auto_corrected`, `text_overflow`, …) travel as diagnostics, not errors"),
each `{ code, item: Option<u32>, path, message }`, plus the leaf facts the
commit shape carries (`changed`, `operations`); `Diagnostics::is_exact_noop()`
is true when nothing changed. `Patch` is the batch's ordered list of leaf
patch steps (§4.1).

**`Transaction`.** The proxy of `edit` has the same five verbs plus `get` and
`find`, all `&mut self`; each call plans and stages its item immediately
(errors surface at the call, with the item's index), and `add` returns the
new element's path. `get`/`find` read the state as staged so far, through
the same staged views and barriers as a batch (§5.3): a read that needs a
barrier closes the run first. The closure's `Ok` commits the open run; `Err`
or a panic drops everything (ADR 0003, `docs/adr/0003-snapshots-edits-and-patches.md:13-16`, transaction-scoped proxies).

### 2.2 Python (the M3 surface, fixed here)

M3 implements the bindings (goal "EXECUTION PLAN"); M2 fixes their shape so
the Rust API serves it without copies: the handle is `Send + Sync`, so every
call below releases the GIL (`py.detach`, as the read bindings already do,
`crates/pyo3-litchi/src/sheet.rs:102-149`), and batch payloads are converted
to `BatchItem`s once, before the GIL is released.

```python
class File:
    def set(self, path: str, props: Mapping[str, object]) -> File: ...
    def add(self, parent: str, type: str, props: Mapping[str, object] | None = None, *,
            index: int | None = None, after: str | None = None,
            before: str | None = None) -> File: ...
    def remove(self, path: str) -> File: ...
    def move(self, path: str, *, to: str | None = None, index: int | None = None,
             after: str | None = None, before: str | None = None) -> File: ...
    def swap(self, path: str, path2: str) -> File: ...
    def apply(self, command: Mapping[str, object]) -> ApplyResult: ...   # .file, .path, .diagnostics
    def batch(self, items: Sequence[Mapping[str, object]], *,
              mode: Literal["atomic", "best_effort"] = "atomic") -> BatchResult: ...
              # .file, .results (list of {"index", "success", "path"?, "code"?, "message"?, "suggestion"?}),
              # .diagnostics
    def save(self, path: str | os.PathLike[str] | None = None, *,
             durability: Literal["full", "file-only", "no-sync"] = "full",
             password: str | None = None) -> None: ...
    def to_bytes(self) -> bytes: ...
    def into_plaintext(self) -> File: ...
    @property
    def is_sensitive(self) -> bool: ...

def create(format: FileFormat | str) -> File: ...
```

- Every write returns a new `File`; none takes the inner handle out of
  `self` (record 0009 §7). `add` returns the `File`; the new element's path
  comes back from `apply` and `batch` results, as OfficeCLI prints it.
- Advisory diagnostics are raised as `warnings.warn(..., LitchiWarning)`
  from every write, so `doc = doc.set(...)` stays a one-liner.
- An atomic batch failure raises the item's exception with `.item` (the
  index) set; `PropertyError` (already in the stub,
  `crates/pyo3-litchi/python/litchi_py/__init__.pyi:51`) carries `.property`.
- `props` values are Python scalars, `Length`/`RGBColor` objects (in the stub
  today, `:102-174`) or strings; dotted keys are plain dict keys.

### 2.3 The raw layer (design §12) in M2

`BatchItem` carries OfficeCLI's raw keys (`part`, `xpath`, `action`, `xml`)
from W1 on, so dumps parse, but the raw commands are not part of M2's core:

| Command | What it needs | M2 status |
|---|---|---|
| `raw` (read a part) | the leaves' fallible part accessors (ADR 0030) | M2-extended step W10, read-only, no transaction |
| `add-part` | a part plus its relationship, through the eager packages (`OpcPackage::add_part`, `OPC/package.rs:1666`) or a source-backed overlay | W10, as a family of its own (one run, barrier on both sides) |
| `raw-set` (XPath edit) | an XPath engine; the workspace has none (grep: no XPath evaluator in any crate) | new engine (c), scheduled with M4's engines |
| `validate` (schema validation, never mutating, ADR 0006) | an OOXML schema validator; none exists | new engine (c), M4 |

Until then each refuses with `unsupported_capability` and its follow-up, and
`capabilities()` lists them (decision D19). The 42 `raw-set` lines of the
OfficeCLI examples are outside W8's corpus subset.

## 3. The property registry

### 3.1 What M1 left, and the seam

M1 declares, per format, one `static FormatSpec<Kind>` whose `ElementSpec`s
carry the element's readable properties as `&'static [PropSpec]`, a key and
a `ValueType` (`O/node/spec.rs:36-58`, `:60-95`); the node view reads them by
index (`index_of`, e.g. `O/format/docx/props.rs:555-597`), the query
compiler compiles literals against their `ValueType`, and R5's capability
report prints them with `get: true` and `set`/`add` false
(`m1@056bc55b30:crates/litchi-ops/src/capabilities.rs:1110-1124`). Record
0010 §7.3 and R5's module documentation fix the seam: M2 "turns the `set`
and `add` flags on where a typed setter exists" and answers
`Element::operations` from its handler table, without changing the report's
shape. `litchi_ops::value` already holds every coercion a `set` needs:
`ValueType::coerce` (`O/value/ty.rs:70`) is allocation-free, and
`ValueError::into_error_for(property)` (`O/value/error.rs:276`) builds the
typed `invalid_value` with its suggestion on the error path only.

### 3.2 Entry shape

A registry entry is the M1 `PropSpec` plus four write-side fields, declared
in the same row so that one table drives reading, writing, schemas,
capabilities, suggestions and goldens:

```rust
/// One property of one element kind: what M1 reads plus what M2 writes.
pub(crate) struct PropEntry<S: Copy + 'static> {
    pub(crate) spec: PropSpec,                     // key + ValueType (M1, unchanged)
    pub(crate) access: Access<S>,                  // how `set`/`add` reach the leaf
    pub(crate) constraint: Constraint,             // domain beyond the ValueType
    pub(crate) aliases: &'static [&'static str],   // accepted input spellings, sorted
}

pub(crate) enum Access<S: 'static> {
    ReadOnly,                         // `get` only (OfficeCLI's get-only keys)
    Set(S),                           // `set` on an existing element
    Add(S),                           // only when the element is created (`add`)
    SetAndAdd(S),
    Gap(&'static GapSpec),            // the format has it; the leaf cannot write it yet
}

pub(crate) enum Constraint {
    None,
    Int { min: i64, max: i64 },
    Number { min: f64, max: f64 },
    Length { min: Length, max: Length },
    Chars { max: u32 },               // e.g. sheet names: 31 characters
    NotEmpty,
}
```

`S` is a per-format `Copy` enum of setters (`docx::Setter::RunBold`,
`xlsx::Setter::CellValue`, …). The format's handler turns a
`(S, Scalar)` pair into a leaf operation with one exhaustive `match`, so a
declared setter without an implementation, or an implementation without a
declaration, does not compile. No function pointer and no trait object is
stored (goal constraint 2, ADR 0004).

The tables are generated by one macro per format, `registry!`, which emits
from each row both M1's `&'static [PropSpec]` (so the node view, the query
compiler and their allocation tests do not change) and the parallel
`&'static [PropEntry<S>]`, with `const` assertions that the two agree index
by index. A `const`-evaluated sorted permutation of each table (the
`aliases_valid`/`compare` machinery of `O/node/spec.rs:176-252`) gives the
binary search below.

### 3.3 Lookup and coercion (success path allocates nothing)

1. **Key.** Input keys are matched ASCII case-insensitively, as OfficeCLI's
   batch dictionary does (`OrdinalIgnoreCase`, `OC/BatchTypes.cs:14`):
   binary search of the element's sorted permutation; on a miss, the
   format's sorted alias table (`(&'static str alias, canonical)`), then a
   miss. Typed keys (`prop::font::BOLD`) carry the canonical spelling and
   skip the alias table.
2. **Access.** `Access::Set` for `set`, `Add`/`SetAndAdd` for `add`;
   `ReadOnly` refuses with `unsupported_property` ("`paraId` is read-only");
   `Gap` refuses with the gap's `unsupported_capability` and follow-up.
3. **Value.** `PropValue::Text` goes through `ValueType::coerce`; a typed
   `PropValue` of the property's type is used as is; a typed value of
   another type is rendered to its canonical text and coerced (an `Int` for
   a `Length` key means the key's bare-number unit, as the text `"12"`
   would: points for font sizes and line widths, EMU otherwise, design §6);
   `PropValue::Null` clears the property where the entry allows clearing.
4. **Constraint.** Checked on the coerced `Scalar`; a failure is
   `invalid_value` naming the key and the domain.
5. **Setter.** The `(S, Scalar)` pair is handed to the format's run (§5).

Nothing in steps 1–5 allocates when it succeeds: keys and texts are
borrowed, tables are `static`, `Scalar` borrows its text, and errors are
built only on failure. The allocation tests of M1
(`crates/litchi-ops/tests/*_allocations.rs`) gain a write twin that counts
allocations of parse + resolve + lookup + coerce + stage minus the leaf's own
staging, with the expectation zero.

### 3.4 Errors and suggestions

| Situation | Code | Message and suggestion |
|---|---|---|
| Unknown key | `unsupported_property` | "unknown property `font.bolld` for `run`"; suggestion "did you mean `font.bold`?" from `value::suggest::nearest` over the element's keys and aliases (`O/value/suggest.rs:44`), OfficeCLI's threshold `max(2, len/3)`; when the key exists on a related kind, "`fill` is a property of `paragraph`, not `run`" |
| Read-only key | `unsupported_property` | names the key as read-only |
| Known key the leaf cannot write | `unsupported_capability` | the gap's reason and follow-up |
| Value does not coerce | `invalid_value` | from `ValueError::into_error_for` (nearest token for vocabularies, the accepted forms for lengths and colours) |
| Value outside the constraint | `invalid_value` | the domain |
| Missing required key on `add` (e.g. `src` of a picture) | `missing_property` | the key |

Suggestions are computed only on the error path (goal, "Allocation
discipline").

### 3.5 Invalid values: the write policy

M1 recorded two read policies: DOCX leaves a malformed stored value unset
and lists it (`crates/litchi-docx/src/properties.rs:31-40`), while the XLSX
reader refuses (m1-log, L-DOCX-R4/R8 note). For writes, M2 decides:
**refuse, never repair** (decision D4). A `set` or `add` with an invalid
value fails with `invalid_value` before anything is staged; in an atomic
batch the batch fails at that item; in best-effort mode that item fails and
the others proceed. Nothing is clamped, rounded beyond the unit precision
the format stores (Word half-points and twips; the docx reader's own rule,
`crates/litchi-docx/src/properties.rs:42-50`), or silently dropped. This is
ADR 0006's "save never silently repairs" applied to input, and it gives an
agent a typed error with a suggestion to retry on. Reads keep each leaf's
policy until the leaf exposes per-property validity; the node view then
adopts DOCX's "absent and listed" rule for every format (a follow-up outside
M2's steps).

### 3.6 Schemas, capabilities and goldens

R5 already serializes `Capabilities` as OfficeCLI's help schema per element
(`m1@056bc55b30:crates/litchi-ops/src/capabilities.rs:47-79`), from the same
spec tables the resolver runs on. M2 changes only what the flags say:

- `Element::operations()` answers `add`, `set` and `remove` from the
  handler's static table (`const OPS: &[(Kind, Verb, Support)]` per format,
  checked against the handler's exhaustive `match` by a test), instead of the
  constant `false` of M1 (`:1115-1124`).
- Each property's `set`/`add` flags come from its `Access` field; `Gap`
  entries also join the report's `gaps` list with their follow-up.
- Each property gains the schema fields OfficeCLI's `_schema.json` defines
  and the registry can state truthfully: `values` (from the `Tokens`
  vocabulary), `aliases`, and the constraint bounds (`3rdparty/OfficeCLI/schemas/help/_schema.json:121-142`).
  Litchi does not emit `appliesWhen`, `modifiers` or `examples` in M2.
- The `Verb` entries `Set`, `Add`, `Remove`, `Move`, `Swap`, `Batch` and
  `Save` (`:221-261`) report `supported` per format as its handler lands.

Drift guards, added to R5's (`:24-46`):

1. `registry!` `const` checks: one `PropEntry` per `PropSpec`, same order;
   setters unique per element; constraint bounds ordered; alias tables sorted
   and disjoint from canonical keys.
2. Exhaustive `match` on each format's `Setter` enum in its handler.
3. **Write liveness**: for every settable key of every element, on a fixture
   (a generated one where the corpus has none), `set` a sample value of the
   key's type, `get` the node, and require the key to read back as the
   canonical form of the sample; for every `Gap` key, require the gap's code.
   This is what keeps "what you write is what you read" (design §5) true.
4. Golden JSON per format (`crates/litchi-ops/tests/golden/capabilities/`,
   `LITCHI_BLESS=1`), re-blessed per handler step.
5. `tools/officecli_capability_delta.py` (R5) gains write columns: for each
   OfficeCLI property, Litchi's `set` and `add` status against OfficeCLI's
   flags, with the statuses `supported`, `gap` (with follow-up), `read-only`
   and `missing` (untriaged), and the summary counts per format; its `--check`
   mode keeps the committed table in sync with the goldens.
6. The typed `prop` module is generated from the same tables by the macro,
   so a typed key cannot name a property no element has.

Python introspection (M3) reads `capabilities()`; it has no table of its own.

## 4. Per-format write handlers

### 4.1 Shape of a handler

One module per format, `O/format/<format>/write.rs`, next to M1's `nav.rs`,
`node.rs` and `spec.rs`. A handler is not a trait object and not a trait the
facade dispatches through (ADR 0004; design §7 keeps "the handler trait …
internal"): `O/write.rs` holds one `match` per step over a crate-internal
enum, exactly as R5's `read.rs` does for reads
(`m1@056bc55b30:crates/litchi-ops/src/read.rs:76-135`). Each handler provides:

| Piece | Job |
|---|---|
| `Target` | the resolved element plus its leaf coordinate (DOCX `Position`s, XLSX sheet + zero-based cell, PPTX slide index + shape id), built by M1's `Navigate` on the current state |
| `classify(target, verb, keys) -> Plan` | the item's `(family, engine)` and its addressing effect (§5.3) |
| `Run` | the open leaf transaction of one family plus the per-item records needed for results and fusion |
| `open(state, family)` / `stage(run, op)` / `commit(run)` | build the editor from `&self` state; stage one item; commit into the normalized shape |
| `Setter` | the `Copy` enum the registry's `Access::Set(S)` names (§3.2); one exhaustive `match` maps `(Setter, Scalar)` to a leaf call |

**The commit-shape adapter.** Every handler's `commit` returns
`(State, PatchStep, Diagnostics)`:

| Leaf | Leaf commit | Normalized |
|---|---|---|
| DOCX `document::Edit::commit` (`D/document/transaction.rs:5067`) | `Commit { snapshot, patch, diagnostics }` (`:5274-5301`); `Diagnostics { operations, changed }` (`:5253-5268`) | pending source state or forked eager package; `PatchStep::Docx(patch)`; diagnostics carried over |
| XLSX `Edit::commit` (`X/workbook/edit/semantic/transaction.rs:1501`) | `Commit { workbook, patch }`, **no diagnostics** (`X/workbook/edit/model.rs:2909-2951`) | new eager `Workbook`; `PatchStep::Xlsx`; diagnostics synthesized (`changed = !patch.changes().is_empty()`, `operations` = staged item count) |
| PPTX `opened::Transaction::commit` (`P/opened/transaction.rs:1390`) | `Commit { snapshot, patch }`, **no diagnostics** (`:2176-2225`) | pending commit list, applied into one forked package at the end of the batch (§5.4); synthesized diagnostics |
| XLSB `cell_values::Edit::commit` (`B/cell_values/worksheet.rs:1088`) | `Commit { snapshot, patch }`, no diagnostics, **applied in place** by `apply_cell_values(&mut self)` (`B/workbook/package.rs:431-450`) | a new `Workbook` through the non-mutating apply of L-XLSB-W1 (§6); synthesized diagnostics |
| DOC eager `body_text::Edit::commit` (`W/body_text.rs:2115`) | `Commit { snapshot, patch }`, no diagnostics (`:2324-2353`) | the commit's snapshot (chains); synthesized |
| XLS `cell_values` `commit` (`L/cell_values/mod.rs:2723-2876`) | `Commit { snapshot, patch, diagnostics }` (`:3313`) | the commit's snapshot; diagnostics carried over |
| PPT `slide_order` `commit` (`T/slide_order.rs:2121`) | `Commit { snapshot, patch }`, no diagnostics (`:2419-2443`) | the commit's snapshot; synthesized |

The unified `Patch` is the ordered list of `PatchStep`s of a batch, with
`inverse()` (the reversed list of leaf inverses). Applying a patch to a
`File` (`File::apply_patch`) is offered where every step's leaf apply checks
its own source and no rebase separates the steps (eager XLSX
`Workbook::apply`, `X/workbook/model.rs:1395`; PPTX
`apply_opened_presentation_patch`, `P/package/model.rs:582`; DOCX eager
`apply_document_patch`, `D/package/package/document.rs:155`); across a DOCX
source-backed rebase it is `unsupported_capability` until option D (decision
D16).

**Owner state after a write.** Three cases (decision D5):

1. *Eager engines* (XLSX `Workbook`, PPTX and DOCX eager `Package`, XLSB, the
   OLE snapshots): the commit yields a complete new owner, wrapped in a new
   handle state that shares the origin path, the open's limits and the
   sidecars. No pending state.
2. *Source-backed engines* (DOCX, and PPTX runs on its source text editor,
   §4.4): the handle holds `Pending { base, commits }`: a share of the base
   package (record 0012 made it O(1) `Clone`) and the leaf commits not yet
   published. `save`/`write_to` publish base + commits in one
   leaf publication. The first read of, or next write to, a pending handle
   *rebases* once: publish into memory, re-open source-backed over an
   `OwnedSource` (PPTX: into the eager owner, case 3), memoized in a
   `OnceLock` on the handle state.
3. *Promotion* (XLSX and PPTX source owners, XLSB source owner): the first
   write copies the exact source artifact (`OPC/source_backed.rs:7765`,
   `:3832`) and opens the eager owner from those bytes (record 0009 D4),
   memoized per open (not per handle), then proceeds as case 1. The eager
   writer then raw-copies every untouched member and refuses rather than
   regenerates when raw copying is impossible (`PreservationUnavailable`,
   `OPC/pkgwriter.rs:1268-1273`); an exact no-op still writes the source
   bytes (`OPC/pkgwriter.rs:1257-1265`).

Memory for a rebase or a promotion is charged to the handle's execution
context when the open had one (ADR 0031, `Resource::Memory`, refused with
`limit_exceeded` when the budget is spent); without a context the bytes count
against the open's input ceiling, as an `open_bytes` of them would (record
0009 §9's open point, decided here as D17).

### 4.2 DOCX (tier 1)

**Owners and engine.** The handle's owner decides the engine, and M2 does
not promote DOCX (D7). Three reasons: the section-layout and story-text
families exist only on the source-backed owner (table below); the
source-backed route admits one more corpus file than the owning route (59
against 58 edits, 52 against 51 changed publications;
`docs/unified-ops/0011-docx-mce-admission.md:130-145`); and the eager open
grows with the whole package (S1: `docx_file_eager_open` 11,498 µs against
`docx_file_source_open` 316 µs on the media corpus,
`docs/unified-ops/results/0003/before/README.md`, "open" rows). The price is
the rebase between runs: on a 10,000-paragraph text document §5.6 measured
9.4 ms per chained write on the source route, where the eager owner would
chain for about 4.3 ms per run with no publication (an eager open from bytes
costs only 176–282 µs on the three documents measured). Choosing the engine
per run, eager for multi-run main-story batches, is the follow-up once W7 has
measured media-heavy documents too (§12.3 Q13).

| Owner | Edit entry | Publication | Chaining |
|---|---|---|---|
| Source-backed (default) | `source_backed::Package::edit_document(&self)` (`D/source_backed.rs:1253`) | `publish_document_commit_to_stream(self, W, &Commit)` (`:1619`): exact no-op, topology or overlay door (`:1653-1671`) | none: a chained commit is `StaleSource` (`:1625-1630`); the handler rebases (§5.4) |
| Eager | `Package::edit_document(&self)` (`D/package/package/document.rs:36`) | `publish_document_commit(&mut self, Commit)` (`:227`), then `save`/`to_stream` | native on a forked package (L-DOCX-W1) |

Both entries refuse a package whose `MutableDocument` was modified, whose
properties are dirty, or that was decrypted (`ensure_opc_current`,
`D/package/codec.rs:706-743`); the unified layer never touches
`document_mut`, so only the last one can fire (§8.2). Budget-managed
snapshots accept only `replace_paragraph_text` and
`replace_body_paragraph_texts` (`ensure_unmanaged`,
`D/document/transaction.rs:5220-5229`); every other operation on a handle
opened with an execution context is refused with its leaf error until the
leaf admits it (listed per row below as "unmanaged only").

**Families.**

| Family | Leaf | Owner | Ops M2 uses |
|---|---|---|---|
| `Main` | `document::Edit` | both | the 24 operations of `D/document/transaction.rs:2281-4457`: paragraph, run, hyperlink, field-result, revision, content-control and table-cell text; `insert_paragraph` (plain); `insert_paragraph_transfer`; revision accept/reject |
| `Section` | `edit_section_layout(&self, selector)` (`D/source_backed.rs:1123`), ops page size, margins, start, columns (`D/section/layout.rs:255-316`) | source only | page setup keys of `/body/section[N]` |
| `Story` | `edit_story_text(&self, Selector::{Header, Footer, Footnote, Endnote, Comment, Glossary})` (`D/source_backed/story_text.rs:177`, `:1212`), op `replace_paragraph_text` (`:1007`) | source only | text of story paragraphs |
| `Variables` | `edit_document_variables(&self)` (`D/source_backed.rs:1285`); eager setters (`D/package/package/settings.rs:162-209`, no patch) | both | not in OfficeCLI's vocabulary; exposed only through the raw layer later |
| `Properties` | eager `props_mut`/`put_props`/`custom_props_mut` (`D/package/model.rs:911-986`), staged until save and with no patch (`D/package/codec.rs:1307-1310`); source-backed: none | — | core/custom keys of `/`, after L-DOCX-W6 |
| `Numbering` | `apply_numbering_patch` over raw `edit_opc` (`D/package/package/parts.rs:54-69`); ops only `restartNumberingAfterBreak` (`D/numbering/transaction.rs:126`, `:163`) | eager | numbering basics need L-DOCX-W5/W10 |

**Addressing: M1 path → leaf coordinate.**

| Path (M1, record 0010 §4.5) | Leaf coordinate | Rule |
|---|---|---|
| `/body/p[N]`, `/body/p[@paraId=…]` | `Position` = `IndexedBlock::transaction_position()` (`D/document/blocks.rs:612-634`): the raw direct-body ordinal | refused (`unsupported_capability`) when the paragraph sits in `w:customXml` (`:620`) or the part's raw and visible positions are not proven equal (`:628`, record 0011); the staged view of §5.3 then maps it through the run's inserts |
| `/body/p[N]/r[M]` | OfficeCLI's run set (`Paragraph::addressable_runs`, `D/paragraph/codec/addressable.rs:95-124`) | `replace_run_text` addresses the *direct* `w:r` children (`select_direct_child`, `D/document/transaction.rs:4944`); the handler maps run M to its direct ordinal when the run is a direct child and refuses runs inside `w:hyperlink` until L-DOCX-W2 (M2-Δ3) |
| `/body/p[N]/hyperlink[M]` | `ParagraphHyperlinkAddress` (paragraph + direct `w:hyperlink` ordinal) | `replace_hyperlink_text` (`:3454`) |
| `/body/tbl[T]/tr[R]/tc[C]` and `…/p[J]` | `(table, row, cell[, paragraph])` direct ordinals; nested tables as a `TableCellAddress` path (`:215-230`) | M1 refuses cell and content-control positions until the nested containers carry the part's verdict (`O/format/docx/node.rs:432-442`): L-DOCX-W3 |
| `/body/sdt[S]/p[J]` | `(&[Position], paragraph)` | `replace_block_content_control_paragraph_text` (`:4015`); same proof rule as cells |
| `/header[N]/p[J]`, `/footer…`, `/footnote…`, `/endnote…`, `/comments/comment[K]/p[J]` | `story_text::Selector` + paragraph | source owner only; a story that uses markup compatibility is "admitted for reading only; story text publication does not apply the document protection policy" (`D/source_backed/story_text.rs:834-842`), which is every story part of the Word fixtures (a test pins the refusal for each, `crates/litchi-docx/tests/story_documents.rs:458-494`): L-DOCX-W4 |
| `/body/section[N]`, `/section[N]` | `section::Selector` | source owner only |
| `/` | the package properties | L-DOCX-W6 |

**Fusion.** Paragraph text sets in one run go to one
`replace_body_paragraph_texts` call (`D/document/transaction.rs:2958`;
strictly increasing positions, so the handler sorts and keeps the last value
per paragraph); hyperlink text sets to `replace_body_hyperlink_texts`
(`:3544`). Each `Edit` operation otherwise rewrites the projected snapshot
(`insert_paragraph` replaces a byte range of the whole projected XML and
re-reads it, `:4390-4446`), so an unfused run of n operations costs n passes
over the part (§5.6 measures it).

**Verb × element (DOCX).** Class (a) wiring only, (b) small leaf API,
(c) new leaf engine; "—" the format or OfficeCLI has no such operation.

| Element | set | add | remove | move | swap |
|---|---|---|---|---|---|
| document `/` | metadata keys: (b) L-DOCX-W6 | (parent only) | — (OfficeCLI refuses, `OC/Handlers/Word/WordHandler.Mutations.cs:71`) | — | — |
| paragraph (body) | `text`: (a) `replace_paragraph_text`; `style`, `align`, spacing, indents, keep flags, `numId`/`numLevel`, `outlineLvl`: (c) L-DOCX-W5 | plain text: (a) `insert_paragraph` (`:4390`); with properties or runs: (c) L-DOCX-W5 | (b) L-DOCX-W5 (the operation and its inverse exist privately, `:1801`, `:2052-2058`, `:4829`); the narrow source family `remove_plain_paragraph` (`D/source_backed/paragraph_copy.rs:658`) refuses most real documents (`:44-62`) | (b) L-DOCX-W5 | (b) as two moves |
| paragraph (cell, content control) | `text`: (a) after L-DOCX-W3 | (c) | (c) | (c) | (c) |
| paragraph (story) | `text`: (a) after L-DOCX-W4 | (c) | (c) | (c) | (c) |
| run | `text`: (a) direct runs; hyperlinked runs (b) L-DOCX-W2; formatting keys (`bold`, `italic`, `size`, `font.*`, `color`, `underline`, `highlight`, `caps`, …, 39 readable keys, `O/format/docx/spec.rs:457-496`): (c) L-DOCX-W5 | (c) L-DOCX-W5 | (c) L-DOCX-W5 | (c) | (c) |
| hyperlink | `text`: (a); `url`, `anchor`, `tooltip`: (b) L-DOCX-W9 | (b) L-DOCX-W9 | (b) L-DOCX-W9 | — | — |
| table, row | properties: (c) L-DOCX-W8 | (c) L-DOCX-W8 | (c) L-DOCX-W8 | (c) | (c) |
| cell | `text`: (a) (`:4130`, `:4198`, `:4278`); properties: (c) L-DOCX-W8 | (c) | (c) | — | — |
| content control | `text`: (a) (`:3867-4099`); `alias`/`tag`: (b) | (c) | (c) | (c) | (c) |
| section | page size, margins, start, columns: (a) source owner; eager (b) L-DOCX-W1 | (c) | (c) | — | — |
| header, footer | `text`: (a) after W4; `type`: — | (c) | (c) | — | — |
| footnote, endnote | `text`: (a) after W4 | (c) | (c) | — | — |
| comment | `text`: (a) after W4; author, date: (b) L-DOCX-W7 | (b) L-DOCX-W7 (fixes the comments-part overwrite, §12.2 F-4) | (b) L-DOCX-W7 | — | — |
| revision | accept/reject: (a) `apply_revision` (`:3754`) | — | — | — | — |
| field (simple, complex result) | result text: (a) (`:3602`, `:3632`) | (c) | (c) | — | — |
| picture | SVG attach/detach: (a) source owner (`D/source_backed/svg_lifecycle.rs:1877`); replace: (b); geometry: (c) | (c) | (c) | (c) | (c) |
| style, numbering (abstractNum, num, level) | numbering restart: (a); definitions: (c) L-DOCX-W10 | (c) L-DOCX-W10 | (c) | — | — |
| bookmark, formfield, toc, watermark, textbox, shape, chart, equation, ole | (c) or M4 (charts, TOC refresh) | (c) | (c) | (c) | (c) |

Signed and protected inputs: the source-backed publisher refuses a changed
publication under enforced protection or tracked revisions
(`D/source_backed.rs:1650-1651`; record 0011), and the eager owner refuses a
changed signed package with `UnsafeEdit`
(`D/package/package/access.rs:240-246`), which the unified classifier maps
to `unsupported_capability` today (`O/error.rs:854-865`); M2 maps that
refusal to `document_protected` with the `Signed::unsign` suggestion (step
W2, finding F-6).

### 4.3 XLSX (tier 1)

**Owners and engine.** Every M2 write runs on the eager transaction,
`Workbook::edit(&self)` (`X/workbook/model.rs:1361-1364`), whose `Workbook`
is an `Arc` snapshot (`:426-431`) and whose commits chain
(`committed.workbook().edit()`, test `X/workbook/edit/tests/tabs.rs:74`). A
source-backed handle is promoted once per open through the exact source
(§4.1 case 3; L-XLSX-W1 = record 0009 S1, which also deletes
`Xlsx::into_editor(self)`'s `try_unwrap` and the `materialize` fallback,
`O/format/xlsx.rs:441-472`, `:504-506`). The 16 source-backed per-feature
editors are not used in M2: each publishes one family to a stream and
cannot chain with another commit (`X/cell_values/source.rs:537`, `:600`;
record 0009 §2.4), the value editor takes 256 cells of a one-worksheet
workbook and refuses range formulas (`X/cell_values/source.rs:24`;
`X/cell_values/snapshot.rs:560-563`; §5.6 hit the range-formula refusal on
119 of 362 numeric cells of `MatrixFormulaEvalTestData.xlsx`), and the
families M2 needs all exist on the eager `Edit`. Over 137 corpus workbooks
the probe of §5.6 found the source-backed value editor accepting a first edit
on 45 (79 refused as multi-sheet, 12 for worksheet relationships, 1 for the
shared-string count), while the eager `edit` → `commit` → `to_bytes` route
accepted 132 (3 protected sheets, one invalid dimension, one "owned OPC
source preservation is unavailable"). Data validation, conditional formats,
autofilter and sheet protection exist *only* as source-backed editors
(`X/data_validation/source.rs:115` and siblings); they are P1 (design §13)
and stay listed gaps.

**What the eager `Edit` offers** (`X/workbook/edit/semantic/transaction.rs`,
"T"; `X/workbook/edit/semantic/worksheet.rs`, "WS"):

| Group | Operations | Caps and refusals |
|---|---|---|
| sheets | `add`, `add_before`, `add_after` (T:544-577), `move_before`/`move_after`/`move_to` (T:632-666), `remove` (T:1206), `TabEdit::{rename, activate, show, hide, very_hide}` (WS:57-91) | 32,767 sheets (T:3344-3350); `remove` cannot share an `Edit` with any other change (`SheetRemoveBlocked{MixedEdit}`, T:1213-1220); selectors resolve against the *source* snapshot, so a sheet added or renamed in the same `Edit` is not found by its new name (`X/workbook/edit/semantic/transaction/snapshot.rs:8-10`) |
| cells | `WorksheetEdit::{set, clear, remove, style, reset_style, merge, unmerge}` (WS:297-430) | `set` takes `Content = Value | Formula` (`X/cell.rs:342-346`); array and data-table formulas and formulas with a cached value are refused (`:348-363`); text ≤ 32,767 characters (`:19`); at commit: protected sheet, a cell under data validation, a group formula, a covered merge cell, markup-compatibility content (`X/raw/worksheet/edit/validation.rs:121-160`); `remove` does not shift (WS:429-430); `style` takes an existing `Style` of the same lineage (`Error::ForeignStyle`, WS:397-403) |
| layout | `RowEdit`, `ColumnEdit`, `DefaultsEdit` (`X/workbook/edit/semantic/layout.rs:45-378`) | row height 0–409 pt (`X/row.rs:12-19`), column width 0–255 characters (`X/column.rs:12-19`) |
| hyperlinks | `put_hyperlink`, `remove_hyperlink` (T:922, :936) | 4,096 operations per worksheet (T:158) |
| names | `replace_defined_names`, `clear_defined_names` (T:517, :535): whole catalog | cannot combine with sheet add, move or rename (T:4238-4253) |
| cell transfer | `copy_cells`, `move_cells`, `copy_cells_from` (T:1015-1078) | 65,536 cells (T:67); merges, range formulas and drawings refused (T:1126-1174) |

Staging a cell is O(log n) (binary search for the sheet, then a `BTreeMap`
insert, T:3394-3396, WS:357-366) and commit rewrites only the staged
worksheets once each, clones the OPC graph and re-parses only the workbook
catalog (T:1806, :2178-2191, :3189-3271). §5.6 measured it.

**Addressing.**

| Path | Leaf coordinate |
|---|---|
| `/Sheet1/B2`, `Sheet1!B2` | the sheet by caseless name (the leaf's own key, `X/workbook/model.rs:893-911`) + `At::Cell` zero-based (`crates/litchi-sheet/src/lib.rs:333-338`, `:474-482`); a missing cell resolves (record 0010 D6), and `set` creates it, as OfficeCLI's does |
| `/Sheet1/A1:C3` | `Area` (`crates/litchi-sheet/src/lib.rs:542`) |
| `/Sheet1/row[5]`, `/Sheet1/col[B]` | `RowAt` (zero-based), `ColumnAt` |
| `/Sheet1`, `/sheet[N]` | the sheet's position; `TabEdit` |
| `/namedrange[NAME]` | the handler computes the new whole catalog from the base and the run's staged name items, then stages one `replace_defined_names` |
| `/` | workbook keys; metadata after L-XLSX-W4 |

Staged names (sheet add or rename in the same run) are tracked by the
handler: a later item addressing the new name uses the `NewSheet` handle the
`add` returned (WS:443-654) or forces a barrier (§5.3); L-XLSX-W5 removes the
barrier by letting `Edit::sheet`/`tab` see staged names.

**Verb × element (XLSX).**

| Element | set | add | remove | move | swap |
|---|---|---|---|---|---|
| workbook `/` | `activeTab`: (a) `TabEdit::activate`; metadata: (b) L-XLSX-W4 | sheet: (a) `Edit::add` | — | — | — |
| sheet | `name`: (a) `rename` (rewrites modeled references); `hidden`/`visibility`: (a); `tabColor`, panes, zoom: (b) L-XLSX-W6 (P1) | (a) | (a) `Edit::remove`, alone in its run | (a) `move_to`/`move_before`/`move_after` | (a) two moves in one `Edit` (inferred from the order plan, `X/workbook/edit/semantic/transaction/order.rs:6-75`; tested in W5b) |
| cell | `value`, `formula`, `type`: (a) `set`; clear: (a) `clear`; formatting keys (`font.*`, `fill*`, `border.*`, `alignment.*`, `numberformat`/`numFmtId`, `protection.*`, `quotePrefix`; `O/format/xlsx/spec.rs:327-377`): (b) L-XLSX-W2; `arrayformula`: (c) | (a) = `set` | (a) `remove` (no shift); OfficeCLI's `--shift`: (c) L-XLSX-W8 | (a) `move_cells` (≤ 65,536) | (b) |
| range | `value` (every cell), `merge`: (a); formatting: (b) L-XLSX-W2 | — | (a) per cell | (a) `move_cells` | — |
| row | `height`, `hidden`, `outlineLevel`, `collapsed`: (a) `RowEdit` | insert with shift: (c) L-XLSX-W8 | shift-delete: (c) L-XLSX-W8 | (c) | (c) |
| column | `width`, `hidden`, `outlineLevel`, `collapsed`: (a) `ColumnEdit` | (c) L-XLSX-W8 | (c) L-XLSX-W8 | (c) | — |
| named range | `ref`, `name`: (a) whole catalog | (a) | (a) | — | — |
| hyperlink (cell `link`) | (a) `put_hyperlink` | (a) | (a) `remove_hyperlink` | — | — |
| merge cell | (a) `merge`/`unmerge` | (a) | (a) | — | — |
| comment | (b) L-XLSX-W7 (the raw-OPC comments transaction, `X/workbook/comments/transaction.rs:29`, needs a `Workbook` route and VML for new notes, `:66-68`) | (b) | (b) | — | — |
| table, data validation, conditional format, autofilter, page breaks | (c)/P1: source-backed editors only, or none | (c) | (c) | — | — |
| chart, picture, shape, OLE, pivot, slicer, sparkline | (c), M4 or P2 | (c) | (c) | — | — |
| rich-text run | (b): `intern_rich` exists, not wired (`X/raw/shared_strings.rs:586`) | (b) | — | — | — |

**Errors.** The unified classifier maps only seven `litchi_xlsx::Error`
variants and sends the rest, among them `EncryptionPolicy`, the
`EditBlocked` family, `Coordinate`, `Range`, `ForeignStyle` and
`PatchConflict`, to `corrupt_file` (`O/error.rs:919-931`). M2 classifies
them (step W2): protection and encryption → `document_protected`;
coordinates and ranges → `invalid_path`/`out_of_range` (also S8d's finding,
`docs/unified-ops/m0a-log.md:247-248`); a covered merge or data-validation cell →
`invalid_argument` with the reason; `PatchConflict` → `internal_error` in a
batch (it cannot happen on a chain the executor built).

### 4.4 PPTX (tier 1)

**Owners and engine.** The general engine is the eager opened-presentation
transaction, `Package::opened_presentation_transaction(&self)`
(`P/package/model.rs:310`); the source-backed text editor takes the runs
described below. A transaction is `Clone` and clones the
`OpcPackage` it edits (`P/opened/transaction.rs:21-57`); a commit's snapshot
starts the next transaction without being applied, and two chained commits
apply in sequence (test `P/opened/slide_root_memo_tests.rs:1040-1045`;
application checks only the write set's before-states,
`P/opened/patch.rs:909`, `:973-990`). The handler therefore chains runs on
commit snapshots and applies the whole chain once, into one fork of the
package, at the end of the batch (L-PPTX-W1 = record 0009 S3: a `Clone` for
the eager `Package` that re-projects its digest memo, `P/package/model.rs:44`,
`:2231-2240`, plus `&self` fork-apply and `&self` `to_bytes`/`write_to`).
On a source-backed handle, a run of shape-text (and transition) items that
fits the source-backed slide editor runs there: one operation per
single-slide edit (`P/presentation/source.rs:1913`, guards `:2520-2525`,
`:2572-2577`), or `set_shape_texts` over at most 32 slides in one batch edit
(`:1942`, `:2949`, cap `:106`). §5.6 measured it 7–18× cheaper than the eager
transaction for text (397 µs against 7.0 ms for one shape; 4.1 ms against
about 31 ms for 256 shapes on 32 slides). Its commit stays pending (§4.1
case 2) and `save` publishes it (`publish_slide_commit_to_stream`, `:1956`;
batch `:2042`). Any other run, or a second run, promotes the open once
through the exact source to the eager engine. Building the source editor
still needs `Arc::try_unwrap` (`:1250-1293`); L-PPTX-W1 gives it a `&self`
constructor.

**Addressing.** The R4c adapter (`uo/m1-r4c-pptx`, `5e9137372c`) resolves
`/slide[N]` and per-kind shape counters over the direct children of each
shape tree, with `shape[@id=N]` (the `cNvPr` id) as the canonical step
(record 0010 D5). The transaction addresses slides by `slide::Selector::Position`
(zero-based, *staged* order) or by name (`P/slide/model.rs:18-23`;
`P/opened/transaction.rs:1472-1484`) and shapes by `shape::Key::Name` (exact,
ambiguity refused) or `Index` (depth-first pre-order position)
(`P/shape/reader.rs:50-53`, `:414-430`). No verb accepts the `cNvPr` id,
although every add verb returns it (`P/opened/transaction.rs:300-361`).

| Path | Leaf coordinate |
|---|---|
| `/slide[N]` | `slide::Selector::Position(Position::new(N-1))`: the staged order makes a moved or added slide addressable in the same run |
| `/slide[N]/shape[K]`, `…/shape[@id=I]`, `picture[K]`, `table[K]`, … | today: the pre-order index the scene gives for that record, recomputed when the run changed the slide's tree (a barrier); after L-PPTX-W2: `shape::Key::Id(I)`, stable across staged adds and removes |
| `/slide[N]/notes` | the slide's notes (existing notes only, `P/opened/transaction.rs:657`, refusal `:663-669`) |
| `/` | presentation keys; metadata after L-PPTX-W6 |

**Fusion.** Text sets on shapes of one slide go to one `set_shape_texts`
call (≤ 256 pairs, `P/opened/transaction.rs:270`; `P/opened/model.rs:16`).

**A text-set risk to settle first.** `set_shape_text` writes the text into the
first `a:t`, empties the others (`P/opened/xml.rs:852-973`), and refuses the
edit unless the re-read flattened text equals the request
(`P/opened/transaction.rs:246-252`). The reader inserts `\n` between further
paragraphs (`P/shape/reader.rs:1085-1091`) and does not trim it
(`:1576-1580`), so a multi-paragraph target may always be refused (read from
code; no test has one). L-PPTX-W7 adds the test and the fix before W5c relies
on it.

**Verb × element (PPTX).**

| Element | set | add | remove | move | swap |
|---|---|---|---|---|---|
| presentation `/` | metadata: (b) L-PPTX-W6 (custom properties exist, `P/package/model.rs:74-114`; core properties only through `edit_opc`); slide size: (b) | slide: (b/c) L-PPTX-W4 (`insert_slide_binding` is crate-private, `P/opened/xml.rs:577`, `:601`) | — | — | — |
| slide | `hidden`, `name`: (b) L-PPTX-W6; transition: (b) | (b/c) L-PPTX-W4; duplicate: (a) copy plan (`P/opened/copy_plan.rs:188`; no inverse, `:163`) | (a) `remove_slide`, dependency-free slides only (`P/opened/remove_plan.rs:270-282`) | (a) `move_slide` (`P/opened/transaction.rs:116`) | (a) two moves |
| shape, textbox, placeholder | `text`: (a) `set_shape_text(s)` (after L-PPTX-W7); `x`, `y`, `width`, `height`, `rotation`, flips, `name`, `alt`: (b) L-PPTX-W3 (DrawingML has a detached transform transaction, `DML/transform/transaction.rs:15-256`); `fill`, `line`: (b) L-PPTX-W3; text formatting and paragraph/run addressing: (c) L-PPTX-W5 (the text-body module has a reader and a fresh-body writer but no splicer, `DML/text/body/mod.rs:1-10`) | text box, rectangle, ellipse: (a) (`P/opened/transaction.rs:300-343`; fill only at creation) | (a) `remove_shape`, top-level only (`:439`) | z-order: (b); to another slide: (a) `transfer_shape` (`:493`) then `remove_shape` | (b) |
| picture | geometry: (b) L-PPTX-W3; replace: (b) | (a) `add_picture` (`:361`) | (a) | as shape | (b) |
| table | cell text: (b) L-PPTX-W6 (tables are read-only today, `P/table/shape.rs:63-210`); rows, columns: (c) | (a) `add_table` (`:738`) | (a) | as shape | (b) |
| chart | data and titles: (b) wire `litchi_drawingml::chart::edit` (`DML/chart/edit/transaction.rs:91-212`), M4 typed charts | (a) `add_chart` (`:776`) | (a) | as shape | (b) |
| paragraph, run | (c) L-PPTX-W5 | (c) | (c) | — | — |
| notes | `text`: (a) existing notes | creating notes: (b) L-PPTX-W6 | — | — | — |
| comment, modern comment, reply | (a) modern replace (`:958-1304`); legacy update: (b) (free functions only, `P/comments/package.rs:348-698`) | (a) `add_comment`, modern add | (a) | — | — |
| master, layout | — (P1) | layout: (a) `add_slide_layout` (`:844`) | unused layout: (a) | — | — |
| group, connector, SmartArt, OLE, media, animation, transition | (c) or P2 | (c) | (c) | — | — |

**Commit and apply.** `Commit { snapshot, patch }` has no diagnostics
(`P/opened/transaction.rs:2176-2225`). The handler applies commits only
through `apply_opened_presentation_commit` (`P/package/model.rs:564`), never
the bare-patch route, which has no signature check of its own (§8.1).

### 4.5 XLSB (tier 1, scheduled in M2b)

**Engine.** The eager `Workbook` (`B/workbook/model.rs:37`, not `Clone`);
the source owner has no edit API and is promoted once per open through the
exact source (`O/format/xlsb.rs:345-368` calls `to_workbook()` on every
`editor()` call today and loses the exact source, `B/workbook/source.rs:1076-1101`).

| Family | Entry | Ops | Publication |
|---|---|---|---|
| cell values | `Workbook::edit_cell_values(&self, sheet)` (`B/workbook/package.rs:393`, `_by_name` `:403`) | `set_value` keeps the storage family; typed setters; `set_style(StyleIndex)` (existing XF only); `insert`, `insert_formula`, `remove` (`B/cell_values/worksheet.rs:769-1006`); changing family = remove + insert (`:1039-1054`) | `apply_cell_values(&mut self)` (`B/workbook/package.rs:431-450`), which already computes a new workbook in the crate-private `apply_retaining_parse` (`B/cell_values/workbook.rs:121-136`): L-XLSB-W1 exposes it as a non-mutating apply (record 0009 S4) |
| workbook structure | `edit_workbook_structure(&self)` (`B/workbook/package.rs:338-349`) | `rename_sheet`, authored strings and formulas with `AuthoredStyle` (new cells only), images, transfers (`B/cell_values/root.rs:187-459`) | `WorkbookPatch::apply(&self, &mut Workbook)`; the commit carries no snapshot (`:461`, `:502-514`) |
| merged cells | `set_merged_cell_ranges`, `add_`, `remove_` (`B/workbook/codec/package.rs:236-337`) | through `edit_opc`, no patch | in place |

Save: `Package::save`/`save_with_durability` are atomic
(`B/package/mod.rs:529-550`); `Workbook::save<W: Write + Seek>` is
stream-only (`B/workbook/package.rs:1115-1119`), so the handler saves through
`Package::from(workbook.opc_package().clone())` until L-XLSB-W1 adds a
`Workbook` path save.

| Element | set | add | remove | move | swap |
|---|---|---|---|---|---|
| cell | value: (a) same family; family change: (a) remove + insert; existing XF: (a); new formatting: (b) L-XLSB-W2 (`AuthoredStyle` for existing cells, `B/cell_values/root.rs:34-46`) | (a) `insert` | (a) | (c) | (c) |
| sheet | `name`: (a) `rename_sheet`; hide: (c) | (c) | (c) | (c) | (c) |
| merge | (a) | (a) | (a) | — | — |
| workbook metadata | (b) L-XLSB-W2 (raw `edit_opc` + `properties::write` only) | — | — | — | — |
| hyperlinks, comments, tables, validation, formats, names | (c) (record 0002 XLSB §5) | (c) | (c) | — | — |

### 4.6 DOC (tier 2, M2b)

M0a's builder hands out the narrow source-backed editor
(`O/format/doc.rs:187-203`), which takes equal-UTF-16-length replacements
only and refuses whole files with fields, drawings, revisions, macros or an
ObjectPool (`W/body_text/source.rs:745-756`, `:1087-1125`). M2b's handler
uses the eager `body_text::Snapshot` over the retained bytes instead: its
main story may change length; other stories need equal length
(`W/body_text.rs:344-406`); character properties are bold, italic and
underline only (`:1683-1818`); commits chain (`:2339`). Each `edit()`
re-parses a full copy of the bytes (`W/body_text.rs:1354-1364`), so after
the first write the handle keeps the leaf snapshot (record 0009 S5).
Publication is `finish()`/`bytes()` only (`:859-865`, `:1351`), hence
L-OLE-W1. Paths: `/body/p[N]` maps to the paragraph projection's `Position`
after a corpus check that M1's DOC story blocks (L-DOC-R1) and
`body_text` count the same paragraphs (record 0002 DOC §5 asked for it).

| Element | set | add | remove | move | swap |
|---|---|---|---|---|---|
| paragraph (main) | `text`: (a) `replace_paragraph`/`replace_text`; `bold`: (a) `set_paragraph_bold` | (c) structural body editing | (c) | (c) | (c) |
| paragraph (other stories, table cell, field result) | `text`: (a), equal length outside the main story | (c) | (c) | — | — |
| run | `bold`, `italic`, `underline`: (a) `set_character_property` on uniform targets; others (c) | (c) | (c) | — | — |
| document `/` | metadata: (a) property-set transactions for existing SI/DSI streams; user-defined: (a) (`W/package/property_set.rs:300-375`) | — | — | — | — |
| revision | accept insert / reject delete: (a) `dispose_revision` | — | — | — | — |

### 4.7 XLS (tier 2, M2b)

The `cell_values::Snapshot` (`L/cell_values/mod.rs:549-655`): `set_value`
with its family rule, inserts into empty cells, styles by XF index or raw XF
bytes, row and column insert/delete, `rename_sheet`; at most 4,096 staged
changes (`:73`); commits carry diagnostics (`:3313`) and chain (`:3301`).
`Xls::cell_values_editor` copies and re-indexes the input on every call
(`O/format/xls.rs:246-273`), so after the first write the handle keeps the
snapshot (record 0009 S5). Ordinary commits publish bytes only (`:805`):
L-OLE-W1. Sheet visibility, comment text and property sets are separate
families with their own commits (`L/sheet_visibility.rs:298-373`;
`L/comments/transaction.rs:334-449`; `L/workbook/package/property_set.rs:140-233`).

| Element | set | add | remove | move | swap |
|---|---|---|---|---|---|
| cell | value: (a) within the family rule, else (b) `replace_cell`; style: (a) XF index; new formatting: (b) typed XF builder | (a) empty cells | (a) (formula cells refused) | (c) | (c) |
| row, column | (c) | (a) `insert_rows`/`insert_columns` | (a) `delete_*` (durable only for empty spans) | — | — |
| sheet | `name`: (a); `hidden`: (a) visibility family | (c) | chart sheets only: (a); worksheets (c) | reorder-all: (b) (bytes only, `L/chart/package.rs:868-905`) | (b) |
| comment | author, text: (a) | (c) | (c) (always refused) | — | — |
| workbook metadata | (a) property sets (creates missing sections) | — | — | — | — |

### 4.8 PPT (tier 2, M2b)

M0a's builder hands out the narrow `text_edit::SourceSnapshot`
(`O/format/ppt.rs:241-272`): equal encoded length only
(`T/text_edit.rs:612-621`). M2b's handler uses `slide_order::Snapshot`
(`T/slide_order.rs:587-606`): `move_slide`, `remove_slide` (payload kept as
unreachable history), hidden, advance, transition, `set_shape_text(s)` under
the eager text rule (one text atom, one paragraph run, one character run),
anchors, external media (`:1599-2056`); commits chain; publication is
`bytes()` only (`:762`), hence L-OLE-W1. A file with review history, a root
storage other than one `PP97_DUALSTORAGE`, or a signature, encryption or
DRM storage is refused (`:1051-1064`).

| Element | set | add | remove | move | swap |
|---|---|---|---|---|---|
| slide | `hidden`, transition: (a) | (c) | (a) | (a) `move_slide` | (a) two moves |
| shape | `text`: (a) within the eager text rule; anchor (`x`, `y`, `width`, `height`): (a) `set_shape_anchor`; formatting (c) | (c) | (c) | (c) | (c) |
| presentation metadata | (a) property sets (existing SI/DSI; the writer emits them as templates) | — | — | — | — |

### 4.9 Tier 3

Out of M2 (goal: tier 3 starts after tiers 1–2 have a working read+write
loop). `capabilities()` keeps reporting every edit verb on RTF, ODT, ODS and
ODP as `unsupported_capability` with follow-up M6, and the write module's
`match` refuses them in one arm.

## 5. The batch executor

One executor serves `set`/`add`/`remove`/`move_to`/`swap` (a batch of one),
`edit` (a batch fed by the closure) and `batch`/`apply` (OfficeCLI items).
It lives in `O/batch.rs` (a scaffold today, `O/batch.rs:1-4`), is
format-agnostic, and drives the per-format handlers of §4 through one
`match` per step (static dispatch).

### 5.1 Semantics

1. **Sequential.** Item *k* sees the effects of items *1…k−1*: its paths
   resolve against the document as those items left it. This is OfficeCLI's
   behaviour (items run one by one against the live document,
   `OC/CommandBuilder.Batch.cs:62-153`), and the only reading under which
   OfficeCLI's scripts and dumps replay: they add a paragraph, then address
   it as `/body/p[last()]`
   (`3rdparty/OfficeCLI/examples/word/run-formatting.sh:50-54`).
2. **Coalescing never changes a result.** Grouping items into one leaf
   transaction is an optimization the executor applies only where it cannot
   change what any item resolves to or what any item writes (§5.3). A
   coalesced batch publishes the same bytes as the same items applied one
   per transaction.
3. **Atomic by default.** `BatchMode::Atomic`: the first failing item ends
   the batch; the caller's handle is untouched, because nothing is mutated in
   place (record 0009 §5.5; ADR 0003, `docs/adr/0003-snapshots-edits-and-patches.md:23`, `:272`); the error carries the item's
   index. `BatchMode::BestEffort`: a failing item is reported and skipped,
   and the items before and after it apply; the result lists every item.
4. **No publication.** A batch returns a new handle; `save` publishes it
   (record 0009 §5.6).

### 5.2 Plan

For each item, in order:

| Step | Work | Allocation on success |
|---|---|---|
| parse | `PathArg` from text or a typed `Path` (record 0005: about 57 ns per short path, zero allocations) | none |
| resolve | the format adapter's `Navigate` on the *current* state (§5.3), to a typed target carrying the leaf coordinate (§4) | none on a warm index |
| classify | the handler's `(family, engine, effect)` for the target kind and verb (§4) | none |
| props | registry lookup, access check, coercion, constraint (§3.3) | none |
| stage | into the open run, or close the run and open the next (§5.3) | the leaf's own |

Errors at any step are attributed to the item (§5.5).

### 5.3 Runs, fusion and barriers

A **run** is a maximal sequence of consecutive items that the same leaf
transaction can take. The executor keeps one open run; an item joins it when
all four hold:

1. **Same family and engine**: the handler classifies the item into the run's
   `(family, engine)` (for example DOCX main-document `Edit` on the
   source-backed owner; XLSX eager `Workbook::edit`).
2. **Within the transaction's caps**: DOCX 4,096 operations per `Edit`
   (`MAX_OPERATIONS`, `D/document/transaction.rs:47`); PPTX source-backed
   editor one operation per single-slide edit and 32 slides per batch; PPTX
   `set_shape_texts` 256 pairs; XLSX hyperlinks 4,096 operations per
   worksheet; XLS 4,096 staged changes (§4). The executor prefers an engine
   without a cap for a run that would exceed one (record 0009 §5.3).
3. **No addressing conflict**: the item's path reads no container whose
   child numbering an earlier item of the run changed (an `add`, `remove`,
   `move` or `swap` there), and no name scope an earlier item changed (a
   sheet rename, a style or name addition), unless the handler resolves
   through the run's staged view (below).
4. **No read-after-write on content** the handler cannot see staged: e.g. a
   `set` whose value is computed from text an earlier item replaced (none in
   M2's verbs; listed for `find`/`replace` in M4).

Otherwise the executor **closes** the run (one leaf commit), **chains** to
the committed state, and opens a new run. Closing a run costs one commit;
chaining costs nothing on eager engines and OLE snapshots, which chain
natively (§5.4), and one *rebase* on an OOXML source-backed engine (publish
the pending commit into memory, re-open those bytes; record 0009 D6), until
the leaf supports overlay chaining (option D).

**Staged views.** Two handlers resolve through the run instead of closing it,
because their leaf transactions already number positions in the evolving
state:

- DOCX main-document `Edit` positions are *projected*: `insert_paragraph`
  inserts "at a projected zero-based position" and every operation reads and
  rewrites the projected snapshot (`D/document/transaction.rs:4381-4446`).
  The handler keeps the run's insert/remove list for `/body` and maps a
  visible `p[N]` through it, so "append a paragraph, then format
  `/body/p[last()]`" stays in one run.
- XLSX cell coordinates and sheet names: a cell set never renumbers
  anything; a sheet add/rename/remove changes only the name scope, which the
  handler tracks for the run.

Every other conflict closes the run.

**Fusion.** Where the leaf offers a batched form of an operation, the
handler hands it the whole run at commit instead of one call per item, with
the same result: DOCX paragraph text sets become one
`replace_body_paragraph_texts` call (strictly increasing positions; a later
set of the same paragraph supersedes an earlier one,
`D/document/transaction.rs:2946-2965`); PPTX shape text sets on one slide
become `set_shape_texts` (≤ 256 pairs, `P/opened/transaction.rs:270`;
`P/opened/model.rs:16`). Each item still gets its own result.

### 5.4 Chaining between runs, per engine

| Engine | Next run starts from | Cost |
|---|---|---|
| XLSX eager `Workbook::edit` | the commit's `Workbook` (`Edit::commit(self)` returns a new snapshot; `Workbook` is an `Arc` snapshot) | none |
| PPTX eager `opened::Transaction` | `commit.snapshot().edit()`; the commits are applied in order at the end into one fork of the package (record 0009 S3) | none per run; one fork at the end |
| DOCX eager `Package` | the forked package after `publish_document_commit` (S2's fork-apply) | one OPC graph clone per run (O(parts), parts shared) |
| DOCX source-backed `Edit` | rebase: `publish_document_commit_to_stream` into memory, re-open source-backed over `OwnedSource` | O(document) per boundary (§5.6) |
| PPTX source-backed editor | a second run promotes: publish the pending commit into memory and open the eager owner from those bytes | one promotion |
| XLSX source-backed editors | not used for batches in M2 (caps, single-worksheet rule; record 0009 §2.4) | — |
| OLE snapshots (M2b) | the commit's snapshot | none |

### 5.5 Errors and attribution

- **Planning and staging errors** (syntax, resolution, unknown key,
  coercion, a leaf refusal while staging) belong to the item being planned.
  Leaf staging errors leave the leaf transaction unchanged (DOCX: "without
  changing the projected snapshot", `D/document/transaction.rs:2276-2279`
  and every op's docs), so best-effort mode continues on the same run.
- **Commit errors** belong to the run. Atomic mode reports the first item of
  the run and the run's range (`items 12–40`); to name the item, the executor
  re-plays the run in halves on the error path only (bisection: O(r log r)
  re-stagings for a run of r items, never on the success path). Best-effort
  mode bisects too, commits the largest prefix that succeeds, reports the
  failing item, and continues after it.
- **The error type.** `Error` gains `item(): Option<u32>` and
  `property(): Option<&str>` (the box keeps it one pointer,
  `O/error.rs:435-452`); `BatchError` is not a separate type. Python maps
  them to `LitchiError.item` and `.property`.
- **OfficeCLI shape.** The command-dict channel and the M3 CLI render a
  failed atomic batch as OfficeCLI does (`success: false`,
  `atomicRolledBack: true`, per-item `index`, `success`, `code`, `error`;
  `OC/CommandBuilder.Batch.cs:675-681`, `OC/BatchTypes.cs:254-290`); the Rust
  API returns the typed error.

### 5.6 Cost model, with leaf numbers

A batch of n items that the executor splits into r runs costs

  T(n) = Σ_runs [ open + Σ_items (parse + resolve + coerce + stage) + commit ]
         + Σ_boundaries chain + wrap,

and the hand-written leaf leg doing the same work costs the same runs without
parse, resolve and coerce, and without the unified `wrap` (one `Arc` state).
The 2× batch budget therefore holds when, per item,
parse + resolve + coerce ≤ stage + (open + commit)/n, and when the
executor needs no more runs than the leaf leg (fusion and staged views make
that true for homogeneous batches; leaf caps split both legs alike).

**What the leaves cost.** A throwaway probe measured the leaf legs only
(release build with the workspace's `lto = true`, `panic = "abort"`; one
pinned CPU; median of 11–201 repetitions per point, each from the same base;
host load 2–8 on 32 CPUs). Tables, fixtures and the exact leaf calls:
`docs/unified-ops/results/0015/README.md`. Times are medians in µs.

*XLSX eager `Workbook::edit` → n × `WorksheetEdit::set` → `commit`*
(`X/workbook/model.rs:1361`; `X/workbook/edit/semantic/worksheet.rs:352`;
`X/workbook/edit/semantic/transaction.rs:1501`):

| Base | stage n=1 | stage n=1k | stage n=10k | commit n=1 | commit n=1k | commit n=10k | `to_bytes` n=1 |
|---|---:|---:|---:|---:|---:|---:|---:|
| `Workbook::create()` round-tripped (2 KB) | < 1 | 65 | 588 | 23 | 481 | 4,640 | 24 |
| generated 10k-cell sheet (30 KB) | < 1 | 65 | 593 | 1,090 | 1,820 | 8,310 | 6,780 |
| `test-data/ooxml/xlsx/MatrixFormulaEvalTestData.xlsx` (15 KB) | < 1 | 68 | 601 | 1,860 | 3,410 | 17,400 | 387 |
| `test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx` (47 KB) | < 1 | 68 | 600 | 6,710 | 7,730 | 15,900 | 2,540 |
| `test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx` (75,770 cells) | 1.7 | 81 | — | 53,500 | 54,500 | — | 41,500 |

Staging is about 0.06 µs per cell at every size; commit is a fixed cost that
grows with the touched worksheet (it rewrites and re-reads it, §4.3) plus
about 0.4–1.7 µs per changed cell. *Without* coalescing, n one-cell commits
chained one after another cost 113 ms (10 cells), 1.19 s (100 cells) and
12.6 s (1,000 cells) on the 47 KB fixture, 11.3–12.6 ms per item, because each
commit re-parses the sheet it starts from; 1,000 chained cells cost 2.28 s on
the generated 10k-cell sheet against 1.89 ms in one transaction (1,206×):
coalescing is what makes M2 usable on XLSX, not an optimization. Promotion: re-opening the
published bytes with `Workbook::from_bytes` costs 32 µs (2 KB), 150 µs
(15–30 KB), 290 µs (47 KB) and 3,280 µs (75,770 cells); the first commit on a
freshly opened workbook additionally pays the worksheet parse (for example
11.8 ms against 6.7 ms on the 47 KB fixture). The source-backed value editor
on the 15 KB fixture costs 3.0 ms for one cell (edit start 1.33 ms, commit
1.06 ms, publish 0.52 ms) against the eager route's 2.9 ms first commit plus
0.39 ms `to_bytes`, refused 119 of its 362 numeric cells (range formulas) and
stops at 256 cells; on the 75,770-cell sheet it costs 202 ms for one cell.
Hence D7: XLSX promotes and stays eager.

*DOCX `document::Edit`, n × `replace_paragraph_text` on a generated
10,000-paragraph document (37 KB zipped, 851 KB main part)*
(`D/source_backed.rs:1253`, `:1619`; `D/document/transaction.rs:2281`, `:5067`):

| Owner | open | edit start | stage n=1 | stage n=100 | stage n=1k | stage n=4,096 | commit n=1 | commit n=4,096 | publish |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| source-backed | 58 | 2,170 | 29 | 2,650 | 26,200 | 105,300 | 2,040 | 3,100 | 5,020–7,210 (`publish_document_commit_to_stream` into a `Vec`) |
| eager (`Package` from the bytes) | 210 | 2,110 | 30 | 2,650 | 26,200 | 104,900 | 2,090 | 5,000 | 41–56 (`publish_document_commit`) + 800–850 (`to_stream`) |

Each unfused operation costs about 26 µs here, one pass over the projected
part (§4.2); one `Edit` takes at most 4,096 operations (the 4,097th was
refused with `Limit { resource: "operations", max: 4096 }`), so a 10k-item
run is three runs on both legs. A rebase (publish into memory, re-open
source-backed, start the next edit) costs 9.4 ms per chained one-paragraph
write on this document (9,384 µs for 1, 93.9 ms for 10, 943 ms for 100
chained writes); source-backed open costs 35 µs here, 69 µs on
`test-data/ooxml/docx/drawing.docx` (95 KB) and 50 µs on
`testComment.docx` (65 KB), with edit starts of 2.25 ms, 1.12 ms and 29 µs.
The fused `replace_body_paragraph_texts` was not measured. Each unfused
operation copies and re-splices the whole main part (the probe traced it to
`D/document/transaction.rs:2332`, `:2338`), and the source-backed publisher
re-reads the main part first (`D/source_backed.rs:1625`), which is why its
publication (5–7 ms) costs more than the eager `publish_document_commit` plus
`to_stream` (0.85 ms) on this document.

*PPTX on a generated 200-slide × 8-text-box deck (230 KB)*
(`P/package/model.rs:310`; `P/opened/transaction.rs:217`, `:1390`;
`P/package/model.rs:564`; `P/presentation/source.rs:1913`, `:1942`):

| Route | open | edit start | stage | commit | apply | publish / `to_bytes` | total |
|---|---:|---:|---:|---:|---:|---:|---:|
| eager, 1 shape text | 2,250 | 4,820 (cold), 3,570 (warm) | 93 | 1,520 | 295 | 300 | 7,030 |
| eager, 100 texts on 100 slides | 2,260 | 4,780 | 6,940 | 4,820 | 1,300 | 2,740 | 20,600 |
| eager, 1,000 texts on 125 slides | 2,280 | 4,790 | 69,160 | 5,660 | 1,550 | 3,380 | 84,580 |
| source-backed editor, 1 shape text | 1,840 (open + editor) | 47 | 53 | 0.1 | — | 297 | 397 |
| source-backed batch, 64 texts on 8 slides | 1,830 | — | 700 | 0.9 | — | 554 | 1,254 (after open) |
| source-backed batch, 256 texts on 32 slides | 1,840 | — | 2,710 | 2.8 | — | 1,340 | 4,060 (after open) |

The source-backed slide editor takes one operation per single-slide edit
(the second was refused, "source-backed slide edits support one atomic shape
text operation") and 32 slides per batch (the 33rd was refused). For text
on a source handle it is 7–18× cheaper than the eager transaction, whose
edit start alone costs 3.6–4.8 ms and whose staging costs about 69 µs per
shape. Hence D7's PPTX rule: a run of shape-text (and transition) items that
fits the source editor's caps runs there; any other run promotes the open to
the eager engine once.

**Unified overhead per item.** Parse: about 57 ns for a typical short path,
zero allocations (record 0005); resolve: M1's warm `get` costs +0.27 µs over
the leaf for an XLSX cell and +0.29 µs for a DOCX paragraph (m1-log R4b,
R4a), and the write path skips the node view; registry lookup and coercion:
binary search plus an allocation-free parse (§3.3). So about 0.3–0.5 µs per
item, provisionally, against leaf per-item costs of 0.06 µs (XLSX staging
alone), 26 µs (DOCX) and 11–69 µs (PPTX).

**Against the budgets** (provisional; W7 measures the real thing):

| Scenario | Leaf leg (measured) | Unified overhead (estimated) | Ratio | Budget |
|---|---:|---:|---:|---|
| XLSX 1 cell, real fixtures | 1,860–53,500 µs commit | < 2 µs | ≈ 1.00 | ≤ 1.10 |
| XLSX 1k cells, `create()` base | 547 µs | 300–500 µs | 1.5–1.9 | ≤ 2: tight on near-empty workbooks |
| XLSX 1k cells, real fixtures | 3,480–54,600 µs | 300–500 µs | ≤ 1.15 | ≤ 2 |
| XLSX 10k cells | 5,230–18,000 µs | 3–5 ms | 1.2–1.9 | ≤ 2 |
| DOCX 1 paragraph (source) | 4,240 µs (edit start + stage + commit) | ~2 µs warm; + body-index build when cold (Risk 10) | ≈ 1.00 warm | ≤ 1.10 |
| DOCX 1k paragraphs | 30,700 µs unfused (source); less fused | 0.3–0.5 ms | ≤ 1.02 | ≤ 2 |
| PPTX 1 shape text (source) | 397 µs | ~2 µs | ≈ 1.01 | ≤ 1.10 |
| PPTX 1k shape texts | 84,580 µs eager (one transaction) | 0.5 ms | ≤ 1.01 | ≤ 2 |

The one tight case is a near-empty XLSX workbook, where the leaf does almost
nothing per item; W2 and W3 hold the unified per-item cost under 0.5 µs
with an allocation test, and W7 reports this row separately.

## 6. Leaf prerequisites, in order

Named like M1's leaf requests (`L-<FORMAT>-W<n>`; record 0010 §4.11 used
`R<n>` for reads). Each lands in its leaf crate with its own tests (goal
constraint 7: "prefer adding the API in the leaf crate"), on a branch of
`uo/m2`, before the handler step that needs it. Sizes as in record 0010
§10.2: 1 = one implementer for one change.

### 6.1 Sharing (record 0009 S1–S5) — first, they gate every handler

| Id | 0009 | Change | Leaf / ops | Size | Gates |
|---|---|---|---|---|---|
| L-OPC-W0 | S0 | `SourceBackedPackage: Clone` | done (record 0012, `OPC/source_backed.rs:6531`) | — | — |
| L-XLSX-W1 | S1 | `Xlsx::editor(&self)` replaces `into_editor(self)`; delete the `try_unwrap` branch and `materialize` (`O/format/xlsx.rs:441-472`, `:504-506`); promote through the exact source once per open (memoized), also for `open_prepared` under `SourceMode::Eager`, which calls `materialize` at open today (`:233-234`) | ops only | 1 | W5b |
| L-DOCX-W1 | S2 | publishers M2 uses take `&self` (main `D/source_backed.rs:1619`, section `:1386`, story `D/source_backed/story_text.rs:1225`; the other 29 in a sweep); eager `Package` gains a fork-apply returning a new `Package` (`apply_document_patch` already forks inside `edit_opc_candidate`) and a `&self` save of a clean package; `Docx` builds and publishes from `&self` | leaf + ops | 1–2 | W5a |
| L-PPTX-W1 | S3 | `SourceBackedPresentation::editor(&self)` (today `try_into_editor(self)` needs `Arc::try_unwrap`, `P/presentation/source.rs:1250-1293`); eager `Package: Clone` with memo re-projection (`P/package/model.rs:2231-2240`); `&self` fork-apply and `to_bytes`/`write_to`; delete `EditorRefused` (`O/format/pptx.rs:169-181`) | leaf + ops | 2 | W5c |
| L-XLSB-W1 | S4 | public non-mutating cell-values apply (expose `apply_retaining_parse`, `B/cell_values/workbook.rs:121-136`) and the same for each in-place family as it is wired; exact-source promotion once per open; a `Workbook` path save | leaf + ops | 1–2 | W9a (M2b) |
| L-OLE-W0 | S5 | after the first write a handle holds the leaf edit snapshot (DOC eager `body_text::Snapshot`, XLS `cell_values::Snapshot`, PPT `slide_order::Snapshot`) instead of rebuilding it per write | ops only | 1 | W9b–d (M2b) |

### 6.2 Save

| Id | Change | Size | Gates |
|---|---|---|---|
| L-OLE-W1 | `litchi-cfb` exports its sibling-temporary publisher for caller-staged bytes (`publish_staged` is `pub(crate)`, `crates/litchi-cfb/src/writer/publish.rs:93-112`), and DOC `body_text::Snapshot`, XLS `cell_values::Snapshot`/`Commit` and PPT `slide_order::Snapshot` gain `save_with_durability` over it (record 0003 R10; §7.2, D13) | 1 | W4 (legacy saves), W9 |
| L-OOXML-W1 | the encrypted saves take a `Durability` (they are `Full` only: `crates/litchi-docx/src/package/codec.rs:673-683`; `X/writer.rs:47-53`; `P/package/encryption.rs:360-365`) | 1 | `save_with` + re-encryption |

### 6.3 Addressing and admission

| Id | Change | Size | Gates |
|---|---|---|---|
| L-DOCX-W2 | `document::Edit` run operations addressed by OfficeCLI's run set (`addressable_runs`), or an owner path for each addressable run, so runs inside `w:hyperlink` are editable (`replace_run_text` reaches direct `w:r` only, `D/document/transaction.rs:3571`, `:4944`) | 1 | run rows of §4.2 |
| L-DOCX-W3 | nested containers (cells, block content controls) inherit the part's MCE verdict, so `transaction_position` answers for their blocks (M1 left them refused, `O/format/docx/node.rs:432-442`) | 1 | cell and control rows |
| L-DOCX-W4 | record 0011's admission plus the document protection policy for story publication, so `edit_story_text` accepts the Word fixtures' headers, footers, notes and comments (today a markup-compatibility story is read-only because story publication skips the protection policy, `D/source_backed/story_text.rs:834-842`; pinned for every fixture story by `crates/litchi-docx/tests/story_documents.rs:458-494`) | 1 | story rows |
| L-PPTX-W2 | `shape::Key::Id(u32)`: address a shape by its `cNvPr` id (`P/shape/reader.rs:50-53`) | 1 | stable shape paths across staged edits |
| L-PPTX-W7 | test and fix the multi-paragraph `set_shape_text` read-back (§4.4) | 1 | shape `text` |
| L-XLSX-W5 | `Edit::sheet`/`tab` resolve sheets added or renamed in the same `Edit` (`X/workbook/edit/semantic/transaction/snapshot.rs:8-10`) | 1 | removes a barrier (optimization) |
| L-DOCX-MCE4 | record 0008 stage 4: admit run-level `mc:AlternateContent` with a post-commit projection check (`docs/unified-ops/0008-docx-source-backed-mce-edits.md:487-497`; not implemented, `docs/unified-ops/0011-docx-mce-admission.md:414`, `:427`) | 2 | **not required for M2**: one corpus file of 61 needs it (0011's counts: 59 admitted for source-backed edits); scheduled after M2 unless the formatting engine's tests show run-level MCE in the OfficeCLI corpus outputs |

### 6.4 Formatting and structure (the parity engines)

The OfficeCLI example corpus is almost entirely `create` + `add` + `set` with
formatting properties (about 2,430 `add` and 1,180 `set` calls against 11
`remove` calls and no `move` or `swap`, §1.3; e.g. every Word run-formatting
line adds a paragraph with `bold`, `size`, `color`,
`3rdparty/OfficeCLI/examples/word/run-formatting.sh:21-54`, and every Excel
cell line sets `value` with `font.*` and `fill`,
`3rdparty/OfficeCLI/examples/excel/cell-formatting.sh:33-41`). None of the
three tier-1 leaves can format existing content today (§1.2). These are the
largest items and the critical path of M2's parity:

| Id | Change | Class | Size | Gates |
|---|---|---|---|---|
| L-DOCX-W5 | the DOCX formatting engine in `document::Edit`: set run properties and paragraph properties on existing content (a splice of `w:rPr`/`w:pPr` that keeps unknown children and change-tracking elements, driven by the typed models of `D/properties.rs` that L-DOCX-R4 added for reading); insert a paragraph with properties and runs; append or insert a run; public `remove_paragraph` (the operation and its inverse exist, `D/document/transaction.rs:1801`, `:2052-2058`, `:4829`); `move_paragraph`. Covers style application (`pStyle`/`rStyle`) and numbering application (`numPr`) | (c) | 3 | §4.2 formatting rows; P0 style and numbering |
| L-XLSX-W2 | XLSX style allocation in `Edit`: `WorksheetEdit::format(At, FormatPatch)` (and rows, columns) interning font, fill, border, number format, alignment and protection into the stylesheet with the lineage kept; the stylesheet writer exists and is unused (`X/style/stylesheet/writer.rs:25`, `:35`) | (b) | 2 | cell formatting rows; P0 style application |
| L-PPTX-W3 | shape geometry, name, alt text, fill and line on existing shapes in `opened::Transaction` (wire `DML/transform/transaction.rs:15-256`) | (b) | 2 | shape property rows |
| L-PPTX-W5 | a DrawingML text-body splicer: paragraph and run addressing and run/paragraph properties on existing `txBody`s | (c) | 3 | paragraph/run rows; deferred within M2 (§12.3 Q8) |
| L-PPTX-W4 | add a blank or layout-based slide to an opened deck | (b/c) | 2 | slide `add` |
| L-DOCX-W6 / L-XLSX-W4 / L-PPTX-W6 | typed core, app and custom properties with patches on each owner (DOCX source-backed overlay publication of `/docProps/core.xml`; XLSX through `Workbook::edit`; PPTX typed wrappers over `properties::write`, `crates/litchi-ooxml-common/src/properties/write.rs:38`), plus PPTX slide hide/rename, notes creation and table-cell text | (b) | 1 each | P0 metadata |
| L-DOCX-W10 | numbering definitions on opened documents: add a list definition (`abstractNum` + `num`) for bullets and decimals | (b) | 1 | P0 numbering basics |
| L-DOCX-W7 | comments add/remove on opened documents with id allocation, fixing the comments-part overwrite (F-4) | (b) | 1 | after the P0 core (§11) |
| L-DOCX-W8, L-DOCX-W9, L-XLSX-W6, L-XLSX-W7, L-XLSX-W8 | table structure (DOCX), hyperlink targets (DOCX), sheet views, comments (XLSX), row/column shift (XLSX) | (b)/(c) | 1–2 each | listed gaps; P1 or M2-extended |

### 6.5 Legacy creators (M2b)

| Id | Change | Size |
|---|---|---|
| L-DOC-W1 | `Writer::set_property` writes a SummaryInformation stream (today a silent no-op, `W/writer/core/model/semantic.rs:1046-1048`), or the SI/DSI property-set transactions create a missing stream (today `StreamNotFound`, `W/package/property_set.rs:300-375`) | 1 |
| L-PPT-W1 | `Writer::set_property` serializes into the SI/DSI streams (today dropped, `T/writer/core/model/semantic.rs:1764-1766`; the streams are fixed templates, `T/writer/core/package.rs:129-160`) | 1 |
| L-XLS-W1 | the `Writer` writes property streams (none today, `L/writer/core/package.rs:93-131`); chaining the property-set transaction after the writer already works for XLS, which creates missing sections (`L/workbook/package/property_set.rs:140-233`) | 1 |

### 6.6 Encryption (P1, listed so M2 can refuse with a follow-up)

| Id | Change |
|---|---|
| L-OOXML-W2 | edit a decrypted OOXML package in memory and publish it encrypted with the captured profile (the leaves refuse every edit of a decrypted package, `crates/litchi-docx/src/package/codec.rs:707-714`, and their `to_reencrypted` re-encrypts only the package as decrypted) |

## 7. Save and create

### 7.1 Contract

| Verb | Writes | Receiver | Notes |
|---|---|---|---|
| `save()` | the handle's origin path | `&self` | `file_required` when the handle has no origin (opened from bytes, created) |
| `save_as(path)` | `path` | `&self` | the handle's origin is unchanged: handles are immutable |
| `save_with(SaveTo, &SaveOptions)` | origin or a path | `&self` | `SaveOptions { durability: Durability, overwrite: Overwrite, reencrypt: Option<Password> }` |
| `write_to(sink)` | any `io::Write` | `&self` | not atomic; returns the byte count; reports partial output as the leaf does (ADR 0005, `docs/adr/0005-io-memory-and-performance.md:35-37`) |
| `to_bytes()` | a `Vec<u8>` | `&self` | `write_to` into a vector sized from the source when known |

Every form may be called any number of times on any handle (record 0009
§7). A save never changes what the handle reads, and never repairs (ADR 0006,
`docs/adr/0006-validation-security-and-compatibility.md:62`).

**Same format only.** `save_as("x.docx")` on an XLSX handle writes XLSX
bytes; a path extension never selects or changes a format (ADR 0006, `docs/adr/0006-validation-security-and-compatibility.md:73-74`:
"Format and flavor are content-derived, never inferred from a filename").
M2 refuses a `save_as` whose extension names another *family* with
`invalid_argument` and the suggestion "conversion is explicit: use
`export` (not in M2)", and accepts flavour extensions of the same format
(`.docm`, `.dotx` for DOCX; flavour promotion stays explicit per ADR 0006,
`docs/adr/0006-validation-security-and-compatibility.md:122-126`). Cross-format export is design §9's future `export(target)` with a
fidelity report.

### 7.2 What a save does per owner state

`save` is the leaf save: the bytes are what the leaf writes for the handle's
current state, through the leaf's own sibling-temporary replacement, with no
copy added by the unified layer (goal, PERFORMANCE BAR "Save").

| Handle state | Publication | Atomic path writer |
|---|---|---|
| Source-backed OOXML, no pending commit (unchanged) | the source artifact, byte for byte (`source_artifact()`, `OPC/source_backed.rs:7765`; `write_to_stream`, `:3832`) | `litchi_opc::atomic::replace_with_durability` (`OPC/atomic.rs:132`) |
| Source-backed DOCX or PPTX with a pending commit (§4.1) | `publish_document_commit_to_stream` and the family's other stream publishers (PPTX: `publish_slide_commit_to_stream`, `P/presentation/source.rs:1956`): untouched members raw-copied | `litchi_opc::atomic::replace_with_durability` around the stream publisher, as change 0497 already does for tail append (`crates/litchi-docx/src/source_backed/tail_append_stream.rs:2282-2299`) |
| Eager OOXML (xlsx `Workbook`, pptx/docx/xlsb `Package`) | the leaf's `save_with_durability` / `write_to`: exact-source chunk copy for an exact no-op and raw copy of untouched members when the package retains its owned source (`OPC/pkgwriter.rs:1257-1277`) | the leaf's own (`OPC/pkgwriter.rs:1215-1223`) |
| Legacy edit snapshots with a path save (doc source commit, xls source commits, ppt text edit) | the leaf commit's `save_with_durability` | the leaf's own, through CFB `publish_staged` (`crates/litchi-cfb/src/writer/publish.rs:93-112`) |
| Legacy bytes-only owners (doc eager `body_text::Snapshot::finish`, xls `cell_values::Snapshot::bytes`, ppt `slide_order::Snapshot::bytes`) | the leaf's bytes | **new:** a public CFB entry over `publish_staged` (L-OLE-W1, §6.2) |

**The legacy atomic writer (L-OLE-W1).** Record 0003 R10 asked for "a shared facade
atomic writer in `litchi-ops`". Two complete sibling-temporary
implementations already exist, one per container family:
`litchi_opc::atomic` (public; `OPC/atomic.rs:88-140`, with symlink refusal,
permission preservation and `Committed`) and CFB's `publish_staged`
(`pub(crate)`; `crates/litchi-cfb/src/writer/publish.rs:1-16`, `:93-112`).
A third copy in `litchi-ops` would split ADR 0005's one replacement rule
(`docs/adr/0005-io-memory-and-performance.md:32-34`, `:2438-2456`) three
ways. Decision D13: the facade writer is a policy layer in `litchi-ops`
(`O/save.rs`: origin path, durability, overwrite policy, source-changed
checks, error mapping, re-encryption) over the two leaf helpers; L-OLE-W1 makes
CFB's helper callable for caller-staged bytes (`litchi_cfb::publish_bytes`
or `save_bytes_with_durability`, a small leaf API, class (b)) and gives the
bytes-only legacy snapshots a `save_with_durability` that uses it, so a
legacy save is a leaf save too. The unified layer writes no rename or fsync
code of its own.

**Durability.** `litchi_core::Durability` (`crates/litchi-core/src/durability.rs:67-83`)
is the only knob: `save` and `save_as` use `Full`, as every leaf `save`
does; `save_with` passes the caller's level through. Encrypted publication
is `Full` only in the leaves today (`crates/litchi-docx/src/package/codec.rs:673-683`;
`crates/litchi-pptx/src/package/encryption.rs:360-365`), so `save_with` with
a weaker level and re-encryption is refused with `invalid_argument` until
the leaves take the level (follow-up).

**Overwrite policy.** `save()` replaces its origin; `save_as` replaces an
existing destination unless `Overwrite::Refuse` is set (`file_exists`).
Symbolic-link and non-file destinations are refused by the leaf helpers
(`OPC/atomic.rs:259-271`).

**Saving over the file other handles read.** On Unix a path source pins its
open descriptor (`crates/litchi-core/src/source/file.rs:59-61`, version
`:348-357`), so after `h1.save()` over the origin every other handle keeps
reading the bytes it opened (record 0009 T6; core test
`crates/litchi-core/tests/file_source.rs:108-135`; DOCX precedent
`crates/litchi-docx/tests/source_backed_tail_append_stream.rs:2120-2154`).
Windows replacement over an open source is not claimed (record 0009 §9;
change 0497, `docs/performance/changes/0497-docx-atomic-publication.md:42-43`).

**Origin path.** No handle records its path today (grep: no `PathBuf` or
`origin` field in `O/format/*.rs` or `O/open*.rs`). M2 adds
`origin: Option<Arc<std::path::Path>>` to the shared state each path opener
builds (one `Arc` per open, cloned into every handle a write derives), set by
`open`/`open_with` and unset by `open_bytes*` and `create`.

### 7.3 `create(format)`

`create` builds the leaf's blank document, serializes it once in memory and
opens those bytes through the format's normal constructor, so a created file
is an ordinary handle over an `OwnedSource`: it reads, edits and saves
exactly like an opened one, its exact no-op writes the created bytes, and the
pptx "save and reopen before an opened transaction" rule
(`P/package/model.rs:282-287`) never surfaces. The extra round trip is small
for a blank document and is measured (route `*_unified_create`).

| Format | Builder | Notes |
|---|---|---|
| docx | `litchi_docx::Package::new()` (template: `new_template()`) | the blank main part has no paragraph and declares `mc:Ignorable="w14 wp14"`, which record 0011's admission accepts (`D/resources/generated/document.xml`) |
| xlsx | `litchi_xlsx::Workbook::create()` | one empty sheet |
| pptx | `litchi_pptx::Package::new()` → `to_bytes` | the opened-transaction refusal for authored packages (`P/package/model.rs:282-287`) is why the reopen is needed |
| xlsb | `WorkbookWriter` / `Package::create()` | document properties are a fixed template (record 0002) |
| doc, xls, ppt | the leaves' `Writer`s | metadata is dropped by doc and ppt `Writer::set_property` and absent from xls `Writer` output (record 0002 finding 7); fixed by L-DOC-W1, L-PPT-W1 and L-XLS-W1 (§6.5) |

`CreateOptions` carries the flavour (template vs document), the OOXML
conformance (Transitional default, ADR 0006, `docs/adr/0006-validation-security-and-compatibility.md:71-72`) and optional initial
metadata. Deterministic output (ADR 0006, `docs/adr/0006-validation-security-and-compatibility.md:116-120`): `create` never reads the
clock; `created`/`modified` stay unset unless the caller passes them.

## 8. Signatures and encryption at write time

### 8.1 Signatures: when `Signed::unsign` takes effect (record 0013 F5)

Record 0013 left one question: a file `Signed::unsign` returned and never
edited still holds the signed source, because the removal happens when its
first editor is built (`O/format/signature.rs:1-13`); what does its `save`
write? Two readings:

- **(a) the source bytes.** Nothing was edited, so the exact no-op of ADR
  0003 (`docs/adr/0003-snapshots-edits-and-patches.md:136-137`) writes the
  input, signature included.
- **(b) the unsigned copy.** `unsign` is the explicit, consuming "remove the
  signature" of ADR 0006 (`docs/adr/0006-validation-security-and-compatibility.md:136-144`),
  and the handle already reports `is_signed() == false`
  (`O/file.rs:235-281`).

**Decision D14: (b).** The removal is a change the caller asked for; the
handle's state is "unsigned", and a save must publish that state, or a
re-open would return `File::Signed` for a handle that said it was not
signed. The exact no-op is measured against the handle's state, not the
input's. The removal is still lazy and fallible where it was: it happens at
the first write or at the first save or `write_to`, whichever comes first,
through the same leaf entries record 0013 §1 lists (OOXML `unsign` on a
package opened over the exact input bytes; `litchi_sign::cfb::Editor::clear`
for the compound files), and `unsign()` itself stays O(1) and infallible
(`O/signed.rs:164`). The removal's output is memoized on the handle state
(one copy per open, not per save). Edits after the removal publish no
signature parts (record 0013 §4's audit test, extended to the save path).

Write verbs on a still-signed file keep record 0013's refusal:
`document_protected`, suggestion naming `Signed::unsign()`; the M2 verbs
reach it through the existing builders, so no new signed check is written.
Two leaf gaps the write kernel must not route through (record 0013 F1;
record 0014a's follow-ups F1-late and F1-semantic): pptx stores that
fail only at save (`store_presentation_comments`, `store_modern_comment`,
`media_parts::store`, `slide_sync::store`) and the bare-patch route
`apply_opened_presentation_patch` (`P/package/model.rs:582`), which has no
signature check of its own (none in `P/opened/patch.rs`, grep; the signature
tests cover only the commit route, `crates/litchi-pptx/tests/opened_signature_policy.rs:39-101`).
M2's PPTX handler applies commits
only through `apply_opened_presentation_commit` (`P/package/model.rs:564`),
whose commit refuses a changed signed package (`P/opened/transaction.rs:1403-1405`).

`resign` stays P1 (record 0013 §1, F6): it needs a credential type the
workspace does not have (no public credential type was found; OPC `sign`
and `resign` are crate-private, `crates/litchi-opc/src/sign.rs:601`, `:629`).

### 8.2 Encryption: re-encrypt on save, or refuse

ADR 0006 is explicit: "Encryption uses consuming `Locked<T> -> Sensitive<T>`
transitions. Sensitive documents save encrypted unless explicitly consumed
into plaintext. Credentials are non-clone, redacted, zeroized, and
non-serializable" (`docs/adr/0006-validation-security-and-compatibility.md:133-135`),
and "Weak legacy encryption is decode-only by default" (`:147-148`). Today:

- `Locked::unlock(self, Password)` decrypts once and drops the password
  (`O/locked.rs:242-254`); the unlocked eager owner keeps the encryption
  profile (`O/locked.rs:214-217`).
- The OOXML leaves refuse to edit or plain-save a decrypted package and
  offer `to_reencrypted`/`save_reencrypted(password: &str)` with the
  preserved profile and `to_encrypted`/`save_encrypted` with an explicit mode
  (`crates/litchi-docx/src/package/codec.rs:640-690`, plain-save refusal
  `:745-753`; `crates/litchi-xlsx/src/workbook/model.rs:1211-1350`;
  `P/package/encryption.rs:157-306`). They publish at `Full` only.
- The legacy leaves refuse edits on encrypted inputs (record 0002 DOC, XLS,
  PPT §6).

**Decision D15 (M2 policy):**

1. An unlocked file is *sensitive*: `File::is_sensitive()` reports it, and
   `capabilities()` states it.
2. **Saving re-encrypts or refuses.** `save*`/`write_to` of a sensitive OOXML
   handle requires `SaveOptions::reencrypt(Password)` and publishes through
   the leaf's `to_reencrypted` with the preserved profile, atomically
   (`litchi_opc::atomic`); without a password it refuses with
   `document_protected` and the suggestion "pass the password to re-encrypt,
   or consume the file with `File::into_plaintext()`". The unified layer
   never writes a sensitive document in the clear on its own, and never
   spills plaintext to a temporary file (ADR 0005, `docs/adr/0005-io-memory-and-performance.md:25-26`): the ciphertext is
   what is staged.
3. **Plaintext is explicit.** `File::into_plaintext(self) -> File` consumes
   the sensitive state; the result edits and saves like any plain file.
4. **Edits of a sensitive file are P1.** M2's write verbs refuse a sensitive
   handle with `document_protected` and the same two suggestions. Editing in
   memory and re-encrypting at save needs a plaintext working copy that the
   leaves refuse to make from a decrypted package, plus encrypting a
   *changed* package with the captured profile; both are listed as P1 leaf
   work (§6.6, L-OOXML-W2).
5. **Legacy encryption** (DOC, XLS, PPT): saving a sensitive legacy handle
   refuses with `unsupported_capability` ("weak legacy encryption is
   decode-only", ADR 0006, `docs/adr/0006-validation-security-and-compatibility.md:147-148`) unless `into_plaintext` was called.
6. **No password retention.** The unified layer does not keep the
   password after `unlock`, as today; the caller supplies it again to
   re-encrypt. Retaining it in the sensitive state is open question Q4.

## 9. The P0 structural surface of M2

Design §15 puts four structural items in M2: worksheet lifecycle, metadata,
style application and numbering basics (design §13 P0). Each maps as
follows; "M2b" rows are the tier-2 and XLSB counterparts.

### 9.1 Worksheet lifecycle (add, remove, rename, move, hide)

| Format | add | remove | rename | move | hide | Notes |
|---|---|---|---|---|---|---|
| XLSX | (a) `Edit::add`/`add_before`/`add_after` (`X/workbook/edit/semantic/transaction.rs:544-577`) | (a) `Edit::remove` (`:1206`), alone in its run (`:1213-1220`) | (a) `TabEdit::rename` (`X/workbook/edit/semantic/worksheet.rs:57`), rewrites modeled references | (a) `move_to`/`move_before`/`move_after` (`X/workbook/edit/semantic/transaction.rs:632-666`) | (a) `show`/`hide`/`very_hide` (`X/workbook/edit/semantic/worksheet.rs:69-91`); refuses hiding the last visible tab (`X/error.rs:306-315`) | all in M2 core |
| XLSB (M2b) | (c) | (c) | (a) `rename_sheet` (`B/cell_values/root.rs:187`) | (c) | (c) | record 0002 XLSB (c): catalog + XTI/name remap |
| XLS (M2b) | (c) | chart sheets only (a) | (a) `rename_sheet` (`L/cell_values/mod.rs:2398-2450` region) | reorder-all, bytes only (b) | (a) visibility family (`L/sheet_visibility.rs:298-373`) | |

OfficeCLI parity notes: OfficeCLI refuses removing the last sheet
(`OC/Handlers/Excel/ExcelHandler.Remove.cs:113`) and remaps sheet-scoped
names on move (`OC/Handlers/Excel/ExcelHandler.Add.cs:179-269`); the leaf
refuses the first case itself and the M2 tests cover the second.

### 9.2 Document metadata (core, app, custom properties)

The node view already reads ten core keys (`title`, `author`, `subject`,
`keywords`, `description`, `category`, `lastModifiedBy`, `revisionNumber`,
`created`, `modified`) on DOCX and XLSX roots (`O/format/docx/spec.rs:409-420`;
`O/format/xlsx/spec.rs:300-317`). M2 makes them writable:

| Format | Today | M2 route |
|---|---|---|
| DOCX | eager `props_mut`/`put_props`/`custom_props_mut`, staged until save, no patch, and they block `edit_document` while dirty (`D/package/codec.rs:727-740`); source-backed: read-only (`D/source_backed.rs:1165`) | L-DOCX-W6: a properties family with a patch on both owners (source-backed: one-part overlay of `/docProps/core.xml` and `custom.xml`) |
| XLSX | custom: `Package::put_custom_props` (`X/package.rs:664`); core: raw `litchi_ooxml_common::properties::write` only; no `Workbook ⇄ Package` route but bytes (`X/workbook/model.rs:383`) | L-XLSX-W4: core and custom properties on `Workbook::edit` |
| PPTX | custom: `put_custom_props` (`P/package/model.rs:95`); core: `edit_opc` + `properties::write` | L-PPTX-W6: typed core properties on the package, applied in the fork of §4.4 |
| DOC, PPT (M2b) | property-set transactions on existing SI/DSI streams; user-defined creates the section | (a); created DOC files need L-DOC-W1 |
| XLS (M2b) | property-set transactions create missing sections | (a) |
| XLSB (M2b) | raw `edit_opc` only | L-XLSB-W2 |

`app.xml` (extended properties) is template-only in DOCX
(`D/package/codec.rs:177-186`) and has no typed writer in any OOXML leaf;
M2 exposes it read-only and lists the writable keys as gaps (follow-up in
L-DOCX-W6 and siblings). The determinism rule holds: no write sets
`modified` or `lastModifiedBy` unless the caller passes them (ADR 0006,
`docs/adr/0006-validation-security-and-compatibility.md:116-120`); OfficeCLI rewrites `custom.xml` audit keys on every save
(`OC/Core/OfficeCliMetadata.cs:180-249`), which Litchi does not (M2-Δ7).

### 9.3 Style application

| Format | What "apply a style" means | Route |
|---|---|---|
| DOCX | paragraph `style` (`w:pStyle`) and run `rStyle` (`w:rStyle`), both readable today (`O/format/docx/spec.rs:426-427`, `:496`) | L-DOCX-W5 (paragraph and run properties); style *definitions* stay read-only (CRUD is not P0) |
| XLSX | a cell's formatting: OfficeCLI's `font.*`, `fill`, `border.*`, `alignment.*`, `numberformat` keys; and applying an existing named cell style | L-XLSX-W2 (style allocation); applying an existing `Style` handle is (a) (`WorksheetEdit::style`) |
| PPTX | no style application in OfficeCLI's sense beyond shape properties | shape fill/line via L-PPTX-W3; text via L-PPTX-W5 |

### 9.4 Numbering and list basics

| Need | DOCX route |
|---|---|
| put a paragraph in a list (`numId`, `numLevel`/`ilvl`) | L-DOCX-W5 (`w:numPr` in the paragraph splice) |
| create a simple bullet or decimal list definition | L-DOCX-W10 (`abstractNum` + `num` added to an opened document's numbering part; today only `restartNumberingAfterBreak` is editable, `D/numbering/transaction.rs:126`, `:163`) |
| restart numbering | (a) the numbering transaction |

XLSX and PPTX have no numbering in OfficeCLI's P0 sense (PPTX bullets are
text-body properties, L-PPTX-W5).

## 10. Performance plan

### 10.1 Paired routes

Every route is opt-in (outside `Case::DEFAULT`, `tools/perf-baseline/src/lib.rs:1763`),
named with the S8b/M1 infixes, `<prefix>_unified_<rest>` and
`<prefix>_leaf_<rest>` (`docs/unified-ops/s8b-perf-routes.md:44-52`, `:58`),
both legs in one binary, on one corpus, alternated ABBA, and registered in the
`unified_ops` catalog (`tools/perf-baseline/src/unified_ops.rs:152`). The leaf
leg makes exactly the leaf calls the handler makes, with no detection, no
path parsing and no registry (`docs/unified-ops/s8b-perf-routes.md:67-85`). Each route's help text names
`litchi_ops::File` (record 0010 §9.1). Registering them updates
`EXPECTED_SELECTABLE_CASE_COUNT` (652 today, `tools/perf-baseline/src/lib.rs:63151`),
the registry test (`tools/perf-baseline/src/lib.rs:63222-63279`), the pinned
name lists (`tools/perf-baseline/src/lib.rs:64184-64226`),
`tools/test_crud_coverage_index.py:31-34` and the source-policy pins of
`tools/test_perf_baseline_source_policy.py:543-598` (62 unified / 56 leaf
names, `:555-556`).

| Scenario | Unified leg | Leaf leg | Corpus |
|---|---|---|---|
| one paragraph text set, warm source handle | `file.set("/body/p[N]", props!{"text" => …})` | `edit_document` → `replace_paragraph_text(Position)` → `commit` | `docx-semantic-{medium,large}`; real control `test-data/ooxml/docx/drawing.docx` |
| one run property set (after L-DOCX-W5) | `file.set("/body/p[N]/r[1]", props!{"bold" => true})` | the L-DOCX-W5 operation on the same position | same |
| one cell set | `file.set("/Sheet1/B2", props!{"value" => 42})` | `Workbook::edit` → `WorksheetEdit::set` → `commit` on the promoted workbook | generated `XlsxShape::DenseWide`; real control `no_drawing_patriarch.xlsx` |
| one cell format (after L-XLSX-W2) | `props!{"font.bold" => true, "fill" => "1F4E79"}` | the L-XLSX-W2 operation | same |
| one shape text set, source handle | `file.set("/slide[100]/shape[4]", props!{"text" => …})` | `edit_slide` → `set_shape_text` → `commit` on the source-backed editor | `pptx-source-backed-media` (200 × 8 text boxes) |
| one shape text set, eager handle | the same | `opened_presentation_transaction` → `set_shape_text` → `commit` | same |
| one shape geometry set (after L-PPTX-W3) | `props!{"x" => "2cm"}` | the L-PPTX-W3 operation | same |
| promotion (first write on a source handle) | first `set` on a cold source handle | `source_artifact` copy + eager open | xlsx, pptx corpora |
| rebase (DOCX, second write) | second `set` on the handle the first returned | publish into memory + source-backed open | docx corpora |
| batch scaling | `file.batch(items, Atomic)` with N homogeneous items, N ∈ {1, 10, 100, 1 000, 10 000} | one hand-written leaf transaction doing the same N operations, using the leaf's batched form where one exists | per family: XLSX cell values; DOCX paragraph text; PPTX shape text; DOCX run properties (after L-DOCX-W5) |
| batch, mixed families | alternating families every k items | per-run leaf transactions plus the same rebases | reported, no budget (§10.2) |
| save, unchanged | `file.save_as(tmp)` | leaf save of the same owner | each format |
| save after one set | `set` then `save_as(tmp)` | leaf commit + publish/save | each format |
| create | `litchi::create(f)` then `to_bytes` | leaf builder `to_bytes` | each format |

Untimed gates, as tests on every change (record 0010 §9.2 style):

- **Allocation counts** for parse + resolve + registry lookup + coercion +
  staging of a typed `set`, minus the leaf's own staging: zero (goal,
  "Allocation discipline").
- **Decode counts**: a `set` on `/slide[3]/shape[1]` decodes slide 3 only; a
  cell set decodes its worksheet only.
- **Byte equality**: every coalesced batch equals the same items applied one
  per transaction (§5.1 rule 2), on a real-producer and a Litchi-written
  fixture each (record 0009 §8 T3).

### 10.2 Budgets (review triggers, not claims)

| Scenario | Trigger | Source |
|---|---|---|
| single-op set, warm, same owner and engine | unified ≤ leaf + 10% | `docs/UNIFIED_OPS_GOAL.md:132-134` |
| batch of 1k homogeneous ops | unified ≤ 2 × leaf | `docs/UNIFIED_OPS_GOAL.md:135-141` |
| batch scaling | unified per-op cost flat (± noise) from 100 to 10k | `docs/UNIFIED_OPS_GOAL.md:135-136` ("scale linearly") |
| save | unified = leaf save: ≤ leaf + max(5%, 20 µs), zero added copies (allocation count of the publication path equal) | `docs/UNIFIED_OPS_GOAL.md:142-143` |
| promotion, rebase | reported per format, no trigger (one-time costs outside the budgets, §5.6) | record 0009 §9 |
| create | ≤ leaf builder + one in-memory open | — |
| existing read routes | no route slower than 5% (ABBA) | `docs/UNIFIED_OPS_GOAL.md:161-163` |

### 10.3 Method

The program's rules (`docs/GOAL.md:678-701`): before and after on one build
and machine, ABBA where noise warrants (`docs/UNIFIED_OPS_GOAL.md:159-163`),
machine-readable results, every result reported. Capture with the S1 driver
(`docs/unified-ops/results/0003/before/capture.py`: fresh child per group,
pinned CPU, `/usr/bin/time -v`); analyze with `analyze.py` extended to report
the median of per-child p50s, which the S10 rule uses but the script does not
compute yet (`docs/unified-ops/m0a-log.md:34`; finding F-9). Trigger
`max(5%, route spread)`, a 0.5 µs floor under 5 µs, full separation above
10% (`docs/unified-ops/s8b-perf-routes.md:365-371`). Results under `docs/unified-ops/results/0015-*/`, each
file under 1 MiB; raw captures outside the repository with name, size,
SHA-256 and the regeneration command (AGENTS.md).

## 11. Step breakdown

### 11.1 Branching and merge order

- **Integration branch.** `uo/m2` is cut from `uo/m1` at or after
  `056bc55b30`, the R5 merge: M2's write dispatch mirrors R5's read dispatch
  and flips R5's capability flags, so it needs R5, and its handlers resolve
  through the R4a/R4b adapters (R4c for PPTX). `uo/m1` is merged into `uo/m2`
  periodically, as `uo/m0a` is merged into `uo/m1` today (m1-log). Every step
  branches from `uo/m2` (leaf steps too, so their tests run against the
  current adapters) and is merged back by the coordinator after review.
- **Merge order into the trunk** (`feat/unified-ops-api`): M0a first (its
  single `--no-ff` merge, record 0007 Q5), then M1 (record 0010 §10.1), then
  M2. M2 merges only after M1; if M1 slips, M2 keeps merging `uo/m1` in and
  waits. Leaf-only prerequisite steps that M1 does not need still land on
  `uo/m2`, not on `uo/m1`, so the M1 merge carries no M2 code.
- **M2b** (XLSB and the OLE formats) runs on the same branch after the M2
  core exit (§11.3), as the goal's milestone order puts it, and can start
  its leaf steps earlier.
- **Green at every step** (ADR 0008): a verb whose handler has not landed
  returns `unsupported_capability` with its step id as follow-up, and
  `capabilities()` says so; nothing is exposed half-done.

### 11.2 Steps

Size as in record 0010 §10.2 (1 = one implementer for one change, 2 = twice
that or two implementers). "Leaf" steps change a leaf crate; "ops" steps
change `litchi-ops` (and the facade re-exports).

| Step | Content | Kind | Depends on | Parallel with | Size |
|---|---|---|---|---|---|
| **W1** | Verb surface types: `Commit<T>`, `Patch`/`PatchStep`, `Diagnostics`, `BatchItem` (21 keys, `serde::Deserialize` with OfficeCLI's aliases and both `props` forms), `BatchMode`, `BatchCommit`/`ItemResult`, `PropValue`/`IntoProps`/`props!`, `Placement`/`IntoPlacement`, `NewNode`/`IntoNode`, `IntoTarget`, `SaveOptions`/`SaveTo`/`Saved`, `CreateOptions`; `Error::{item, property}`; `File` verb methods that return `unsupported_capability` per format; rustdoc with runnable examples on a fake format | ops | `uo/m2` cut | W2, W3, all leaf steps | 1 |
| **W2** | Registry core: the `registry!` macro (extends XLSX's `props!`, `O/format/xlsx/spec.rs:170-201`) emitting M1's `PropSpec` slices plus `PropEntry` tables and sorted permutations; lookup, aliases, access, coercion of `PropValue`, constraints; typed key module `prop` generated from the tables; suggestion errors; classifier fixes (XLSX `EditBlocked` family, DOCX signed `UnsafeEdit`, §4.2–4.3); allocation tests | ops | W1 types | W3, leaf steps | 2 |
| **W3** | Batch executor core over the test-only fake formats (the M1 pattern, `O/path/fake.rs`, `O/query/fake.rs`): planning, runs, barriers, staged views, fusion hooks, atomic and best-effort modes, attribution and bisection, `edit` proxy, OfficeCLI result rendering for the command channel; the byte-equality and attribution tests of §10.1 against the fake | ops | W1 | W2, leaf steps | 2 |
| **W4** | Save and create: origin path in every handle state; `save`/`save_as`/`save_with`/`write_to`/`to_bytes` per owner state; `create(format)` round trip; the signature rule of §8.1 (removal at first write or save) and the encryption rule of §8.2 (`is_sensitive`, `into_plaintext`, re-encrypting save) | ops (+ L-OLE-W1 for legacy saves) | W1 | W2, W3 | 1–2 |
| **L-XLSX-W1** | record 0009 S1 (§6.1) | ops | `uo/m2` | everything | 1 |
| **L-DOCX-W1** | record 0009 S2 | leaf + ops | `uo/m2` | everything | 1–2 |
| **L-PPTX-W1** | record 0009 S3 | leaf + ops | `uo/m2` | everything | 2 |
| **L-DOCX-W2, W3, W4** | run addressing, nested-container verdict, story admission (§6.3) | leaf | — | everything | 1 each |
| **L-PPTX-W2, W7** | shape id key; multi-paragraph text fix (§6.3) | leaf | — | everything | 1 each |
| **L-DOCX-W5** | DOCX formatting engine (§6.4) | leaf | — | everything | 3 |
| **L-XLSX-W2** | XLSX style allocation (§6.4) | leaf | — | everything | 2 |
| **L-PPTX-W3** | PPTX shape properties (§6.4) | leaf | L-PPTX-W2 | everything | 2 |
| **L-DOCX-W6, L-XLSX-W4, L-PPTX-W6, L-DOCX-W10** | metadata families; PPTX slide hide/rename, notes creation, table-cell text; DOCX list definitions (§6.4) | leaf | — | everything | 1 each |
| **W5a** | DOCX handler: families, addressing, fusion, pending state and rebase, commit adapter, the §4.2 rows whose leaf API exists; each later leaf step turns its rows on | ops | W1–W3, L-DOCX-W1 | W5b, W5c | 2 |
| **W5b** | XLSX handler (§4.3), including sheet lifecycle and staged-name tracking | ops | W1–W3, L-XLSX-W1 | W5a, W5c | 2 |
| **W5c** | PPTX handler (§4.4) | ops | W1–W3, L-PPTX-W1, R4c merged into `uo/m1` | W5a, W5b | 2 |
| **W6** | Capabilities, schemas, goldens: `operations` and property `set`/`add` flags from the handler and registry tables; per-property `values` (from `Tokens`), `aliases`, constraint bounds; re-bless `tests/golden/capabilities/*.json`; extend `tools/officecli_capability_delta.py` with write columns (supported / gap / write-only, against OfficeCLI's `add`/`set` flags); write-liveness round trips (every settable key: `set` a sample, `get` it back) | ops + tools | W2, W5a–c | W7 | 1 |
| **W7** | Performance routes and the M2 measurement record (§10) | harness | W5a–c | W6, W8 | 1–2 |
| **W8** | P0 closure (§9) and the CRUD gate evidence for the write rows (`docs/CRUD_Scenario_Checklist.md:55-58`, `:67-68`, `:80`); record 0009 §8's T1–T10 on each tier-1 format, each on a real-producer and a Litchi-written fixture; the OfficeCLI write corpus subset (the `create`/`add`/`set` lines of `3rdparty/OfficeCLI/examples/{word,excel,ppt}/**/*.sh` that the landed rows cover) as an integration test that never panics or returns `internal_error` | ops | W5a–c, the P0 leaf steps | W6, W7 | 1–2 |
| **L-XLSB-W1/W2, L-OLE-W0/W1, L-DOC-W1, L-PPT-W1, L-XLS-W1** | M2b leaf work (§6.1, §6.2, §6.5) | leaf/ops | — | anything | 1–2 each |
| **W9a / W9b / W9c / W9d** | XLSB, DOC, XLS, PPT handlers (§4.5–4.8), with their capability goldens, perf routes and T1–T10 | ops | M2 core exit; their leaf steps | each other | 1–2 each |
| **W10** (M2-extended) | raw layer reads and `add-part` (§2.3) | ops | W3, W5a–c | anything | 1 |

Critical path: L-DOCX-W5 (the largest leaf item) → W5a's formatting rows →
W8's OfficeCLI corpus subset. Everything else is parallel: in the first wave
up to about twelve implementers (W1–W4 plus the tier-1 leaf steps), as in
M1's first wave (record 0010 §10.2).

### 11.3 Exit checklist (M2 core; M2b repeats it for its four formats)

1. `set`, `add`, `remove`, `move_to`, `swap`, `edit`, `batch`, `apply`,
   `save`, `save_as`, `save_with`, `write_to`, `to_bytes` and `create` work on
   DOCX, XLSX and PPTX for every (a) row of §4.2–4.4 and every row its landed
   leaf steps turned on; every other row is a listed gap with a follow-up in
   `capabilities()` and in the generated delta table.
2. The P0 items of §9 work for the three formats.
3. Atomicity: a failed batch leaves the caller's handle byte-identical
   (record 0009 T10); best-effort results name every failing item.
4. Sharing: record 0009 T1–T9 pass on each format, on a real-producer and a
   Litchi-written fixture; the XLSX `materialize` fallback and the PPTX
   `EditorRefused` path are deleted.
5. Coalescing is invisible: every batch test also runs one item per
   transaction and compares bytes.
6. The registry is the single source: schemas, `capabilities()`, suggestions
   and Python introspection read it; goldens are blessed; the write-liveness
   test covers every settable key.
7. The goal's verification gates: fmt, clippy `-D warnings` over the feature
   matrix, rustdoc `-D warnings`, the crate-boundary checker (no new edge is
   expected: `litchi-ops` already depends on every leaf it writes), unit,
   integration, doc and compile-fail tests, the M0b and M1 fuzz targets, and
   pytest for the read surface (the Python write surface is M3).
8. The §10 routes are recorded with every trigger met or explained in the
   record, with a follow-up.
9. No accepted ADR, preservation, security, determinism or boundary
   requirement regressed; exact no-ops stay exact (`docs/CRUD_Scenario_Checklist.md:55`).

## 12. Risks, findings, open questions and deltas

### 12.1 Risks

- **Risk 1 (parity engines).** The OfficeCLI corpus formats almost
  everything it creates (§6.4), and none of the three tier-1 leaves can
  format existing content. L-DOCX-W5 is the largest new leaf engine of M2 and
  sits on its critical path. Mitigation: it starts in the first wave; until
  it lands, M2 ships text, structure and metadata, with every formatting key
  a listed gap; W8's corpus subset grows as the engines land.
- **Risk 2 (DOCX chained writes).** Until overlay chaining (record 0009
  option D) lands in `litchi-docx`, the second write to a source-backed DOCX
  handle, and every family boundary inside a batch, costs a publication into
  memory plus a re-open: O(document), outside the handle-cost budget
  (§5.6 measures it). Interactive sessions that alternate `set` and `get`
  pay it on every pair. Mitigation: fusion keeps homogeneous batches in one
  run; the rebase is memoized per handle; option D is the follow-up.
- **Risk 3 (XLSX commit cost).** One eager commit rewrites each touched
  worksheet in full and reads it back (§4.3); §5.6 measured 1.9 ms on a
  15 KB workbook and 53.5 ms on a 75,770-cell sheet for a one-cell commit,
  and a chain of one-cell commits costs that per item. Single-op writes
  stay within their budget (the overhead is the unified layer's, not the
  commit's), but an agent that issues many single sets on a large sheet is
  slow. Mitigation: `edit`/`batch` in the docs and examples; a leaf follow-up
  for incremental worksheet rewrites, outside M2.
- **Risk 4 (sequential semantics).** Coalescing must never change a result
  (§5.1). The barrier rules are the risk. Mitigation: every executor test
  compares a coalesced batch with the same items one per transaction, byte
  for byte, on real-producer and Litchi-written fixtures (§10.1).
- **Risk 5 (fidelity of promotion).** XLSX and PPTX writes on source handles
  run on an eager owner opened from the exact source. It raw-copies
  untouched members and refuses rather than regenerates when it cannot
  (`OPC/pkgwriter.rs:1268-1273`), but whether it rewrites more members than
  the source-backed editors would is not measured for PPTX. Mitigation:
  record 0009 T3/T5 on real-producer fixtures in W5b/W5c, counting raw
  members as record 0009 §2.4 did.
- **Risk 6 (R4c).** The PPTX handler resolves through R4c, which is not
  merged into `uo/m1` yet. Mitigation: W5c depends on it explicitly; L-PPTX
  leaf steps do not.
- **Risk 7 (memory).** Pending commits, memoized rebases and promoted
  owners raise the resident memory of written handles; Python lists of
  handles keep all of it (record 0009 §9). Mitigation: rebases and
  promotions are charged to the execution context when there is one (D17);
  W7 measures peak RSS on the write routes.
- **Risk 8 (classifier drift).** Leaf errors that fall through to
  `corrupt_file` (XLSX, `O/error.rs:919-931`) or `unsupported_capability`
  (DOCX signed `UnsafeEdit`, `O/error.rs:854-865`) would mislead agents on
  every write. Mitigation: W2 classifies them, with a test per variant.
- **Risk 9 (Windows).** Atomic replacement over a file another handle holds
  open is untested on Windows (record 0009 §9; change 0497, `docs/performance/changes/0497-docx-atomic-publication.md:42-43`). Not
  claimed.
- **Risk 10 (cold DOCX resolution).** A cold `set` on `/body/p[N]` builds
  M1's body index (a whole-part scan, memoized per part generation; the
  first `/body/p[N]` on `drawing.docx` took 3.45 ms before the memo, m1-log
  L-DOCX-R1) and then the `Edit` scans the part again for its own snapshot
  (§5.6: 1.12 ms edit start on `drawing.docx`). The single-op budget is for
  warm handles; the cold path is reported, and sharing one scan between the
  index and the transaction snapshot is a leaf follow-up.
- **Risk 11 (leaf publication limits).** Large batches can hit leaf limits
  that hand-written code also hits: the XLSX eager commit of 10,000 cells on
  the 75,770-cell sheet was refused by the publication audit ("XML Attributes
  limit 250000 exceeded"), and a DOCX `Edit` stops at 4,096 operations. The
  executor splits runs at known caps (§5.3) and reports audit refusals with
  the item range; it never loosens a limit.
- **Risk 12 (noise).** Single-op budgets on millisecond commits are
  insensitive to the unified overhead, while batch budgets on tiny
  documents are dominated by it (created-empty XLSX: 547 µs for 1,000
  cells). Mitigation: report both, with ABBA and per-route floors (§10.3).

### 12.2 Findings (disagreements with records or code)

- **F-1 (record 0003 R10).** "A shared facade atomic writer in
  `litchi-ops`" would be a third sibling-temporary implementation; two exist
  (`litchi_opc::atomic`, CFB `publish_staged`). M2 builds a facade policy
  layer over them and exports the CFB one (L-OLE-W1, D13).
- **F-2 (design §3).** "Batches sequence per-operation commits across Edit
  families … every per-op commit pays dependency-closure validation." With
  fusion and staged views most homogeneous batches need one commit per run,
  not per operation; per-op commits remain only at barriers.
- **F-3 (record 0002, DOCX).** Record 0002 lists "~20 scoped transaction
  families"; the `document::Edit` itself has 24 text and insertion
  operations and no property, removal or move operation
  (`D/document/transaction.rs:2281-4457`); the families that touch other
  parts are mostly eager-only or source-only (§4.2).
- **F-4 (record 0002, DOCX comments).** The suspected comments overwrite is
  real by code path and still untested: `from_xml` starts with an empty
  comment list (`D/writer/doc/model.rs:299`) and save replaces
  `/word/comments.xml` (`D/package/package/parts.rs:148`;
  `OPC/package.rs:1672`); authored footnotes and endnotes follow the same
  pattern (`D/package/codec.rs:895-896`, `:1104-1115`). M2 never routes
  through `document_mut` (L-DOCX-W7 replaces it).
- **F-5 (record 0010 §4.5.1).** "M2 maps `/body/p[N]` to the transaction's
  ordinal" holds for paragraphs; *runs* differ: `/p[N]/r[M]` counts
  OfficeCLI's run set (runs inside `w:hyperlink` included) while
  `replace_run_text` addresses direct `w:r` children only
  (`D/document/transaction.rs:3571`, `:4944`). L-DOCX-W2.
- **F-6 (unified errors).** DOCX's eager signed refusal is `UnsafeEdit`
  (`D/package/package/access.rs:240-246`), which the classifier maps to
  `unsupported_capability`; record 0013's policy says `document_protected`.
- **F-7 (record 0009 §2.4).** XLSX `materialize` is still called at open
  under `SourceMode::Eager` from `open_prepared` (`O/format/xlsx.rs:233-234`),
  not only in the editor fallback record 0009 F3 named. L-XLSX-W1 replaces
  both.
- **F-8 (PPTX).** `set_shape_text`'s read-back may refuse every
  multi-paragraph target (§4.4); unverified.
- **F-9 (method).** The S10 rule's "median of per-child p50s"
  (`docs/unified-ops/m0a-log.md:34`) is not computed by `analyze.py`, which pools samples;
  W7 extends the script.
- **F-10 (legacy builders).** M0a's DOC and PPT builders hand out the narrow
  source-backed editors (equal-length text only); the general editors are
  the eager snapshots, which publish bytes only (§4.6, §4.8).
- **F-11 (XLSB).** `Xlsb::editor` materializes on every call and loses the
  exact source (`O/format/xlsb.rs:363-365`); a non-mutating apply already
  exists, crate-private (`B/cell_values/workbook.rs:121-136`).
- **F-12 (record 0009 §2.4, extended).** The probe's survey of 137 XLSX
  fixtures (§4.3) confirms record 0009's conclusion that most M2 XLSX writes
  must run on the eager engine: the source value editor admits 45, the eager
  route 132.
- **F-13 (PPTX costs).** The eager opened transaction costs 3.6–4.8 ms to
  start and about 70 µs per `set_shape_text` (the probe traced two scene
  parses per call, `P/opened/transaction.rs:228`, `:246`), so the source text
  editor, not the eager engine, is the cheap route for text (D7).
### 12.3 Open questions (provisional defaults adopted, for owner confirmation)

| Id | Question | Default |
|---|---|---|
| Q1 | Single-op verbs return `File` (diagnostics dropped) or `Commit`? | `File`; `edit` returns `Commit` (record 0009 §7) |
| Q2 | Unknown or unsupported keys: refuse, or OfficeCLI's apply-the-rest-and-warn? | refuse in the Rust and Python verbs; the command channel (`apply`, `batch`, the M3 CLI) takes `unsupported = "warn"` to replay OfficeCLI scripts that rely on it (M2-Δ1) |
| Q3 | What does a never-edited `Signed::unsign` file save? | the unsigned copy (D14) |
| Q4 | Keep the password after `unlock` to re-encrypt on `save()`? | no: the caller passes it to `save_with` (D15) |
| Q5 | Edit a sensitive (decrypted) file in M2? | no; `into_plaintext` or P1 (L-OOXML-W2) |
| Q6 | `save_as` with another family's extension? | refused (`invalid_argument`); export is explicit (§7.1) |
| Q7 | Apply a `Patch` across a DOCX rebase? | `unsupported_capability` until option D (D16) |
| Q8 | PPTX text-body splicer (L-PPTX-W5) inside M2 or after? | after the core exit, as M2-extended; shape-level text and properties are M2 |
| Q9 | Promote source handles to the eager engine at the first write? | XLSX yes; PPTX for any run the source text editor cannot take; DOCX no (D7) |
| Q10 | Use the XLSX source-backed editors for small single-sheet edits? | no in M2 (one engine per format; their caps and refusals, §4.3) |
| Q11 | Python `add` returns `File` or `(File, path)`? | `File`; paths through `apply`/`batch` |
| Q12 | Query targets for `set`/`remove` (OfficeCLI's fan-out) in M2? | yes, guarded by `bare_selector_rejected` |
| Q13 | Run DOCX main-story runs of multi-run batches on the eager owner (cheaper chaining) instead of rebasing the source owner? | not in M2; decide after W7 measures media-heavy documents |
| Q14 | Raw layer timing: `raw`/`add-part` in M2-extended, `raw-set`/`validate` in M4? | yes (D19) |

### 12.4 OfficeCLI deltas introduced by M2

| Id | OfficeCLI | M2 | Disposition |
|---|---|---|---|
| Δ1 | unknown keys are applied as raw XML when they match an OOXML name, otherwise reported as `unsupported_property` warnings while the rest applies (exit 2) (`OC/CommandBuilder.cs:764-806`, `:1763-1821`; `OC/CommandBuilder.Set.cs:338-351`, `:435`) | typed refusal with a suggestion; `unsupported = "warn"` on the command channel only | improvement (strict API), compatibility switch |
| Δ2 | a key one edit away from a known name is silently corrected with an `auto_corrected` warning (`OC/CommandBuilder.Set.cs:329-337`) | never corrected; the suggestion names it | deliberate |
| Δ3 | `r[N]` edits any run of the set | runs inside hyperlinks refused until L-DOCX-W2 | deferred |
| Δ4 | error codes inferred from message text; many write failures are `internal_error` (`OC/Core/OutputFormatter.cs:317-715`) | one typed code per failure | improvement |
| Δ5 | best-effort is the CLI default; atomic still runs every item before rolling back unless `--stop-on-error` (`OC/CommandBuilder.Batch.cs:119-131`, `:177-193`) | atomic is the default and stops at the first failure; best effort is explicit | improvement (design §3) |
| Δ6 | batch numbers are stringified in the current culture (`OC/BatchTypes.cs:47`) | invariant canonical forms | improvement |
| Δ7 | every modified save rewrites `custom.xml` audit keys (`OC/Core/OfficeCliMetadata.cs:180-249`) and `create` stamps `creator=OfficeCLI` (`:113-117`) | never | deliberate (ADR 0006 determinism) |
| Δ8 | atomic save without fsync (`OC/Core/AtomicPackageWriter.cs:45-51`); write-back errors swallowed at dispose | `Durability::Full` by default; every failure typed | improvement |
| Δ9 | opening for edit may repair the source file in place (`OC/Handlers/DocumentHandlerFactory.cs:84-131`) | never (ADR 0006, `docs/adr/0006-validation-security-and-compatibility.md:62`) | improvement |
| Δ10 | no encryption or signature support: an encrypted file fails to open as `corrupt_file` and nothing detects signatures (`OC/Handlers/DocumentHandlerFactory.cs:139-155`, `:175`) | `Locked`/`Signed`/sensitive states, explicit consumption | superset |
| Δ11 | Excel column `--index` is 1-based against help that says 0-based (`OC/Handlers/Excel/ExcelHandler.Add.cs:420-450`) | zero-based everywhere | improvement |
| Δ12 | sets on unknown OOXML elements are silent no-ops (`OC/Handlers/Word/WordHandler.Set.cs:645`) | `unsupported_path`/`unsupported_capability` | improvement |

## What is not claimed

- No production code was written. Every unified-layer cost in this record
  (parse, resolve, coerce, wrap) is an estimate from M0b and M1 measurements
  of reads, not a measurement of writes; W7 measures them.
- The leaf numbers of §5.6 come from one throwaway probe: a standalone Cargo
  package in the session scratchpad with path dependencies on this tree,
  locked to `tools/perf-baseline/Cargo.lock` (the worktree keeps no root lock
  file), built in release with the workspace's `lto = true`,
  `panic = "abort"`, pinned to one CPU of a shared 32-CPU host (load 2–8 while
  it ran), median of 11–201 repetitions per point, each from the same base;
  each edited output was re-opened and read back once per data point. They
  are single-host, single-session figures, not ABBA comparisons, and they
  ranked engines (§4, §5.6, D7); they are not budgets. The condensed tables,
  the fixtures, the method and the exact leaf calls are committed in
  `docs/unified-ops/results/0015/README.md` (well under 1 MiB). The probe
  sources were throwaway and were deleted with the scratchpad; no binary was
  committed and the build directory was removed.
- Line counts and sizes in §6 and §11 are estimates.
- Statements marked "(inferred)" were read from code and not executed: DOCX
  eager chaining, XLSX sheet swap as two moves, the PPTX multi-paragraph
  `set_shape_text` refusal, the comments-part overwrite.
- OfficeCLI behaviour is described from its source at `dced0d74`; where its
  help text and code disagree (Excel column `--index`, `trackChange.*` keys),
  this record follows the code.
- No measurement covered Windows; no claim is made about replacement over an
  open file there.
