# 0353 — Static dispatch on the `litchi-core::sheet` traits

- Status: implemented on `uo2/sheet-dispatch`, cut from `feat/unified-ops-api`
  at `5ff605692e`.
- Date: 2026-10-20
- Input: owner decision [0350](0350-owner-api-debt-decision.md), memo
  [0331](0331-owner-decision-memo.md) item 29. Goal constraint 2: no
  `Box<dyn Trait>` on a public or per-operation path.
- Owned: `crates/litchi-core/src/sheet/traits.rs`, the trait implementations
  in `litchi-xls`, `litchi-xlsb`, `litchi-ops` (`calc.rs`, `view/`,
  `format/{xlsx,xlsb,xls}`) and the facade's `sheet` module, the
  `litchi-eval` engine's cell-read path, and every consumer (tests,
  examples). Shared dispatch files: one `[[example]]` hunk in
  `crates/litchi-ops/Cargo.toml`.
- `performance_claim`: measured (§4). Every measured scenario is at or below
  its baseline; per-cell iteration is up to 2.8 times faster.

## 1. The traits

`Cell`, `Worksheet` and `WorkbookTrait` used to return `Box<dyn Cell>`,
`Box<dyn Worksheet>` and boxed `CellIterator`/`RowIterator`/
`WorksheetIterator` objects. Every cell read through the traits allocated
one box, and so did every worksheet lookup. The formula engine looks the
worksheet up for each cell it reads, so every evaluated reference cost at
least two allocations.

Now each implementation names its concrete types as generic associated
types:

```rust
pub trait Worksheet: Send + Sync {
    type Cell<'a>: Cell where Self: 'a;
    type Cells<'a>: Iterator<Item = Result<Self::Cell<'a>>> where Self: 'a;
    type Rows<'a>: Iterator<Item = Result<Cow<'a, [CellValue]>>> where Self: 'a;
    fn cell(&self, row: u32, column: u32) -> Result<Self::Cell<'_>>;
    fn cells(&self) -> Self::Cells<'_>;
    fn rows(&self) -> Self::Rows<'_>;
    // name, dimensions, cell_value, row, … unchanged
}

pub trait WorkbookTrait: Debug + Send + Sync {
    type Worksheet<'a>: Worksheet where Self: 'a;
    type Worksheets<'a>: Iterator<Item = Result<Self::Worksheet<'a>>> where Self: 'a;
    fn worksheet_by_index(&self, index: usize) -> Result<Self::Worksheet<'_>>;
    // …
}
```

- The three iterator traits are deleted (not aliased). The iterators are
  `std::iter::Iterator`s, so `for cell in sheet.cells()` works.
- The traits are no longer object-safe. Nothing used them as trait objects
  except tests, which now take `&impl WorkbookTrait`. `FormulaEvaluator`
  was already generic over `W: WorkbookTrait`.
- `litchi-core` adds blanket implementations: `Cell for &T` and for
  `Cow<'_, T>`, and `Worksheet for &T` and for `Arc<T>`. A worksheet can
  therefore lend stored cells by reference and absent ones by value, and a
  workbook can lend a cached worksheet without copying it. The XLS
  `Cell for &Cell` and `Worksheet for &Worksheet` impls these supersede are
  removed.
- Method calls on a returned worksheet or cell now need the trait in scope
  (`use litchi_core::sheet::{Cell as _, Worksheet as _}`), because the
  methods are no longer inherent methods of a `dyn` type. Every consumer in
  the workspace was migrated.
- The trait's error type is still `Box<dyn Error + Send + Sync>`; §3 says
  why.

## 2. Implementations

| Implementer | `Worksheet` | `Cell` | Notes |
|---|---|---|---|
| `litchi_xls::raw::Reader` | `&Worksheet` | `Cow<'_, Cell>` | zero-copy, as before. New public iterators `worksheet::{Cells, Rows}`, `workbook::Worksheets` |
| `litchi_xlsb::Workbook` (eager) | `Arc<Worksheet>` from the owner's cache | `Cow<'_, Cell>` | this used to clone the whole decoded worksheet on every lookup, and every cell on every read. `WorksheetIterator` is renamed `Worksheets`; new `sheet::{StoredCells, Rows}` |
| `litchi_ops` `Xlsx`, `Xlsb`, `Xls`, `WorkbookView` | `view::sheet::Worksheet<'a>` | `view::sheet::Cell<'a>` | one public set of enums over the per-format types (`litchi_ops::view::sheet::{Worksheet, Cell, Cells, Rows, Worksheets}`), shared by the handles and the view. A read costs one `match` |
| `litchi_ops` calc `Grid` (evaluation snapshot) | `SheetRef<'g, 'f>` | `GridCell<'g>` borrowing the snapshot | crate-private; `rows()` is `std::iter::Empty` as before |
| facade `sheet::text::TextWorkbook` | `TextWorksheet` (owned, as before) | `TextCell` | minimal migration only (see §5) |
| `litchi-eval` examples | `&MemSheet` | `MemCell<'_>` | the borrowed `SheetRef` wrapper the boxes needed is gone |

The ODS arm of `WorkbookView` still refuses worksheet access. Its
`worksheets()` yields the refusal through a `Refused` variant of
`view::sheet::Worksheets`.

**Keeping the cell enum small.** The first cut used `Cow<'a, xls::Cell>` and
`Cow<'a, xlsb::Cell>` inside `view::sheet::Cell`. That made every cell
272 bytes, and `cell/xlsx` and `cell/xls` were 4–7% slower than with the
boxes (§4, run 1). The BIFF arms now hold a `Stored` value: either a
borrowed stored cell, or the coordinates of an absent position, which read
as empty. The leaves only produce owned cells for absent positions (checked
with `debug_assert!`). For an absent position, `coordinate` and
`is_formula` rebuild the leaf's empty cell, so the results are the leaf's
own.

## 3. `Box<dyn …>` on the evaluate path: removed and kept

The engine's `EngineCtx::get_cell_value` returned a boxed future for every
cell read, including cache hits and literal reads. Now it returns `CellRead`,
a concrete future. `CellRead::Ready` holds a result computed at once:
- a stored value;
- a cached formula result;
- a value already in the evaluation cache;
- the circular-reference, depth-limit and missing-sheet outcomes.

Only an uncached formula is `CellRead::Pending`, a boxed future.
`raw_cell_value` is always ready. Unary and binary operators evaluate
literal and reference operands inline, and box only nested
sub-expressions. The cycle check, depth bound, caching and error behaviour
are unchanged (litchi-eval's 1,163 tests, `litchi-ops` `calc`, and the
facade `eval_xlsx_integration` suite pass unchanged).

Removed from per-cell or per-operation paths (all in this change):

| Site | Before | After |
|---|---|---|
| core traits (`cell`, `cell_by_coordinate`, `cells`, `rows`, `worksheet_by_*`, `active_worksheet`, `worksheets`) | one `Box` per call and per iterated item | none |
| `litchi-ops` `calc.rs` grid, `view/workbook.rs`, `format/{xlsx,xlsb,xls}` (the 72 occurrences of memo item 29) | `Box<dyn CoreWorksheet>` / `Box<dyn CoreCell>` per read | none |
| `litchi-eval` `get_cell_value` / `raw_cell_value` | `BoxFuture` per cell read (each reference, each cell of a range argument) | none unless the cell holds an uncached formula |
| `litchi-eval` `evaluate_expression` operands | `Box::pin` per operand of every operator | none for literal and reference operands |
| `litchi-xlsb` eager workbook lookups | a full worksheet clone per lookup and a cell clone per read | an `Arc` clone and a borrow |

Kept, with the reason (items 1 and 2 are resolved by
[0353a](0353a-sheet-errors-and-sync-engine.md), where F1 and F2 landed):

1. **The trait's error type, `Box<dyn Error + Send + Sync>`**
   (`litchi_core::sheet::Result`). It allocates only on a failure, never on
   the success path. `litchi-ops` carries its typed `crate::Error` through
   it and downcasts on the way out (`calc::sheet_error`,
   `xlsb::sheet_error`), and the engine both raises string errors through
   it and recovers `DepthLimit` by downcast. Replacing it needs a concrete sheet error
   type in `litchi-core`, with a way to carry the typed leaf errors.
   Follow-up **F1**.
2. **One boxed future per evaluated uncached formula, per nested
   sub-expression and per function call** (`CellRead::Pending`,
   `engine::evaluate_operand`, the `registry` function table). The engine
   is `async`: `FormulaEvaluator::evaluate_cell` is a public `async fn`, so
   that `WEBSERVICE` can await the caller's `Fetch` transport. Async
   recursion (formula → expression → function → referenced formula) needs
   one indirection per recursion level, so these boxes cannot be removed
   while the engine is async. They are per formula node, not per cell read.
   A range argument of any size and a reference to a stored or cached cell
   box nothing. Removing them means a synchronous engine: resolve external
   fetches in a prefetch phase, then evaluate synchronously. That changes
   the public `evaluate_cell` and `Fetch` contract. Follow-up **F2**.
3. **The public `Fetch` types** (`FetchFuture`, `FetchError`,
   `web_functions` feature). This is the caller-owned transport contract,
   with one future per external request. It belongs with F2.
4. **`&dyn DispatchCtx` and `&dyn Fetch`**. These are borrowed vtable
   references, not boxes, and they allocate nothing.

## 4. Measurements

Harness: `crates/litchi-ops/examples/sheet_dispatch_bench.rs` (new).

- **Workbook.** For each of XLSX, XLSB and XLS it builds a `rows` × 4 sheet
  through `File::import` with formulas: a number, then `=A*2+1`, an `IF`,
  and a `SUM` over the row. XLSX and XLSB use 2,000 rows; XLS uses 1,000
  (BIFF import cost, as in record 0266).
- **Scenarios:**
  - `recalculate/<f>`: `File::recalculate`, the unified evaluate path;
  - `evaluate/<f>`: `litchi_eval::FormulaEvaluator` over the
    `WorkbookView`, every formula cell;
  - `cell/<f>`: `Worksheet::cell` and `Cell::value` for every cell;
  - `cell_value/<f>`: `Worksheet::cell_value` for every cell;
  - `cells/<f>`: one pass of `Worksheet::cells()`.
- **Build:** release, `cargo build --release -p litchi-ops --features
  xlsx,xlsb,xls,eval --example sheet_dispatch_bench`.
- **Baseline (A):** the same harness at `5ff605692e`, with one line
  different: the baseline needs `CellIterator as _` imported to call
  `next()` on the boxed iterator. Binary SHA-256
  `80f0e0fad53a97f863695972f77f8ecc9be1aef84d46a53273bbc457cccb4dbc`.
- **Candidate (B):** this branch. Binary SHA-256
  `4beb00b0137fd6b9db671a35352ac9bdcb2f150f0624de2a18865747b42bfdf6`. The
  binaries live outside the repository, in the coordinator scratch
  directory `sheet-dispatch/`.
- **Order:** process-level ABBA, 8 rounds (16 processes per leg), pinned
  with `taskset -c 30`. Each process takes 15 samples per scenario after
  one warm-up. The table gives the median of the per-process medians.

| Scenario | Baseline µs | Candidate µs | B/A |
|---|---:|---:|---:|
| `cell/xls` | 123.8 | 106.2 | 0.858 |
| `cell/xlsb` | 1117.5 | 866.2 | 0.775 |
| `cell/xlsx` | 832.9 | 818.4 | 0.983 |
| `cell_value/xls` | 114.5 | 107.4 | 0.938 |
| `cell_value/xlsb` | 1159.2 | 884.8 | 0.763 |
| `cell_value/xlsx` | 836.9 | 811.8 | 0.970 |
| `cells/xls` | 23.4 | 8.5 | 0.362 |
| `cells/xlsb` | 633.3 | 316.2 | 0.499 |
| `cells/xlsx` | 314.3 | 286.0 | 0.910 |
| `evaluate/xls` | 604.0 | 549.2 | 0.909 |
| `evaluate/xlsb` | 2153.7 | 1772.3 | 0.823 |
| `evaluate/xlsx` | 5986.0 | 5888.7 | 0.984 |
| `recalculate/xls` | 3016.0 | 2971.5 | 0.985 |
| `recalculate/xlsb` | 5797.0 | 5710.5 | 0.985 |
| `recalculate/xlsx` | 5767.5 | 5644.5 | 0.979 |

Evidence: [`results/0353/abba.json`](results/0353/abba.json) (this run) and
[`results/0353/abba-run1.json`](results/0353/abba-run1.json). Run 1 was the
first cut with the 272-byte cell enum, which `cell/xls` (1.072) and
`cell/xlsx` (1.042) missed. §2 explains the fix.

- **`recalculate` gains least (about 2%).** Its time goes to loading the
  snapshot grid (a `HashMap` of every cell) and to the formula evaluation
  itself, not to dispatch.
- **`cell/xlsx` and `cell_value/xlsx`** are bound by the XLSX leaf's
  per-cell value conversion.

## 5. Coordination notes

- **`crates/litchi/src/sheet/text`.** The implementations there had to
  change in this commit, because the trait signatures changed and the
  workspace must build. The edits are confined to the trait `impl` blocks
  and their iterator types (`iterators.rs`, `worksheet.rs`, `workbook.rs`,
  and imports in `tests.rs`). No item moved, and no name changed. When the
  agent moving the module merges, the conflicts are the trait `impl`
  bodies: keep the moved location and take this change's bodies.
- `litchi-core` metadata types are untouched.

## 6. Follow-ups

- **F1.** A concrete error type for `litchi_core::sheet::Result` (§3 item 1).
- **F2.** A synchronous formula engine, with external fetches resolved in a
  prefetch phase. This removes the per-formula boxed futures and the
  public boxed `Fetch` future (§3 items 2–3).
