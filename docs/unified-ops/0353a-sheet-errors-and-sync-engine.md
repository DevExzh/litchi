# 0353a — Typed sheet errors and a synchronous formula engine

- Status: implemented on `uo2/sheet-dispatch`, on top of record
  [0353](0353-sheet-trait-static-dispatch.md) (tip `d8c442a413`).
- Date: 2026-10-20
- Input: the coordinator's decision on 0353's follow-ups F1 and F2. The
  owner asked for these boxes to be removed as required work (0350, memo
  item 29).
- Owned: `crates/litchi-core/src/sheet/{error,traits,types,mod}.rs`,
  `crates/litchi-eval`, and the trait implementations and consumers in
  `litchi-xls`, `litchi-xlsb`, `litchi-ops` (`calc.rs`, `view/`,
  `format/{xlsx,xlsb,xls}`, plus `From` arms in `error.rs`) and the facade's
  `sheet` module (trait `impl` bodies only, see §5).
- Also touched: `litchi-xlsx`. It used the boxed
  `litchi_core::sheet::Result` as a general result alias; now it uses a
  local alias with the same type (§2.3).
- `performance_claim`: measured (§4). The 15 scenarios of 0353 plus three
  formula-heavy recalculations are re-run against two baselines: 0353 and
  trunk before 0353. Against trunk every scenario improves, by 2–66%.
  Against 0353, 14 of 18 improve and the formula-heavy recalculations are
  12–14% faster. `evaluate/xls` (+0.3%), `evaluate/xlsb` (+0.2%) and
  `cell_value/xlsx` (+0.4%) are within run-to-run spread. `cells/xlsx` is
  +1.5% against 0353, a reproducible codegen residual explained in §4.3.
  It is still 7% faster than trunk.

## 1. F2 — the formula engine is synchronous

- **No futures in the evaluation path.** The 1,360 `async fn`s of
  `litchi-eval` are plain functions; the 1,775 `.await`s and every
  `Box::pin` are gone. This includes the per-formula box (0353's
  `CellRead::Pending`), the per-sub-expression box
  (`engine::evaluate_operand`, now deleted) and the per-function-call box.
  The function table holds plain `fn(&dyn DispatchCtx, &str, &[Expr]) ->
  Result<CellValue>` pointers. `EngineCtx::get_cell_value` returns the
  value.
- **Recursion and its bound are unchanged.** Recursion is plain native
  recursion, bounded by `with_max_depth`. The check order is unchanged:
  evaluation cache, then the depth bound, then the cycle set, then the
  read, then the evaluation. Stack per formula level, measured with a
  stack-address probe over 400-level chains: 1.8–2.5 KiB optimized and
  6.8–10 KiB in a debug build. 0353 documented about 3 KiB for the async
  engine. The `litchi_ops::calc::Options::max_depth` doc is updated.
- **`WEBSERVICE` goes through the caller's transport in two phases.**
  - When a pass needs a response that has not been fetched yet,
    `EngineCtx::fetched(url)` records the URL in the session and the
    function reads Excel's own placeholder, `#GETTING_DATA`.
  - From then on the pass caches nothing, because later values may rest on
    the placeholder.
  - `FormulaEvaluator::{evaluate_cell, evaluate_sheet}` stay `async` as
    thin wrappers. After a pass they await the caller's `Fetch` for every
    recorded URL, store the responses in the evaluator (`None` for a failed
    or oversized response), and run the pass again. Values computed before
    the first request stay cached.
  - Each distinct URL is fetched once per evaluator.
  - Without the `web_functions` feature, or without a transport, the
    future completes on its first poll. `litchi-ops` drives it with a noop
    waker, as before.
  - The public `Fetch`/`FetchFuture`/`FetchResult` contract is unchanged.
- **New tests** (`litchi-eval` `transport_tests`):
  - a transport that suspends once per request;
  - a chained `WEBSERVICE(WEBSERVICE(…))` that needs two passes, with a
    count of requests;
  - `evaluate_sheet` over the same evaluator;
  - a failing request (`#VALUE!`) and a missing transport (`#CONNECT!`).

  The engine's unit tests run synchronously (`#[test]`). Their stub
  transport is polled at once by the test context.

## 2. F1 — typed errors on the sheet traits

### 2.1 Shape

```rust
pub trait Worksheet: Send + Sync {
    type Error: std::error::Error + Send + Sync + 'static;
    // Cell, Cells, Rows as in 0353; every read returns Result<_, Self::Error>
}
pub trait WorkbookTrait: Debug + Send + Sync {
    type Error: std::error::Error + Send + Sync + 'static;
    type Worksheet<'a>: Worksheet<Error = Self::Error> where Self: 'a;
    // …
}

#[non_exhaustive]
pub enum litchi_core::sheet::Error {
    Refused(litchi_core::edit::Refusal),           // not found, invalid, unsupported
    Limit { kind: LimitKind, observed: usize, maximum: usize },
    SourceChanged,
    Io,
    Malformed(&'static str),
    Internal(&'static str),
}
pub type Result<T, E = Error> = std::result::Result<T, E>;
```

- **Each implementation keeps its own typed error.** The coordinator asked
  for a concrete error type and for the source error to be carried "as a
  typed variant or by kind, not a box". `litchi-core` cannot name the
  format errors, so the traits carry each implementation's own error as an
  associated type:
  - `litchi_ops::Error` for the `Xlsx`, `Xlsb` and `Xls` handles and
    `WorkbookView`;
  - `litchi_xls::Error` for the eager XLS reader;
  - `litchi_xlsb::package::PackageError` for the eager XLSB workbook;
  - `litchi_sheet_text::Error` for the text workbooks (record 0354's leaf
    crate, merged during this change);
  - the new `litchi_core::sheet::Error` for the calc grid and the examples.

  A caller of the format handles gets the typed `litchi_ops::Error` with
  its code, format and leaf. Nothing is flattened, as goal constraint 4
  requires. The tests that used to downcast
  (`format/{xlsx,xlsb}/tests.rs`, `tests/workbook.rs`, the facade
  `xlsb_facade.rs` and the XLSB writer tests) now read the error directly.
- **The engine reports `sheet::Error`, which is content-free.** It has the
  0351 §2.4 shape and holds only static text and counts.
  `FormulaEvaluator` requires `sheet::Error: From<W::Error>`.
  `#[cold] #[inline(never)]` `From` impls classify the three format errors
  by code or variant with static reasons:
  - `litchi_ops::Error` (by `Code`, with the code's own static summary);
  - `litchi_xls::Error`;
  - the XLSB `PackageError` (every variant listed, nothing caught by a
    wildcard);
  - `litchi_sheet_text::Error` (every variant listed).

  The engine's own errors are static refusals:
  - an unindexed sheet;
  - a structured reference without row context, or naming a column the
    table lacks;
  - the `GEOMEAN`/`HARMEAN`/`VARA` argument errors;
  - mismatched array sizes.

  They used to be `format!` strings that could carry column names and
  error-cell text.
- **`litchi_eval::DepthLimit` is deleted.** A too-deep chain is
  `sheet::Error::Limit { kind: LimitKind::Depth, observed, maximum }`.

### 2.2 Unchanged behaviour in `litchi-ops` calc

- `Grid` (the evaluation snapshot) has `Error = sheet::Error`. `trip`
  still keeps the typed `litchi_ops::Error` for `compute`, and hands the
  engine its classification.
- `compute` maps the error kinds as before:
  - `Limit { Depth }` becomes `Code::LimitExceeded` (`LimitKind::Depth`);
  - the grid's missing-sheet refusal (a `const`, matched exactly) becomes
    `Code::NotFound` with the same message as before;
  - every other engine error is "unsupported", as before.
- Cases where the engine folds an error into a value, such as `GEOMEAN`
  reporting `#NUM!` on any argument error, therefore behave exactly as
  before.

### 2.3 Boxed results kept outside the sheet traits

`litchi-xlsx` used `litchi_core::sheet::Result` as a general boxed result,
not as a trait result: its connections, pivot, XML-map, threaded-comment,
sparkline and volatile-dependency reads (19 files). They keep their public
signatures through a local alias with the same
`Box<dyn Error + Send + Sync>` type. These are format APIs with their own
wave (W4), not a sheet-trait or per-cell path.

The facade text readers had the same problem while this change was in
progress. Record 0354 then moved them into `litchi-sheet-text` with a
typed `Error`; the trunk merge resolved them onto that crate, whose trait
`impl`s now name `Error = litchi_sheet_text::Error`.

## 3. `Box<dyn …>` inventory after 0353a

| Site | Status |
|---|---|
| sheet trait results, the engine's `Result`, every per-cell read | none |
| engine evaluation: per cell read, per formula, per sub-expression, per function call | none |
| `view::sheet::Worksheets` ODS refusal | `Option<litchi_ops::Error>` (was a boxed error) |
| `litchi_eval::{Fetch, FetchFuture, FetchError}` | kept: the public, caller-owned async transport, one future per distinct URL per evaluator, outside the evaluation core |
| `&dyn DispatchCtx`, `&dyn Fetch` | borrowed vtable references, no allocation |
| `litchi-xlsx` module results (§2.3) | kept, outside the sheet traits; W4 |

## 4. Measurements

### 4.1 Method

- **Harness.** `crates/litchi-ops/examples/sheet_dispatch_bench.rs` adds
  `recalculate_heavy/<format>`: `File::recalculate` of a sheet with
  `rows / 8` rows, where each row holds three formulas:
  - `=SUM(A1:A{n})`, a running total over a growing range;
  - `=IF(B{n}>A{n}*A{n}/4,B{n}/A{n},MAX(A1:B{n}))`;
  - `=AVERAGE(A1:C{n})+COUNT(B1:B{n})`.

  Each formula reads many cells, and the second and third reference other
  formulas. The XLS formula encoder has no `MOD`, so it is not used.
- **Profiling switch.** `SHEET_DISPATCH_ONLY=<scenario>` limits a process
  to one scenario, for callgrind.
- **Builds.** Release builds (fat LTO), each in its own target directory,
  of the same harness source at three commits:
  - **0353**: tip `d8c442a413`, SHA-256
    `79d743d2924a2513ad10bdc8ad62e054b82a22b8087b1a23c2bd14ee35e3e6d6`;
  - **trunk before 0353**: `5ff605692e`, with the one-line `CellIterator`
    import the old traits need, SHA-256
    `55067d6b227db49772c6dec9b1bdbb13b390751fd53538673ebe8adcf6da88f4`;
  - **candidate**: this branch, SHA-256
    `00b05af5d210794082b7e00a380f53bc5e84dd159903a792379be8728d362b19`.

  The binaries are in the coordinator scratch directory `sheet-dispatch/`.
- **Order.** Process-level ABBA, 12 rounds (24 processes per leg), pinned
  with `taskset -c 20`. Each process takes 15 samples per scenario. The
  tables give the median of the per-process medians in µs.

### 4.2 Results

| Scenario | 0353 | candidate | ratio | trunk | candidate | ratio |
|---|---:|---:|---:|---:|---:|---:|
| `cell/xls` | 105.3 | 92.3 | 0.876 | 130.2 | 92.4 | 0.709 |
| `cell/xlsb` | 891.6 | 776.3 | 0.871 | 1118.5 | 778.2 | 0.696 |
| `cell/xlsx` | 825.0 | 808.6 | 0.980 | 850.0 | 817.6 | 0.962 |
| `cell_value/xls` | 114.3 | 105.5 | 0.923 | 117.8 | 107.7 | 0.914 |
| `cell_value/xlsb` | 859.0 | 799.6 | 0.931 | 1165.7 | 800.3 | 0.686 |
| `cell_value/xlsx` | 817.9 | 821.2 | 1.004 | 845.9 | 829.6 | 0.981 |
| `cells/xls` | 8.4 | 8.2 | 0.974 | 23.4 | 8.0 | 0.344 |
| `cells/xlsb` | 309.2 | 297.0 | 0.961 | 628.8 | 296.0 | 0.471 |
| `cells/xlsx` | 290.6 | 294.9 | 1.015 | 317.4 | 294.8 | 0.929 |
| `evaluate/xls` | 555.5 | 557.0 | 1.003 | 599.7 | 556.9 | 0.929 |
| `evaluate/xlsb` | 1752.5 | 1756.4 | 1.002 | 2165.9 | 1764.4 | 0.815 |
| `evaluate/xlsx` | 5994.1 | 5622.9 | 0.938 | 6044.9 | 5644.3 | 0.934 |
| `recalculate/xls` | 3048.6 | 2793.4 | 0.916 | 3045.4 | 2799.2 | 0.919 |
| `recalculate/xlsb` | 5861.6 | 5368.5 | 0.916 | 5866.8 | 5356.7 | 0.913 |
| `recalculate/xlsx` | 5800.3 | 5331.0 | 0.919 | 5799.7 | 5330.7 | 0.919 |
| `recalculate_heavy/xls` | 2035.4 | 1785.4 | 0.877 | 2337.4 | 1781.0 | 0.762 |
| `recalculate_heavy/xlsb` | 6974.1 | 6030.6 | 0.865 | 8181.6 | 6021.8 | 0.736 |
| `recalculate_heavy/xlsx` | 7007.6 | 5996.1 | 0.856 | 8176.3 | 6006.5 | 0.735 |

Evidence: [`results/0353/abba-0353a-vs-0353.json`](results/0353/abba-0353a-vs-0353.json)
and [`results/0353/abba-0353a-vs-trunk.json`](results/0353/abba-0353a-vs-trunk.json).

### 4.3 The XLSX iteration residual

- **Size.** `cells/xlsx` is +1.5% against 0353 (−7% against trunk); over
  three ABBA runs of the landed code it ranged from +1.5% to +1.9%.
  - `cell_value/xlsx` ranged from +0.4% to +1.2%. Its callgrind
    instruction count is 0.4% lower than 0353's (220.1 M against 221.1 M).
  - `evaluate/xls` ranged from −0.5% to +0.3%, and `evaluate/xlsb` from
    −0.2% to +0.4%.

  These three are within the run-to-run spread.
- **Cause.** Callgrind, with collection limited to the measured closure, is
  deterministic: 137.7 M instructions for 0353 and 142.4 M for the
  candidate, about +28 per cell. A build with F2 alone (`21bb454bc8`)
  measures 137.7 M, so F1 is the cause. Under fat LTO, the item
  `Option<Result<view::sheet::Cell, litchi_ops::Error>>` is re-encoded
  through more stack temporaries than with the old boxed error: three
  48-byte copies per cell. A build with aligned functions and blocks
  reproduces it, so it is not an alignment artefact.
- **Variants tried, each built and measured.** Records of the attempts are
  in the scratch directory.
  - Building the XLSX item in place (`Cells::next_with`, inline): −2.9%
    instructions and −9% time on `cells/xlsx`. But the generated code of
    the shared `Cells::next` then made `cells/xls` +11–35%.
  - Moving the step for each format out of line: XLSX stays better, but the
    returned item is written with scalar stores and read with vector loads
    (a store-forwarding stall). Up to +250% on `cells/xls`, and +11% on
    `cells/xlsx` when only the XLSX step is moved.
  - A separate error field in the XLSX iterator: instructions −0.4%, time
    worse.
- **What landed** removes every larger loss instead:
  - the BIFF lookups read the leaf's `get_cell` directly;
  - the XLSB source-backed `value_at` and `cell` read `get_cell`;
  - error conversions are cold;
  - the small leaf `next`/`get_cell` and the view's `Cell` accessors are
    `#[inline]`.

  The XLSX iteration residual is the only scenario still above 0353, and
  it is below trunk. Follow-up **F3**: the XLSX iterator converts
  `xlsx::SourceCell` values by cloning (`convert_cell`). An owning
  conversion (an `xlsx::formula::Formula::into_text`-style accessor, W4's
  crate) would remove an allocation per formula cell, which is far larger
  than this residual.

## 5. Coordination

- **`litchi-sheet-text`** (record 0354, merged from trunk during this
  change). Its `Workbook`/`Worksheet` trait `impl`s gain `type Error =
  Error`; the `.into()` boxing of `MissingWorksheet` and `RowOutOfRange`
  is dropped; and `error.rs` gains the `From<Error> for
  litchi_core::sheet::Error` classification. The old
  `crates/litchi/src/sheet/text` edits are superseded by the move.
- **`litchi-xlsx`.** One `pub(crate) type BoxedResult<T>` in `lib.rs`
  replaces 19 imports of the old alias. Public signatures are unchanged.
- **`litchi-xls` and `litchi-xlsb`.** Each gains a `From<Error> for
  litchi_core::sheet::Error` classification, and `#[inline]` on the
  stored-cell `next`/`get_cell`.

## 6. Follow-ups

- **F3**: §4.3.
- **W4**: replace the boxed results kept in `litchi-xlsx` (§2.3) with
  typed module errors.
