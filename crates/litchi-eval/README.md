# litchi-eval

Spreadsheet formula evaluation engine for the Litchi office-formats library.

## Overview

`litchi-eval` is the engine behind `=SUM(A1:A10)` and friends in `.xlsx`, `.xlsb`, `.ods`, and `.numbers` workbooks parsed by Litchi. It operates on top of `litchi-core`'s `WorkbookTrait`, prefers cached results embedded in files when present, and falls back to evaluating the formula text. It is used through the `eval_engine` feature of the `litchi` umbrella crate.

## Usage

```toml
[dependencies]
litchi-eval = "0.0.1"
```

```rust
use litchi_eval::FormulaEvaluator;
use litchi_core::sheet::WorkbookTrait;

async fn sum_a1(workbook: &impl WorkbookTrait) -> litchi_core::sheet::Result<()> {
    let evaluator = FormulaEvaluator::new(workbook);
    let value = evaluator.evaluate_cell("Sheet1", 1, 1).await?;
    println!("A1 = {:?}", value);
    Ok(())
}
```

## Features

- `FormulaEvaluator::evaluate_cell` and `evaluate_sheet` for per-cell or whole-sheet evaluation.
- Defined names and Excel-style structured table references via `define_name`, `define_name_local`, `define_table`.
- Circular-reference detection and result caching.
- Optional `web_functions` feature gates network-bound functions (uses `reqwest`).

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
