//! Regressions of the crate-private cell-value engine (record 0359a).
//!
//! These suites drove `cell_values::{Snapshot, Edit, Commit, Patch}` and the
//! semantic patch replay as integration tests. Record 0359a makes the engine
//! crate-private (the public routes are `Package::edit` and the package
//! semantic routes), so the suites compile here as unit tests of the same
//! code. The crate is named `litchi_xls` inside them (`extern crate self`),
//! so their paths are unchanged.

#[path = "../tests/internal/xls_cell_values.rs"]
mod xls_cell_values;
#[path = "../tests/internal/xls_cross_workbook_scalar.rs"]
mod xls_cross_workbook_scalar;
#[path = "../tests/internal/xls_formula_ancillary.rs"]
mod xls_formula_ancillary;
