//! Formula evaluation over workbooks opened with `litchi::open`.
//!
//! `litchi::eval::FormulaEvaluator` is generic over `WorkbookTrait`, which
//! the `WorkbookView` of `File::workbook` implements by static dispatch, so
//! the evaluator reads an opened workbook directly (M0a step S8c; these tests
//! evaluated the boxed workbook of the facade's generic workbook opener).

#![cfg(all(feature = "eval", feature = "xlsx"))]

mod aggregate_logical;
mod financial;
mod lookup_text;
mod routes;
