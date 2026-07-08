//! Internal helpers for parsing simple Excel-style formulas.
//!
//! This module is intentionally small and focused. It currently provides
//! minimal support for Phase 1 of the evaluator:
//!
//! - Literal constants: numbers, strings, booleans.
//! - Single-cell references in A1 notation, with optional sheet qualifiers.
pub mod ast;
pub mod expr;
pub mod literal;
pub mod reference;
pub mod structured_ref;

pub use ast::{BinaryOp, Expr, RangeRef};
pub use expr::parse_expression;
pub use reference::{parse_range_reference, parse_single_cell_reference};
pub use structured_ref::{StructuredReference, parse_structured_reference};
