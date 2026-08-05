//! Numbers semantic value models.
//!
//! Archive parsing, protobuf decoding, and package mutation remain owned by
//! the Numbers implementation. This crate starts the downward migration with
//! the dependency-free cell vocabulary used by Numbers, Pages table editing,
//! and the shared structured extractor.

#![forbid(unsafe_code)]

/// Cell-level Numbers vocabulary.
pub mod cell;
/// Dependency-free formula vocabulary shared by Numbers, Pages, and Keynote.
pub mod formula;
/// Semantic sheet containers.
pub mod sheet;
/// Sparse semantic table vocabulary.
pub mod table;

pub use formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};
pub use sheet::{Builder as SheetBuilder, Sheet};
pub use table::{
    Builder as TableBuilder, Cell, Dimensions, Error as TableError, Grid, GridBudget, InsertError,
    InsertResult, Position, Range, Table, View,
};
