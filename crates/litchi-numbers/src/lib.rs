//! Numbers semantic value models.
//!
//! Archive parsing, protobuf decoding, and package mutation remain owned by
//! the Numbers implementation. This crate starts the downward migration with
//! the dependency-free cell vocabulary used by Numbers, Pages table editing,
//! and the shared structured extractor.

#![forbid(unsafe_code)]

/// Cell-level Numbers vocabulary.
pub mod cell;
/// Immutable, archive-free Numbers document snapshots.
pub mod document;
/// Dependency-free formula vocabulary shared by Numbers, Pages, and Keynote.
pub mod formula;
/// Human-readable and checked positional selectors for Numbers objects.
pub mod selector;
/// Semantic sheet containers.
pub mod sheet;
/// Sparse semantic table vocabulary.
pub mod table;

pub use document::{
    DEFAULT_MAX_TEXT_BYTES, Document, Error as DocumentError, Limits as DocumentLimits,
    MAX_MATERIALIZED_CELLS, MAX_SHEETS, MAX_TABLES, Result as DocumentResult,
};
pub use formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};
pub use selector::{SheetSelector, TableSelector};
pub use sheet::{Builder as SheetBuilder, SelectorError as TableSelectorError, Sheet};
pub use table::dimension::{Dimension, Points, Size};
pub use table::title::Settings;
pub use table::topology::{ColumnDeletion, RowDeletion};
pub use table::{
    AddressError, Builder as TableBuilder, Cell, CellPosition, CellRange, CoordinateError,
    Dimensions, Error as TableError, Grid, GridBudget, InsertError, InsertResult, Position, Range,
    Table, View,
};
