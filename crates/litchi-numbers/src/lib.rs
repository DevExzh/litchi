//! Numbers semantic value models.
//!
//! This crate owns archive-free Numbers semantics and bounded native package
//! reading. [`Package::document`] exposes the strict rooted workbook, while
//! [`compatibility_tables_from_bytes`] is an explicitly allocating global
//! projection for historical structured-data migration, including detached
//! table models.
//!
//! The native cell wire codec is intentionally excluded from the supported
//! API. The semantic cell API remains available through [`cell`].
//!
//! ```compile_fail,E0603
//! use litchi_numbers::cell::wire::BncCell;
//! ```
//!
//! # Reorder sheets
//!
//! Use the direct [`sheet::order`] namespace for the transaction types. A move
//! selects a sheet in the immutable source snapshot and supplies its final
//! zero-based position: conceptually, `remove(source); insert(destination,
//! sheet)`. Only destinations already in the source sequence are valid.
//! Selecting position `n` and supplying destination `n` is an exact no-op.
//!
//! ```no_run
//! use litchi_numbers::{Package, SheetSelector};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::open("input.numbers")?;
//! let commit = package
//!     .edit_sheet_order()
//!     .move_sheet(SheetSelector::name("Summary"), 0)?
//!     .commit()?;
//!
//! // Changed moves require all three canonical previews and remove them.
//! // Patches authorize only their retained exact source; this inverse restores
//! // the original artifact and its previews exactly.
//! let restored = commit
//!     .package()
//!     .apply_sheet_order(&commit.patch().inverse())?;
//!
//! let mut output = Vec::new();
//! restored.package().write_to(&mut output)?;
//! # Ok(())
//! # }
//! ```

#![forbid(unsafe_code)]

/// Cell-level Numbers vocabulary.
pub mod cell;
/// Immutable, archive-free Numbers document snapshots.
pub mod document;
/// Dependency-free formula vocabulary shared by Numbers, Pages, and Keynote.
pub mod formula;
/// Atomic, preservation-safe Numbers sheet and table name transactions.
pub mod names;
/// Native Numbers package parsing and semantic projection.
pub mod package;
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
pub use litchi_iwa_common::table::title::Settings;
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::__compatibility_tables_from_prepared_source;
pub use package::{
    Error as PackageError, Limits as PackageLimits, MAX_OBJECTS, MAX_REFERENCES, Package,
    PayloadLimitKind as PackagePayloadLimitKind, ReadOptions as PackageReadOptions,
    ResourceError as PackageResourceError, Result as PackageResult, SemanticLimitKind,
    SemanticLimits as PackageSemanticLimits, SemanticLimitsError as PackageSemanticLimitsError,
    SemanticPath as PackageSemanticPath, TableLockCommit, TableLockDiagnostics, TableLockEdit,
    TableLockError, TableLockLimitKind, TableLockPatch, WriteError,
    compatibility_tables_from_bytes, compatibility_tables_from_bytes_with_options,
};
pub use selector::{SheetSelector, TableSelector};
pub use sheet::{Builder as SheetBuilder, SelectorError as TableSelectorError, Sheet};
pub use table::dimension::{Dimension, Points, Size};
pub use table::topology::{ColumnDeletion, RowDeletion};
pub use table::{
    AddressError, Builder as TableBuilder, Cell, CellPosition, CellRange, CoordinateError,
    Dimensions, Error as TableError, Grid, GridBudget, InsertError, InsertResult, Position, Range,
    Table, View,
};
