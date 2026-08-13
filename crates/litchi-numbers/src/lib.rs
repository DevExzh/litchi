//! Numbers semantic value models.
//!
//! This crate owns archive-free Numbers semantics, bounded native package
//! reading, and focused exact package transactions. [`Package::document`]
//! exposes the strict rooted workbook, while
//! [`compatibility_tables_from_bytes`] is an explicitly allocating global
//! projection for historical structured-data migration, including detached
//! table models.
//!
//! Native cell wire records are intentionally excluded from the supported API.
//! The selector-first semantic read and edit API is available through
//! [`table::cells`].
//!
//! ```compile_fail,E0603
//! use litchi_numbers::cell::wire::BncCell;
//! ```
//!
//! # Table-cell reads
//!
//! Use [`table::cells`] to read one cell or a bounded dense range from a
//! selector-first table. Sheet and table names are exact; duplicate names in
//! malformed input are rejected instead of selecting an arbitrary match.
//! Index selectors are checked zero-based positions. Single coordinates and
//! range endpoints are checked against the selected table.
//!
//! A [`table::cells::State`] keeps physical cell presence explicit:
//! [`table::cells::Storage::Missing`] differs from
//! [`table::cells::Storage::Stored`] containing [`cell::Value::Empty`]. Range
//! reads are half-open, row-major, and dense—each requested coordinate gets a
//! state, including missing cells. Their retained element count is bounded by
//! the package's semantic materialized-cell limit, so oversized ranges fail
//! before a partial result is returned.
//!
//! The read behavior is grounded in Apple-authored Numbers workbooks with
//! materialized text and number cells. Native IDs and raw archive payloads
//! remain private implementation details.
//!
//! ```no_run
//! use litchi_numbers::{CellPosition, CellRange, Package, SheetSelector, TableSelector};
//! use litchi_numbers::table::cells::Storage;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::open("input.numbers")?;
//! let sheet = SheetSelector::name("Summary");
//! let table = TableSelector::name("Revenue");
//!
//! let b2 = package.table_cell(sheet, table, CellPosition::from_a1("B2")?)?;
//! if let Storage::Stored(value) = b2.storage() {
//!     println!("B2 is a {} cell", value.cell_type().name());
//! }
//!
//! let range = CellRange::from_a1("B2:C3")?;
//! let states = package.table_cells(sheet, table, range)?;
//! assert_eq!(states.len(), 4); // Dense, row-major: B2, C2, B3, C3.
//! # Ok(())
//! # }
//! ```
//!
//! # Table-cell transactions
//!
//! Use [`Package::edit_table_cells`] to stage one bounded, selector-first
//! batch. [`table::cells::Input`] accepts plain text, Boolean values, and
//! finite number, date, and duration scalars; every [`table::cells::Change`]
//! is an explicit set or clear. Supported changed sources include admitted
//! sparse storage, in-place authored-text replacement in a uniquely owned
//! rich backing, and the narrow
//! final-overlay formula-cache subset. Formula construction and unsupported
//! rich-text ownership, formula graphs or dependencies outside that subset,
//! merged cells, and other dependencies fail with typed errors before
//! publication.
//!
//! The complete batch is atomic: selected storage, shared-string refcounts,
//! supported rich-text ownership, and affected final-overlay caches are
//! validated together. An unchanged batch is an exact no-op. A changed
//! [`table::cells::Patch`] privately retains its exact source/target package
//! pair as a process-local capability; [`Package::apply_table_cells`] accepts
//! it only for that exact source, and [`table::cells::Patch::inverse`] restores
//! the original package and any stale previews deleted in the forward
//! direction. Candidate verification proves focused storage/cache locality.
//!
//! Native Numbers 14.4 evidence covers the admitted rich-text no-impact case.
//! It does not establish a UI oracle for formulas whose displayed result is
//! affected by a cell edit. The legacy raw-ID migration-host cell writer has
//! been retired; host comment/reply APIs and formula authoring remain separate
//! migration-host scope.
//!
//! ```no_run
//! use litchi_numbers::{Package, SheetSelector, TableSelector};
//! use litchi_numbers::table::cells::Input;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::open("input.numbers")?;
//! let commit = package
//!     .edit_table_cells(SheetSelector::name("Summary"), TableSelector::name("Revenue"))?
//!     .set_a1("B3", Input::number(43.0)?)?
//!     .set_a1("C3", Input::text("updated")?)?
//!     .clear_a1("D3")?
//!     .commit()?;
//! let restored = commit
//!     .package()
//!     .apply_table_cells(&commit.patch().inverse())?;
//! let mut source_bytes = Vec::new();
//! package.write_to(&mut source_bytes)?;
//! let mut restored_bytes = Vec::new();
//! restored.package().write_to(&mut restored_bytes)?;
//! assert_eq!(restored_bytes, source_bytes);
//! # Ok(())
//! # }
//! ```
//!
//! # Table-title transactions
//!
//! Use [`table::title`] to read and transactionally update a single table
//! title. The public boundary is selector-first: choose a sheet with
//! [`SheetSelector`] and then choose a table on that sheet with
//! [`TableSelector`]. The API deliberately accepts neither native object IDs
//! nor raw package/archive bytes.
//!
//! [`table::title::Settings`] is lossless for the two native optional Boolean
//! fields. `None` means the field is absent; it is distinct from
//! `Some(false)`, which retains an explicitly stored false value. A commit
//! whose requested settings equal the source is an exact byte and shared
//! snapshot no-op. A changed commit is source-bound, removes every existing
//! canonical root preview, and its inverse restores the original artifact
//! exactly. Changed publication refuses an effectively locked table. When a
//! requested title is visible, the native title height and both canonical
//! paragraph and shape style prerequisites must be valid.
//!
//! ```no_run
//! use litchi_numbers::{Package, SheetSelector, TableSelector};
//! use litchi_numbers::table::title::Settings;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::open("input.numbers")?;
//! let sheet = SheetSelector::name("Summary");
//! let table = TableSelector::name("Revenue");
//!
//! let before = package.table_title_settings(sheet, table)?;
//! let commit = package
//!     .edit_table_title(sheet, table)?
//!     .set(Settings::new(Some(false), before.outlined()))
//!     .commit()?;
//!
//! // A patch is authorized only for its exact source. Its inverse restores
//! // the original package and any previews deleted by this changed commit.
//! let restored = commit
//!     .package()
//!     .apply_table_title(&commit.patch().inverse())?;
//! # let _ = restored;
//! # Ok(())
//! # }
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
    DEFAULT_MAX_TEXT_BYTES, Document, DocumentReadOptions, DocumentSourceLimitKind,
    DocumentSourceLimits, DocumentSourceLimitsError, Error as DocumentError, IoKind,
    LimitKind as DocumentLimitKind, Limits as DocumentLimits, LimitsError as DocumentLimitsError,
    MAX_MATERIALIZED_CELLS, MAX_SHEETS, MAX_TABLES, ReadError as DocumentReadError,
    ReadLimitKind as DocumentReadLimitKind, Result as DocumentResult, Stats as DocumentStats,
};
pub use formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::{
    __compatibility_tables_from_prepared_source, __semantic_document_from_prepared_source,
};
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
