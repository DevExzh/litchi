//! OpenDocument Spreadsheet (.ods) implementation.
//!
//! This module provides comprehensive support for parsing, creating, and manipulating
//! OpenDocument Spreadsheet documents (.ods files), which are the open standard
//! equivalent of Microsoft Excel spreadsheets.
//!
//! # Implementation Progress
//!
//! ## ✅ Reading (`spreadsheet.rs`, `parser.rs`, `sheet.rs`, `cell.rs`) - COMPLETE
//! - ✅ `Spreadsheet::open()` - Load from file path
//! - ✅ `Spreadsheet::from_bytes()` - Load from memory
//! - ✅ `sheets()` - Get all sheets
//! - ✅ `sheet_by_name()` / `sheet_by_index()` - Access specific sheets
//! - ✅ `Sheet::cell()` - Access cells by A1 notation or row/col
//! - ✅ `Cell::value()` - Get cell value (String, Number, Boolean, Date, DateTime, Duration, %)
//! - ✅ `Cell::formula()` - Get cell formula
//! - ✅ `Cell::style()` - Get cell style
//! - ✅ `to_csv()` - Export to CSV format
//! - ✅ Repeated cell/row expansion
//! - ✅ Merged cell handling
//! - ✅ Metadata extraction
//! - ✅ Global and sheet-local named ranges and named expressions
//! - ✅ Cell annotations with metadata, rich text/lists, extensions, and drawing geometry
//!
//! ## ✅ Formula Support (`formula.rs`) - PARTIAL
//! - ✅ Formula string representation
//! - ✅ Basic formula parsing
//! - ⚠️ Formula evaluation (not implemented)
//! - ⚠️ Formula dependency tracking
//!
//! ## ✅ Writing (`builder.rs`, `mutable.rs`) - COMPLETE
//! - ✅ `SpreadsheetBuilder::new()` - Create new spreadsheets
//! - ✅ `add_sheet()` - Add sheets with names
//! - ✅ `set_cell_value()` - Set cell values (all types)
//! - ✅ `set_cell_formula()` - Set cell formulas
//! - ✅ `set_cell_style()` - Apply cell styling
//! - ✅ `insert_row()` / `delete_row()` - Row operations
//! - ✅ `insert_column()` / `delete_column()` - Column operations
//! - ✅ `save()` / `to_bytes()` - Write to file or bytes
//! - ✅ `MutableSpreadsheet` - Modify existing spreadsheets
//! - ✅ Create, replace, edit, remove, and round-trip cell annotations
//! - ✅ Content-validation definitions, prompts, cell bindings, and inert event metadata
//! - ✅ Document/sheet keys, direct cell protection flags, and LibreOffice permissions
//!
//! ## 🚧 TODO - Advanced Features
//! - ⚠️ Chart creation and parsing (embedded charts)
//! - ⚠️ Conditional formatting
//! - ⚠️ Pivot tables
//! - ⚠️ Cell-style protection resolution
//! - ⚠️ Filter and sort criteria
//! - ⚠️ Sparklines
//! - ⚠️ Data tables and scenarios
//! - ⚠️ External data connections
//!
//! # References
//! - ODF Specification: §9 (Spreadsheet Content)
//! - odfpy: `odf/table.py`, `odf/chart.py`
//! - calamine: Spreadsheet parsing patterns
//! - ODF Toolkit: Simple API - Spreadsheet class

mod annotation;
mod builder;
mod cell;
mod data_validation;
/// OpenFormula parsing and support
pub mod formula;
mod mutable;
mod named_expression;
mod parser;
mod protection;
mod row;
mod sheet;
mod spreadsheet;

pub use annotation::{AnnotationElement, AnnotationNode, CellAnnotation};
pub use builder::SpreadsheetBuilder;
pub use cell::{Cell, CellValue};
pub use data_validation::{
    ContentValidation, ValidationDisplayList, ValidationErrorMacro, ValidationErrorMessage,
    ValidationEventListener, ValidationMessage, ValidationMessageType,
    ValidationPresentationEventListener, ValidationPresentationSound,
    ValidationScriptEventListener,
};
pub use mutable::MutableSpreadsheet;
pub use named_expression::{
    FormulaNamespace, NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange,
    NamedRangeUsage,
};
pub use protection::{
    ProtectionKey, SheetProtection, SheetProtectionOptions, SpreadsheetProtection,
};
pub use row::Row;
pub use sheet::Sheet;
pub use spreadsheet::Spreadsheet;

// Re-export formula types for public API
#[allow(unused_imports)] // Public API exports
pub use formula::{CellRef, Formula, RangeRef, Token};
