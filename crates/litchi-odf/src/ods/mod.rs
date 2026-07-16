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
//! - ✅ Database ranges, recursive filters, sort keys, and subtotal rules
//! - ✅ Inert database query/table/SQL source metadata
//! - ✅ What-if scenarios and inert external cell-range sources
//! - ✅ Inert formula-auditing highlights and operations (`table:detective`)
//! - ✅ Formula calculation, null-date, and iteration settings
//! - ✅ Row and column label ranges
//! - ✅ Inert spreadsheet consolidation declarations
//! - ✅ Inert DDE source declarations and document-stored cached tables
//! - ✅ Data-pilot (pivot-table) sources, fields, levels, references, and groups
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
//! - ✅ Database ranges, filters, sorting, subtotals, and inert source metadata
//! - ✅ Create, edit, remove, and round-trip scenarios and external range links
//! - ✅ Create, edit, remove, and round-trip cell formula-auditing metadata
//! - ✅ Create, edit, clear, and round-trip calculation settings
//! - ✅ Create, edit, remove, and round-trip row/column label ranges
//! - ✅ Create, edit, clear, and round-trip consolidation declarations
//! - ✅ Create, edit, remove, and round-trip inert DDE caches
//! - ✅ Create, edit, remove, and round-trip data-pilot tables
//!
//! ## 🚧 TODO - Advanced Features
//! - ⚠️ Chart creation and parsing (embedded charts)
//! - ⚠️ Conditional formatting
//! - ⚠️ Sparklines
//! - ⚠️ External data connections
//!
//! # References
//! - ODF Specification: §9 (Spreadsheet Content)
//! - odfpy: `odf/table.py`, `odf/chart.py`
//! - calamine: Spreadsheet parsing patterns
//! - ODF Toolkit: Simple API - Spreadsheet class

mod annotation;
mod builder;
mod calculation;
mod cell;
mod consolidation;
mod data_pilot;
mod data_validation;
mod database_range;
mod dde;
mod detective;
/// OpenFormula parsing and support
pub mod formula;
mod label_range;
mod mutable;
mod named_expression;
mod parser;
mod protection;
mod row;
mod scenario;
mod sheet_image;
mod sheet;
mod source;
mod spreadsheet;
mod structure;
mod style_protection;
mod table_template;
mod tracked_changes;

pub use annotation::{AnnotationElement, AnnotationNode, CellAnnotation};
pub use builder::SpreadsheetBuilder;
pub use calculation::{
    CalculationIteration, CalculationNullDate, CalculationSettings, IterationStatus,
};
pub use cell::{Cell, CellMatrixSpan, CellMerge, CellValue};
pub use consolidation::{Consolidation, ConsolidationUseLabels};
pub use data_pilot::{
    DataPilotDisplayInfo, DataPilotDisplayMemberMode, DataPilotField, DataPilotFieldReference,
    DataPilotGrandTotal, DataPilotGroup, DataPilotGroupBoundary, DataPilotGroupBy, DataPilotGroups,
    DataPilotLayoutInfo, DataPilotLayoutMode, DataPilotLevel, DataPilotMember,
    DataPilotOrientation, DataPilotReferenceMemberType, DataPilotReferenceType, DataPilotSortInfo,
    DataPilotSortMode, DataPilotSortOrder, DataPilotSource, DataPilotTable,
};
pub use data_validation::{
    ContentValidation, ValidationDisplayList, ValidationErrorMacro, ValidationErrorMessage,
    ValidationEventListener, ValidationMessage, ValidationMessageType,
    ValidationPresentationEventListener, ValidationPresentationSound,
    ValidationScriptEventListener,
};
pub use database_range::{
    DatabaseFilter, DatabaseOrientation, DatabaseRange, DatabaseSort, DatabaseSortKey,
    DatabaseSource, EmbeddedNumberBehavior, FilterCondition, FilterConditionSource, FilterDataType,
    FilterExpression, SortOrder, SubtotalField, SubtotalRule, SubtotalRules, SubtotalSortGroups,
};
pub use dde::{DdeConversionMode, DdeLink, DdeSource};
pub use detective::{
    CellDetective, DetectiveDirection, DetectiveHighlightedRange, DetectiveOperation,
    DetectiveOperationKind,
};
pub use label_range::{LabelRange, LabelRangeOrientation};
pub use mutable::MutableSpreadsheet;
pub use named_expression::{
    FormulaNamespace, NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange,
    NamedRangeUsage,
};
pub use protection::{
    ProtectionKey, SheetProtection, SheetProtectionOptions, SpreadsheetProtection,
};
pub use row::Row;
pub use scenario::SheetScenario;
pub use sheet::Sheet;
pub use source::{CellRangeSource, SheetTableSource, TableSourceMode};
pub use spreadsheet::Spreadsheet;
pub use structure::{
    Column, SheetPrintSettings, SheetStyle, SheetStyleUsage, TableGroup, TableRange,
    TableStructure, TableVisibility,
};
pub use style_protection::{ConditionalCellStyle, ConditionalCellStyleRule, CellStyleProtection};
pub use table_template::{TableTemplate, TableTemplateAxis, TableTemplateStyle};
pub use tracked_changes::{
    SpreadsheetCellContentChange, SpreadsheetChangeAcceptance, SpreadsheetChangeCutOff,
    SpreadsheetChangeDimension, SpreadsheetChangeInfo, SpreadsheetChangeMetadata,
    SpreadsheetDeletion, SpreadsheetInsertion, SpreadsheetMovement, SpreadsheetNestedDeletion,
    SpreadsheetTrackedCell, SpreadsheetTrackedCellAddress, SpreadsheetTrackedCellValue,
    SpreadsheetTrackedChange, SpreadsheetTrackedChanges, SpreadsheetTrackedRangeAddress,
};

// Re-export formula types for public API
#[allow(unused_imports)] // Public API exports
pub use formula::{CellRef, Formula, RangeRef, Token};
