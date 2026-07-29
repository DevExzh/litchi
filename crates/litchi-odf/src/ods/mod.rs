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
//! - ✅ Standard `style:map` conditional cell styles and inert LibreOffice
//!   `calcext` conditional formats (condition, color-scale, data-bar,
//!   icon-set, and date-is rules)
//! - ✅ Inert LibreOffice `calcext` sparkline groups
//!
//! ## ✅ Formula Support (`formula.rs`) - PARTIAL
//! - ✅ Formula string representation
//! - ✅ Basic formula parsing
//! - ✅ Immutable `WorkbookTrait` adapter for shared formula evaluation
//! - ⚠️ Formula dependency tracking
//!
//! ## ✅ Writing (`builder.rs`, `mutable.rs`) - COMPLETE
//! - ✅ `SpreadsheetBuilder::new()` - Create new spreadsheets
//! - ✅ `add_sheet()` - Add sheets with names
//! - ✅ `set_cell_value()` - Set cell values (all types)
//! - ✅ `set_cell_hyperlink()` / `add_cell_hyperlink()` - Add inert full-cell or text-range hyperlinks
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
//! - ✅ Create, edit, remove, and round-trip conditional formats
//! - ✅ Create, edit, remove, and round-trip sparkline groups
//!
//! ## 🚧 TODO - Advanced Features
//! - ⚠️ Chart creation and parsing (embedded charts)
//! - ⚠️ `calcext` sparkline complex (theme) colors
//! - ⚠️ External data connections
//!
//! # References
//! - ODF Specification: §9 (Spreadsheet Content)
//! - odfpy: `odf/table.py`, `odf/chart.py`
//! - calamine: Spreadsheet parsing patterns
//! - ODF Toolkit: Simple API - Spreadsheet class

pub(crate) mod annotation;
mod builder;
pub(crate) mod calculation;
mod cell;
mod conditional_format;
mod consolidation;
pub(crate) mod data_pilot;
mod data_validation;
pub(crate) mod database_range;
mod dde;
mod detective;
mod evaluation;
/// OpenFormula parsing and support
pub mod formula;
mod hyperlink;
mod rich_text;
mod label_range;
mod mutable;
pub(crate) mod named_expression;
pub(crate) mod parser;
mod protection;
mod row;
mod scenario;
mod shape;
mod sheet;
mod sheet_image;
mod source;
mod sparkline;
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
pub use conditional_format::{
    ConditionalColorScale, ConditionalColorScaleEntry, ConditionalDataBar, ConditionalDataBarEntry,
    ConditionalDateIs, ConditionalDateType, ConditionalFormat, ConditionalFormatCondition,
    ConditionalFormatEntryType, ConditionalFormatRule, ConditionalIconSet, ConditionalIconSetEntry,
    DataBarAxisPosition, IconSetType,
};
pub use consolidation::{Consolidation, ConsolidationUseLabels};
pub use data_pilot::{
    DataPilotDisplayInfo, DataPilotDisplayMemberMode, DataPilotField, DataPilotFieldReference,
    DataPilotGrandTotal, DataPilotGrandTotalElement, DataPilotGrandTotalOrientation,
    DataPilotGroup, DataPilotGroupBoundary, DataPilotGroupBy, DataPilotGroups,
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
pub use evaluation::{OdsWorkbook, normalize_open_formula};
pub use hyperlink::CellHyperlink;
pub use rich_text::CellTextContent;
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
pub use shape::{SheetShape, SheetShapeAnchor};
pub use sheet::Sheet;
pub use source::{CellRangeSource, SheetTableSource, TableSourceMode};
pub use sparkline::{
    Sparkline, SparklineAxisType, SparklineColors, SparklineEmptyCells, SparklineFlags,
    SparklineGroup, SparklineType,
};
pub use spreadsheet::Spreadsheet;
pub use structure::{
    Column, SheetPrintSettings, SheetStyle, SheetStyleUsage, TableGroup, TableRange,
    TableStructure, TableVisibility,
};
pub use style_protection::{
    CellStyleProtection, ConditionalCellStyle, ConditionalCellStyleRule, TableCellProtectionStyle,
};
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
