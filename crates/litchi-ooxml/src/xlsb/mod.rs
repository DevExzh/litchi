//! Excel XLSB (.xlsb) file format reader and writer
//!
//! This module provides comprehensive functionality to read and write Microsoft Excel XLSB files,
//! which are Excel 2007+ binary format files stored in ZIP containers.
//! The implementation follows the [MS-XLSB] specification.
//!
//! # Features
//!
//! ## Reading
//! - **Cell Data**: Read all cell types including numbers, strings, booleans, errors, and formulas
//! - **Formulas**: Parse formula records with cached values (FMLA_STRING, FMLA_NUM, FMLA_BOOL, FMLA_ERROR)
//! - **Styles**: Parse fonts, fills, borders, number formats, and cell formats from styles.bin
//! - **Shared Strings**: Efficient shared string table parsing
//! - **Workbook Structure**: Parse workbook properties, sheet metadata, and relationships
//! - **Column Information**: Support for column widths, hidden columns, and custom widths
//! - **Merged Cells**: Parse and handle merged cell ranges
//! - **Hyperlinks**: Parse hyperlink data with locations and tooltips
//! - **Named Ranges**: Parse defined names and ranges
//! - **Chart Sheets**: Typed chart sheet stream parsing (code name, tab color, views,
//!   protection, page setup, drawing links)
//! - **Drawings and Charts**: SpreadsheetDrawing inventory (anchors, shapes, pictures,
//!   graphic frames) with embedded charts parsed through the shared typed chart model
//! - **Error Handling**: Comprehensive error types with detailed context
//!
//! ## Writing
//! - **Workbook Creation**: Create complete XLSB files with multiple worksheets
//! - **Cell Writing**: Write all cell types (numbers, strings, booleans, errors, dates)
//! - **Styles**: Write custom fonts, fills, borders, and number formats
//! - **Shared Strings**: Automatic shared string table management
//! - **Advanced Features**: Write merged cells, hyperlinks, and comments
//! - **CRUD Operations**: Full support for creating, reading, updating, and deleting cells
//!
//! # Supported Record Types
//!
//! The module supports parsing of 100+ record types from the MS-XLSB specification, including:
//! - Cell records (blank, RK, error, bool, real, string, ISST)
//! - Formula records (string, numeric, boolean, error)
//! - Style records (fonts, fills, borders, number formats, XF)
//! - Workbook structure records (sheets, properties, views)
//! - Worksheet records (dimensions, columns, rows)
//! - Advanced features (hyperlinks, merged cells, named ranges, comments)
//!
//! # Examples
//!
//! ## Reading an XLSB File
//!
//! ```rust,no_run
//! use litchi_ooxml::xlsb::XlsbWorkbook;
//! use litchi_core::sheet::WorkbookTrait;
//! use std::fs::File;
//!
//! // Open an XLSB file
//! let file = File::open("workbook.xlsb")?;
//! let workbook = XlsbWorkbook::new(file)?;
//!
//! // Access worksheets
//! for name in workbook.worksheet_names() {
//!     println!("Sheet: {}", name);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ## Writing an XLSB File
//!
//! ```rust,no_run
//! use litchi_ooxml::xlsb::writer::{XlsbWorkbookWriter, MutableXlsbWorksheet};
//! use litchi_ooxml::xlsb::merged_cells::MergedCell;
//! use std::fs::File;
//!
//! // Create a new workbook
//! let mut workbook = XlsbWorkbookWriter::new();
//!
//! // Create a worksheet
//! let mut sheet = MutableXlsbWorksheet::new("Sheet1");
//! sheet.set_cell(0, 0, "Hello");
//! sheet.set_cell(0, 1, "World");
//! sheet.set_cell(1, 0, 42.0);
//! sheet.set_cell(1, 1, true);
//!
//! // Add merged cells
//! sheet.add_merged_cell(MergedCell::new(2, 3, 0, 1));
//!
//! // Add worksheet to workbook
//! workbook.add_worksheet(sheet);
//!
//! // Save to file
//! let file = File::create("output.xlsb")?;
//! workbook.save(file)?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Performance Considerations
//!
//! - Zero-copy parsing where possible using `Bytes` and borrows
//! - Lazy loading of worksheets (worksheets are parsed on-demand)
//! - Efficient binary record parsing with variable-length encoding
//! - UTF-16LE string decoding using `encoding_rs`
//! - Optimized memory layout for cache efficiency
//! - Minimal allocations and preference for move semantics
//!
//! # Reference
//!
//! - [MS-XLSB]: Excel Binary File Format (.xlsb) Structure Specification
//!   https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/

/// Error types for XLSB parsing
mod error;
mod frt;

/// XLSB record parsing utilities
mod records;

/// Workbook parsing implementation
mod workbook;

/// Worksheet parsing implementation
mod worksheet;

mod calculation;
/// Cell value parsing and representation
mod cell;

/// XLSB cells reader
mod cells_reader;

/// Shared parsing utilities
mod utils;

mod shared_strings;
/// Styles parsing (fonts, fills, borders, number formats)
mod styles;
mod styles_table;
pub mod vba_project;

/// Date and time utilities
pub mod date_utils;

/// Writer modules for creating XLSB files
pub mod writer;

/// Merged cell support
pub mod merged_cells;

/// Hyperlink support
pub mod hyperlinks;

/// Comment support
pub mod comments;

/// Named range support
pub mod named_ranges;

/// Data validation support
pub mod data_validation;

/// Conditional formatting support
pub mod conditional_formatting;

/// PivotCache definition stream parsing (MS-XLSB 2.1.7.38)
pub mod pivot;

/// Table (ListObject) stream parsing (MS-XLSB 2.1.7.51)
pub mod table;
pub mod connections;
pub(crate) mod walker;

/// Chart sheet stream parsing (MS-XLSB 2.1.7.7)
pub mod chartsheet;

/// SpreadsheetDrawing XML inventory for XLSB Drawings parts (MS-XLSB 2.1.7.23)
pub mod drawing;
pub(crate) mod drawing_write;

/// Formula parsing and generation
pub mod formula;
pub mod web_extension_bindings;

pub use calculation::{CalculationMode, CalculationProperties};
pub use cell::XlsbCell;
pub use chartsheet::{
    XlsbChartSheet, XlsbChartSheetColor, XlsbChartSheetColorType, XlsbChartSheetPageSetup,
    XlsbChartSheetProtection, XlsbChartSheetState, XlsbChartSheetView, parse_chart_sheet_part,
};
pub use data_validation::{DataValidation, DataValidationRecordKind, DataValidationSettings};
pub use drawing::{
    CHART_GRAPHIC_DATA_URI, XlsbDrawing, XlsbDrawingAnchor, XlsbDrawingAnchorKind,
    XlsbDrawingCellMarker, XlsbDrawingEmuPoint, XlsbDrawingEmuSize, XlsbDrawingGraphicFrame,
    XlsbDrawingNonVisual, XlsbDrawingObject, XlsbEmbeddedChart, XlsbSheetDrawing,
    parse_drawing_part,
};
pub use error::{XlsbError, XlsbResult};
pub use formula::{XlsbExternalLink, XlsbExternalLinkKind};
pub use pivot::{
    CalculatedItem, CalculatedMember, CalculatedMemberExt14, PivotCacheConsolidationPage,
    PivotCacheConsolidationSet, PivotCacheConsolidationSource, PivotCacheDateTime,
    PivotCacheDefinition, PivotCacheDefinitionExt14, PivotCacheDiscreteGrouping,
    PivotCacheErrorCode, PivotCacheField, PivotCacheFieldGrouping, PivotCacheGroupBy,
    PivotCacheGroupingGroup, PivotCacheGroupingGroupMember, PivotCacheGroupingLevel,
    PivotCacheHierarchy, PivotCacheHierarchyExt14, PivotCacheItem, PivotCacheItemInfo,
    PivotCacheItemValue, PivotCacheRange, PivotCacheRangeGrouping, PivotCacheSharedItems,
    PivotCacheSharedItemsStats, PivotCacheSource, PivotCacheSourceType, PivotCacheTupleCache,
    PivotCacheTupleCacheSet, PivotCacheWorksheetSource, PivotName, PivotNameFunction,
    PivotNamePair, PivotParsedFormulaData, PivotRuleFilter, parse_pivot_cache_definition,
};
pub use shared_strings::{
    PhoneticAlignment, PhoneticRun, PhoneticString, PhoneticType, SharedString, SharedStringRun,
};
pub use styles::{
    Alignment, Border, BorderSide, BorderStyle, HorizontalAlignment, VerticalAlignment,
};
pub use styles_table::{CellFormat, Fill, Font, StylesTable};
pub use table::{
    XlsbTable, XlsbTableColumn, XlsbTableFormula, XlsbTableRange, XlsbTableStyleInfo,
    XlsbTableTotalsRowFunction, XlsbTableType, parse_table_part,
};
pub use vba_project::{VbaProject, VbaProjectSignature, VbaProjectSignatureKind};
pub use web_extension_bindings::{
    XlsbWebExtensionBinding, XlsbWebExtensionRange, parse_xlsb_web_extension_bindings,
    validate_xlsb_web_extension_apprefs, write_xlsb_web_extension_bindings,
};
pub use workbook::XlsbWorkbook;
pub use worksheet::{
    XlsbAutoFilter, XlsbColumnInfo, XlsbRowInfo, XlsbSheetProtection, XlsbStrongProtection,
    XlsbWorksheet,
};
// Re-export low-level record iterator types for diagnostics and advanced users.
// The actual implementations live in records.rs; this keeps the module private
// while still allowing external tooling (like examples) to traverse raw records.
pub use records::{XlsbRecord, XlsbRecordHeader, XlsbRecordIter};

// Re-export writer types for convenience
pub use writer::{
    MutableSharedStringsWriter, MutableXlsbWorksheet, RecordWriter, SheetProtection, StylesWriter,
    XlsbWorkbookWriter,
};
