//! Excel (.xlsx) spreadsheet support.
//!
//! This module provides parsing and manipulation of Microsoft Excel spreadsheets
//! in the Office Open XML (OOXML) format (.xlsx files).
//!
//! # Architecture
//!
//! The module follows a similar pattern to other OOXML modules:
//!
//! - `Workbook`: The main workbook content and API
//! - `Worksheet`: Individual sheet content and data access
//! - Various internal parsers for styles, shared strings, etc.
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ooxml::xlsx::Workbook;
//!
//! // Open a workbook
//! let workbook = Workbook::open("workbook.xlsx")?;
//!
//! // Access worksheets
//! for worksheet in workbook.worksheets() {
//!     println!("Sheet: {}", worksheet.name());
//!
//!     // Access cells
//!     let cell = worksheet.cell(1, 1)?;
//!     println!("A1 value: {:?}", cell.value());
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod cell;
pub mod chart;
pub mod format;
pub mod parsers;
pub mod pivot;
pub mod shared_strings;
pub mod sort;
pub mod sparkline;
pub mod styles;
pub mod table;
pub mod template;
pub mod threaded_comments;
pub mod views;
pub mod workbook;
pub mod worksheet;
pub mod writer;

// Re-export main types for convenience
pub use cell::Cell;
pub use chart::{ChartAnchor, WorksheetChart};
// Re-export shared formatting types
pub use format::{
    CellBorder, CellBorderLineStyle, CellBorderSide, CellFill, CellFillPatternType, CellFont,
    CellFormat, DataValidation, DataValidationOperator, DataValidationType,
};
pub use shared_strings::SharedStrings;
pub use sort::{SortBy, SortCondition, SortMethod, SortState};
pub use sparkline::{
    Sparkline, SparklineAxisMinMax, SparklineColor, SparklineDisplayEmptyCellsAs, SparklineGroup,
    SparklineGroupColors, SparklineGroupOptions, SparklineType,
};
pub use styles::{Alignment, Border, BorderStyle, CellStyle, Fill, Font, NumberFormat, Styles};
pub use table::{Table, TableColumn, TableFormula, TableStyleInfo, TableType, TotalsRowFunction};
pub use views::{SheetView, SheetViewType};
pub use workbook::Workbook;
pub use worksheet::{
    AutoFilter, ColumnInfo, Comment, ConditionalFormatRule, DataValidationRule, Hyperlink,
    PageBreak, PageSetup, RowInfo, Worksheet, WorksheetInfo,
};
// Re-export pivot types
pub use pivot::{
    AxisType, DataField, FieldItem, ItemType, Location, PageField, PivotArea, PivotCacheDefinition,
    PivotCacheField, PivotCacheRecords, PivotField, PivotFilter, PivotTableDefinition,
    PivotTableStyle, Reference, RowColField, RowColItem, SharedItem, SortType, Subtotal,
    read_pivot_cache_definition, read_pivot_table_definition, read_pivot_tables,
    write_pivot_cache_definition, write_pivot_cache_records, write_pivot_table,
};
// Re-export writer types
pub use writer::{
    AutoFilter as WriterAutoFilter, CellComment as WriterCellComment, ConditionalFormat,
    ConditionalFormatType, FreezePanes, HeaderFooter, Hyperlink as WriterHyperlink, Image,
    MutableSharedStrings, MutableWorkbookData, MutableWorksheet, NamedRange,
    PageBreak as WriterPageBreak, PageSetup as WriterPageSetup, RichTextRun, SheetProtection,
    StylesBuilder, WorkbookProtection,
};
// Re-export threaded comments types
pub use threaded_comments::{
    Mention, Person, PersonList, ThreadedComment, ThreadedComments, read_persons,
    read_threaded_comments, write_persons, write_threaded_comments,
};
