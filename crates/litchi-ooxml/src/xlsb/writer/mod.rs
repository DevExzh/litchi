//! XLSB binary format writer modules
//!
//! This module provides functionality to write XLSB files (Excel Binary Workbook).
//! XLSB files are Excel 2007+ binary format files stored in ZIP containers.
//!
//! # Features
//!
//! - **Binary Record Writing**: Variable-length encoded records according to MS-XLSB spec
//! - **Workbook Writing**: Complete workbook structure with properties and sheets
//! - **Worksheet Writing**: Cell data with all types (numbers, strings, booleans, errors, formulas)
//! - **Shared Strings**: Efficient shared string table generation
//! - **Styles**: Fonts, fills, borders, and number formats
//! - **Advanced Features**: Comments, hyperlinks, merged cells, data validation, and typed worksheet drawings
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_ooxml::xlsb::writer::{XlsbWorkbookWriter, MutableXlsbWorksheet};
//! use std::fs::File;
//!
//! // Create a new workbook
//! let mut workbook = XlsbWorkbookWriter::new();
//!
//! // Create a worksheet
//! let mut sheet = MutableXlsbWorksheet::new("Sheet1");
//! sheet.set_cell(0, 0, "Hello");
//! sheet.set_cell(0, 1, 42.0);
//! sheet.set_cell(1, 0, true);
//!
//! // Add worksheet to workbook
//! workbook.add_worksheet(sheet);
//!
//! // Save to file
//! let file = File::create("output.xlsb")?;
//! workbook.save(file)?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

/// Shared strings table writer
mod shared_strings;

/// Styles writer (fonts, fills, borders, number formats)
mod styles;

/// Mutable worksheet with CRUD operations
mod worksheet;

/// Workbook writer for creating complete XLSB files
mod workbook;

/// Typed chart-sheet authoring
mod chartsheet;

/// Binary cell range serialization helpers (shared by DV and CF writers)
pub(crate) mod bin_range;

/// Data validation writer (BrtBeginDVals / BrtDVal / BrtEndDVals)
pub(crate) mod data_validation;

// Re-export main types for public API
pub use crate::xlsb::drawing_image::{Image, ImageFormat};
pub use crate::xlsb::external_link::{
    AreaReference, CachedValue, CellLocation, CellReference, DdeItem, DefinedName, ErrorValue,
    Kind, Link, MAX_XLSB_EXTERNAL_CACHE_COLUMNS, MAX_XLSB_EXTERNAL_CACHE_ROWS,
    MAX_XLSB_EXTERNAL_CACHED_VALUES, NameFormula, NameFormulaKind, OleItem, SheetRange,
    ValueMatrix,
};
pub use crate::xlsx::writer::{
    ConnectionEndSpec, ConnectionShapeSpec, DrawingObjectSpec, Geometry, GroupSpec, ShapeSpec,
};
pub use crate::xlsx::{
    BodyProperties, CellMarker, ChartAnchor, Columns, Coordinate32, EditAs, Emu, EmuExtent,
    EmuOffset, GroupTransform, Paragraph, Preset, Run, ShapeAnchor, TextInsets, TextSize,
    WorksheetChart, XlsxTextAutofit, XlsxTextDirection, XlsxTextUnderline, XlsxTextVerticalAnchor,
    XlsxTextWrap,
};
pub use chartsheet::MutableXlsbChartSheet;
pub use litchi_xlsb::pivot_view::Part;
pub use shared_strings::MutableSharedStringsWriter;
pub use styles::{DxfStyle, StylesWriter};
pub use workbook::XlsbWorkbookWriter;
pub use worksheet::{CellData, MutableXlsbWorksheet, SheetProtection};
