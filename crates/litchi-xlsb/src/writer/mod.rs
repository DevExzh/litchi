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
//! use litchi_xlsb::writer::{WorkbookWriter, MutableWorksheet};
//! use std::fs::File;
//!
//! // Create a new workbook
//! let mut workbook = WorkbookWriter::new();
//!
//! // Create a worksheet
//! let mut sheet = MutableWorksheet::new("Sheet1");
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
pub(crate) mod shape;
mod styles;

/// Mutable worksheet with CRUD operations
mod worksheet;

/// Workbook writer for creating complete XLSB files
mod workbook;

/// Typed chart-sheet authoring
mod chartsheet;

/// Binary cell range serialization helpers (shared by DV and CF writers)
#[allow(dead_code, unreachable_pub)]
pub(crate) mod bin_range;

/// Data validation writer (BrtBeginDVals / BrtDVal / BrtEndDVals)
#[allow(dead_code, unreachable_pub)]
pub(crate) mod data_validation;

// Re-export main types for public API
pub use crate::chart::{Anchor as ChartAnchor, Chart};
pub use crate::external_link::{
    AreaReference, CachedValue, CellLocation, CellReference, DdeItem, DefinedName, ErrorValue,
    Kind, Link, MAX_XLSB_EXTERNAL_CACHE_COLUMNS, MAX_XLSB_EXTERNAL_CACHE_ROWS,
    MAX_XLSB_EXTERNAL_CACHED_VALUES, NameFormula, NameFormulaKind, OleItem, SheetRange,
    ValueMatrix,
};
pub use crate::package::drawing_image::{Image, ImageFormat};
pub use crate::pivot_view::Part;
pub use crate::shapes::{
    Autofit, Body, CellMarker, Columns, Coordinate32, Direction, EditAs, Emu, EmuExtent, EmuOffset,
    Geometry, GroupTransform, Insets, Paragraph, Preset, Properties, Run, ShapeAnchor, TextSize,
    Underline, VerticalAnchor, Wrap,
};
pub use chartsheet::MutableChartSheet;
pub use shape::{ConnectionEndSpec, ConnectionShapeSpec, DrawingObjectSpec, GroupSpec, ShapeSpec};
pub use shared_strings::MutableSharedStringsWriter;
pub use styles::{DxfStyle, StylesWriter};
pub use workbook::WorkbookWriter;
pub use worksheet::{CellData, MutableWorksheet, SheetProtection};
