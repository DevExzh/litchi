//! Mutable worksheet and workbook writer components for XLSX.

pub mod chart_sheet;
pub mod shape;
pub mod sheet;
pub mod strings;
pub mod styles;
pub mod table;
pub mod workbook;

// Re-export main types
pub use crate::xlsx::sheet_protection::WorksheetProtection as SheetProtection;
pub use chart_sheet::MutableChartSheet;
pub use shape::{
    XlsxConnectionEndSpec, XlsxConnectionShapeSpec, XlsxDrawingObjectSpec, XlsxGroupSpec,
    XlsxShapeSpec,
};
pub use sheet::{
    AutoFilter, CellComment, ConditionalFormat, ConditionalFormatType, DefinedNameBuiltIn,
    FreezePanes, HeaderFooter, Hyperlink, Image, MutableWorksheet, NamedRange, PageBreak,
    PageSetup, PageSetupProperties, RichTextRun,
};
pub use strings::MutableSharedStrings;
pub use styles::StylesBuilder;
pub use workbook::{MutableWorkbookData, WorkbookProtection};
