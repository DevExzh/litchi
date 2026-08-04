//! Mutable worksheet and workbook writer components for XLSX.

pub mod chart_sheet;
pub mod shape;
pub mod sheet;
pub mod strings;
pub mod styles;
pub mod table;
pub mod workbook;

// Re-export main types
pub use crate::xlsx::conditional_formatting::{IconSet, Operator};
pub use crate::xlsx::shapes::Geometry;
pub use chart_sheet::MutableChartSheet;
pub use shape::{ConnectionEndSpec, ConnectionShapeSpec, DrawingObjectSpec, GroupSpec, ShapeSpec};
pub use sheet::{
    AutoFilter, CellComment, ConditionalFormat, ConditionalFormatType, DefinedNameBuiltIn,
    FreezePanes, HeaderFooter, Hyperlink, Image, MutableWorksheet, NamedRange, PageBreak,
    PageSetupProperties, ParseTokenError, RichTextRun, Visibility,
};
pub use strings::MutableSharedStrings;
pub use styles::StylesBuilder;
pub use workbook::{MutableWorkbookData, WorkbookProtection};
