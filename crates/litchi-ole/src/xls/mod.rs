//! Legacy Excel (.xls) file format reader
//!
//! This module provides functionality to parse Microsoft Excel files
//! in the legacy binary format (.xls files), which are OLE2-based files.
//! The implementation is based on the BIFF (Binary Interchange File Format)
//! specification and draws inspiration from other spreadsheet libraries.

/// Error types for XLS parsing
mod error;

/// BIFF8 password-to-open encryption support
mod encryption;

/// BIFF record parsing utilities
pub mod records;

/// Workbook parsing implementation
mod workbook;

/// Worksheet parsing implementation
mod worksheet;

/// Cell value parsing and representation
mod cell;

/// BIFF8 worksheet data-validation records.
mod data_validation;

/// BIFF formula token rendering
mod formula;

/// Internal workbook and sheet defined names (`Lbl`).
mod defined_names;

/// BIFF8 number formats, XF slots, and workbook date system.
mod number_format;

/// BIFF8 workbook custom and default color palettes.
mod palette;

/// BIFF8 workbook font table.
mod font;

/// BIFF8 XF cell and style alignment metadata.
mod alignment;
/// BIFF8 XF border and fill metadata.
mod border_fill;

/// BIFF8 worksheet row heights and column widths/formatting.
mod layout;

/// BIFF8 worksheet window, zoom, pane, and selection state.
mod view;

/// BIFF8 worksheet print and page setup.
mod page_setup;

/// Legacy BIFF8 conditional formatting.
mod conditional_format;

/// Workbook sheet directory metadata.
mod sheet_metadata;

/// Shape extraction
pub mod shapes;

/// Shared parsing utilities
mod utils;

/// Merged cell range parsing (MERGECELLS 0x00E5)
pub mod merged_cells;

/// Hyperlink parsing (HLINK 0x01B8)
pub mod hyperlinks;

/// Comment/note parsing (NOTE 0x001C)
pub mod comments;

/// AutoFilter and sort parsing (AUTOFILTERINFO 0x009D, AUTOFILTER 0x009E, SORT 0x0090)
pub mod autofilter;

/// Pivot table parsing (SXVIEW, SXVD, SXVI, SXDI, SXVS, SXPI)
pub mod pivot_table;

/// Sheet protection parsing (PROTECT, OBJECTPROTECT, SCENPROTECT, PASSWORD)
pub mod protection;

/// XLS file writing
pub mod writer;

pub use cell::XlsCell;
pub use data_validation::{
    XlsDataValidationErrorStyle, XlsDataValidationFormula, XlsDataValidationImeMode,
    XlsDataValidationKind, XlsDataValidationOperator, XlsDataValidationRange,
    XlsDataValidationRule, XlsDataValidationSettings,
};
pub use error::{XlsEncryptionKind, XlsError, XlsResult};
pub use defined_names::{
    XlsBuiltInName, XlsDefinedName, XlsDefinedNameKind, XlsNameScope,
};
pub use number_format::{
    XlsDateSystem, XlsEffectiveExtendedFormat, XlsExtendedFormat,
    XlsExtendedFormatApplications, XlsExtendedFormatKind, XlsFormatting, XlsNumberFormat,
};
pub use palette::{XlsColor, XlsPalette};
pub use font::{XlsFont, XlsFontCharset, XlsFontEscapement, XlsFontFamily, XlsFontUnderline};
pub use alignment::{
    XlsCellAlignment, XlsHorizontalAlignment, XlsReadingOrder, XlsTextRotation,
    XlsVerticalAlignment,
};
pub use border_fill::{
    XlsBorderSide, XlsBorderStyle, XlsCellBorders, XlsCellFill, XlsFillPattern,
};
pub use layout::{XlsColumnLayout, XlsRowLayout};
pub use view::{XlsPane, XlsPaneType, XlsSelection, XlsSelectionRange, XlsWorksheetView};
pub use page_setup::{
    XlsPageSetup, XlsPrintComments, XlsPrintErrors, XlsPrintOrder, XlsPrintOrientation,
    XlsPrintSetup,
};
pub use conditional_format::{
    XlsConditionalAlignment, XlsConditionalBorder, XlsConditionalComparison,
    XlsConditionalFont, XlsConditionalFormatRange, XlsConditionalFormatting,
    XlsConditionalNumberFormat, XlsConditionalPattern, XlsConditionalProtection,
    XlsConditionalRule, XlsConditionalRuleKind, XlsConditionalStyle,
};
pub use sheet_metadata::{XlsSheetKind, XlsSheetMetadata, XlsSheetVisibility};
pub use records::{
    PhoneticAlignment, PhoneticRun, PhoneticString, PhoneticType, SharedStringFormatRun,
    SharedStringProperties,
};
pub use shapes::XlsShape;
pub use workbook::{XlsOpenOptions, XlsWorkbook};
pub use worksheet::XlsWorksheet;
pub use writer::XlsWriter;
