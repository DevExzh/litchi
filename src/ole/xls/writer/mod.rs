//! XLS file writing module
//!
//! This module provides comprehensive support for creating and modifying
//! Microsoft Excel files in the legacy binary format (.xls files).

/// BIFF8 record generation
pub(crate) mod biff;

/// Core XLS writer implementation
mod core;

/// Cell formatting (fonts, fills, borders)
pub mod formatting;

/// Formula tokenization
pub mod formula;

// Re-export public types
pub use core::{
    XlsCellValue, XlsConditionalFormat, XlsConditionalFormatType, XlsConditionalPattern,
    XlsDataValidation, XlsDataValidationOperator, XlsDataValidationType, XlsWriter,
};
pub use formatting::{
    BorderStyle, Borders, CellStyle, ExtendedFormat, Fill, FillPattern, Font, FormattingManager,
    HorizontalAlignment, VerticalAlignment,
};
pub use formula::{FormulaTokenizer, Ptg};
