//! Immutable spreadsheet-domain vocabulary.
#![allow(
    dead_code,
    reason = "model codecs are retained for the package parser migration"
)]

mod calculation;
mod consolidation;
mod detective;
mod label_range;
mod named_expression;
mod protection;
mod source;
mod structure;
mod style_protection;
mod table_template;

pub use calculation::{
    CalculationIteration, CalculationNullDate, CalculationSettings, IterationStatus,
};
pub use consolidation::{Consolidation, ConsolidationUseLabels};
pub use detective::{
    CellDetective, DetectiveDirection, DetectiveHighlightedRange, DetectiveOperation,
    DetectiveOperationKind,
};
pub use label_range::{LabelRange, LabelRangeOrientation};
pub use named_expression::{
    FormulaNamespace, NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange,
    NamedRangeUsage,
};
pub use protection::{
    ProtectionKey, SheetProtection, SheetProtectionOptions, SpreadsheetProtection,
};
pub use source::{CellRangeSource, SheetTableSource, TableSourceMode};
pub use structure::{
    Column, SheetPrintSettings, SheetStyle, SheetStyleUsage, TableGroup, TableRange,
    TableStructure, TableVisibility,
};
pub use style_protection::{
    CellStyleProtection, ConditionalCellStyle, ConditionalCellStyleRule, TableCellProtectionStyle,
};
pub use table_template::{TableTemplate, TableTemplateAxis, TableTemplateStyle};
