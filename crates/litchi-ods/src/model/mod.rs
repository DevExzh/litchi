//! Immutable spreadsheet-domain vocabulary.
#![allow(
    dead_code,
    reason = "model codecs are retained for the package parser migration"
)]

mod calculation;
mod conditional_format;
mod consolidation;
mod data_pilot;
mod data_validation;
mod database_range;
mod detective;
mod hyperlink;
mod label_range;
mod named_expression;
mod protection;
mod source;
mod sparkline;
mod structure;
mod style_protection;
mod table_template;
mod tracked_changes;

pub use litchi_odf_common::annotation::{AnnotationElement, AnnotationNode, CellAnnotation};

pub use calculation::{
    CalculationIteration, CalculationNullDate, CalculationSettings, IterationStatus,
};
pub use conditional_format::{
    ConditionalColorScale, ConditionalColorScaleEntry, ConditionalCustomIcon, ConditionalDataBar,
    ConditionalDataBarEntry, ConditionalDateIs, ConditionalDateType, ConditionalFormat,
    ConditionalFormatCondition, ConditionalFormatEntryType, ConditionalFormatRule,
    ConditionalIconSet, ConditionalIconSetEntry, DataBarAxisPosition, IconSetType,
};
pub use consolidation::{Consolidation, ConsolidationUseLabels};
pub use data_pilot::*;
pub use data_validation::*;
pub use database_range::*;
pub use detective::{
    CellDetective, DetectiveDirection, DetectiveHighlightedRange, DetectiveOperation,
    DetectiveOperationKind,
};
pub use hyperlink::{CellHyperlink, HyperlinkActuate, HyperlinkShow};
pub use label_range::{LabelRange, LabelRangeOrientation};
pub use litchi_odf_common::rdf::{Graph, Object, Subject, Triple};
pub use named_expression::{
    FormulaNamespace, NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange,
    NamedRangeUsage,
};
pub use protection::{
    ProtectionKey, SheetProtection, SheetProtectionOptions, SpreadsheetProtection,
};
pub use source::{CellRangeSource, SheetTableSource, TableSourceMode};
pub use sparkline::*;
pub use structure::{
    Column, SheetPrintSettings, SheetStyle, SheetStyleUsage, TableGroup, TableRange,
    TableStructure, TableVisibility,
};
pub use style_protection::{
    CellStyleProtection, ConditionalCellStyle, ConditionalCellStyleRule, TableCellProtectionStyle,
};
pub use table_template::{TableTemplate, TableTemplateAxis, TableTemplateStyle};
pub use tracked_changes::*;
