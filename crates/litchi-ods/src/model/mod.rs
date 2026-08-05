//! Immutable spreadsheet-domain vocabulary.
#![allow(
    dead_code,
    reason = "model codecs are retained for the package parser migration"
)]

mod conditional_format;
mod consolidation;
mod data_pilot;
pub mod data_validation;
mod database_range;
mod detective;
pub mod hyperlink;
pub mod label_range;
pub mod names;
mod protection;
mod source;
mod sparkline;
mod structure;
mod style_protection;
mod table_template;
mod tracked_changes;

pub use litchi_odf_common::annotation::{AnnotationElement, AnnotationNode, CellAnnotation};

pub use conditional_format::{
    ConditionalColorScale, ConditionalColorScaleEntry, ConditionalCustomIcon, ConditionalDataBar,
    ConditionalDataBarEntry, ConditionalDateIs, ConditionalDateType, ConditionalFormat,
    ConditionalFormatCondition, ConditionalFormatEntryType, ConditionalFormatRule,
    ConditionalIconSet, ConditionalIconSetEntry, DataBarAxisPosition, IconSetType,
};
pub use consolidation::{Consolidation, ConsolidationUseLabels};
pub use data_pilot::*;
pub use database_range::*;
pub use detective::{
    CellDetective, DetectiveDirection, DetectiveHighlightedRange, DetectiveOperation,
    DetectiveOperationKind,
};
pub use litchi_odf_common::calculation::{Iteration, IterationStatus, NullDate, Settings};
pub use litchi_odf_common::rdf::{Graph, Object, Subject, Triple};
pub use protection::{Protection, ProtectionKey, SheetProtection, SheetProtectionOptions};
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
