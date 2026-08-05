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

/// Typed worksheet view-state writing options
pub mod view;

/// Checked extended range-sort configuration.
pub mod sort;

/// Checked worksheet-shape anchors and primitive shape values.
pub use core::shape;

/// Checked BIFF8 INDEX/DBCELL worksheet layout generation
pub mod row_blocks;

// Re-export public types
pub use crate::EncryptionProfile;
pub use crate::{
    AutoFilter12Criterion, AutoFilter12Icon, AutoFilter12IconSet, AutoFilter12Operator,
    AutoFilter12Value, ExternalTableField, ExternalTableMetadata, ExternalTableVersion,
    ListColumnId, ListObject, ListObjectColumn, ListObjectFeatureVersion, ListObjectId,
    ListObjectRange, ListObjectSourceMetadata, ListObjectStyleOptions, ListTotalAggregation,
    TableAutoFilter12, WebColumnType, WebDefaultValue, WebEditMode, WebFieldInfo, WebInvalidCell,
    WebReadingOrder, WebTableField, WebTableMetadata, XmlColumnMapping, XmlDataType, XmlTableField,
    XmlTableMetadata,
};
pub use crate::{
    Consolidation, ConsolidationBuiltInName, ConsolidationFile, ConsolidationFunction,
    ConsolidationRange, ConsolidationSource,
};
pub use crate::{DefinedNameFutureRecords, NameFnGrp12, NamePublish};
pub use biff::{AutoFilterConditionWrite, write_cfex12_marker, write_cfheader};
pub use core::{
    AddInFunctionOptions, CalculationSettings, CellValue, CommentTextRunWrite, CommentWriteOptions,
    ConditionalFormat, ConditionalFormat12Group, ConditionalFormat12Rule, ConditionalFormat12Type,
    ConditionalFormatGroup, ConditionalFormatOperator, ConditionalFormatRange,
    ConditionalFormatRule, ConditionalFormatType, ConditionalPattern, CustomTableStyles,
    DataValidation, DataValidationErrorStyle, DataValidationFormulaKind, DataValidationImeMode,
    DataValidationOperator, DataValidationOptions, DataValidationRange, DataValidationTableOptions,
    DataValidationType, DdeOrOleItemOptions, DdeOrOleLinkOptions, DefinedName,
    DefinedNameRecordOptions, ExternalCacheRowOptions, ExternalDefinedNameOptions,
    ExternalSheetOptions, ExternalWorkbookOptions, FunctionGroupOptions, PageSetupOptions,
    PivotCacheValue, PivotDataItemConfig, PivotFieldConfig, PivotItemConfig, PivotTableConfig,
    ShapeColor, ShapeFill, ShapeGroupChild, ShapeGroupWrite, ShapeKind, ShapeLine, ShapeText,
    ShapeTextRun, ShapeWrite, WorkbookEnvironmentOptions, WorkbookWindowOptions,
    WorksheetLayoutOptions, Writer,
};
pub use formatting::{
    BorderStyle, Borders, CellStyle, ExtendedFormat, Fill, FillPattern, Font, FormattingManager,
    HorizontalAlignment, VerticalAlignment,
};
pub use formula::{Area, FormulaTokenizer, Ptg, Ref};
