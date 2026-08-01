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

/// Checked BIFF8 INDEX/DBCELL worksheet layout generation
pub mod row_blocks;

// Re-export public types
pub use crate::XlsEncryptionProfile;
pub use crate::{
    XlsAutoFilter12Criterion, XlsAutoFilter12Icon, XlsAutoFilter12IconSet, XlsAutoFilter12Operator,
    XlsAutoFilter12Value, XlsExternalTableField, XlsExternalTableMetadata, XlsExternalTableVersion,
    XlsListColumnId, XlsListObject, XlsListObjectColumn, XlsListObjectFeatureVersion,
    XlsListObjectId, XlsListObjectRange, XlsListObjectSourceMetadata, XlsListObjectStyleOptions,
    XlsListTotalAggregation, XlsTableAutoFilter12, XlsWebColumnType, XlsWebDefaultValue,
    XlsWebEditMode, XlsWebFieldInfo, XlsWebInvalidCell, XlsWebReadingOrder, XlsWebTableField,
    XlsWebTableMetadata, XlsXmlColumnMapping, XlsXmlDataType, XlsXmlTableField,
    XlsXmlTableMetadata,
};
pub use crate::{
    XlsConsolidation, XlsConsolidationBuiltInName, XlsConsolidationFile, XlsConsolidationFunction,
    XlsConsolidationRange, XlsConsolidationSource,
};
pub use crate::{XlsDefinedNameFutureRecords, XlsNameFnGrp12, XlsNamePublish};
pub use biff::{AutoFilterConditionWrite, write_cfex12_marker, write_cfheader};
pub use core::{
    PivotCacheValue, XlsAddInFunctionOptions, XlsCalculationSettings, XlsCellValue,
    XlsCommentAnchor, XlsCommentTextRunWrite, XlsCommentWriteOptions, XlsConditionalFormat,
    XlsConditionalFormat12Group, XlsConditionalFormat12Rule, XlsConditionalFormat12Type,
    XlsConditionalFormatGroup, XlsConditionalFormatOperator, XlsConditionalFormatRange,
    XlsConditionalFormatRule, XlsConditionalFormatType, XlsConditionalPattern,
    XlsCustomTableStyles, XlsDataValidation, XlsDataValidationErrorStyle,
    XlsDataValidationFormulaKind, XlsDataValidationImeMode, XlsDataValidationOperator,
    XlsDataValidationOptions, XlsDataValidationRange, XlsDataValidationTableOptions,
    XlsDataValidationType, XlsDdeOrOleItemOptions, XlsDdeOrOleLinkOptions, XlsDefinedName,
    XlsDefinedNameRecordOptions, XlsExternalCacheRowOptions, XlsExternalDefinedNameOptions,
    XlsExternalSheetOptions, XlsExternalWorkbookOptions, XlsFunctionGroupOptions, XlsGroupRect,
    XlsPageSetupOptions, XlsPivotDataItemConfig, XlsPivotFieldConfig, XlsPivotItemConfig,
    XlsPivotTableConfig, XlsShapeAnchor, XlsShapeColor, XlsShapeFill, XlsShapeGroupChild,
    XlsShapeGroupWrite, XlsShapeKind, XlsShapeLine, XlsShapeText, XlsShapeTextRun, XlsShapeWrite,
    XlsWorkbookEnvironmentOptions, XlsWorkbookWindowOptions, XlsWorksheetLayoutOptions, XlsWriter,
};
pub use formatting::{
    BorderStyle, Borders, CellStyle, ExtendedFormat, Fill, FillPattern, Font, FormattingManager,
    HorizontalAlignment, VerticalAlignment,
};
pub use formula::{FormulaTokenizer, Ptg};
pub use view::{
    XlsPaneMode, XlsViewScale, XlsWorksheetPaneOptions, XlsWorksheetSelectionOptions,
    XlsWorksheetViewOptions,
};
