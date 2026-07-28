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

/// BIFF8 calculation and recalculation records.
mod calculation;
mod chart;

/// BIFF8 worksheet scenario manager records.
mod scenario;

/// Inert BIFF8 VBA project markers and object code names.
mod vba;

/// BIFF8 workbook-global environment and behavioral options.
mod environment;

/// Inert BIFF8 workbook access-provenance metadata.
mod access;

/// BIFF8 default table and PivotTable style catalog metadata.
mod table_styles;

/// BIFF8 global differential formatting records and typed XF properties.
mod differential_format;

/// BIFF8 extended shared-string table lookup index.
mod shared_string_index;

/// BIFF8 worksheet INDEX and DBCELL row-block lookup metadata.
mod row_block_index;

/// BIFF8 formula error-checking shared features.
mod formula_errors;

mod autofilter12;
/// BIFF8 worksheet tables and their List12 formatting records.
mod list_object;

/// BIFF8 workbook windows and stable sheet-tab identifiers.
mod workbook_view;

/// BIFF8 built-in and user-defined function categories.
mod function_group;

/// Inert BIFF8 supporting-book links and external cell caches.
mod external_link;

/// Inert BIFF8 worksheet data-consolidation directories and sources.
mod consolidation;

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

/// Opt-in tolerance for non-structural formatting defects.
mod leniency;

/// BIFF8 XF border and fill metadata.
mod border_fill;

/// BIFF8 worksheet row heights and column widths/formatting.
mod layout;

/// BIFF8 `BookExt` record: workbook extension flags.
mod book_ext;

/// BIFF8 `Table` record: what-if data tables.
mod data_table;

/// BIFF8 `XFExt` record: formatting property extensions for XF records.
mod xf_ext;

/// BIFF8 `StyleExt` record: cell-style extensions.
mod style_ext;

/// BIFF8 `Theme` record: the document theme.
mod theme;

/// BIFF8 `PhoneticInfo` record: phonetic-string format and visible ranges.
mod phonetic_info;

/// BIFF8 custom-view records (`UserBView`, `UserSViewBegin`, `UserSViewEnd`).
mod custom_view;

/// BIFF8 worksheet default dimensions and outline workspace metadata.
mod sheet_layout;

/// BIFF8 `SheetExt` record: sheet tab color and publish state.
mod sheet_ext;

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

/// Extended BIFF8 range-sort metadata (`SortData` and `SortCond12`).
mod sort_data;

/// BIFF8 `QUERYTABLE` sequence: typed, inert query tables and connections.
mod query_table;

/// Pivot table parsing (SXVIEW, SXVD, SXVI, SXDI, SXVS, SXPI)
#[forbid(unsafe_code)]
pub mod pivot_table;
#[forbid(unsafe_code)]
mod pivot_editor;
pub use pivot_editor::XlsPivotViewEditor;
mod ole_object;
pub use ole_object::{
    XlsCheckState, XlsDropDownStyle, XlsEditBoxValidation, XlsFormControl, XlsFtCblsData,
    XlsFtCmo, XlsFtEdoData, XlsFtGboData, XlsFtLbsData, XlsFtPictFmla, XlsFtPioGrbit,
    XlsFtRboData, XlsFtSbs, XlsLbsDropData, XlsLbsItem, XlsListBehaviorClass,
    XlsListSelectionType, XlsObjSubrecord, XlsObjectType, XlsOleObjectEditor,
    XlsOleObjectRecord,
};
pub use pivot_table::{
    PageFieldEntry, PivotAdditionalExtension, PivotAxis, PivotAxisField, PivotCache,
    PivotCacheDateGroupUnit, PivotCacheDateGrouping, PivotCacheDateTime,
    PivotCacheDiscreteGrouping, PivotCacheError, PivotCacheField, PivotCacheGrouping,
    PivotCacheItem, PivotCacheNumericGrouping, PivotDataItem, PivotFunction, PivotItemType,
    PivotLayoutLine, PivotPageSelection, PivotQueryTag, PivotSourceType, PivotTable, PivotViewDef, PivotViewEx9,
    PivotViewExtension, PivotViewField, PivotViewFieldExtension, PivotViewItem,
};

/// Sheet protection parsing (PROTECT, OBJECTPROTECT, SCENPROTECT, PASSWORD)
pub mod protection;

/// XLS file writing
pub mod writer;

pub use access::{XlsWriteAccess, XlsWriteAccessEncoding};
pub use alignment::{
    XlsCellAlignment, XlsHorizontalAlignment, XlsReadingOrder, XlsTextRotation,
    XlsVerticalAlignment,
};
pub use autofilter12::{
    AUTO_FILTER12_RECORD_TYPE, XlsAutoFilter12Criterion, XlsAutoFilter12DateGroup,
    XlsAutoFilter12DateLevel, XlsAutoFilter12DifferentialFormat, XlsAutoFilter12DynamicType,
    XlsAutoFilter12FormatKind, XlsAutoFilter12Icon, XlsAutoFilter12IconSet,
    XlsAutoFilter12Operator, XlsAutoFilter12Value, XlsTableAutoFilter12,
};
pub use border_fill::{XlsBorderSide, XlsBorderStyle, XlsCellBorders, XlsCellFill, XlsFillPattern};
pub use calculation::{
    XlsCalculationMode, XlsMultithreadedCalculation, XlsReferenceMode, XlsWorkbookCalculation,
    XlsWorksheetCalculation,
};
pub use chart::*;
pub use cell::XlsCell;
pub use comments::CommentVisibility;
pub use conditional_format::{
    XlsConditionalAlignment, XlsConditionalBorder, XlsConditionalComparison,
    XlsConditionalExtension, XlsConditionalFont, XlsConditionalFormatRange,
    XlsConditionalFormatting, XlsConditionalFormatting12, XlsConditionalNumberFormat,
    XlsConditionalPattern, XlsConditionalProtection, XlsConditionalRule, XlsConditionalRule12,
    XlsConditionalRule12Kind, XlsConditionalRuleKind, XlsConditionalStyle,
};
pub use consolidation::{
    XlsConsolidation, XlsConsolidationBuiltInName, XlsConsolidationFile, XlsConsolidationFunction,
    XlsConsolidationRange, XlsConsolidationSource,
};
pub use data_validation::{
    XlsDataValidationErrorStyle, XlsDataValidationFormula, XlsDataValidationImeMode,
    XlsDataValidationKind, XlsDataValidationOperator, XlsDataValidationRange,
    XlsDataValidationRule, XlsDataValidationSettings,
};
pub use defined_names::{
    XlsBuiltInName, XlsDefinedName, XlsDefinedNameFutureRecords, XlsDefinedNameKind,
    XlsNameFnGrp12, XlsNamePublish, XlsNameScope,
};
pub use differential_format::{
    XlsDifferentialFormat, XlsThemeColor, XlsXfBorder, XlsXfColor, XlsXfColorSource,
    XlsXfFontScheme, XlsXfFontWeight, XlsXfGradient, XlsXfGradientStop, XlsXfProperties,
    XlsXfProperty,
};
pub use environment::{XlsLinkUpdateMode, XlsObjectDisplayMode, XlsWorkbookEnvironment};
pub use error::{XlsEncryptionKind, XlsError, XlsResult};
pub use encryption::XlsEncryptionProfile;
pub use external_link::{
    XlsDdeOleValueMatrix, XlsExternalCacheRow, XlsExternalCachedError, XlsExternalCachedValue,
    XlsExternalClipboardFormat, XlsExternalLinks, XlsExternalName, XlsExternalNameBody,
    XlsExternalSheet, XlsExternalSheetReference, XlsExternalWorkbook, XlsSupportingBook,
};
pub use font::{XlsFont, XlsFontCharset, XlsFontEscapement, XlsFontFamily, XlsFontUnderline};
pub use formula_errors::{
    XlsFormulaErrorChecks, XlsFormulaErrorFeature, XlsFormulaErrorHeader, XlsFormulaErrorRange,
};
pub use function_group::{XlsBuiltInFunctionCategories, XlsFunctionGroups};
pub use layout::{XlsColumnLayout, XlsRowLayout};
pub use leniency::{XlsFormattingDefect, XlsLeniency, XlsToleranceReport, XlsToleratedDefect};
pub use list_object::{
    XlsCachedDiskHeader, XlsExternalTableField, XlsExternalTableMetadata, XlsExternalTableVersion,
    XlsListColumnId, XlsListObject, XlsListObjectColumn, XlsListObjectFeatureVersion,
    XlsListObjectId, XlsListObjectRange, XlsListObjectSourceMetadata, XlsListObjectStyleOptions,
    XlsListTotalAggregation, XlsOpaqueListObjectFeature, XlsOpaqueListObjectFutureRecord,
    XlsWebColumnType, XlsWebDefaultValue, XlsWebEditMode, XlsWebFieldInfo, XlsWebInvalidCell,
    XlsWebReadingOrder, XlsWebTableField, XlsWebTableMetadata, XlsXmlColumnMapping, XlsXmlDataType,
    XlsXmlTableField, XlsXmlTableMetadata,
};
pub use number_format::{
    XlsDateSystem, XlsEffectiveExtendedFormat, XlsExtendedFormat, XlsExtendedFormatApplications,
    XlsExtendedFormatKind, XlsFormatting, XlsNumberFormat,
};
pub use page_setup::{
    XlsHeaderFooter, XlsPageBreak, XlsPageSetup, XlsPrintComments, XlsPrintErrors, XlsPrintOrder,
    XlsPrintOrientation, XlsPrintSetup,
};
pub use palette::{XlsColor, XlsPalette};
pub use query_table::{
    XlsHtmlFormatting, XlsOleDbConnection, XlsQueryParameter, XlsQueryParameterType,
    XlsQuerySource, XlsQueryTable, XlsTextCodePage, XlsTextDelimiter, XlsTextField,
    XlsTextFieldFormat, XlsTextQuery,
};
pub use records::{
    PhoneticAlignment, PhoneticRun, PhoneticString, PhoneticType, SharedStringFormatRun,
    SharedStringProperties,
};
pub use row_block_index::{
    XlsDbCellRecord, XlsIndexedRow, XlsRowBlock, XlsRowBlockIndex, XlsWorksheetIndexRecord,
};
pub use scenario::{XlsScenario, XlsScenarioCell, XlsScenarioManager, XlsScenarioRange};
pub use shapes::XlsShape;
pub use shared_string_index::{XlsSharedStringBucket, XlsSharedStringIndex};
pub use xf_ext::{XlsExtProp, XlsFullColorExt, XlsFullColorType, XlsXfExt};
pub use style_ext::{XlsStyleCategory, XlsStyleExt};
pub use theme::XlsTheme;
pub use phonetic_info::{
    XlsPhoneticAlignment, XlsPhoneticFormat, XlsPhoneticInfo, XlsPhoneticRange, XlsPhoneticType,
};
pub use custom_view::{
    XlsChartSheetCustomViewBegin, XlsCustomViewHiddenRows, XlsCustomViewNoteDisplay,
    XlsCustomViewTopLeft, XlsSheetCustomView, XlsSheetCustomViewBegin, XlsSheetCustomViewEnd,
    XlsWorkbookCustomView,
};
pub use data_table::{
    XlsDataTable, XlsDataTableInputCell, XlsDataTableKind, XlsDataTableRange,
};
pub use book_ext::{
    XlsBookExt, XlsBookExtConditional11, XlsBookExtConditional12, XlsFactoidDisplay,
};
pub use sheet_ext::{XlsSheetExt, XlsSheetExtOptional};
pub use sheet_layout::XlsWorksheetLayout;
pub use sheet_metadata::{XlsSheetKind, XlsSheetMetadata, XlsSheetVisibility};
pub use sort_data::{
    CONTINUE_FRT12_RECORD_TYPE, SORT_DATA_RECORD_TYPE, XlsDifferentialFormatIndex,
    XlsSortCondition, XlsSortData, XlsSortIcon, XlsSortIconSet, XlsSortMethod, XlsSortOn,
    XlsSortOrientation, XlsSortParent, XlsSortRange, parse_sort_data,
};
pub use table_styles::{
    XlsDifferentialFormatId, XlsTableStyle, XlsTableStyleElement, XlsTableStyleRegion,
    XlsTableStyles,
};
pub use vba::{XlsVbaMetadata, XlsVbaProjectStorage};
pub use view::{XlsPane, XlsPaneType, XlsSelection, XlsSelectionRange, XlsWorksheetView};
pub use workbook::{XlsOpenOptions, XlsWorkbook};
pub use workbook_view::{XlsWorkbookView, XlsWorkbookWindow};
pub use worksheet::XlsWorksheet;
pub use writer::{
    XlsShapeAnchor, XlsShapeColor, XlsShapeFill, XlsShapeKind, XlsShapeLine, XlsShapeText,
    XlsShapeTextRun, XlsShapeWrite, XlsWriter,
};
