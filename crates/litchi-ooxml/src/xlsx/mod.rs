//! Excel (.xlsx) spreadsheet support.
//!
//! This module provides parsing and manipulation of Microsoft Excel spreadsheets
//! in the Office Open XML (OOXML) format (.xlsx files).
//!
//! # Architecture
//!
//! The module follows a similar pattern to other OOXML modules:
//!
//! - `Workbook`: The main workbook content and API
//! - `Worksheet`: Individual sheet content and data access
//! - Various internal parsers for styles, shared strings, etc.
//!
//! # Example
//!
//! ```ignore
//! use litchi_ooxml::xlsx::Workbook;
//! use litchi_core::sheet::WorkbookTrait;
//!
//! // Open a workbook
//! let workbook = Workbook::open("workbook.xlsx")?;
//!
//! // Access worksheet names
//! for name in workbook.worksheet_names() {
//!     println!("Sheet: {}", name);
//! }
//! # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
//! ```

pub mod active_x;
pub mod auto_filter;
pub mod calculation_chain;
pub mod calculation_properties;
pub mod cell;
pub mod chart;
pub mod chartsheet;
mod comments;
pub mod conditional_formatting;
pub mod connections;
pub mod data_consolidation;
pub mod data_validation;
mod drawing;
pub mod external_links;
pub mod format;
pub mod header_footer;
pub mod ignored_errors;
pub mod named_sheet_view;
mod namespace;
pub mod ole_objects;
pub mod outline_properties;
pub mod page_margins;
pub mod page_setup;
pub mod parsers;
pub mod phonetic_properties;
pub mod pivot;
pub mod pivot_chart;
pub mod print_options;
pub mod printer_settings;
pub mod query_table;
pub mod revisions;
mod shared_formula;
pub mod shared_strings;
pub mod sheet_format;
pub mod sheet_properties;
pub mod sheet_protection;
pub mod sheet_view;
pub mod shapes;
pub mod slicer_cache;
pub mod slicers;
pub mod slicer_timeline_crud;
pub mod sort;
pub mod sparkline;
pub mod styles;
pub mod table;
pub mod template;
pub mod threaded_comments;
pub mod views;
pub mod volatile_dependencies;
pub mod vba_project;
pub mod workbook;
pub mod workbook_metadata;
pub mod workbook_protection;
pub mod worksheet;
pub mod xml_maps;
pub use chartsheet::{
    ChartSheet, ChartSheetBackgroundPicture, ChartSheetChartCompanionResource,
    ChartSheetChartEmbeddedPackageContentType, ChartSheetChartEmbeddedPackageResource,
    ChartSheetChartImageContentType, ChartSheetChartImageResource, ChartSheetChartOutboundResource,
    ChartSheetChartResource, ChartSheetChartResourceKind, ChartSheetChartThemeOverrideResource,
    ChartSheetChartUserShapeImageContentType, ChartSheetChartUserShapeImageResource,
    ChartSheetChartUserShapesResource, ChartSheetColor, ChartSheetConformance,
    ChartSheetCustomView, ChartSheetDrawingResource, ChartSheetEntry, ChartSheetExtension,
    ChartSheetExtensionList, ChartSheetExtensionRelationship,
    ChartSheetExtensionRelationshipTarget, ChartSheetHeaderFooter, ChartSheetImageContentType,
    ChartSheetMargins, ChartSheetPackage, ChartSheetPageSetup, ChartSheetPrinterSettings,
    ChartSheetProperties, ChartSheetProtection, ChartSheetState, ChartSheetView,
    ChartSheetVmlDrawingResource, ChartSheetWebPublishItem, ChartSheetWebPublishItems,
    ChartSheetWebSourceType, PageOrientation, load_chartsheet, parse_chartsheet, store_chartsheet,
    write_chartsheet,
};
pub mod writer;

// Re-export main types for convenience
pub use auto_filter::{
    AutoFilterDefinition, CalendarType, ColorFilter, CustomFilter, CustomFilterOperator,
    CustomFilters, DateGroupItem, DateTimeGrouping, DynamicFilter, DynamicFilterType,
    FilterColumnDefinition, FilterColumnPayload, FilterIconSet, FilterItem, FilterRange,
    FilterValues, IconFilter, SortConditionDefinition, SortStateDefinition, Top10Filter,
};
pub use calculation_chain::{
    CalculationCell, CalculationChain, CalculationChainConformance,
    CalculationChainExtensionAttribute, load_calculation_chain_from_package,
    parse_calculation_chain, remove_calculation_chain, store_calculation_chain,
};
pub use calculation_properties::{
    WorkbookCalculationMode, WorkbookCalculationProperties, WorkbookReferenceMode,
    parse_workbook_calculation_properties,
};
pub use cell::Cell;
pub use chart::{
    ChartAnchor, ChartExternalDataPart, ChartExternalDataTarget, ChartRelationship,
    ChartRelationshipTarget, ChartUserShapesPart, ChartUserShapesRelationship,
    ChartUserShapesRelationshipTarget, WorksheetChart,
};
pub use conditional_formatting::{
    CellRangeRef, ColorScale, ConditionalFormatOperator, ConditionalFormatPayload,
    ConditionalFormatValue, ConditionalFormatValueType, ConditionalFormatting,
    ConditionalFormattingRule, ConditionalFormattingSource, ConditionalRuleType, DataBar,
    DifferentialFormat, DifferentialFormatComponent, DifferentialFormatRef,
    DifferentialNumberFormat, ExtensionAssociation, IconSet, NamedColor, SpreadsheetColor,
};
pub use data_consolidation::{
    WorksheetDataConsolidation, WorksheetDataConsolidationConformance,
    WorksheetDataConsolidationFunction, WorksheetDataConsolidationRangeReference,
    WorksheetDataReference, WorksheetDataReferenceSource, WorksheetDataReferences,
    parse_worksheet_data_consolidation, write_worksheet_data_consolidation,
};
pub use data_validation::{
    DataValidationCollection, DataValidationConformance, DataValidationFormula,
    DataValidationRange, DataValidationSource,
    DataValidationSqref, ParsedDataValidation, ParsedDataValidationErrorStyle,
    ParsedDataValidationImeMode, ParsedDataValidationOperator, ParsedDataValidationType,
    ValidationListSource, parse_data_validation_collections,
    replace_data_validation_collections,
    validate_data_validation_collections, write_data_validation_collections,
};
pub use external_links::{
    ExternalCell, ExternalCellType, ExternalDdeItem, ExternalDdeLink, ExternalDdeValue,
    ExternalLinkConformance,
    ExternalDdeValueType, ExternalDdeValues, ExternalDefinedName, ExternalLinkEntry,
    ExternalLinkKind, ExternalOleItem, ExternalOleItemSource, ExternalOleLink, ExternalOleTarget,
    ExternalRow, ExternalSheetData, ExternalWorkbookLink, ExternalWorkbookTarget,
};
pub use query_table::{
    add_worksheet_query_table, find_worksheet_query_table, load_worksheet_query_tables,
    remove_worksheet_query_table, reorder_worksheet_query_tables, replace_worksheet_query_table,
    update_worksheet_query_table,
    QueryTable, QueryTableConformance, QueryTableExtensionAttribute, QueryTableExtensionList,
    QueryTableField, QueryTableGrowShrinkType, QueryTableIconSet, QueryTableRefresh,
    QueryTableSortBy, QueryTableSortCondition, QueryTableSortMethod, QueryTableSortState,
    WorksheetQueryTable, parse_query_table, write_query_table,
};
pub use revisions::{
    RevisionAttribute, RevisionAttributeNamespace, RevisionConformance, RevisionHeader,
    RevisionHeaderProperties, RevisionHeaders, RevisionLog, RevisionLogPart, RevisionRecord,
    RevisionRecordKind, RevisionUser, RevisionUsers, RevisionXmlElement, WorkbookRevisions,
    load_workbook_revisions, parse_revision_headers, parse_revision_log, parse_revision_users,
    store_workbook_revisions, write_revision_headers, write_revision_log, write_revision_users,
};
// Re-export shared formatting types
pub use format::{
    CellBorder, CellBorderLineStyle, CellBorderSide, CellFill, CellFillPatternType, CellFont,
    CellFormat, DataValidation, DataValidationOperator, DataValidationType,
};
pub use header_footer::{
    HeaderFooterSectionKind, HeaderFooterText, WorksheetHeaderFooter, parse_worksheet_header_footer,
};
pub use ignored_errors::{
    IgnoredErrorRangeReference, WorksheetIgnoredError, WorksheetIgnoredErrorType,
    WorksheetIgnoredErrors, WorksheetIgnoredErrorsExtension, parse_worksheet_ignored_errors,
};
pub use named_sheet_view::{
    NamedSheetView, NamedSheetViewColumnFilter, NamedSheetViewDifferentialFormat,
    NamedSheetViewExtension, NamedSheetViewFilter, NamedSheetViewGuid, NamedSheetViewIconSet,
    NamedSheetViewMarkup, NamedSheetViewRange, NamedSheetViewSortCondition,
    NamedSheetViewSortConditionKind, NamedSheetViewSortRule, NamedSheetViewSortRules,
    NamedSheetViews, load_worksheet_named_sheet_views, parse_named_sheet_views,
    remove_worksheet_named_sheet_views, store_worksheet_named_sheet_views, write_named_sheet_views,
};
pub use xml_maps::{
    XmlMap, XmlMapConformance, XmlMapDataBinding, XmlMapInfo, XmlMapSchema,
    load_from_package as load_xml_maps_from_package,
    load_from_package_with_conformance as load_xml_maps_from_package_with_conformance,
    remove_from_package as remove_xml_maps_from_package,
    store_in_package as store_xml_maps_in_package,
};
pub use volatile_dependencies::{
    VolatileDependencies, VolatileDependenciesConformance, VolatileDependencyType, VolatileMain,
    VolatileReference, VolatileTopic, VolatileType, VolatileValue,
    load_from_package as load_volatile_dependencies_from_package,
    load_from_package_with_conformance as load_volatile_dependencies_from_package_with_conformance,
    remove_from_package as remove_volatile_dependencies_from_package,
    store_in_package as store_volatile_dependencies_in_package,
};
pub use ole_objects::{
    OleObjectAnchor, OleObjectAspect, OleObjectConformance, OleObjectMarker, OleObjectProperties,
    OleObjectRelationshipKind, OleObjectResource, OleObjectTarget, OleObjectUpdate,
    WorksheetOleObject, WorksheetOleObjects, load_worksheet_ole_objects,
    parse_worksheet_ole_objects, store_worksheet_ole_objects, write_worksheet_ole_objects,
};
pub use outline_properties::{WorksheetOutlineProperties, parse_worksheet_outline_properties};
pub use page_margins::{PageMargin, WorksheetPageMargins, parse_worksheet_page_margins};
pub use page_setup::{
    PageSetupCellComments, PageSetupOrder, PageSetupOrientation, PageSetupPrintErrors,
    PositiveUniversalMeasure, UniversalMeasureUnit, WorksheetPageSetup,
    parse_complete_worksheet_page_setup,
};
pub use phonetic_properties::{
    WorksheetPhoneticAlignment, WorksheetPhoneticProperties, WorksheetPhoneticType,
    parse_worksheet_phonetic_properties,
};
pub use print_options::{WorksheetPrintOptions, parse_worksheet_print_options};
pub use printer_settings::{
    PrinterSettingsConformance, PrinterSettingsResource, WorksheetPrinterSettings,
    WorksheetPrinterSettingsReference, load_worksheet_printer_settings,
    parse_worksheet_printer_settings_reference, store_worksheet_printer_settings,
    write_worksheet_printer_settings_reference,
};
pub use shared_strings::SharedStrings;
pub use sheet_format::{WorksheetSheetFormatProperties, parse_worksheet_sheet_format_properties};
pub use sheet_properties::{
    WorksheetPageSetupProperties, WorksheetSheetProperties, WorksheetSynchronizationReference,
    WorksheetTabColor, parse_worksheet_sheet_properties,
};
pub use sheet_protection::{
    ProtectedRangeSource, ProtectionPasswordVerifier, ProtectionRangeReference,
    ProtectionRangeReferenceKind, ProtectionRangeSqref, StrongProtectionPasswordVerifier,
    WorksheetProtectedRange, WorksheetProtectedRangeCollection, WorksheetProtection,
    WorksheetProtectionConformance, WorksheetProtectionMetadata, parse_worksheet_protection,
    replace_worksheet_protection, validate_worksheet_protection_metadata,
    write_worksheet_protection,
};
pub use sheet_view::{
    PivotAreaType, PivotSelectionAxis, WorksheetCellReference, WorksheetPanePosition,
    WorksheetPaneState, WorksheetPivotArea, WorksheetPivotSelection, WorksheetRangeReference,
    WorksheetViewCollection, WorksheetViewDefinition, WorksheetViewExtension, WorksheetViewPane,
    WorksheetViewSelection, WorksheetViewSqref, WorksheetViewType, parse_worksheet_views,
};
pub use shapes::{
    XlsxAnchoredObject, XlsxCellMarker, XlsxClientData, XlsxConnectionShape, XlsxDrawingObject,
    XlsxDrawingOleObject, XlsxEditAs, XlsxEmu, XlsxEmuExtent, XlsxEmuOffset, XlsxGroupTransform,
    XlsxShape, XlsxShapeAnchor, XlsxShapeBodyProperties, XlsxShapeConnectionEnd,
    XlsxShapeGroup, XlsxShapeNonVisual, XlsxShapeParagraph, XlsxShapePreset, XlsxShapeRun,
    XlsxShapeTextBody, XlsxTextAutofit, XlsxTextDirection, XlsxTextInsets,
    XlsxTextVerticalAnchor, XlsxTextWrap, XlsxWorksheetShapes, load_shapes,
    load_worksheet_shapes, parse_drawing_shapes,
};
pub use slicer_cache::{
    SLICER_CACHE_CONTENT_TYPE, SLICER_CACHE_RELATIONSHIP_TYPE, SlicerCacheData,
    SlicerCacheDataKind, SlicerCacheDefinition, SlicerCacheExtensionList, SlicerCachePivotTable,
    WorkbookSlicerCache, load_slicer_caches, parse_slicer_cache_definition, store_slicer_cache,
    write_slicer_cache_definition,
};
pub use slicers::{
    SLICERS_CONTENT_TYPE, SLICERS_RELATIONSHIP_TYPE, Slicer, SlicerExtensionList, Slicers,
    WorksheetSlicers, load_worksheet_slicers, parse_slicers, store_worksheet_slicers,
    write_slicers,
};
pub use slicer_timeline_crud::{
    add_slicer, add_slicer_cache, add_timeline, add_timeline_cache, find_slicer,
    find_slicer_cache, find_timeline, find_timeline_cache, remove_slicer,
    remove_slicer_cache, remove_timeline, remove_timeline_cache, reorder_slicer_caches,
    reorder_slicers, reorder_timeline_caches, reorder_timelines, replace_slicer,
    replace_slicer_cache, replace_timeline, replace_timeline_cache, update_slicer,
    update_slicer_cache, update_timeline, update_timeline_cache,
};
pub use sort::{SortBy, SortCondition, SortMethod, SortState};
pub use sparkline::{
    Sparkline, SparklineAxisMinMax, SparklineColor, SparklineDisplayEmptyCellsAs, SparklineGroup,
    SparklineGroupColors, SparklineGroupOptions, SparklineType,
};
pub use styles::{Alignment, Border, BorderStyle, CellStyle, Fill, Font, NumberFormat, Styles};
pub use table::{
    Table, TableColumn, TableFormula, TableStyleInfo, TableType, TotalsRowFunction, parse_table_xml,
};
pub use views::{
    SheetPane, SheetPanePosition, SheetPaneState, SheetSelection, SheetView, SheetViewType,
};
pub use workbook::Workbook;
pub use workbook_protection::{WorkbookProtectionMetadata, parse_workbook_protection};
pub use vba_project::VbaProject;
pub use worksheet::{
    AutoFilter, ColumnInfo, Comment, ConditionalFormatRule, DataValidationRule, Hyperlink,
    PageBreak, PageSetup, RowInfo, Worksheet, WorksheetInfo,
};
// Re-export pivot types
pub use pivot::{
    AxisType, DataField, FieldItem, ItemType, Location, PageField, PivotArea, PivotCacheDefinition,
    PivotCacheField, PivotCacheRecords, PivotField, PivotFilter, PivotTableDefinition,
    PivotTableStyle, Reference, RowColField, RowColItem, SharedItem, SortType, Subtotal,
    read_pivot_cache_definition, read_pivot_cache_records, read_pivot_table_definition,
    read_pivot_tables, write_pivot_cache_definition, write_pivot_cache_records, write_pivot_table,
};
// Re-export pivot-chart binding types
pub use pivot_chart::{
    DEFAULT_PIVOT_CHART_FORMAT_ID, PIVOT_OPTIONS_EXTENSION_URI, PivotChart, PivotChartBinding,
    PivotChartDropZoneVisibility, PivotChartFieldType, PivotChartPivotOptions, PivotChartSeries,
    PivotChartSheetKind, PivotChartSource, WorksheetPivotCharts, load_pivot_charts,
    load_worksheet_pivot_charts, parse_pivot_chart_binding,
};
// Re-export writer types
pub use writer::{
    AutoFilter as WriterAutoFilter, CellComment as WriterCellComment, ConditionalFormat,
    ConditionalFormatType, DefinedNameBuiltIn, FreezePanes, HeaderFooter,
    Hyperlink as WriterHyperlink, Image, MutableSharedStrings, MutableWorkbookData,
    MutableWorksheet, NamedRange, PageBreak as WriterPageBreak, PageSetup as WriterPageSetup,
    PageSetupProperties as WriterPageSetupProperties, RichTextRun, SheetProtection, StylesBuilder,
    WorkbookProtection, XlsxConnectionEndSpec, XlsxConnectionShapeSpec, XlsxDrawingObjectSpec,
    XlsxGroupSpec, XlsxShapeSpec,
};
// Re-export threaded comments types
pub use threaded_comments::{
    Mention, Person, PersonList, ThreadedComment, ThreadedCommentGraph, ThreadedComments,
    WorkbookPersonPart, WorksheetThreadedCommentPart, add_threaded_comment,
    add_threaded_comment_person, add_threaded_comment_reply, find_threaded_comment,
    find_threaded_comment_person, load_threaded_comment_graph, read_persons,
    read_threaded_comments, remove_threaded_comment, remove_threaded_comment_person,
    reorder_threaded_comment_persons, reorder_threaded_comments, replace_threaded_comment,
    replace_threaded_comment_person, update_threaded_comment, update_threaded_comment_person,
    validate_threaded_comment_graph, write_persons, write_threaded_comments,
};
pub mod timelines;
pub use timelines::{
    PivotFilterType, TIMELINE_CACHE_CONTENT_TYPE, TIMELINE_CACHE_EXTENSION_URI,
    TIMELINE_CACHE_RELATIONSHIP_TYPE, TIMELINES_CONTENT_TYPE, TIMELINES_EXTENSION_URI,
    TIMELINES_RELATIONSHIP_TYPE, Timeline, TimelineCacheDefinition, TimelineCachePivotTable,
    TimelineLevel, TimelineOpaqueXml, TimelinePivotFilter, TimelineRange, TimelineState, Timelines,
    WorkbookTimelineCache, WorksheetTimelines, load_timeline_caches, load_timelines,
    parse_timeline_cache_definition, parse_timelines, store_timeline_caches,
    store_worksheet_timelines, write_timeline_cache_definition, write_timelines,
};
pub mod custom_data;
pub use custom_data::{
    CUSTOM_DATA_CONTENT_TYPE, CUSTOM_DATA_PROPERTIES_CONTENT_TYPE,
    CUSTOM_DATA_PROPERTIES_RELATIONSHIP_TYPE, CUSTOM_DATA_RELATIONSHIP_TYPE,
    CustomDataExtensionList, CustomDataPayload, CustomDataProperties, WorkbookCustomData,
    load_custom_data, parse_custom_data_properties, store_custom_data,
    write_custom_data_properties,
};
pub mod data_model;
pub use data_model::{
    DATA_MODEL_CONTENT_TYPE, DATA_MODEL_EXTENSION_URI, DATA_MODEL_PART_NAME, DataModelDefinition,
    DataModelOpaqueXml, DataModelPayload, DataModelRelationship, DataModelTable, WorkbookDataModel,
    load_data_model, parse_data_model, store_data_model, write_data_model,
};
pub mod xldm;
pub use xldm::{
    XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE, XldmBackupLog, XldmCompression, XldmFileEntry,
    XldmFileGroup, XldmFileGroupClass, XldmFileKind, XldmGeneratedNameKind, XldmGeneratedPath,
    XldmHeader, XldmLoggedFile, XldmOffset, XldmPartitionMarker, XldmSize, XldmStorage,
    XldmWriteAccess, XldmXmlEncoding, classify_xldm_generated_path, inspect_xldm, write_xldm,
};
pub mod xldm_native;
pub use xldm_native::{
    XldmDictionaryBody, XldmDictionaryFile, XldmDictionaryType, XldmHashBin, XldmHashHeader,
    XldmHashIndexFile, XldmHashStatistics, XldmHuffmanCharacterSetMode, XldmIdfFile,
    XldmIdfSegment, XldmNativeData, XldmNativeError, XldmNativeFile, XldmNativeModel,
    XldmNativeParseOptions, XldmNativeResult, XldmNumericDictionary, XldmStringDictionary,
    XldmStringHashMode, XldmStringHashOverride, XldmStringPage, XldmStringPageData,
    XldmStringRecordHandle, inspect_xldm_native, parse_xldm_dictionary, parse_xldm_hash_index,
    parse_xldm_idf, write_xldm_native_file,
};
pub mod xldm_generated;
pub use xldm_generated::{
    XldmGeneratedDataError, XldmGeneratedDataResult, XldmSystemGeneratedCompression,
    XldmSystemGeneratedData, XldmSystemGeneratedFile, XldmSystemGeneratedKind,
    XldmSystemGeneratedModel, inspect_xldm_system_generated, parse_xldm_system_generated_file,
    validate_xldm_system_generated_files, write_xldm_system_generated_file,
};
pub mod xldm_metadata;
pub use xldm_metadata::{
    XldmColumnPolicy, XldmDictionaryPolicy, XldmHierarchyPolicy, XldmMetadataClass,
    XldmMetadataCollection, XldmMetadataDataObject, XldmMetadataError, XldmMetadataFile,
    XldmMetadataFileKind, XldmMetadataMember, XldmMetadataModel, XldmMetadataObject,
    XldmMetadataProperty, XldmMetadataResult, XldmRelationshipIndexKind, XldmRelationshipPolicy,
    inspect_xldm_metadata, parse_xldm_metadata_file, validate_xldm_metadata_files,
    write_xldm_metadata_file,
};
pub mod xldm_olap;
pub use xldm_olap::{
    XldmCubeInformation, XldmDimensionInformation, XldmDimensionInformationMap,
    XldmDimensionInformationProperty, XldmOlapDefinition, XldmOlapDocument, XldmOlapElement,
    XldmOlapError, XldmOlapFile, XldmOlapFileKind, XldmOlapHierarchy, XldmOlapModel,
    XldmOlapObjectKind, XldmOlapParentReference, XldmOlapResult, XldmPartitionInformation,
    XldmTabularExtension, inspect_xldm_olap, parse_xldm_olap_file, validate_xldm_olap_model,
    write_xldm_olap_file,
};
pub mod xldm_compression;
pub mod xldm_crypt;
pub use xldm_compression::*;
pub use xldm_crypt::*;
