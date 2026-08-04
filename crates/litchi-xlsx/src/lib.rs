//! Typed Excel Open XML documents.
//!
//! The ordinary API exposes immutable, cheap-to-share workbook and sheet
//! handles. Package relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]

pub mod active_x;
pub mod auto_filter;
pub mod calculation_properties;
pub mod cell;
pub mod cell_watches;
pub mod chain;
pub mod chart_sheet;
pub mod color;
pub mod column;
pub mod conditional_formatting;
pub mod connections;
pub mod custom_data;
pub mod data_consolidation;
pub mod data_validation;
mod error;
pub mod external_links;
pub mod formula;
pub mod header_footer;
pub mod ignored_errors;
pub mod layout;
pub mod merge;
pub mod named_sheet_view;
mod outline;
pub mod outline_properties;
pub mod page_margins;
pub mod page_setup;
pub mod phonetic_properties;
pub mod print_options;
pub mod query_table;
pub mod raw;
pub mod row;
pub mod scenarios;
pub mod sheet;
pub mod sheet_calculation_properties;
pub mod sheet_protection;
pub mod sheet_view;
pub mod slicer_cache;
pub mod sort;
pub mod style;
pub mod table;
pub mod threaded_comments;
pub mod timelines;
pub mod views;
pub mod volatile_dependencies;
pub mod web;
mod workbook;
pub mod workbook_metadata;
pub mod xml_maps;

pub use active_x::{
    Binary, ControlProperties, ControlSet, Descriptor, Font, LoadedControl, Marker, ObjectAnchor,
    Persistence, Picture, PreviewImage, Property, PropertyObject, WorksheetControl,
    WorksheetControls, load_from_worksheet, remove_from_worksheet, replace_on_worksheet,
    replace_worksheet_controls_xml, store_on_worksheet,
};
pub use auto_filter::{
    AutoFilterDefinition, CalendarType, ColorFilter, CustomFilter, CustomFilterOperator,
    CustomFilters, DateGroupItem, DateTimeGrouping, DynamicFilter, DynamicFilterType,
    FilterColumnDefinition, FilterColumnPayload, FilterIconSet, FilterItem, FilterRange,
    FilterValues, IconFilter, SortConditionDefinition, SortStateDefinition, Top10Filter,
    parse_auto_filter, parse_auto_filter_fragment, write_auto_filter_fragment,
};
pub use calculation_properties::{
    WorkbookCalculationMode, WorkbookCalculationProperties, WorkbookReferenceMode,
    parse_workbook_calculation_properties,
};
pub use cell::{Cell, Cells, Content, Date, ErrorValue, Extents, Number, Text, Value};
pub use cell_watches::{
    CellWatchReference, WorksheetCellWatchConformance, WorksheetCellWatches,
    parse_worksheet_cell_watches, write_worksheet_cell_watches,
};
pub use chart_sheet::{parse_chartsheet, validate_chartsheet, write_chartsheet};
pub use color::{ParseRgbError, Rgb};
pub use column::{Column, Columns, Width, WidthAt};
pub use conditional_formatting::{
    Association, Axis, Color, ColorRole, ColorScale, Component, DataBar, Differential,
    DifferentialRef, Direction, Formatting, IconSet, IconSet14, Icons, Kind, NamedColor,
    NumberFormat, Operator, Payload, Period, Rule, TokenError, ValueKind,
    parse_conditional_formattings, parse_differential_formats,
};
pub use custom_data::{
    ExtensionList, Properties, parse_properties, validate_workbook_root, write_properties,
};
pub use data_consolidation::{
    WorksheetDataConsolidation, WorksheetDataConsolidationConformance,
    WorksheetDataConsolidationFunction, WorksheetDataConsolidationRangeReference,
    WorksheetDataReference, WorksheetDataReferenceSource, WorksheetDataReferences,
    parse_worksheet_data_consolidation, write_worksheet_data_consolidation,
};
pub use data_validation::{
    Collection, Conformance, ListSource, Range, Source, Sqref, Validation, ValidationErrorStyle,
    ValidationImeMode, ValidationOperator, ValidationType, parse_data_validation_collections,
    replace_data_validation_collections, validate_data_validation_collections,
    write_data_validation_collections, write_data_validation_core,
    write_data_validation_extensions,
};
pub use error::{
    ColumnEditBlock, DefaultsEditBlock, EditBlock, Error, MergeEditBlock, RemoveBlock, RenameBlock,
    Result, RowEditBlock, TabEditBlock,
};
pub use external_links::{
    ExternalCell, ExternalCellType, ExternalDdeItem, ExternalDdeLink, ExternalDdeValue,
    ExternalDdeValueType, ExternalDdeValues, ExternalDefinedName, ExternalLinkConformance,
    ExternalLinkEntry, ExternalLinkKind, ExternalOleItem, ExternalOleItemSource, ExternalOleLink,
    ExternalOleTarget, ExternalRow, ExternalSheetData, ExternalWorkbookLink,
    ExternalWorkbookTarget, build_external_link_part, build_external_link_part_with_conformance,
    load_external_link,
};
pub use formula::Formula;
pub use header_footer::{SectionKind, Settings, parse_worksheet_header_footer};
pub use ignored_errors::{
    IgnoredErrorRangeReference, WorksheetIgnoredError, WorksheetIgnoredErrorType,
    WorksheetIgnoredErrors, WorksheetIgnoredErrorsExtension, parse_worksheet_ignored_errors,
};
pub use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
};
pub use named_sheet_view::{
    load_worksheet_named_sheet_views, parse_named_sheet_views, remove_worksheet_named_sheet_views,
    store_worksheet_named_sheet_views, write_named_sheet_views,
};
pub use outline::{Outline, OutlineAt};
pub use outline_properties::{WorksheetOutlineProperties, parse_worksheet_outline_properties};
pub use page_margins::{PageMargin, WorksheetPageMargins, parse_worksheet_page_margins};
pub use page_setup::{
    Comments, Copies, Dpi, ErrorMode, FirstPage, Fit, LexicalError, Measure, Order, Orientation,
    Paper, RangeError, RelId, Scale, Setup, Unit, parse_worksheet_page_setup,
    parse_worksheet_page_setup_relationship_id,
};
pub use phonetic_properties::{
    WorksheetPhoneticAlignment, WorksheetPhoneticProperties, WorksheetPhoneticType,
    parse_worksheet_phonetic_properties,
};
pub use print_options::{WorksheetPrintOptions, parse_worksheet_print_options};
// Query-table semantic types remain under the contextual `query_table` owner;
// package operations are also available at this convenience facade.
pub use query_table::{
    QUERY_TABLE_CONTENT_TYPE, QUERY_TABLE_RELATIONSHIP_TYPE, STRICT_QUERY_TABLE_RELATIONSHIP_TYPE,
    add_worksheet_query_table, find_worksheet_query_table, is_query_table_relationship_type,
    load_worksheet_query_tables, parse_query_table, remove_worksheet_query_table,
    reorder_worksheet_query_tables, replace_worksheet_query_table, update_worksheet_query_table,
    write_query_table,
};
pub use row::{Height, HeightAt, Row, Rows};
pub use scenarios::{
    ScenarioCellReference, ScenarioRangeReference, WorksheetScenario, WorksheetScenarioConformance,
    WorksheetScenarioInputCell, WorksheetScenarios, parse_worksheet_scenarios,
    write_worksheet_scenarios,
};
pub use sheet_calculation_properties::{
    WorksheetSheetCalculationProperties, WorksheetSheetCalculationPropertiesConformance,
    parse_worksheet_sheet_calculation_properties, write_worksheet_sheet_calculation_properties,
};
pub use sheet_protection::{
    ProtectedRangeSource, ProtectionPasswordVerifier, ProtectionRangeReference,
    ProtectionRangeReferenceKind, ProtectionRangeSqref, StrongProtectionPasswordVerifier,
    WorksheetProtectedRange, WorksheetProtectedRangeCollection, WorksheetProtection,
    WorksheetProtectionConformance, WorksheetProtectionMetadata, parse_worksheet_protection,
    replace_worksheet_protection, validate_worksheet_protection_metadata,
    write_worksheet_protection, write_worksheet_protection_core,
    write_worksheet_protection_extensions,
};
pub use sheet_view::parse_worksheet_views;
pub use slicer_cache::{
    SLICER_CACHE_CONTENT_TYPE, SLICER_CACHE_RELATIONSHIP_TYPE, SlicerCacheData,
    SlicerCacheDataKind, SlicerCacheDefinition, SlicerCacheExtensionList, SlicerCachePivotTable,
    WorkbookSlicerCache, parse_slicer_cache_definition, validate_slicer_cache_definition,
    write_slicer_cache_definition,
};
pub use sort::{SortBy, SortCondition, SortMethod, SortState};
pub use style::{LocalStyle, Style, StyleKey, StyleState, Styles, StylesIter};
pub use table::{
    Table, TableColumn, TableFormula, TableStyleInfo, TableType, TotalsRowFunction,
    parse_table_xml, serialize_table, validate_table, write_table_xml,
};
pub use threaded_comments::{
    Comment, Graph, Mention, People, Person, SheetPart, WorkbookPart, parse_comments,
    parse_persons, validate_comments, validate_graph, validate_guid, validate_people,
    validate_timestamp, write_comments, write_persons,
};
pub use timelines::{
    Cache, CacheDefinition, CachePivotTable, FilterType, Level, OpaqueXml, PivotFilter,
    TIMELINE_CACHE_CONTENT_TYPE, TIMELINE_CACHE_EXTENSION_URI, TIMELINE_CACHE_RELATIONSHIP_TYPE,
    TIMELINES_CONTENT_TYPE, TIMELINES_EXTENSION_URI, TIMELINES_RELATIONSHIP_TYPE, View, Views,
    WorksheetView, load_timeline_caches, load_timelines, parse_timeline_cache_definition,
    parse_timelines, store_timeline_caches, store_worksheet_timelines,
    write_timeline_cache_definition, write_timelines,
};
pub use views::{
    SheetPane, SheetPanePosition, SheetPaneState, SheetSelection, SheetView, SheetViewType,
};
pub use workbook::{
    ActiveTab, Change, ColumnEdit, Commit, Conflict, ConflictSet, DateSystem, DefaultsEdit, Edit,
    Flavor, JoinError, JoinFailure, NewSheet, PackageChange, Patch, RowEdit, Sheet, SheetEdit,
    SheetKind, SheetSelector, State, TabEdit, Visibility, Workbook,
};
pub use workbook_metadata::{
    FutureMetadata, MetadataBehavior, MetadataBlock, MetadataRecord, MetadataType,
    OpaqueMetadataExtension, WorkbookMetadata,
};
