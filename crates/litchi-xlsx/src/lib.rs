//! Typed Excel Open XML documents.
//!
//! [`Package`] is the immutable, cheap-to-share owner of every state verb:
//! open, create, edit, apply and save. [`Package::workbook`] returns the owned,
//! read-only [`Workbook`] projection with its sheet handles. Package
//! relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]
// The public model deliberately mirrors the names and shapes in the ECMA-376
// schemas. Renaming these items or reshaping their signatures solely for style
// would make the wire model less recognizable and would break the young public
// API without changing correctness.
#![allow(
    clippy::fn_params_excessive_bools,
    clippy::implicit_hasher,
    clippy::module_name_repetitions,
    clippy::needless_pass_by_value,
    clippy::ref_option,
    clippy::return_self_not_must_use,
    clippy::struct_excessive_bools,
    clippy::struct_field_names,
    clippy::trivially_copy_pass_by_ref,
    clippy::unnecessary_wraps,
    reason = "the stable schema-shaped API follows ECMA-376 rather than Clippy's generic API heuristics"
)]
// These documentation suggestions do not affect the package parser or writer.
// Public error behavior is already expressed by the crate-wide Error type and
// adding hundreds of identical boilerplate sections would obscure the schema
// documentation that callers actually need.
#![allow(
    clippy::doc_markdown,
    clippy::empty_docs,
    clippy::missing_errors_doc,
    clippy::missing_panics_doc,
    reason = "the crate documents shared failure contracts centrally and keeps field documentation aligned with schema terminology"
)]
// Streaming XML codecs intentionally keep names aligned with the current wire
// token (`event`, `attribute`, `value`, and similar). Shadowing marks state
// transitions and keeps each token's lifetime as short as possible.
#![allow(
    clippy::many_single_char_names,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::similar_names,
    reason = "short-lived streaming parser bindings mirror wire tokens and their successive decoded forms"
)]
// Codec declarations are ordered by wire traversal and package transaction
// phases, not alphabetically or by Rust item kind.
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "wire codecs keep declarations in parse/write and transaction order for auditability"
)]
// OPC part names and relationship suffixes are case-sensitive ASCII wire names;
// treating them as case-insensitive would accept a different package member.
#![allow(
    clippy::case_sensitive_file_extension_comparisons,
    reason = "OPC part-name suffix comparisons intentionally preserve case-sensitive package identity"
)]
// A map with unit values is used where callers need deterministic keyed lookup,
// not merely set membership, and identical arms keep schema variants explicit.
#![allow(
    clippy::match_same_arms,
    clippy::zero_sized_map_values,
    reason = "schema variants stay explicit and keyed unit maps preserve the required lookup representation"
)]

// Unit tests drive the source-backed families through the integration
// tests' drivers, which name this crate `litchi_xlsx`.
#[cfg(test)]
extern crate self as litchi_xlsx;

pub mod active_x;
pub mod auto_filter;
pub mod builder;
pub mod calculation_properties;
pub mod cell;
pub mod cell_values;
pub mod cell_watches;
pub mod chain;
pub mod chart;
pub mod chart_sheet;
pub mod color;
pub mod column;
pub mod conditional_formatting;
pub mod connections;
pub mod custom;
pub mod custom_data;
pub mod data_consolidation;
pub mod data_type_icons;
pub mod data_validation;
pub mod defined_names;
pub mod drawing;
pub mod edit;
mod error;
pub mod external_links;
pub mod form_control;
pub mod formula;
pub mod header_footer;
pub mod hyperlinks;
pub mod ignored_errors;
pub mod layout;
pub mod leniency;
pub mod merge;
pub mod metadata;
pub mod named_sheet_view;
pub mod ole_objects;
pub mod opc;
mod open_options;
mod outline;
pub mod outline_properties;
pub mod package;
pub mod page_breaks;
pub mod page_margins;
pub mod page_setup;
mod part_edit;
pub mod phonetic_properties;
pub mod pivot;
pub mod print_options;
pub mod query_table;
pub mod raw;
mod record;
pub mod revisions;
pub mod rich_values;
pub mod row;
pub mod row_visibility;
pub mod scenarios;
pub mod semantic_limits;
pub mod shapes;
pub mod sheet;
pub mod sheet_calculation_properties;
pub mod sheet_protection;
pub mod sheet_view;
mod signature_policy;
pub mod slicer;
#[allow(
    dead_code,
    unreachable_pub,
    unused_imports,
    reason = "the legacy slicer-cache wire model is retained for round-trip compatibility"
)]
mod slicer_cache;
pub mod smart_tags;
pub mod sort;
mod source_attributes;
mod source_limits;
mod source_payload;
/// Bounded, sequential creation of one-sheet XLSX workbooks.
pub mod streaming;
pub mod style;
pub mod survey;
pub mod tab_state;
pub mod table;
pub mod task_panes;
pub mod threaded_comments;
pub mod timeline;
#[allow(
    dead_code,
    unreachable_pub,
    unused_imports,
    reason = "the legacy timeline wire model is retained for round-trip compatibility"
)]
mod timelines;
pub mod volatile_dependencies;
pub mod web;

pub use drawing::ContentPartSelector;
pub mod workbook;
pub mod workbook_metadata;
pub mod writer;
pub mod xml_maps;

/// Runtime-neutral `[MS-OFFCRYPTO]` managed-package encryption.
#[cfg(feature = "encryption")]
pub use litchi_crypto::ooxml as encryption;

/// OPC resource limits used by XLSX package and workbook ingress.
pub use litchi_opc::{
    SourceCacheDiagnostics, SourceCacheLimitError, SourceCacheLimits, SourceLimitKind, SourceLimits,
};

pub use active_x::{
    Binary, Control, ControlProperties, ControlSet, Controls, Descriptor, Font, LoadedControl,
    Marker, ObjectAnchor, Persistence, Picture, PreviewImage, Property, PropertyObject,
    load_from_worksheet, replace_controls_xml,
};
pub use auto_filter::{
    Calendar, Condition, Custom, Customs, DateGroup, Definition, Dynamic, DynamicType, Grouping,
    Item, Top10, Values, parse_auto_filter, parse_auto_filter_fragment, write_auto_filter_fragment,
};
pub use calculation_properties::{Mode, ReferenceMode};
pub use cell::{
    Cell, Cells, Content, Date, ErrorValue, Extents, IndexedCells, Number, RecordIndex, RichRun,
    RichText, SharedStringKey, Text, Value,
};
pub use cell_watches::{
    CellWatchConformance, CellWatchReference, CellWatches, parse_cell_watches, write_cell_watches,
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
    CustomData, ExtensionList, Limits as CustomDataLimits, Part as CustomDataPart, Properties,
    Snapshot as CustomDataSnapshot, parse_properties, validate_workbook_root, write_properties,
};
pub use data_consolidation::{
    DataConsolidation, Function, RangeReference, Reference, ReferenceSource, References,
    parse_worksheet_data_consolidation, write_worksheet_data_consolidation,
};
pub use data_type_icons::{SHOW_DATA_TYPE_ICONS_NAMESPACE, ShowDataTypeIcons};
pub use data_validation::{
    Collection, Conformance, ListSource, Range, Source, Sqref, Validation, ValidationErrorStyle,
    ValidationImeMode, ValidationOperator, ValidationType, parse_data_validation_collections,
    replace_data_validation_collections, validate_data_validation_collections,
    write_data_validation_collections, write_data_validation_core,
    write_data_validation_extensions,
};
pub use error::{
    ColumnEditBlock, DefaultsEditBlock, EditBlock, Error, LimitKind, MergeEditBlock, RemoveBlock,
    RenameBlock, Result, RowEditBlock, TabEditBlock,
};
pub use external_links::{
    AlternateUrl, AlternateUrls, CellType, Dde, DdeItem, DdeValue, DdeValueType, DdeValues,
    DefinedName, Entry, ItemSource, Link, Ole, OleItem, SheetData, Target,
    build_external_link_part, build_external_link_part_with_conformance, load_external_link,
};
pub use formula::Formula;
pub use header_footer::{SectionKind, Settings, parse_worksheet_header_footer};
pub use hyperlinks::{Hyperlink, HyperlinkReference};
pub use ignored_errors::{
    IgnoredError, IgnoredErrorRangeReference, IgnoredErrorType, IgnoredErrors,
    IgnoredErrorsExtension, parse_worksheet_ignored_errors,
};
pub use leniency::{
    CatalogDefect, CellDefect, Defect, Leniency, Location as ToleranceLocation,
    SharedStringsDefect, ThemeDefect, ToleranceReport, ToleratedDefect, WorksheetDefect,
};
pub use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
};
pub use named_sheet_view::{
    load_worksheet_named_sheet_views, parse_named_sheet_views, write_named_sheet_views,
};
pub use ole_objects::{
    Aspect, OleObject, OleObjectAnchor, OleObjectConformance, OleObjectMarker, OleObjectProperties,
    OleObjectRelationshipKind, OleObjectResource, OleObjectTarget, OleObjectUpdate, OleObjects,
    load_ole_objects, parse_ole_objects, write_ole_objects,
};
pub use open_options::OpenOptions;
pub use outline::{Outline, OutlineAt};
pub use outline_properties::{OutlineProperties, parse_outline_properties};
pub use package::Package;
pub use page_margins::{
    Margins, PageMargin, parse_page_margins, replace_page_margins, write_page_margins,
};
pub use page_setup::{
    Comments, Copies, Dpi, ErrorMode, FirstPage, Fit, LexicalError, Measure, Order, Orientation,
    Paper, RangeError, RelId, Scale, Setup, Unit, parse_worksheet_page_setup,
    parse_worksheet_page_setup_relationship_id,
};
pub use phonetic_properties::{
    PhoneticAlignment, PhoneticProperties, PhoneticType, parse_phonetic_properties,
};
pub use pivot::{
    PivotCellAddress, PivotCellType, PivotRowView, PivotTableDataDiagnostic, PivotTableDataView,
    PivotTableSelector, PivotTableServerFormats, PivotTableView, PivotValueCellExtraView,
    PivotValueCellView, ServerFormat,
};
pub use print_options::{PrintOptions, parse_print_options};
pub use semantic_limits::SemanticLimits;
pub use streaming::{
    StreamingCell, StreamingCellValue, StreamingWorkbookLimits, StreamingWorkbookWriter,
};
// Query-table semantic types remain under the contextual `query_table` owner;
// its package reads are also available at this convenience facade (edits go
// through `Package::edit_query_tables`).
pub use query_table::{
    QUERY_TABLE_CONTENT_TYPE, QUERY_TABLE_RELATIONSHIP_TYPE, STRICT_QUERY_TABLE_RELATIONSHIP_TYPE,
    find_worksheet_query_table, is_query_table_relationship_type, load_worksheet_query_tables,
    parse_query_table, write_query_table,
};
pub use revisions::{
    RevisionAttribute, RevisionAttributeNamespace, RevisionConformance, RevisionHeader,
    RevisionHeaderProperties, RevisionHeaders, RevisionLog, RevisionLogPart, RevisionRecord,
    RevisionRecordKind, RevisionUser, RevisionUsers, RevisionXmlElement, Revisions,
    load_workbook_revisions, parse_revision_headers, parse_revision_log, parse_revision_users,
    write_revision_headers, write_revision_log, write_revision_users,
};
pub use rich_values::codec::{parse_feature_property_bags, write_feature_property_bags};
pub use rich_values::package::load as load_rich_values;
pub use rich_values::refresh_intervals::{
    REFRESH_INTERVALS_NAMESPACE, RefreshInterval, RefreshIntervals, TypeRefreshIntervals,
    load as load_rich_value_refresh, parse_refresh_intervals, write_refresh_intervals,
};
pub use row::{Height, HeightAt, Row, Rows};
pub use scenarios::{
    CellReference, InputCell, Scenario, UnknownAttribute, UnknownElement,
    parse_worksheet_scenarios, write_worksheet_scenarios,
};
pub use sheet::{Name, NameError};
pub use sheet_protection::{
    Metadata, ProtectedRange, ProtectedRangeCollection, ProtectedRangeSource, Protection,
    ProtectionPasswordVerifier, ProtectionRangeReference, ProtectionRangeReferenceKind,
    ProtectionRangeSqref, StrongProtectionPasswordVerifier, parse_protection, replace_protection,
    validate_metadata, write_core, write_extensions, write_protection,
};
pub use sheet_view::parse_worksheet_views;
pub use sort::{SortBy, SortCondition, SortMethod, SortState};
pub use style::{FontChange, LocalStyle, Style, StyleKey, StyleState, Styles, StylesIter};
pub use survey::{
    Binding as SurveyBinding, ElementProperties as SurveyElementProperties, Guid as SurveyGuid,
    Id as SurveyId, Position as SurveyPosition, Question as SurveyQuestion,
    QuestionFormat as SurveyQuestionFormat, QuestionType as SurveyQuestionType,
    Questions as SurveyQuestions, Survey,
};
pub use table::{
    Table, TableColumn, TableFormula, TableStyleInfo, TableType, TotalsRowFunction,
    parse_table_xml, serialize_table, validate_table, write_table_xml,
};
pub use threaded_comments::{
    Comment, CommentsPart, Graph, Mention, People, PeoplePart, Person, parse_comments,
    parse_persons, validate_comments, validate_graph, validate_guid, validate_people,
    validate_timestamp, write_comments, write_persons,
};
pub use workbook::data_model::{
    Definition as DataModelDefinition, EditOptions as DataModelEditOptions,
    Limits as DataModelLimits, Model as DataModel, OpaqueXml as DataModelOpaqueXml,
    Payload as DataModelPayload, Relationship as DataModelRelationship,
    Snapshot as DataModelSnapshot, Table as DataModelTable, load_data_model, parse_data_model,
    write_data_model,
};
pub use workbook::{
    ActiveTab, Change, Commit, Conflict, ConflictSet, DateSystem, DurablePatch, Edit, EditOptions,
    Flavor, History, HistoryLimits, JoinError, JoinFailure, JoinNormalization, MergeChoice,
    MergeLimits, NewSheet, PackageChange, Patch, SealedPatch, Selector, SourceBackedWorkbook,
    SourceCell, SourceCellView, SourceWorksheet, State, ThreeWayPlan, Visibility, Workbook,
    Worksheet, WorksheetContentPart, WorksheetContentPartOutboundRelationship,
    WorksheetContentPartPayload, WorksheetContentPartRelationship, WorksheetKind,
};
pub use workbook_metadata::{
    FutureMetadata, MetadataBehavior, MetadataBlock, MetadataRecord, MetadataType,
    OpaqueMetadataExtension,
};
