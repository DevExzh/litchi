//! Typed Excel Open XML documents.
//!
//! The ordinary API exposes immutable, cheap-to-share workbook and sheet
//! handles. Package relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]
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

pub mod active_x;
pub mod auto_filter;
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
pub mod data_validation;
pub mod defined_names;
pub mod drawing;
pub mod edit;
mod error;
pub mod external_links;
pub mod formula;
pub mod header_footer;
pub mod ignored_errors;
pub mod layout;
pub mod merge;
pub mod named_sheet_view;
pub mod ole_objects;
mod outline;
pub mod outline_properties;
pub mod package;
pub mod page_breaks;
pub mod page_margins;
pub mod page_setup;
pub mod phonetic_properties;
pub mod pivot;
pub mod print_options;
pub mod query_table;
pub mod raw;
pub mod revisions;
pub mod rich_values;
pub mod row;
pub mod row_visibility;
pub mod scenarios;
pub mod shapes;
pub mod sheet;
pub mod sheet_calculation_properties;
pub mod sheet_protection;
pub mod sheet_view;
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
pub mod workbook;
pub mod workbook_metadata;
pub mod writer;
pub mod xml_maps;

/// Runtime-neutral `[MS-OFFCRYPTO]` managed-package encryption.
#[cfg(feature = "encryption")]
pub use litchi_crypto::ooxml as encryption;

/// OPC resource limits used by XLSX package and workbook ingress.
pub use litchi_opc::ReadLimits;

pub use active_x::{
    Binary, Control, ControlProperties, ControlSet, Controls, Descriptor, Font, LoadedControl,
    Marker, ObjectAnchor, Persistence, Picture, PreviewImage, Property, PropertyObject,
    load_from_worksheet, remove_from_worksheet, replace_controls_xml, replace_on_worksheet,
    store_on_worksheet,
};
pub use auto_filter::{
    Calendar, Condition, Custom, Customs, DateGroup, Definition, Dynamic, DynamicType, Grouping,
    Item, Top10, Values, parse_auto_filter, parse_auto_filter_fragment, write_auto_filter_fragment,
};
pub use calculation_properties::{Mode, ReferenceMode};
pub use cell::{
    Cell, Cells, Content, Date, ErrorValue, Extents, Number, SharedStringKey, Text, Value,
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
    ExtensionList, Properties, parse_properties, validate_workbook_root, write_properties,
};
pub use data_consolidation::{
    DataConsolidation, Function, RangeReference, Reference, ReferenceSource, References,
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
    CellType, Dde, DdeItem, DdeValue, DdeValueType, DdeValues, DefinedName, Entry, ItemSource,
    Link, Ole, OleItem, SheetData, Target, build_external_link_part,
    build_external_link_part_with_conformance, load_external_link,
};
pub use formula::Formula;
pub use header_footer::{SectionKind, Settings, parse_worksheet_header_footer};
pub use ignored_errors::{
    IgnoredError, IgnoredErrorRangeReference, IgnoredErrorType, IgnoredErrors,
    IgnoredErrorsExtension, parse_worksheet_ignored_errors,
};
pub use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
};
pub use named_sheet_view::{
    load_worksheet_named_sheet_views, parse_named_sheet_views, remove_worksheet_named_sheet_views,
    store_worksheet_named_sheet_views, write_named_sheet_views,
};
pub use ole_objects::{
    Aspect, OleObject, OleObjectAnchor, OleObjectConformance, OleObjectMarker, OleObjectProperties,
    OleObjectRelationshipKind, OleObjectResource, OleObjectTarget, OleObjectUpdate, OleObjects,
    load_ole_objects, parse_ole_objects, store_ole_objects, write_ole_objects,
};
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
pub use print_options::{PrintOptions, parse_print_options};
pub use streaming::{
    StreamingCell, StreamingCellValue, StreamingValue, StreamingWorkbookLimits,
    StreamingWorkbookWriter,
};
// Query-table semantic types remain under the contextual `query_table` owner;
// package operations are also available at this convenience facade.
pub use data_validation::{
    Commit as DataValidationCommit, Diagnostics as DataValidationDiagnostics,
    Patch as DataValidationPatch, Snapshot as DataValidationSnapshot,
    SourceBackedEditor as SourceBackedDataValidationEditor,
    SourceEdit as SourceBackedDataValidationEdit,
};
pub use query_table::{
    QUERY_TABLE_CONTENT_TYPE, QUERY_TABLE_RELATIONSHIP_TYPE, STRICT_QUERY_TABLE_RELATIONSHIP_TYPE,
    add_worksheet_query_table, find_worksheet_query_table, is_query_table_relationship_type,
    load_worksheet_query_tables, parse_query_table, remove_worksheet_query_table,
    reorder_worksheet_query_tables, replace_worksheet_query_table, update_worksheet_query_table,
    write_query_table,
};
pub use revisions::{
    Commit as RevisionCommit, Patch as RevisionPatch, RevisionAttribute,
    RevisionAttributeNamespace, RevisionConformance, RevisionHeader, RevisionHeaderProperties,
    RevisionHeaders, RevisionLog, RevisionLogPart, RevisionRecord, RevisionRecordKind,
    RevisionUser, RevisionUsers, RevisionXmlElement, Revisions, Snapshot as RevisionSnapshot,
    Transaction as RevisionTransaction, load_workbook_revisions, parse_revision_headers,
    parse_revision_log, parse_revision_users, remove_workbook_revisions, store_workbook_revisions,
    write_revision_headers, write_revision_log, write_revision_users,
};
pub use rich_values::codec::{parse_feature_property_bags, write_feature_property_bags};
pub use rich_values::package::load as load_rich_values;
pub use row::{Height, HeightAt, Row, Rows};
pub use scenarios::{
    CellReference, InputCell, Scenario, UnknownAttribute, UnknownElement,
    parse_worksheet_scenarios, write_worksheet_scenarios,
};
pub use sheet::{Name, NameError};
pub use sheet_protection::{
    Commit as SheetProtectionCommit, Diagnostics as SheetProtectionDiagnostics, Metadata,
    Patch as SheetProtectionPatch, ProtectedRange, ProtectedRangeCollection, ProtectedRangeSource,
    Protection, ProtectionPasswordVerifier, ProtectionRangeReference, ProtectionRangeReferenceKind,
    ProtectionRangeSqref, Snapshot as SheetProtectionSnapshot,
    SourceBackedEditor as SourceBackedSheetProtectionEditor,
    SourceEdit as SourceBackedSheetProtectionEdit, StrongProtectionPasswordVerifier,
    parse_protection, replace_protection, validate_metadata, write_core, write_extensions,
    write_protection,
};
pub use sheet_view::parse_worksheet_views;
pub use sort::{SortBy, SortCondition, SortMethod, SortState};
pub use style::{LocalStyle, Style, StyleKey, StyleState, Styles, StylesIter};
pub use survey::{
    Binding as SurveyBinding, ElementProperties as SurveyElementProperties, Guid as SurveyGuid,
    Id as SurveyId, Part as SurveyPart, Position as SurveyPosition, Question as SurveyQuestion,
    QuestionFormat as SurveyQuestionFormat, QuestionType as SurveyQuestionType,
    Questions as SurveyQuestions, Survey, load as load_surveys, parse as parse_survey,
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
pub use workbook::{
    ActiveTab, Change, ColumnEdit, Commit, Conflict, ConflictSet, DateSystem, DefaultsEdit,
    DurablePatch, Edit, Flavor, History, HistoryLimits, JoinError, JoinFailure, MergeChoice,
    MergeLimits, NewSheet, PackageChange, Patch, RowEdit, SealedPatch, Selector,
    SourceBackedWorkbook, SourceCell, SourceCellView, SourceWorksheet, State, TabEdit,
    ThreeWayPlan, Visibility, Workbook, Worksheet, WorksheetEdit, WorksheetKind,
};
pub use workbook_metadata::{
    FutureMetadata, MetadataBehavior, MetadataBlock, MetadataRecord, MetadataType,
    OpaqueMetadataExtension,
};
