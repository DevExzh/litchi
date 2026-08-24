#![forbid(unsafe_code)]
// BIFF record models and their public facades intentionally retain names,
// boolean fields, and value widths from the specification. Reshaping these
// APIs solely for generic style guidance would break compatibility without
// improving the bounded codecs.
#![allow(
    clippy::double_must_use,
    clippy::fn_params_excessive_bools,
    clippy::large_types_passed_by_value,
    clippy::missing_fields_in_debug,
    clippy::module_name_repetitions,
    clippy::needless_pass_by_value,
    clippy::ref_option,
    clippy::return_self_not_must_use,
    clippy::struct_excessive_bools,
    clippy::struct_field_names,
    clippy::trivially_copy_pass_by_ref,
    clippy::unused_self,
    reason = "the established XLS API and fixed-width BIFF models deliberately mirror specification vocabulary and ownership"
)]
// Streaming BIFF parsers reuse short record-local names as each wire value is
// decoded. The shadowing is confined to small scopes and makes the source map
// directly to the record currently being handled.
#![allow(
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::similar_names,
    reason = "short-lived parser bindings track successive BIFF wire values and decoded projections"
)]
// Test fixtures and assertions deliberately fail fast; production code is no
// longer covered by a crate-wide `unwrap`/`expect` correctness exemption.
#![cfg_attr(
    test,
    allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test fixtures and assertions intentionally fail fast"
    )
)]
// The codec is ordered by BIFF stream traversal and keeps spec cases explicit.
// The remaining suggestions are behavior-neutral style alternatives whose
// application would obscure that audit order.
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::default_trait_access,
    clippy::empty_line_after_doc_comments,
    clippy::float_cmp,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::match_same_arms,
    clippy::match_wildcard_for_single_variants,
    clippy::nonminimal_bool,
    clippy::only_used_in_recursion,
    clippy::single_match_else,
    clippy::unnecessary_wraps,
    reason = "legacy BIFF codec declarations follow stream order and retain explicit specification branches for auditability"
)]
//! Legacy Excel (.xls) file format reader
//!
//! This module provides functionality to parse Microsoft Excel files
//! in the legacy binary format (.xls files), which are OLE2-based files.
//! The implementation is based on the BIFF (Binary Interchange File Format)
//! specification and draws inspiration from other spreadsheet libraries.

/// Error types for XLS parsing
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "error conversion preserves unknown dependency error variants as inert InvalidData"
)]
mod error;

/// BIFF8 password-to-open encryption support
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "bounded encryption state and fixed-width fields retain checked legacy invariants"
)]
mod encryption;

/// BIFF record parsing utilities
pub mod records;

/// Bounded, non-mutating validation of CFB-backed XLS ingress and BIFF
/// ownership metadata.
pub mod validation;

/// Workbook parsing implementation
pub mod workbook;

/// Worksheet parsing implementation and worksheet-owned semantic views.
pub mod worksheet;

/// Cell value parsing and representation
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "cell projection deliberately preserves non-target record/value variants"
)]
mod cell;

/// Source-checked edits of existing BIFF8 `Number` cell values.
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "the cell transaction accepts forward-compatible public sheet values"
)]
pub mod cell_values;
/// Lossless, source-backed BIFF8 worksheet visibility transactions.
pub mod sheet_visibility;

/// BIFF8 worksheet data-validation records.
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "unknown validation kinds remain inert and are rejected by typed mutation"
)]
mod data_validation;

/// BIFF8 calculation and recalculation records.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "calculation records use fixed-width fields and all-or-none collector state"
)]
mod calculation;
/// BIFF8 chart-sheet and embedded-chart metadata and mutation.
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "chart projections retain unsupported chart and location variants inertly"
)]
pub mod chart;

/// BIFF8 chart layout future records (`CrtLayout12`, `CrtLayout12A`).
mod chart_layout;

/// BIFF8 `MarkerFormat` record: data-marker color, size, and shape.
#[allow(
    clippy::expect_used,
    reason = "marker fields are sliced only after their fixed record length is checked"
)]
mod marker_format;

/// BIFF8 `ObjectLink` record: the chart object a text is linked to.
mod object_link;

/// BIFF8 `UsesELFs` record: the natural language formula flag.
#[allow(
    clippy::expect_used,
    reason = "UsesElfs is a fixed two-byte record checked before conversion"
)]
mod uses_elfs;

/// BIFF8 `SerAuxErrBar` record: error bar properties.
#[allow(
    clippy::expect_used,
    reason = "error-bar fields are sliced only after fixed-length validation"
)]
mod ser_aux_err_bar;

/// BIFF8 `SerAuxTrend` record: a trendline.
#[allow(
    clippy::expect_used,
    reason = "trend-line fields are sliced only after fixed-length validation"
)]
mod ser_aux_trend;

/// BIFF8 `Fbi` record: scalable chart font information.
mod fbi;

/// BIFF8 `EntExU2` record: an application-specific cache, preserved opaquely.
mod ent_ex_u2;

/// BIFF8 chart axis-group records (`AxesUsed`, `AxisParent`).
#[allow(
    clippy::expect_used,
    reason = "AxesUsed reserved bytes are copied only after fixed-length validation"
)]
mod axes_used;

/// BIFF8 `BopPop` record: bar of pie / pie of pie chart group attributes.
mod bop_pop;

/// BIFF8 `PlotGrowth` record: plot area font-scaling factors.
#[allow(
    clippy::expect_used,
    reason = "PlotGrowth fields are sliced only after fixed-length validation"
)]
mod plot_growth;

/// BIFF8 `Scl` record: the view zoom fraction.
mod scl;

/// BIFF8 `Chart3d` record: 3-D plot area attributes.
mod chart_3d;

/// BIFF8 `BopPopCustom` record: a custom pie split bit sequence.
mod bop_pop_custom;

/// BIFF8 extended data label records (`DataLabExt`, `DataLabExtContents`).
mod data_label_ext;

/// BIFF8 chart property-stream future records (`ShapePropsStream`,
/// `TextPropsStream`, `RichTextStream`).
#[allow(
    clippy::expect_used,
    reason = "chart property fields are sliced only after bounded header validation"
)]
mod chart_property_stream;

/// BIFF8 `ForceFullCalculation` record: the forced calculation mode.
#[allow(
    clippy::expect_used,
    reason = "ForceFullCalculation fields are sliced only after fixed-length validation"
)]
mod force_full_calculation;

/// BIFF8 `Backup` record: the workbook save-backup flag.
#[allow(
    clippy::expect_used,
    reason = "Backup is a fixed two-byte record checked before conversion"
)]
mod backup;

/// BIFF8 `BkHim` record: sheet background image data.
#[allow(
    clippy::expect_used,
    reason = "background picture framing is indexed only after complete record validation"
)]
mod background_picture;

/// BIFF8 `CellWatch` record: a watched-cell reference.
#[allow(
    clippy::expect_used,
    reason = "CellWatch fields are sliced only after fixed-length validation"
)]
mod cell_watch;

/// BIFF8 user-interface collection markers (`InterfaceHdr`, `InterfaceEnd`).
#[allow(
    clippy::expect_used,
    reason = "interface record fields are sliced only after fixed-length validation"
)]
mod interface_records;

/// BIFF8 `HFPicture` record: a sheet header/footer picture.
#[allow(
    clippy::expect_used,
    reason = "header/footer picture framing is indexed only after bounded validation"
)]
mod header_footer_picture;

/// BIFF8 `Pls` record: printer driver `DEVMODE` data spanning `Continue`
/// records.
mod printer_driver;

/// BIFF8 worksheet scenario manager records.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "scenario continuation chunks are initialized and bounded before access"
)]
mod scenario;

/// Inert BIFF8 VBA project markers and object code names.
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "VBA storage traversal deliberately ignores unrelated directory entry kinds"
)]
mod vba;

/// Typed, inert legacy XLS XML maps and their root-level `XML` stream.
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "XML token parsing explicitly ignores non-semantic token variants"
)]
pub mod xml_map;

/// BIFF8 workbook-global environment and behavioral options.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "environment fields are sliced only after fixed-length validation"
)]
mod environment;

/// Inert BIFF8 workbook access-provenance metadata.
mod access;

/// BIFF8 default table and PivotTable style catalog metadata.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "table-style fields are sliced only after their declared lengths are checked"
)]
mod table_styles;

/// BIFF8 global differential formatting records and typed XF properties.
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "differential formatting preserves unsupported property variants inertly"
)]
mod differential_format;

/// BIFF8 extended shared-string table lookup index.
#[allow(
    clippy::unwrap_used,
    reason = "ExtSST bucket chunks have exact widths established before conversion"
)]
mod shared_string_index;

/// BIFF8 worksheet INDEX and DBCELL row-block lookup metadata.
#[allow(
    clippy::unwrap_used,
    reason = "row-block fields are read only after fixed-width framing validation"
)]
mod row_block_index;

/// BIFF8 formula error-checking shared features.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "formula-error feature fields are sliced after fixed header validation"
)]
mod formula_errors;

#[allow(
    clippy::unwrap_used,
    clippy::wildcard_enum_match_arm,
    reason = "bounded AutoFilter12 fields and inert unmatched values retain legacy invariants"
)]
mod autofilter12;
/// BIFF8 worksheet tables and their List12 formatting records.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "list-object codecs index only validated fixed-width and counted structures"
)]
mod list_object;

/// BIFF8 workbook windows and stable sheet-tab identifiers.
mod workbook_view;

/// BIFF8 built-in and user-defined function categories.
mod function_group;

/// Inert BIFF8 supporting-book links and external cell caches.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::wildcard_enum_match_arm,
    reason = "external-link codecs retain bounded legacy invariants and inert unknown variants"
)]
mod external_link;

/// Inert BIFF8 worksheet data-consolidation directories and sources.
#[allow(
    clippy::expect_used,
    reason = "consolidation fields are sliced only after declared-length validation"
)]
mod consolidation;

/// BIFF formula token rendering
mod formula;

/// Typed metadata surrounding a BIFF8 `Formula` record.
pub mod formula_metadata;

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

/// BIFF8 worksheet row and column formatting records.
pub mod layout;

/// BIFF8 `BookExt` record: workbook extension flags.
#[allow(
    clippy::expect_used,
    reason = "BookExt fields are sliced only after fixed-length validation"
)]
mod book_ext;

/// Typed BIFF8 `CompressPictures` recommendation and lossless snapshots.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "picture-compression headers are indexed only after bounded framing validation"
)]
pub mod picture_compression;

/// BIFF8 `Table` record: what-if data tables.
mod data_table;

/// BIFF8 `XFExt` record: formatting property extensions for XF records.
#[allow(
    clippy::expect_used,
    reason = "XF extension fields are sliced only after declared-length validation"
)]
mod xf_ext;

/// BIFF8 `StyleExt` record: cell-style extensions.
mod style_ext;

/// BIFF8 `Theme` record: the document theme.
#[allow(
    clippy::expect_used,
    reason = "theme records are indexed only after complete framing validation"
)]
mod theme;

/// BIFF8 `PhoneticInfo` record: phonetic-string format and visible ranges.
mod phonetic_info;

/// BIFF8 custom-view records (`UserBView`, `UserSViewBegin`, `UserSViewEnd`).
mod custom_view;

/// BIFF8 `SheetExt` record: sheet tab color and publish state.
mod sheet_ext;

/// BIFF8 worksheet window, zoom, pane, and selection state.
pub mod view;

/// BIFF8 worksheet print and page setup.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "page-setup fields are sliced only after fixed-length validation"
)]
mod page_setup;

/// BIFF8 worksheet print display flags (`PrintRowCol`, `GridSet`).
mod print_flags;

/// BIFF8 `SXViewLink` record: the PivotTable view linked to a Pivot Chart.
#[allow(
    clippy::expect_used,
    reason = "SXViewLink fields are sliced only after fixed-length validation"
)]
mod sxview_link;

/// Legacy BIFF8 conditional formatting.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "conditional-format fields are indexed only after bounded owner validation"
)]
mod conditional_format;

/// Workbook sheet directory metadata.
mod sheet_metadata;

/// Shape extraction
pub mod shapes;

/// Typed, inert worksheet OfficeArt anchor metadata.
pub mod drawing_metadata;

/// Shared parsing utilities
mod utils;

/// Merged cell range parsing (MERGECELLS 0x00E5)
pub mod merged_cells;

/// Hyperlink parsing (HLINK 0x01B8)
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "hyperlink moniker fields are indexed only after bounded framing validation"
)]
pub mod hyperlinks;

/// Comment/note parsing (NOTE 0x001C)
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "comment owners are indexed only after checked Note/Obj/TXO linkage"
)]
pub mod comments;
pub mod compatibility;

/// Inert BIFF8 Office Toolbars stream (`XCB`) metadata.
pub mod toolbar;

/// AutoFilter and sort parsing (AUTOFILTERINFO 0x009D, AUTOFILTER 0x009E, SORT 0x0090)
pub mod autofilter;

/// Extended BIFF8 range-sort metadata (`SortData` and `SortCond12`).
mod sort_data;

/// BIFF8 `QUERYTABLE` sequence: typed, inert query tables and connections.
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "query-table codecs index only validated fixed-width and counted fields"
)]
mod query_table;

/// BIFF8 `RealTimeData` record: typed, inert real-time data (RTD) topics.
pub mod real_time_data;

/// Typed, inert readers for the shared-workbook RRD revision record family.
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "revision records preserve unsupported action variants inertly"
)]
pub mod revision_records;

/// BIFF8 shared-workbook `Revision Log` stream (MS-XLS 2.1.7.14).
pub mod revision_log;

/// BIFF8 `WebPub` record: typed, inert Web publishing metadata.
mod web_pub;

/// BIFF8 `METADATA` production: typed, inert MDX (OLAP cube) metadata records.
pub mod mdx_metadata;

/// BIFF8 `SXVIEWEX` sequence: typed, inert PivotTable OLAP extension records.
mod pivot_olap;

/// BIFF8 shared-workbook user records (`CUsr`, `CbUsr`, `UsrInfo`) and
/// routing-slip records (`DocRoute`, `RecipName`).
#[allow(
    clippy::expect_used,
    reason = "user-routing fields are sliced only after declared-length validation"
)]
mod user_routing;

#[forbid(unsafe_code)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "pivot editing indexes only validated owner and stream structures"
)]
mod pivot_editor;
/// Pivot table parsing (SXVIEW, SXVD, SXVI, SXDI, SXVS, SXPI)
#[forbid(unsafe_code)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::wildcard_enum_match_arm,
    reason = "pivot codecs retain bounded legacy invariants and inert unmatched variants"
)]
pub mod pivot_table;
pub use pivot_editor::PivotViewEditor;
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "OLE semantic projections preserve unrelated and unknown Obj variants inertly"
)]
pub mod ole_object;
pub use ole_object::{
    CheckState, DropDownStyle, EditBoxValidation, Editor, FormControl, FtCblsData, FtCmo,
    FtEdoData, FtGboData, FtLbsData, FtPictFmla, FtPioGrbit, FtRboData, FtSbs, LbsDropData,
    LbsItem, ListBehaviorClass, ListSelectionType, ObjSubrecord, ObjectType, OleObjectRecord,
};
pub use pivot_table::{
    PageFieldEntry, PivotAdditionalExtension, PivotAxis, PivotAxisField, PivotCache,
    PivotCacheDateGroupUnit, PivotCacheDateGrouping, PivotCacheDateTime,
    PivotCacheDiscreteGrouping, PivotCacheError, PivotCacheField, PivotCacheGrouping,
    PivotCacheItem, PivotCacheNumericGrouping, PivotDataItem, PivotFunction, PivotItemType,
    PivotLayoutLine, PivotPageSelection, PivotQueryTag, PivotSourceType, PivotTable, PivotViewDef,
    PivotViewEx9, PivotViewExtension, PivotViewField, PivotViewFieldExtension, PivotViewItem,
};

/// Sheet protection parsing (PROTECT, OBJECTPROTECT, SCENPROTECT, PASSWORD)
pub mod protection;

/// Low-level handoffs for already-indexed positional CFB sources.
pub mod raw;

/// XLS file writing
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::wildcard_enum_match_arm,
    reason = "writer internals retain bounded construction invariants and explicit fallback variants"
)]
pub mod writer;

pub use access::{WriteAccess, WriteAccessEncoding};
pub use alignment::{
    CellAlignment, HorizontalAlignment, ReadingOrder, TextRotation, VerticalAlignment,
};
pub use autofilter12::{
    AUTO_FILTER12_RECORD_TYPE, AutoFilter12Criterion, AutoFilter12DateGroup, AutoFilter12DateLevel,
    AutoFilter12DifferentialFormat, AutoFilter12DynamicType, AutoFilter12FormatKind,
    AutoFilter12Icon, AutoFilter12IconSet, AutoFilter12Operator, AutoFilter12Value,
    TableAutoFilter12,
};
pub use axes_used::{AxesUsed, AxesUsedCount, AxisGroupPosition, AxisParent};
pub use background_picture::{BackgroundImage, BackgroundImageFormat};
pub use backup::Backup;
pub use book_ext::{BookExt, BookExtConditional11, BookExtConditional12, FactoidDisplay};
pub use bop_pop::{BopPop, BopPopSplit, BopPopSubtype};
pub use bop_pop_custom::BopPopCustom;
pub use border_fill::{BorderSide, BorderStyle, CellBorders, CellFill, FillPattern};
pub use calculation::{
    CalculationMode, MultithreadedCalculation, ReferenceMode, WorkbookCalculation,
    WorksheetCalculation,
};
pub use cell::Cell;
pub use cell_watch::CellWatch;
pub use chart_3d::Chart3d;
pub use chart_layout::{CrtLayout12, CrtLayout12A, CrtLayout12Mode};
pub use chart_property_stream::{RichTextStream, ShapePropsStream, TextPropsStream};
pub use comments::Visibility;
pub use compatibility::CompatibilityProfile;
pub use conditional_format::{
    ConditionalAlignment, ConditionalBorder, ConditionalComparison, ConditionalExtension,
    ConditionalFont, ConditionalFormatRange, ConditionalFormatting, ConditionalFormatting12,
    ConditionalNumberFormat, ConditionalPattern, ConditionalProtection, ConditionalRule,
    ConditionalRule12, ConditionalRule12Kind, ConditionalRuleKind, ConditionalStyle,
};
pub use consolidation::{
    Consolidation, ConsolidationBuiltInName, ConsolidationFile, ConsolidationFunction,
    ConsolidationRange, ConsolidationSource,
};
pub use custom_view::{
    ChartSheetCustomViewBegin, CustomViewHiddenRows, CustomViewNoteDisplay, CustomViewTopLeft,
    SheetCustomView, SheetCustomViewBegin, SheetCustomViewEnd, WorkbookCustomView,
};
pub use data_label_ext::{DataLabExt, DataLabExtContents};
pub use data_table::{DataTable, DataTableInputCell, DataTableKind, DataTableRange};
pub use data_validation::{
    DataValidationErrorStyle, DataValidationFormula, DataValidationImeMode, DataValidationKind,
    DataValidationOperator, DataValidationRange, DataValidationRule, DataValidationSettings,
};
pub use defined_names::{
    BuiltInName, DefinedName, DefinedNameFutureRecords, DefinedNameKind, NameFnGrp12, NamePublish,
    NameScope,
};
pub use differential_format::{
    DifferentialFormat, ThemeColor, XfBorder, XfColor, XfColorSource, XfFontScheme, XfFontWeight,
    XfGradient, XfGradientStop, XfProperties, XfProperty,
};
pub use drawing_metadata::{AnchorBehavior, AnchorPoint, SheetAnchor};
pub use encryption::{EncryptionProfile, WeakEncryptionPolicy};
pub use ent_ex_u2::EntExU2;
pub use environment::{LinkUpdateMode, ObjectDisplayMode, WorkbookEnvironment};
pub use error::{EncryptionKind, Error, Result};
pub use external_link::{
    CacheRow, CachedValue, ClipboardFormat, ErrorValue, Links, Name, NameBody, Sheet,
    SheetReference, SupportingBook, ValueMatrix, Workbook as ExternalWorkbook,
};
pub use fbi::{Fbi, FontScaleBasis};
pub use font::{Font, FontCharset, FontEscapement, FontFamily, FontUnderline};
pub use force_full_calculation::ForceFullCalculation;
pub use formula_errors::{
    FormulaErrorChecks, FormulaErrorFeature, FormulaErrorHeader, FormulaErrorRange,
};
pub use formula_metadata::Metadata as FormulaMetadata;
pub use function_group::{BuiltInFunctionCategories, FunctionGroups};
pub use header_footer_picture::HeaderFooterPicture;
pub use interface_records::{InterfaceEnd, InterfaceHdr};
pub use layout::{Column, Row};
pub use leniency::{FormattingDefect, Leniency, ToleranceReport, ToleratedDefect};
pub use list_object::{
    CachedDiskHeader, ExternalTableField, ExternalTableMetadata, ExternalTableVersion,
    ListColumnId, ListObject, ListObjectColumn, ListObjectFeatureVersion, ListObjectId,
    ListObjectRange, ListObjectSourceMetadata, ListObjectStyleOptions, ListTotalAggregation,
    OpaqueListObjectFeature, OpaqueListObjectFutureRecord, WebColumnType, WebDefaultValue,
    WebEditMode, WebFieldInfo, WebInvalidCell, WebReadingOrder, WebTableField, WebTableMetadata,
    XmlColumnMapping, XmlDataType, XmlTableField, XmlTableMetadata,
};
pub use marker_format::{ChartRgb, DataMarkerKind, MarkerFormat};
pub use mdx_metadata::{
    CubeFunction, KpiProperty, Mdb, MdtInfo, MdtInfoFlags, MdxKpi, MdxMetadata, MdxMetadataDir,
    MdxMetadataRecord, MdxProp, MdxSet, MdxSetSortOrder, MdxTuple,
};
pub use number_format::{
    DateSystem, EffectiveExtendedFormat, ExtendedFormat, ExtendedFormatApplications,
    ExtendedFormatKind, Formatting, NumberFormat,
};
pub use object_link::{ObjectLink, ObjectLinkTarget};
pub use page_setup::{
    HeaderFooter, PageBreak, PageSetup, PrintComments, PrintErrors, PrintOrder, PrintOrientation,
    PrintSetup,
};
pub use palette::{Color, Palette};
pub use phonetic_info::{
    PhoneticAlignment, PhoneticFormat, PhoneticInfo, PhoneticRange, PhoneticType,
};
pub use pivot_olap::{
    HiddenMemberSet, OlapSequence, PivotFieldOlapExt, PivotHierarchy, PivotHierarchyAxis,
    PivotItemOlapFlags, PivotPageItemOlapExt, PivotViewOlapHeader,
};
pub use plot_growth::{FixedPoint, PlotGrowth};
pub use print_flags::{GridSet, PrintRowCol};
pub use printer_driver::PrinterDriverData;
pub use query_table::{
    HtmlFormatting, OleDbConnection, QueryParameter, QueryParameterType, QuerySource, QueryTable,
    TextCodePage, TextDelimiter, TextField, TextFieldFormat, TextQuery,
};
pub use records::{PhoneticRun, PhoneticString, SharedStringFormatRun, SharedStringProperties};
pub use revision_log::{
    OpaqueRevisionRecord, REVISION_LOG_STREAM_NAME, Revision, RevisionChange, RevisionHeader,
    RevisionLog, RrdChgCellRevision, RrdInsDelRevision, RrdMoveRevision,
};
pub use revision_records::{
    FileLock, FileLockPurpose, RevisionCellContent, RevisionCellLocation, RevisionCellRange,
    RevisionRecordHeader, RevisionType, RrInsertSh, RrTabId, RrdChgCell, RrdConflict, RrdHead,
    RrdInfo, RrdInsDel, RrdMove, RrdRenSheet, RrdUserView, ShortDtr, UsrExcl,
};
pub use row_block_index::{
    DbCellRecord, IndexedRow, RowBlock, RowBlockIndex, WorksheetIndexRecord,
};
pub use scenario::{Scenario, ScenarioCell, ScenarioManager, ScenarioRange};
pub use scl::Scl;
pub use ser_aux_err_bar::{ErrorBarDirection, ErrorBarSource, SerAuxErrBar};
pub use ser_aux_trend::{SerAuxTrend, TrendlineKind};
pub use shapes::Shape;
pub use shared_string_index::{SharedStringBucket, SharedStringIndex};
pub use sheet_ext::{SheetExt, SheetExtOptional};
pub use sheet_metadata::{SheetKind, SheetMetadata, SheetVisibility};
pub use style_ext::{StyleCategory, StyleExt};
pub use sxview_link::SXViewLink;
pub use table_styles::{
    DifferentialFormatId, TableStyle, TableStyleElement, TableStyleRegion, TableStyles,
};
pub use theme::Theme;
pub use toolbar::{Control, Toolbar, ToolbarSet, VisualData, Wrapper};
pub use user_routing::{CUsr, CbUsr, DocRoute, RecipName, RoutingDelivery, UsrInfo};
pub use uses_elfs::UsesElfs;
pub use vba::{VbaMetadata, VbaProjectStorage};
pub use web_pub::{WebPageType, WebPub, WebPubRange, WebSourceType};
pub use workbook::{
    OpenOptions, SourceBackedCell, SourceBackedError, SourceBackedLimits, SourceBackedWorkbook,
    SourceBackedWorksheet, SourceBackedWorksheetDescriptor, Workbook,
};
pub use workbook_view::{WorkbookView, WorkbookWindow};
pub use worksheet::Worksheet;
pub use worksheet::layout::Layout;
pub use writer::{
    ShapeColor, ShapeFill, ShapeKind, ShapeLine, ShapeText, ShapeTextRun, ShapeWrite, Writer,
};
pub use xf_ext::{ExtProp, FullColorExt, FullColorType, XfExt};
pub use xml_map::{
    DataBinding, LoadMode, Map, MapId, MapInfo, NamespaceDeclaration, OpaqueXml, Schema, SchemaId,
    XPath,
};
