#![forbid(unsafe_code)]
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
pub mod workbook;

/// Worksheet parsing implementation and worksheet-owned semantic views.
pub mod worksheet;

/// Cell value parsing and representation
mod cell;

/// BIFF8 worksheet data-validation records.
mod data_validation;

/// BIFF8 calculation and recalculation records.
mod calculation;
/// BIFF8 chart-sheet and embedded-chart metadata and mutation.
pub mod chart;

/// BIFF8 chart layout future records (`CrtLayout12`, `CrtLayout12A`).
mod chart_layout;

/// BIFF8 `MarkerFormat` record: data-marker color, size, and shape.
mod marker_format;

/// BIFF8 `ObjectLink` record: the chart object a text is linked to.
mod object_link;

/// BIFF8 `UsesELFs` record: the natural language formula flag.
mod uses_elfs;

/// BIFF8 `SerAuxErrBar` record: error bar properties.
mod ser_aux_err_bar;

/// BIFF8 `SerAuxTrend` record: a trendline.
mod ser_aux_trend;

/// BIFF8 `Fbi` record: scalable chart font information.
mod fbi;

/// BIFF8 `EntExU2` record: an application-specific cache, preserved opaquely.
mod ent_ex_u2;

/// BIFF8 chart axis-group records (`AxesUsed`, `AxisParent`).
mod axes_used;

/// BIFF8 `BopPop` record: bar of pie / pie of pie chart group attributes.
mod bop_pop;

/// BIFF8 `PlotGrowth` record: plot area font-scaling factors.
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
mod chart_property_stream;

/// BIFF8 `ForceFullCalculation` record: the forced calculation mode.
mod force_full_calculation;

/// BIFF8 `Backup` record: the workbook save-backup flag.
mod backup;

/// BIFF8 `BkHim` record: sheet background image data.
mod background_picture;

/// BIFF8 `CellWatch` record: a watched-cell reference.
mod cell_watch;

/// BIFF8 user-interface collection markers (`InterfaceHdr`, `InterfaceEnd`).
mod interface_records;

/// BIFF8 `HFPicture` record: a sheet header/footer picture.
mod header_footer_picture;

/// BIFF8 `Pls` record: printer driver `DEVMODE` data spanning `Continue`
/// records.
mod printer_driver;

/// BIFF8 worksheet scenario manager records.
mod scenario;

/// Inert BIFF8 VBA project markers and object code names.
mod vba;

/// Typed, inert legacy XLS XML maps and their root-level `XML` stream.
pub mod xml_map;

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
mod book_ext;

/// Typed BIFF8 `CompressPictures` recommendation and lossless snapshots.
pub mod picture_compression;

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

/// BIFF8 `SheetExt` record: sheet tab color and publish state.
mod sheet_ext;

/// BIFF8 worksheet window, zoom, pane, and selection state.
pub mod view;

/// BIFF8 worksheet print and page setup.
mod page_setup;

/// BIFF8 worksheet print display flags (`PrintRowCol`, `GridSet`).
mod print_flags;

/// BIFF8 `SXViewLink` record: the PivotTable view linked to a Pivot Chart.
mod sxview_link;

/// Legacy BIFF8 conditional formatting.
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
pub mod hyperlinks;

/// Comment/note parsing (NOTE 0x001C)
pub mod comments;
pub mod compatibility;

/// Inert BIFF8 Office Toolbars stream (`XCB`) metadata.
pub mod toolbar;

/// AutoFilter and sort parsing (AUTOFILTERINFO 0x009D, AUTOFILTER 0x009E, SORT 0x0090)
pub mod autofilter;

/// Extended BIFF8 range-sort metadata (`SortData` and `SortCond12`).
mod sort_data;

/// BIFF8 `QUERYTABLE` sequence: typed, inert query tables and connections.
mod query_table;

/// BIFF8 `RealTimeData` record: typed, inert real-time data (RTD) topics.
pub mod real_time_data;

/// Typed, inert readers for the shared-workbook RRD revision record family.
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
mod user_routing;

#[forbid(unsafe_code)]
mod pivot_editor;
/// Pivot table parsing (SXVIEW, SXVD, SXVI, SXDI, SXVS, SXPI)
#[forbid(unsafe_code)]
pub mod pivot_table;
pub use pivot_editor::PivotViewEditor;
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

/// XLS file writing
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
pub use workbook::{OpenOptions, Workbook};
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
