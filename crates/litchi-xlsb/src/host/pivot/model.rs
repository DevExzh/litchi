//! Typed model for the XLSB PivotCache definition stream (MS-XLSB 2.1.7.38).
//!
//! All types are inert data snapshots: relationship identifiers, external
//! connection identifiers, and MDX/formula payloads are stored verbatim and
//! are never dereferenced, contacted, or executed.

use crate::package::error::Error;

/// A PivotCache parsed from one `pivotCacheDefinition` part (MS-XLSB 2.4.168
/// `BrtBeginPivotCacheDef` and its record collection).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct PivotCacheDefinition {
    /// Data functionality level the cache was last refreshed with (`bVerCacheLastRefresh`).
    pub version_last_refresh: u8,
    /// Lowest data functionality level required to refresh (`bVerCacheRefreshableMin`).
    pub version_refreshable_min: u8,
    /// Data functionality level the cache was created with (`bVerCacheCreated`).
    pub version_created: u8,
    /// Cache records exist (`fSaveData`). Always `false` for OLAP caches.
    pub save_data: bool,
    /// Cache records are invalid and must be ignored (`fInvalid`).
    pub invalid: bool,
    /// Refresh the cache on load (`fRefreshOnLoad`).
    pub refresh_on_load: bool,
    /// Memory optimizations are applied (`fOptimizeCache`).
    pub optimize_cache: bool,
    /// Refresh is enabled (`fEnableRefresh`).
    pub enable_refresh: bool,
    /// Refresh asynchronously (`fBackgroundQuery`).
    pub background_query: bool,
    /// Upgrade to functionality level 3 on next refresh (`fUpgradeOnRefresh`).
    pub upgrade_on_refresh: bool,
    /// Cache stores information for cube functions (`fSheetData`, OLAP only).
    pub cube_functions: bool,
    /// Source supports OLAP subselect (`fSupportSubquery`).
    pub support_subquery: bool,
    /// Source supports attribute drilldown (`fSupportAttribDrill`).
    pub support_attrib_drill: bool,
    /// Unused (ghost) cache items retained before discarding on refresh
    /// (`citmGhostMax`; `-1` = application-optimized, `0` = discard all).
    pub ghost_items_max: i32,
    /// Last refresh time as an Excel serial date (`xnumRefreshedDate`, `DateAsXnum`).
    pub refreshed_date_serial: f64,
    /// Number of cache records (`cRecords`; meaningful only when `save_data`).
    pub record_count: u32,
    /// Name of the user who last refreshed the cache (`stRefreshedWho`).
    pub refreshed_by: Option<String>,
    /// Inert relationship identifier of the PivotCache Records part (`stRelIDRecords`).
    pub records_rel_id: Option<String>,
    /// Source data of the cache (`BrtBeginPCDSource` collection).
    pub source: Option<PivotCacheSource>,
    /// Cache fields in cache field index order (`BrtBeginPCDFields` collection).
    pub fields: Vec<PivotCacheField>,
    /// OLAP cache hierarchies (`BrtBeginPCDHierarchies` collection).
    pub hierarchies: Vec<PivotCacheHierarchy>,
    /// OLAP sheet-data tuple cache (`BrtBeginPCDSDTupleCache` collection).
    pub tuple_cache: Option<PivotCacheTupleCache>,
    /// Calculated items (`BrtBeginPCDCalcItems` collection).
    pub calculated_items: Vec<CalculatedItem>,
    /// OLAP calculated members and named sets (`BrtBeginPCDCalcMems` collection).
    pub calculated_members: Vec<CalculatedMember>,
    /// Excel 2013 extension data (`BrtBeginPCD14`).
    pub ext14: Option<PivotCacheDefinitionExt14>,
}

/// A typed snapshot of one `pivotCacheRecords*.bin` part (MS-XLSB 2.1.7.39).
///
/// The `record_count` value is retained from `BrtBeginPivotCacheRecords` so a
/// caller can distinguish the wire declaration from the number of records
/// materialized in `records`. For the valid, fully parsed representation the
/// two values are equal. The same snapshot corresponds to the OOXML
/// `pivotCacheRecords` part and its `count` attribute.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct PivotCacheRecords {
    /// Declared number of cache records (`crecords`).
    pub record_count: u32,
    /// Cache records in source-row order.
    pub records: Vec<PivotCacheRecord>,
}

/// One source row from a PivotCache records part.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct PivotCacheRecord {
    /// Values in the order of source cache fields (`fSrcField = 1`).
    /// `Index` values reference the corresponding field's shared items.
    pub values: Vec<PivotCacheItemValue>,
}

/// PivotCache source data type (`iSrcType`, MS-XLSB 2.4.166).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum PivotCacheSourceType {
    /// Sheet source data (`BrtBeginPCDSRange` present).
    Worksheet = 0,
    /// External source data; `dwConnID` identifies the external connection.
    External = 1,
    /// Multiple consolidation ranges (`BrtBeginPCDSConsol` present).
    Consolidation = 2,
    /// Scenario source data.
    Scenario = 3,
}

impl TryFrom<u32> for PivotCacheSourceType {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Worksheet),
            1 => Ok(Self::External),
            2 => Ok(Self::Consolidation),
            3 => Ok(Self::Scenario),
            _ => Err(Error::Unrecognized {
                typ: "PivotCache source type".to_string(),
                val: format!("0x{value:08X}"),
            }),
        }
    }
}

/// PivotCache source data properties (`BrtBeginPCDSource`, MS-XLSB 2.4.166).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheSource {
    /// Source data type.
    pub source_type: PivotCacheSourceType,
    /// Identifier of the external connection (`dwConnID`); `Some` only for
    /// [`PivotCacheSourceType::External`]. Inert; never resolved.
    pub connection_id: Option<u32>,
    /// Workbook-contained source (`BrtBeginPCDSRange`).
    pub worksheet: Option<PivotCacheWorksheetSource>,
    /// Consolidation source (`BrtBeginPCDSConsol` collection).
    pub consolidation: Option<PivotCacheConsolidationSource>,
}

/// A cell range (`UncheckedRfX`, MS-XLSB 2.5.154).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PivotCacheRange {
    /// First row (`rwFirst`).
    pub first_row: i32,
    /// Last row (`rwLast`).
    pub last_row: i32,
    /// First column (`colFirst`).
    pub first_column: i32,
    /// Last column (`colLast`).
    pub last_column: i32,
}

/// Source data contained in a workbook (`BrtBeginPCDSRange`, MS-XLSB 2.4.167).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheWorksheetSource {
    /// Defined name that is the source (`namedRange`); `Some` iff located by name.
    pub named_range: Option<String>,
    /// The defined name is a built-in name (`fBuiltIn`).
    pub built_in_name: bool,
    /// Sheet the source is scoped to (`sheetName`).
    pub sheet_name: Option<String>,
    /// Inert relationship identifier of an external workbook (`relId`); never resolved.
    pub external_rel_id: Option<String>,
    /// Cell range that is the source (`range`); `Some` iff not located by name.
    pub range: Option<PivotCacheRange>,
}

/// Multiple-consolidation-ranges source (`BrtBeginPCDSConsol`, MS-XLSB 2.4.150).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheConsolidationSource {
    /// Exactly one automatically-created page field exists (`fAutoPage`).
    pub auto_page: bool,
    /// Consolidation ranges (`BrtBeginPCDSCSets` collection).
    pub sets: Vec<PivotCacheConsolidationSet>,
    /// Page fields (`BrtBeginPCDSCPages` collection, at most four).
    pub pages: Vec<PivotCacheConsolidationPage>,
}

/// One consolidation range (`BrtBeginPCDSCSet`, MS-XLSB 2.4.154).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheConsolidationSet {
    /// Per-page item indexes (`rgiItem`); `u32::MAX` marks an unused page slot.
    pub item_indexes: [u32; 4],
    /// Defined name of the consolidation range (`irstName`).
    pub named_range: Option<String>,
    /// The defined name is a built-in name (`fBuiltIn`).
    pub built_in_name: bool,
    /// Sheet the range is scoped to (`irstSheet`).
    pub sheet_name: Option<String>,
    /// Inert relationship identifier of an external workbook (`irstRelId`).
    pub external_rel_id: Option<String>,
    /// Cell range of the consolidation (`rfx`).
    pub range: Option<PivotCacheRange>,
}

/// One consolidation page field (`BrtBeginPCDSCPage`, MS-XLSB 2.4.151).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheConsolidationPage {
    /// Page item labels (`stName` of each `BrtBeginPCDSCPItem`).
    pub item_names: Vec<String>,
}

/// A cache field (`BrtBeginPCDField` collection, MS-XLSB 2.4.136).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheField {
    /// Unique name of the cache field (`stFldName`).
    pub name: String,
    /// User-facing caption (`stFldCaption`).
    pub caption: Option<String>,
    /// Number format applied to all source values (`ifmt`, a `PivotNumFmt`);
    /// `None` when the default format applies (`PivotNumFmtExt` = `0xFFFFFFFF`).
    pub number_format: Option<u32>,
    /// ODBC SQL data type of the field (`wTypeSql`, `TypeSql`; ODBC caches only).
    pub sql_type: u16,
    /// Cache hierarchy index this field is associated with (`ihdb`; OLAP only).
    pub hierarchy_index: u32,
    /// Cache hierarchy level ordinal (`isxtl`; `0x00007FFF` = whole hierarchy).
    pub level: u32,
    /// Member property cache field indexes (`rgisxtmp`; OLAP only).
    pub member_property_fields: Vec<u32>,
    /// Name of the associated OLAP member property (`stMemPropName`).
    pub member_property_name: Option<String>,
    /// Calculated field formula (`fldFmla`), stored as raw Ptg tokens.
    pub formula: Option<PivotParsedFormulaData>,
    /// Server-based page field (`fServerBased`; ODBC caches only).
    pub server_based: bool,
    /// Unique value list was unavailable at refresh (`fCantGetUniqueItems`).
    pub cant_get_unique_items: bool,
    /// Field corresponds to a source data entity (`fSrcField`).
    pub source_field: bool,
    /// Field is associated with an OLAP member property (`fOlapMemPropField`).
    pub olap_member_property_field: bool,
    /// A `BrtPCDField14` record marked this field as ignorable (MS-XLSB 2.4.725).
    pub ignore: bool,
    /// Raw cache items and statistics (`BrtBeginPCDFAtbl` collection).
    pub shared_items: PivotCacheSharedItems,
    /// Grouping definition (`BrtBeginPCDFGroup` collection).
    pub grouping: Option<PivotCacheFieldGrouping>,
}

/// Raw (ungrouped) cache items of a field plus their summary statistics.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct PivotCacheSharedItems {
    /// Summary statistics from `BrtBeginPCDFAtbl` (MS-XLSB 2.4.131).
    pub stats: Option<PivotCacheSharedItemsStats>,
    /// Raw cache items in cache item index order.
    pub items: Vec<PivotCacheItem>,
}

/// Summary statistics over a field's shared items (`BrtBeginPCDFAtbl`).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PivotCacheSharedItemsStats {
    /// Field contains text items (`fTextEtcField`).
    pub text_field: bool,
    /// Field contains non-date items (`fNonDates`).
    pub non_dates: bool,
    /// Field contains date items (`fDateInField`).
    pub date_in_field: bool,
    /// Field contains at least one text item (`fHasTextItem`).
    pub has_text_item: bool,
    /// Field contains at least one blank item (`fHasBlankItem`).
    pub has_blank_item: bool,
    /// Field contains mixed value types ignoring blanks (`fMixedTypesIgnoringBlanks`).
    pub mixed_types_ignoring_blanks: bool,
    /// Field is numeric (`fNumField`).
    pub numeric_field: bool,
    /// All numeric items are integers (`fIntField`).
    pub integer_field: bool,
    /// Field contains long text items (`fHasLongTextItem`).
    pub has_long_text_item: bool,
    /// Declared number of shared items (`citems`).
    pub item_count: u32,
    /// Minimum numeric/date value (`xnumMin`); `Some` iff `fNumMinMaxValid`.
    pub minimum: Option<f64>,
    /// Maximum numeric/date value (`xnumMax`); `Some` iff `fNumMinMaxValid`.
    pub maximum: Option<f64>,
}

/// One cache item: a value plus optional additional information.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheItem {
    /// Item value and type.
    pub value: PivotCacheItemValue,
    /// Additional information (`PCDIAddlInfo`, MS-XLSB 2.5.100); present for
    /// items carried by `BrtPCDIA*` records, absent otherwise.
    pub additional: Option<PivotCacheItemInfo>,
}

/// Cache item value (`BrtPCDI*`/`BrtPCDIA*` records, MS-XLSB 2.4.728-2.4.740).
#[derive(Debug, Clone, PartialEq)]
pub enum PivotCacheItemValue {
    /// Missing value (`BrtPCDIMissing`).
    Missing,
    /// Numeric value (`BrtPCDINumber`).
    Number(f64),
    /// Boolean value (`BrtPCDIBoolean`).
    Boolean(bool),
    /// Error value (`BrtPCDIError`).
    Error(PivotCacheErrorCode),
    /// Text value, boxed to keep the enum small (`BrtPCDIString`).
    String(Box<str>),
    /// Date/time value (`BrtPCDIDatetime`).
    DateTime(PivotCacheDateTime),
    /// Index into the base field's shared items (`BrtPCDIIndex`);
    /// only valid in discrete groupings.
    Index(u32),
}

/// Error code of an error cache item (`BErr`, MS-XLSB 2.5.98.2).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum PivotCacheErrorCode {
    /// `#NULL!`
    Null = 0x00,
    /// `#DIV/0!`
    Div0 = 0x07,
    /// `#VALUE!`
    Value = 0x0F,
    /// `#REF!`
    Ref = 0x17,
    /// `#NAME?`
    Name = 0x1D,
    /// `#NUM!`
    Num = 0x24,
    /// `#N/A`
    NA = 0x2A,
    /// `#GETTING_DATA`
    GettingData = 0x2B,
}

impl TryFrom<u8> for PivotCacheErrorCode {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::Div0),
            0x0F => Ok(Self::Value),
            0x17 => Ok(Self::Ref),
            0x1D => Ok(Self::Name),
            0x24 => Ok(Self::Num),
            0x2A => Ok(Self::NA),
            0x2B => Ok(Self::GettingData),
            _ => Err(Error::Unrecognized {
                typ: "PivotCache item error code".to_string(),
                val: format!("0x{value:02X}"),
            }),
        }
    }
}

/// Broken-down date/time of a date cache item (`PCDIDateTime`, MS-XLSB 2.5.101).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PivotCacheDateTime {
    /// Year (1900-9999).
    pub year: u16,
    /// Month (1-12).
    pub month: u16,
    /// Day of month (0-31; 0 = only year/month are meaningful).
    pub day: u8,
    /// Hour (0-23).
    pub hour: u8,
    /// Minute (0-59).
    pub minute: u8,
    /// Second (0-59).
    pub second: u8,
}

/// Additional cache item information (`PCDIAddlInfo`, MS-XLSB 2.5.100).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheItemInfo {
    /// Item is unused (no longer present in the source data) (`fGhost`).
    pub ghost: bool,
    /// Item is a calculated item (`fFmla`).
    pub calculated: bool,
    /// Display caption of the item (`stCaption`).
    pub caption: Option<String>,
    /// OLAP member property item indexes (`rgIMemProps`; `-1` = no item).
    pub member_property_items: Vec<i32>,
}

/// Grouping definition of a cache field (`BrtBeginPCDFGroup`, MS-XLSB 2.4.135).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheFieldGrouping {
    /// Grouping parent cache field index (`ifdbParent`).
    pub parent_field: Option<u32>,
    /// Base cache field index whose items are grouped (`ifdbBase`).
    pub base_field: Option<u32>,
    /// Range grouping bounds (`BrtBeginPCDFGRange`).
    pub range: Option<PivotCacheRangeGrouping>,
    /// Discrete grouping (`BrtBeginPCDFGDiscrete`).
    pub discrete: Option<PivotCacheDiscreteGrouping>,
    /// Grouping cache items (`BrtBeginPCDFGItems` collection).
    pub items: Vec<PivotCacheItem>,
}

/// Range grouping bounds (`BrtBeginPCDFGRange`, MS-XLSB 2.4.134).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PivotCacheRangeGrouping {
    /// Grouping unit (`iByType`).
    pub group_by: PivotCacheGroupBy,
    /// Start is taken from the source data (`fAutoStart`).
    pub auto_start: bool,
    /// End is taken from the source data (`fAutoEnd`).
    pub auto_end: bool,
    /// Start/end are date serials (`fDates`).
    pub dates: bool,
    /// Start of the grouping range (`xnumStart`).
    pub start: f64,
    /// End of the grouping range (`xnumEnd`).
    pub end: f64,
    /// Grouping interval (`xnumBy`).
    pub interval: f64,
}

/// Range grouping unit (`iByType`, MS-XLSB 2.4.134).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum PivotCacheGroupBy {
    /// Numeric range grouping.
    NumericRange = 0,
    /// Group by seconds.
    Seconds = 1,
    /// Group by minutes.
    Minutes = 2,
    /// Group by hours.
    Hours = 3,
    /// Group by days.
    Days = 4,
    /// Group by months.
    Months = 5,
    /// Group by quarters.
    Quarters = 6,
    /// Group by years.
    Years = 7,
}

impl TryFrom<u8> for PivotCacheGroupBy {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::NumericRange),
            1 => Ok(Self::Seconds),
            2 => Ok(Self::Minutes),
            3 => Ok(Self::Hours),
            4 => Ok(Self::Days),
            5 => Ok(Self::Months),
            6 => Ok(Self::Quarters),
            7 => Ok(Self::Years),
            _ => Err(Error::Unrecognized {
                typ: "PivotCache grouping type".to_string(),
                val: format!("0x{value:02X}"),
            }),
        }
    }
}

/// Discrete grouping (`BrtBeginPCDFGDiscrete`, MS-XLSB 2.4.132).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheDiscreteGrouping {
    /// Indexes of the base field's cache items that form the group
    /// (`iitem` of each `BrtPCDIIndex`).
    pub item_indexes: Vec<u32>,
}

/// An OLAP cache hierarchy (`BrtBeginPCDHierarchy`, MS-XLSB 2.4.146).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheHierarchy {
    /// MDX unique name (`stUnique`).
    pub unique_name: String,
    /// Display caption (`stCaption`).
    pub caption: String,
    /// MDX unique name of the dimension (`stDimUnq`).
    pub dimension_unique_name: Option<String>,
    /// MDX unique name of the default member (`stDefaultUnq`).
    pub default_member_unique_name: Option<String>,
    /// MDX unique name of the All member (`stAllUnq`).
    pub all_member_unique_name: Option<String>,
    /// Display caption of the All member (`stAllDisp`).
    pub all_member_display: Option<String>,
    /// Display folder (`stDispFld`).
    pub display_folder: Option<String>,
    /// Measure group (`stMeasGrp`).
    pub measure_group: Option<String>,
    /// Hierarchy is a measure collection (`fMeasure`).
    pub measure: bool,
    /// Hierarchy is a named set (`fSet`).
    pub set: bool,
    /// Hierarchy is an attribute hierarchy (`fAttributeHierarchy`).
    pub attribute_hierarchy: bool,
    /// Hierarchy is the measure hierarchy (`fMeasureHierarchy`).
    pub measure_hierarchy: bool,
    /// Hierarchy corresponds to exactly one cache field (`fOnlyOneField`).
    pub only_one_field: bool,
    /// Hierarchy is a time hierarchy (`fTimeHierarchy`).
    pub time_hierarchy: bool,
    /// Hierarchy is the key attribute hierarchy (`fKeyAttributeHierarchy`).
    pub key_attribute_hierarchy: bool,
    /// Hierarchy is hidden (`fHidden`).
    pub hidden: bool,
    /// Unbalanced-real flag, `None` when unknown (`fUnbalancedRealKnown`/`fUnbalancedReal`).
    pub unbalanced_real: Option<bool>,
    /// Unbalanced-group flag, `None` when unknown (`fUnbalancedGroupKnown`/`fUnbalancedGroup`).
    pub unbalanced_group: Option<bool>,
    /// Attribute member value type (`wAttributeMemberValueType`; `0x0007` = date);
    /// `None` when `fAttributeMemberValueTypeKnown` is 0.
    pub attribute_member_value_type: Option<u16>,
    /// Number of OLAP levels (`cLevels`).
    pub level_count: u32,
    /// Parent cache hierarchy index of a named set (`isetParent`).
    pub set_parent_index: Option<u32>,
    /// KPI set identifier (`iconSet`, a `KPISets` value).
    pub icon_set: i32,
    /// Cache field index per level ordinal (`rgifdb` of `BrtBeginPCDHFieldsUsage`;
    /// `-1` = level ordinal unused).
    pub field_usage: Vec<i32>,
    /// User-defined grouping levels (`BrtBeginPCDHGLevels` collection).
    pub grouping_levels: Vec<PivotCacheGroupingLevel>,
    /// User-defined grouping groups (`BrtBeginPCDHGLGroups` collection).
    pub grouping_groups: Vec<PivotCacheGroupingGroup>,
    /// Excel 2013 named-set extension (`BrtPCDH14`, MS-XLSB 2.4.726).
    pub ext14: Option<PivotCacheHierarchyExt14>,
}

/// A user-defined OLAP grouping level (`BrtBeginPCDHGLevel`, MS-XLSB 2.4.139).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheGroupingLevel {
    /// Level is user-defined rather than a source cube level (`fGroupLevel`).
    pub group_level: bool,
    /// MDX unique name of the level (`stUnique`).
    pub unique_name: String,
    /// Caption of the level (`stLevelName`).
    pub caption: String,
}

/// A user-defined OLAP grouping group (`BrtBeginPCDHGLGroup`, MS-XLSB 2.4.143).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheGroupingGroup {
    /// Group number (`iGrpNum`).
    pub group_number: i32,
    /// Group name (`stName`).
    pub name: String,
    /// MDX unique name of the group member (`stUniqueName`).
    pub unique_name: String,
    /// Group caption (`stCaption`).
    pub caption: String,
    /// MDX unique name of the parent member (`stParentUniqueName`).
    pub parent_unique_name: Option<String>,
    /// Group members (`BrtBeginPCDHGLGMembers` collection).
    pub members: Vec<PivotCacheGroupingGroupMember>,
}

/// A member of a user-defined grouping group (`BrtBeginPCDHGLGMember`, MS-XLSB 2.4.141).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheGroupingGroupMember {
    /// `true` when `unique_name` references a group of the subsequent level (`fGroup`).
    pub is_group: bool,
    /// MDX unique name of the member or group (`stUnique`).
    pub unique_name: String,
}

/// Excel 2013 named-set extension of a cache hierarchy (`BrtPCDH14`).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheHierarchyExt14 {
    /// Flatten hierarchies of the named set (`fFlattenHierarchies`).
    pub flatten_hierarchies: bool,
    /// Named set contains measures (`fMeasureSet`).
    pub measure_set: bool,
    /// Hierarchize distinct (`fHierarchizeDistinct`).
    pub hierarchize_distinct: bool,
    /// Hierarchy is an ignorable placeholder (`fIgnorable`).
    pub ignorable: bool,
    /// Cache hierarchy indexes (`rgihdb`; `-2` = measure hierarchy, `-1` = none).
    pub hierarchy_indexes: Vec<i32>,
}

/// A calculated item (`BrtBeginPCDCalcItem`, MS-XLSB 2.4.124).
#[derive(Debug, Clone, PartialEq)]
pub struct CalculatedItem {
    /// Calculated item formula, stored as raw Ptg tokens.
    pub formula: PivotParsedFormulaData,
    /// Field references used by the formula (`BrtBeginPNames` collection).
    pub names: Vec<PivotName>,
    /// Item filters applied by the rule (`BrtBeginPRFilters` collection).
    pub filters: Vec<PivotRuleFilter>,
}

/// A `PivotParsedFormula` (MS-XLSB 2.5.98.15) stored verbatim; never evaluated.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PivotParsedFormulaData {
    /// Ptg token bytes (`rgce`).
    pub tokens: Vec<u8>,
    /// Ancillary bytes (`rgcb`).
    pub extra: Vec<u8>,
}

/// A field reference of a calculated field/item formula (`BrtBeginPName`, MS-XLSB 2.4.176).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotName {
    /// Cache field index (`ifdb`; `0xFFFFFFFF` for calculated items).
    pub field_index: u32,
    /// Aggregation function (`ifn`).
    pub function: PivotNameFunction,
    /// The cache field was not found (`fErrName`).
    pub err_name: bool,
    /// Item references of a calculated item (`BrtBeginPNPairs` collection).
    pub pairs: Vec<PivotNamePair>,
}

/// Aggregation function of a pivot name (`ifn`, MS-XLSB 2.4.176).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum PivotNameFunction {
    /// Sum.
    Sum = 0,
    /// Count of non-blank items.
    CountA = 1,
    /// Average.
    Average = 2,
    /// Maximum.
    Max = 3,
    /// Minimum.
    Min = 4,
    /// Product.
    Product = 5,
    /// Count of numeric items.
    Count = 6,
    /// Sample standard deviation.
    StDev = 7,
    /// Population standard deviation.
    StDevP = 8,
    /// Sample variance.
    Var = 9,
    /// Population variance.
    VarP = 10,
    /// Not specified.
    Unspecified = 255,
}

impl TryFrom<u8> for PivotNameFunction {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Sum),
            1 => Ok(Self::CountA),
            2 => Ok(Self::Average),
            3 => Ok(Self::Max),
            4 => Ok(Self::Min),
            5 => Ok(Self::Product),
            6 => Ok(Self::Count),
            7 => Ok(Self::StDev),
            8 => Ok(Self::StDevP),
            9 => Ok(Self::Var),
            10 => Ok(Self::VarP),
            255 => Ok(Self::Unspecified),
            _ => Err(Error::Unrecognized {
                typ: "PivotName aggregation function".to_string(),
                val: format!("{value}"),
            }),
        }
    }
}

/// An item reference of a calculated item (`BrtBeginPNPair`, MS-XLSB 2.4.178).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PivotNamePair {
    /// `item_index` is not a cache item index (`fPhysical`).
    pub physical: bool,
    /// `item_index` is relative to the calculated item (`fRelative`).
    pub relative: bool,
    /// Cache field index (`ifield`).
    pub field_index: u32,
    /// Cache item or visible item index (`iitem`).
    pub item_index: i32,
}

/// A PivotTable-rule item filter (`BrtBeginPRFilter`, MS-XLSB 2.4.180).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotRuleFilter {
    /// Cache/pivot field index, or `-2` for the data field (`isxvd`).
    pub field: i32,
    /// Included item subtotal types (`itmtypeData` .. `itmtypeVARP`,
    /// bits 0-12 of the `PRFilter` flags).
    pub item_types: u32,
    /// Field header is included (`fSelected`).
    pub selected: bool,
    /// Item indexes (`iitem` of each `BrtBeginPRFItem`).
    pub items: Vec<u32>,
}

/// An OLAP calculated member or named set (`BrtBeginPCDCalcMem`, MS-XLSB 2.4.126;
/// `PCDCalcMemCommon`, MS-XLSB 2.5.99).
#[derive(Debug, Clone, PartialEq)]
pub struct CalculatedMember {
    /// Unique name of the member or set (`stName`).
    pub name: String,
    /// MDX expression (`stMdx`), stored verbatim; never evaluated.
    pub mdx: String,
    /// MDX `SOLVE_ORDER` (`wSolveOrder`).
    pub solve_order: i32,
    /// `true` for a named set, `false` for a calculated member (`fSet`).
    pub is_set: bool,
    /// Member name (`stMemberName`).
    pub member_name: Option<String>,
    /// Source hierarchy (`stSourceHier`).
    pub source_hierarchy: Option<String>,
    /// Parent member unique name (`stParentUnique`).
    pub parent_unique: Option<String>,
    /// Excel 2013 extension (`BrtBeginPCDCalcMem14`).
    pub ext14: Option<CalculatedMemberExt14>,
}

/// Excel 2013 extension of a calculated member (`BrtBeginPCDCalcMem14`, MS-XLSB 2.4.127).
#[derive(Debug, Clone, PartialEq)]
pub struct CalculatedMemberExt14 {
    /// Flatten hierarchies (`fFlattenHierarchies`).
    pub flatten_hierarchies: bool,
    /// Named set is dynamic (`fDynamicSet`).
    pub dynamic_set: bool,
    /// Hierarchize distinct (`fHierarchizeDistinct`).
    pub hierarchize_distinct: bool,
    /// Display folder (`irstDisplayFolder`).
    pub display_folder: String,
    /// Long MDX overflow expression (`irstMDXFormulaLong`); when present it
    /// supersedes [`CalculatedMember::mdx`].
    pub long_mdx: Option<String>,
}

/// OLAP sheet-data tuple cache (`BrtBeginPCDSDTupleCache`, MS-XLSB 2.4.164).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct PivotCacheTupleCache {
    /// Cached tuple values (`BrtBeginPCDSDTCEntries` collection).
    pub entries: Vec<PivotCacheItemValue>,
    /// MDX queries (`irstQuery` of each `BrtBeginPCDSDTCQuery`), stored verbatim.
    pub queries: Vec<String>,
    /// Cached sets (`BrtBeginPCDSDTCSets` collection).
    pub sets: Vec<PivotCacheTupleCacheSet>,
}

/// A cached OLAP set (`BrtBeginPCDSDTCSet`, MS-XLSB 2.4.162).
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheTupleCacheSet {
    /// Number of tuples (`cTuples`); `None` when unknown (`0xFFFFFFFF`).
    pub tuple_count: Option<u32>,
    /// Maximum rank (`iRankMax`).
    pub max_rank: u32,
    /// Set sort order (`ssoType`, a `SdSetSortOrder` value).
    pub sort_order: u32,
    /// The set query failed (`fQueryFailed`).
    pub query_failed: bool,
    /// MDX set definition (`irstDef`), stored verbatim.
    pub definition: String,
}

/// Excel 2013 extension of the cache definition (`BrtBeginPCD14`, MS-XLSB 2.4.123).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PivotCacheDefinitionExt14 {
    /// Cache is used by slicer data (`fSlicerData`).
    pub slicer_data: bool,
    /// Server supports subquery for calculated members (`fSrvSupportSubQueryCalcMem`).
    pub server_support_subquery_calc_mem: bool,
    /// Server supports non-visual subquery (`fSrvSupportSubQueryNonVisual`).
    pub server_support_subquery_non_visual: bool,
    /// Server supports adding calculated members (`fSrvSupportAddCalcMems`).
    pub server_support_add_calc_mems: bool,
    /// Slicer cache identifier (`icacheId`; `0` = no slicer cache).
    pub cache_id: i32,
}
