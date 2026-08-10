//! Record-walking parser for the XLSB PivotCache definition stream
//! (MS-XLSB 2.1.7.38).
//!
//! The parser is strict about record payloads it fully understands and
//! tolerant about everything else: unknown record types are ignored, and
//! known begin/end record pairs that carry no modelled data (KPIs, FRT
//! wrappers, server-format collections, ...) are skipped as balanced
//! collections.

use crate::package::error::{Error, Result};
use crate::package::pivot::model::*;
use crate::package::walker::{RecordWalker, malformed};
use crate::raw::{Cursor, kind as rt};

// Record types used by this stream that have no constant in `records::kind` yet.
/// `BrtPCDField14` (MS-XLSB 2.4.725): marks the preceding cache field as ignorable.
const PCD_FIELD14: crate::raw::Kind = rt::PCD_FIELD14;
/// `BrtPCDH14` (MS-XLSB 2.4.726): named-set extension of a cache hierarchy.
const PCD_H14: crate::raw::Kind = rt::PCD_H14;

// `BrtBeginPivotCacheDef` flags byte 1 (MS-XLSB 2.4.168).
const DEF_SAVE_DATA: u8 = 1 << 0;
const DEF_INVALID: u8 = 1 << 1;
const DEF_REFRESH_ON_LOAD: u8 = 1 << 2;
const DEF_OPTIMIZE_CACHE: u8 = 1 << 3;
const DEF_ENABLE_REFRESH: u8 = 1 << 4;
const DEF_BACKGROUND_QUERY: u8 = 1 << 5;
const DEF_UPGRADE_ON_REFRESH: u8 = 1 << 6;
const DEF_CUBE_FUNCTIONS: u8 = 1 << 7;

// `BrtBeginPivotCacheDef` flags byte 2.
const DEF_LOAD_REFRESHED_WHO: u8 = 1 << 0;
const DEF_LOAD_REL_ID_RECORDS: u8 = 1 << 1;
const DEF_SUPPORT_SUBQUERY: u8 = 1 << 2;
const DEF_SUPPORT_ATTRIB_DRILL: u8 = 1 << 3;

// `BrtBeginPCDSRange` flags bytes (MS-XLSB 2.4.167).
const RANGE_NAME: u8 = 1 << 0;
const RANGE_BUILT_IN: u8 = 1 << 0;
const RANGE_LOAD_REL_ID: u8 = 1 << 0;
const RANGE_LOAD_SHEET: u8 = 1 << 1;

// `BrtBeginPCDSCSet` flags byte (MS-XLSB 2.4.154).
const CONSOL_SET_LOAD_REL_ID: u8 = 1 << 0;
const CONSOL_SET_LOAD_SHEET: u8 = 1 << 1;

// `BrtBeginPCDField` flags word (MS-XLSB 2.4.136).
const FIELD_SERVER_BASED: u16 = 1 << 0;
const FIELD_CANT_GET_UNIQUE_ITEMS: u16 = 1 << 1;
const FIELD_SRC_FIELD: u16 = 1 << 2;
const FIELD_CAPTION: u16 = 1 << 3;
const FIELD_OLAP_MEM_PROP: u16 = 1 << 4;
const FIELD_LOAD_FMLA: u16 = 1 << 8;
const FIELD_LOAD_PROP_NAME: u16 = 1 << 9;

/// `ifmt` value meaning the default number format (`PivotNumFmtExt`, MS-XLSB 2.5.107).
const DEFAULT_NUMBER_FORMAT: u32 = u32::MAX;

// `BrtBeginPCDFAtbl` flags word (MS-XLSB 2.4.131).
const ATBL_TEXT_FIELD: u16 = 1 << 0;
const ATBL_NON_DATES: u16 = 1 << 1;
const ATBL_DATE_IN_FIELD: u16 = 1 << 2;
const ATBL_HAS_TEXT_ITEM: u16 = 1 << 3;
const ATBL_HAS_BLANK_ITEM: u16 = 1 << 4;
const ATBL_MIXED_TYPES: u16 = 1 << 5;
const ATBL_NUM_FIELD: u16 = 1 << 6;
const ATBL_INT_FIELD: u16 = 1 << 7;
const ATBL_NUM_MIN_MAX_VALID: u16 = 1 << 8;
const ATBL_HAS_LONG_TEXT_ITEM: u16 = 1 << 9;

// `BrtBeginPCDFGRange` flags byte (MS-XLSB 2.4.134).
const GROUP_RANGE_AUTO_START: u8 = 1 << 0;
const GROUP_RANGE_AUTO_END: u8 = 1 << 1;
const GROUP_RANGE_DATES: u8 = 1 << 2;

// `BrtBeginPCDIRun` `mdSxoper` values (MS-XLSB 2.4.147).
const RUN_NUMBERS: u16 = 0x0001;
const RUN_STRINGS: u16 = 0x0002;
const RUN_ERRORS: u16 = 0x0010;
const RUN_DATETIMES: u16 = 0x0020;

// `PCDIAddlInfo` flags word (MS-XLSB 2.5.100).
const ADDL_GHOST: u16 = 1 << 0;
const ADDL_CALCULATED: u16 = 1 << 1;
const ADDL_CAPTION: u16 = 1 << 2;

// `BrtBeginPCDHierarchy` flags word 1 (MS-XLSB 2.4.146).
const HIER_MEASURE: u16 = 1 << 0;
const HIER_SET: u16 = 1 << 1;
const HIER_ATTRIBUTE: u16 = 1 << 2;
const HIER_MEASURE_HIERARCHY: u16 = 1 << 3;
const HIER_ONLY_ONE_FIELD: u16 = 1 << 4;
const HIER_TIME: u16 = 1 << 5;
const HIER_KEY_ATTRIBUTE: u16 = 1 << 6;
const HIER_VALUE_TYPE_KNOWN: u16 = 1 << 7;
const HIER_UNBALANCED_REAL_KNOWN: u16 = 1 << 8;
const HIER_UNBALANCED_REAL: u16 = 1 << 9;
const HIER_UNBALANCED_GROUP_KNOWN: u16 = 1 << 10;
const HIER_UNBALANCED_GROUP: u16 = 1 << 11;
const HIER_HIDDEN: u16 = 1 << 12;

// `BrtBeginPCDHierarchy` flags byte 2.
const HIER_LOAD_DIM_UNQ: u8 = 1 << 0;
const HIER_LOAD_DEFAULT_UNQ: u8 = 1 << 1;
const HIER_LOAD_ALL_UNQ: u8 = 1 << 2;
const HIER_LOAD_ALL_DISP: u8 = 1 << 3;
const HIER_LOAD_DISP_FLD: u8 = 1 << 4;
const HIER_LOAD_MEAS_GRP: u8 = 1 << 5;

// `BrtBeginPCDHGLevel` flags byte (MS-XLSB 2.4.139).
const GROUPING_LEVEL_GROUP: u8 = 1 << 0;
// `BrtBeginPCDHGLGroup` flags byte (MS-XLSB 2.4.143).
const GROUPING_GROUP_LOAD_PARENT: u8 = 1 << 0;

// `BrtPCDH14` flags byte (MS-XLSB 2.4.726).
const HIER14_FLATTEN: u8 = 1 << 0;
const HIER14_MEASURE_SET: u8 = 1 << 1;
const HIER14_HIERARCHIZE_DISTINCT: u8 = 1 << 2;
const HIER14_IGNORABLE: u8 = 1 << 3;

// `PCDCalcMemCommon` flags word (MS-XLSB 2.5.99).
const CALC_MEM_LOAD_MEMBER_NAME: u32 = 1 << 0;
const CALC_MEM_LOAD_SOURCE_HIER: u32 = 1 << 1;
const CALC_MEM_LOAD_PARENT_UNIQUE: u32 = 1 << 2;

// `BrtBeginPCDCalcMem14` flags byte (MS-XLSB 2.4.127).
const CALC_MEM14_FLATTEN: u8 = 1 << 0;
const CALC_MEM14_DYNAMIC_SET: u8 = 1 << 1;
const CALC_MEM14_HIERARCHIZE_DISTINCT: u8 = 1 << 2;

// `BrtBeginPName` flags byte (MS-XLSB 2.4.176).
const PNAME_ERR_NAME: u8 = 1 << 0;
// `BrtBeginPNPair` flags byte (MS-XLSB 2.4.178).
const PNPAIR_PHYSICAL: u8 = 1 << 0;
const PNPAIR_RELATIVE: u8 = 1 << 1;

// `PRFilter` 24-bit flags (MS-XLSB 2.4.180).
const PR_FILTER_ITEM_TYPES_MASK: u32 = 0x1FFF;
const PR_FILTER_SELECTED: u32 = 1 << 16;

// `BrtBeginPCDSDTCSet` flags byte (MS-XLSB 2.4.162).
const TUPLE_SET_QUERY_FAILED: u8 = 1 << 0;
/// `cTuples` value meaning the tuple count is unknown (MS-XLSB 2.4.162).
const TUPLE_COUNT_UNKNOWN: u32 = u32::MAX;

// `BrtBeginPCD14` flags byte (MS-XLSB 2.4.123).
const PCD14_SLICER_DATA: u8 = 1 << 0;
const PCD14_SUBQUERY_CALC_MEM: u8 = 1 << 1;
const PCD14_SUBQUERY_NON_VISUAL: u8 = 1 << 2;
const PCD14_ADD_CALC_MEMS: u8 = 1 << 3;

/// `BrtBeginPCDSConsol` flags word (MS-XLSB 2.4.150).
const CONSOL_AUTO_PAGE: u16 = 1 << 0;

/// Parse a PivotCache definition part (`pivotCacheDefinition*.bin`) into a
/// typed [`PivotCacheDefinition`].
///
/// The stream must start with `BrtBeginPivotCacheDef`. Records after
/// `BrtEndPivotCacheDef` are ignored. Unknown record types anywhere in the
/// stream are skipped without failing.
pub(super) fn parse_pivot_cache_definition_binary(data: &[u8]) -> Result<PivotCacheDefinition> {
    let mut walker = RecordWalker::new(data);
    let first = walker.required_begin(rt::BEGIN_PIVOT_CACHE_DEF, "BrtBeginPivotCacheDef")?;
    let mut definition = parse_definition_payload(first.payload())?;
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PIVOT_CACHE_DEF => return Ok(definition),
            rt::BEGIN_PCD_SOURCE => {
                definition.source = Some(parse_source(&mut walker, record.payload())?);
            },
            rt::BEGIN_PCD_FIELDS => {
                parse_fields(&mut walker, record.payload(), &mut definition)?;
            },
            rt::BEGIN_PCD_HIERARCHIES => {
                parse_hierarchies(&mut walker, record.payload(), &mut definition)?;
            },
            rt::BEGIN_PCDSD_TUPLE_CACHE => {
                definition.tuple_cache = Some(parse_tuple_cache(&mut walker)?);
            },
            rt::BEGIN_PCD_CALC_ITEMS => {
                parse_calculated_items(&mut walker, record.payload(), &mut definition)?;
            },
            rt::BEGIN_PCD_CALC_MEMS => {
                parse_calculated_members(&mut walker, record.payload(), &mut definition)?;
            },
            rt::BEGIN_PCD14 => {
                definition.ext14 = Some(parse_pcd14(record.payload())?);
                walker.expect_end(rt::END_PCD14, "BrtBeginPCD14")?;
            },
            other => walker.skip_unhandled(other, "PivotCache definition stream")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPivotCacheDef".to_string(),
    ))
}

trait CursorExt {
    fn read_range(&mut self) -> Result<PivotCacheRange>;
}

impl CursorExt for Cursor<'_> {
    /// Read a `PivotCacheRange` (four signed 32-bit bounds).
    fn read_range(&mut self) -> Result<PivotCacheRange> {
        let first_row = self.read_i32()?;
        let last_row = self.read_i32()?;
        let first_column = self.read_i32()?;
        let last_column = self.read_i32()?;
        Ok(PivotCacheRange {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }
}

/// Read the count carried by a PivotCache collection begin record.
fn parse_collection_count(data: &[u8], context: &'static str) -> Result<u32> {
    let mut cursor = Cursor::new(data, context);
    let count = cursor.read_u32()?;
    cursor.finish()?;
    Ok(count)
}

/// Check a collection's declared count after its children have been parsed.
fn validate_collection_count(declared: u32, actual: usize, context: &'static str) -> Result<()> {
    if u64::from(declared) != actual as u64 {
        return Err(malformed(
            context,
            format!("declared {declared} records, found {actual}"),
        ));
    }
    Ok(())
}

/// `BrtBeginPivotCacheDef` payload (MS-XLSB 2.4.168).
fn parse_definition_payload(data: &[u8]) -> Result<PivotCacheDefinition> {
    let mut cursor = Cursor::new(data, "BrtBeginPivotCacheDef");
    let version_last_refresh = cursor.read_u8()?;
    let version_refreshable_min = cursor.read_u8()?;
    let version_created = cursor.read_u8()?;
    let flags1 = cursor.read_u8()?;
    let ghost_items_max = cursor.read_i32()?;
    let refreshed_date_serial = cursor.read_f64()?;
    let flags2 = cursor.read_u8()?;
    let record_count = cursor.read_u32()?;
    let refreshed_by = if flags2 & DEF_LOAD_REFRESHED_WHO != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let records_rel_id = if flags2 & DEF_LOAD_REL_ID_RECORDS != 0 {
        cursor.read_nullable_wide_string()?
    } else {
        None
    };
    if flags2 & DEF_LOAD_REFRESHED_WHO == 0 {
        // `unused` (4 bytes) exists iff fLoadRefreshedWho is 0.
        cursor.guard(4)?;
        cursor.skip(4)?;
    }
    cursor.finish()?;
    Ok(PivotCacheDefinition {
        version_last_refresh,
        version_refreshable_min,
        version_created,
        save_data: flags1 & DEF_SAVE_DATA != 0,
        invalid: flags1 & DEF_INVALID != 0,
        refresh_on_load: flags1 & DEF_REFRESH_ON_LOAD != 0,
        optimize_cache: flags1 & DEF_OPTIMIZE_CACHE != 0,
        enable_refresh: flags1 & DEF_ENABLE_REFRESH != 0,
        background_query: flags1 & DEF_BACKGROUND_QUERY != 0,
        upgrade_on_refresh: flags1 & DEF_UPGRADE_ON_REFRESH != 0,
        cube_functions: flags1 & DEF_CUBE_FUNCTIONS != 0,
        support_subquery: flags2 & DEF_SUPPORT_SUBQUERY != 0,
        support_attrib_drill: flags2 & DEF_SUPPORT_ATTRIB_DRILL != 0,
        ghost_items_max,
        refreshed_date_serial,
        record_count,
        refreshed_by,
        records_rel_id,
        ..PivotCacheDefinition::default()
    })
}

/// `BrtBeginPCDSource` collection (MS-XLSB 2.4.166).
fn parse_source(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<PivotCacheSource> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDSource");
    let source_type = PivotCacheSourceType::try_from(cursor.read_u32()?)?;
    let connection_id = cursor.read_u32()?;
    cursor.finish()?;
    let mut source = PivotCacheSource {
        source_type,
        connection_id: (source_type == PivotCacheSourceType::External).then_some(connection_id),
        worksheet: None,
        consolidation: None,
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_SOURCE => return Ok(source),
            rt::BEGIN_PCDS_RANGE => {
                source.worksheet = Some(parse_worksheet_range(record.payload())?);
                walker.expect_end(rt::END_PCDS_RANGE, "BrtBeginPCDSRange")?;
            },
            rt::BEGIN_PCDS_CONSOL => {
                source.consolidation = Some(parse_consolidation(walker, record.payload())?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSource collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDSource".to_string()))
}

/// `BrtBeginPCDSRange` payload (MS-XLSB 2.4.167).
fn parse_worksheet_range(data: &[u8]) -> Result<PivotCacheWorksheetSource> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDSRange");
    let flags0 = cursor.read_u8()?;
    let flags1 = cursor.read_u8()?;
    let flags2 = cursor.read_u8()?;
    let by_name = flags0 & RANGE_NAME != 0;
    let sheet_name = if flags2 & RANGE_LOAD_SHEET != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let external_rel_id = if flags2 & RANGE_LOAD_REL_ID != 0 {
        cursor.read_nullable_wide_string()?
    } else {
        None
    };
    let (named_range, range) = if by_name {
        (Some(cursor.read_wide_string()?), None)
    } else {
        (None, Some(cursor.read_range()?))
    };
    cursor.finish()?;
    Ok(PivotCacheWorksheetSource {
        named_range,
        built_in_name: flags1 & RANGE_BUILT_IN != 0,
        sheet_name,
        external_rel_id,
        range,
    })
}

/// `BrtBeginPCDSConsol` collection (MS-XLSB 2.4.150).
fn parse_consolidation(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
) -> Result<PivotCacheConsolidationSource> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDSConsol");
    let flags = cursor.read_u16()?;
    cursor.finish()?;
    let mut consolidation = PivotCacheConsolidationSource {
        auto_page: flags & CONSOL_AUTO_PAGE != 0,
        sets: Vec::new(),
        pages: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDS_CONSOL => return Ok(consolidation),
            rt::BEGIN_PCDSC_SETS => {
                parse_consolidation_sets(walker, record.payload(), &mut consolidation)?
            },
            rt::BEGIN_PCDSC_PAGES => {
                parse_consolidation_pages(walker, record.payload(), &mut consolidation)?
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSConsol collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDSConsol".to_string()))
}

/// `BrtBeginPCDSCSets` collection (MS-XLSB 2.4.155).
fn parse_consolidation_sets(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    consolidation: &mut PivotCacheConsolidationSource,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDSCSets")?;
    let first_set = consolidation.sets.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDSC_SETS => {
                return validate_collection_count(
                    declared,
                    consolidation.sets.len() - first_set,
                    "BrtBeginPCDSCSets",
                );
            },
            rt::BEGIN_PCDSC_SET => {
                consolidation
                    .sets
                    .push(parse_consolidation_set(record.payload())?);
                walker.expect_end(rt::END_PCDSC_SET, "BrtBeginPCDSCSet")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSCSets collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDSCSets".to_string()))
}

/// `BrtBeginPCDSCSet` payload (MS-XLSB 2.4.154).
fn parse_consolidation_set(data: &[u8]) -> Result<PivotCacheConsolidationSet> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDSCSet");
    let item_indexes = [
        cursor.read_u32()?,
        cursor.read_u32()?,
        cursor.read_u32()?,
        cursor.read_u32()?,
    ];
    let by_name = cursor.read_bool8()?;
    let built_in_name = cursor.read_bool8()?;
    let flags = cursor.read_u8()?;
    let sheet_name = if flags & CONSOL_SET_LOAD_SHEET != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let external_rel_id = if flags & CONSOL_SET_LOAD_REL_ID != 0 {
        cursor.read_nullable_wide_string()?
    } else {
        None
    };
    let (range, named_range) = if by_name {
        (None, Some(cursor.read_wide_string()?))
    } else {
        (Some(cursor.read_range()?), None)
    };
    cursor.finish()?;
    Ok(PivotCacheConsolidationSet {
        item_indexes,
        named_range,
        built_in_name,
        sheet_name,
        external_rel_id,
        range,
    })
}

/// `BrtBeginPCDSCPages` collection (MS-XLSB 2.4.152).
fn parse_consolidation_pages(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    consolidation: &mut PivotCacheConsolidationSource,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDSCPages")?;
    let first_page = consolidation.pages.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDSC_PAGES => {
                return validate_collection_count(
                    declared,
                    consolidation.pages.len() - first_page,
                    "BrtBeginPCDSCPages",
                );
            },
            rt::BEGIN_PCDSC_PAGE => {
                let page = parse_consolidation_page(walker, record.payload())?;
                consolidation.pages.push(page);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSCPages collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDSCPages".to_string()))
}

/// `BrtBeginPCDSCPage` collection (MS-XLSB 2.4.151).
fn parse_consolidation_page(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
) -> Result<PivotCacheConsolidationPage> {
    let declared = parse_collection_count(data, "BrtBeginPCDSCPage")?;
    let mut page = PivotCacheConsolidationPage {
        item_names: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDSC_PAGE => {
                validate_collection_count(declared, page.item_names.len(), "BrtBeginPCDSCPage")?;
                return Ok(page);
            },
            rt::BEGIN_PCDSCP_ITEM => {
                let mut cursor = Cursor::new(record.payload(), "BrtBeginPCDSCPItem");
                page.item_names.push(cursor.read_wide_string()?);
                cursor.finish()?;
                walker.expect_end(rt::END_PCDSCP_ITEM, "BrtBeginPCDSCPItem")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSCPage collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDSCPage".to_string()))
}

/// `BrtBeginPCDFields` collection (MS-XLSB 2.4.137).
fn parse_fields(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    definition: &mut PivotCacheDefinition,
) -> Result<()> {
    let declared_fields = parse_collection_count(data, "BrtBeginPCDFields")?;
    let first_field = definition.fields.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_FIELDS => {
                validate_collection_count(
                    declared_fields,
                    definition.fields.len() - first_field,
                    "BrtBeginPCDFields",
                )?;
                return Ok(());
            },
            rt::BEGIN_PCD_FIELD => {
                let field = parse_field(walker, record.payload())?;
                definition.fields.push(field);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFields collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDFields".to_string()))
}

/// `BrtBeginPCDField` collection (MS-XLSB 2.4.136).
fn parse_field(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<PivotCacheField> {
    let mut field = parse_field_payload(data)?;
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_FIELD => return Ok(field),
            rt::BEGIN_PCDF_ATBL => {
                parse_shared_items(walker, record.payload(), &mut field.shared_items)?
            },
            rt::BEGIN_PCDF_GROUP => {
                field.grouping = Some(parse_grouping(walker, record.payload())?);
            },
            PCD_FIELD14 => field.ignore = true,
            other => walker.skip_unhandled(other, "BrtBeginPCDField collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDField".to_string()))
}

/// `BrtBeginPCDField` payload (MS-XLSB 2.4.136).
fn parse_field_payload(data: &[u8]) -> Result<PivotCacheField> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDField");
    let flags = cursor.read_u16()?;
    let number_format = match cursor.read_u32()? {
        DEFAULT_NUMBER_FORMAT => None,
        format => Some(format),
    };
    let sql_type = cursor.read_u16()?;
    let hierarchy_index = cursor.read_u32()?;
    let level = cursor.read_u32()?;
    let member_property_count = cursor.read_u32()?;
    let name = cursor.read_wide_string()?;
    let caption = if flags & FIELD_CAPTION != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let formula = if flags & FIELD_LOAD_FMLA != 0 {
        Some(parse_pivot_formula(&mut cursor)?)
    } else {
        None
    };
    let mut member_property_fields = Vec::new();
    if member_property_count > 0 {
        let byte_count = cursor.read_u32()?;
        if byte_count != member_property_count.saturating_mul(4) {
            return Err(malformed(
                "BrtBeginPCDField",
                "cbRgisxtmp disagrees with cIsxtmps",
            ));
        }
        for _ in 0..member_property_count {
            member_property_fields.push(cursor.read_u32()?);
        }
    }
    let member_property_name = if flags & FIELD_LOAD_PROP_NAME != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    cursor.finish()?;
    Ok(PivotCacheField {
        name,
        caption,
        number_format,
        sql_type,
        hierarchy_index,
        level,
        member_property_fields,
        member_property_name,
        formula,
        server_based: flags & FIELD_SERVER_BASED != 0,
        cant_get_unique_items: flags & FIELD_CANT_GET_UNIQUE_ITEMS != 0,
        source_field: flags & FIELD_SRC_FIELD != 0,
        olap_member_property_field: flags & FIELD_OLAP_MEM_PROP != 0,
        ignore: false,
        shared_items: PivotCacheSharedItems::default(),
        grouping: None,
    })
}

/// `PivotParsedFormula` (MS-XLSB 2.5.98.15), stored verbatim.
fn parse_pivot_formula(cursor: &mut Cursor<'_>) -> Result<PivotParsedFormulaData> {
    let tokens = cursor.read_blob()?;
    let extra = cursor.read_blob()?;
    Ok(PivotParsedFormulaData {
        tokens: tokens.to_vec(),
        extra: extra.to_vec(),
    })
}

/// `BrtBeginPCDFAtbl` collection (MS-XLSB 2.4.131): shared item statistics
/// followed by the raw cache items.
fn parse_shared_items(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    shared_items: &mut PivotCacheSharedItems,
) -> Result<()> {
    shared_items.stats = Some(parse_shared_items_stats(data)?);
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDF_ATBL => return Ok(()),
            rt::BEGIN_PCDI_RUN => {
                parse_item_run(record.payload(), &mut shared_items.items)?;
                walker.expect_end(rt::END_PCDI_RUN, "BrtBeginPCDIRun")?;
            },
            item_type @ (rt::PCDI_MISSING
            | rt::PCDI_NUMBER
            | rt::PCDI_BOOLEAN
            | rt::PCDI_ERROR
            | rt::PCDI_STRING
            | rt::PCDI_DATETIME
            | rt::PCDIA_MISSING
            | rt::PCDIA_NUMBER
            | rt::PCDIA_BOOLEAN
            | rt::PCDIA_ERROR
            | rt::PCDIA_STRING
            | rt::PCDIA_DATETIME) => {
                let item = parse_cache_item(item_type, record.payload(), true)?;
                shared_items.items.push(item);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFAtbl collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDFAtbl".to_string()))
}

/// `BrtBeginPCDFAtbl` payload (MS-XLSB 2.4.131).
fn parse_shared_items_stats(data: &[u8]) -> Result<PivotCacheSharedItemsStats> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDFAtbl");
    let flags = cursor.read_u16()?;
    let item_count = cursor.read_u32()?;
    let (minimum, maximum) = if flags & ATBL_NUM_MIN_MAX_VALID != 0 {
        (Some(cursor.read_f64()?), Some(cursor.read_f64()?))
    } else {
        (None, None)
    };
    cursor.finish()?;
    Ok(PivotCacheSharedItemsStats {
        text_field: flags & ATBL_TEXT_FIELD != 0,
        non_dates: flags & ATBL_NON_DATES != 0,
        date_in_field: flags & ATBL_DATE_IN_FIELD != 0,
        has_text_item: flags & ATBL_HAS_TEXT_ITEM != 0,
        has_blank_item: flags & ATBL_HAS_BLANK_ITEM != 0,
        mixed_types_ignoring_blanks: flags & ATBL_MIXED_TYPES != 0,
        numeric_field: flags & ATBL_NUM_FIELD != 0,
        integer_field: flags & ATBL_INT_FIELD != 0,
        has_long_text_item: flags & ATBL_HAS_LONG_TEXT_ITEM != 0,
        item_count,
        minimum,
        maximum,
    })
}

/// Parse one `BrtPCDI*`/`BrtPCDIA*` record payload into a cache item.
///
/// When `strict` is set the payload must be fully consumed; tuple-cache
/// entries may legally trail a `PCDISrvFmt` (`sxvcellextra`) and are parsed
/// non-strictly.
fn parse_cache_item(
    record_type: crate::raw::Kind,
    data: &[u8],
    strict: bool,
) -> Result<PivotCacheItem> {
    let mut cursor = Cursor::new(data, "BrtPCDI cache item");
    let value = match record_type {
        rt::PCDI_MISSING | rt::PCDIA_MISSING => PivotCacheItemValue::Missing,
        rt::PCDI_NUMBER | rt::PCDIA_NUMBER => PivotCacheItemValue::Number(cursor.read_f64()?),
        rt::PCDI_BOOLEAN | rt::PCDIA_BOOLEAN => PivotCacheItemValue::Boolean(cursor.read_bool8()?),
        rt::PCDI_ERROR | rt::PCDIA_ERROR => {
            PivotCacheItemValue::Error(PivotCacheErrorCode::try_from(cursor.read_u8()?)?)
        },
        rt::PCDI_STRING | rt::PCDIA_STRING => {
            PivotCacheItemValue::String(cursor.read_wide_string()?.into_boxed_str())
        },
        rt::PCDI_DATETIME | rt::PCDIA_DATETIME => {
            PivotCacheItemValue::DateTime(read_date_time(&mut cursor)?)
        },
        rt::PCDI_INDEX => PivotCacheItemValue::Index(cursor.read_u32()?),
        _ => {
            return Err(Error::UnexpectedRecord {
                expected: rt::PCDI_MISSING.get(),
                found: record_type.get(),
            });
        },
    };
    let additional = if matches!(
        record_type,
        rt::PCDIA_MISSING
            | rt::PCDIA_NUMBER
            | rt::PCDIA_BOOLEAN
            | rt::PCDIA_ERROR
            | rt::PCDIA_STRING
            | rt::PCDIA_DATETIME
    ) {
        Some(parse_item_info(&mut cursor)?)
    } else {
        None
    };
    if strict {
        cursor.finish()?;
    }
    Ok(PivotCacheItem { value, additional })
}

/// `PCDIDateTime` (MS-XLSB 2.5.101).
pub(super) fn read_date_time(cursor: &mut Cursor<'_>) -> Result<PivotCacheDateTime> {
    Ok(PivotCacheDateTime {
        year: cursor.read_u16()?,
        month: cursor.read_u16()?,
        day: cursor.read_u8()?,
        hour: cursor.read_u8()?,
        minute: cursor.read_u8()?,
        second: cursor.read_u8()?,
    })
}

/// `PCDIAddlInfo` (MS-XLSB 2.5.100).
fn parse_item_info(cursor: &mut Cursor<'_>) -> Result<PivotCacheItemInfo> {
    let flags = cursor.read_u16()?;
    let caption = if flags & ADDL_CAPTION != 0 {
        cursor.read_nullable_wide_string()?
    } else {
        None
    };
    let member_property_count = cursor.read_u32()?;
    let mut member_property_items = Vec::new();
    for _ in 0..member_property_count {
        member_property_items.push(cursor.read_i32()?);
    }
    Ok(PivotCacheItemInfo {
        ghost: flags & ADDL_GHOST != 0,
        calculated: flags & ADDL_CALCULATED != 0,
        caption,
        member_property_items,
    })
}

/// `BrtBeginPCDIRun` payload (MS-XLSB 2.4.147): a compact run of same-typed
/// cache items.
fn parse_item_run(data: &[u8], items: &mut Vec<PivotCacheItem>) -> Result<()> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDIRun");
    let operation = cursor.read_u16()?;
    let count = cursor.read_u32()?;
    for _ in 0..count {
        let value = match operation {
            RUN_NUMBERS => PivotCacheItemValue::Number(cursor.read_f64()?),
            RUN_STRINGS => PivotCacheItemValue::String(cursor.read_wide_string()?.into_boxed_str()),
            RUN_ERRORS => {
                PivotCacheItemValue::Error(PivotCacheErrorCode::try_from(cursor.read_u8()?)?)
            },
            RUN_DATETIMES => PivotCacheItemValue::DateTime(read_date_time(&mut cursor)?),
            _ => {
                return Err(malformed(
                    "BrtBeginPCDIRun",
                    format!("unknown mdSxoper 0x{operation:04X}"),
                ));
            },
        };
        items.push(PivotCacheItem {
            value,
            additional: None,
        });
    }
    Ok(cursor.finish()?)
}

/// `BrtBeginPCDFGroup` collection (MS-XLSB 2.4.135).
fn parse_grouping(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<PivotCacheFieldGrouping> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDFGroup");
    let parent_field = non_negative_index(cursor.read_i32()?);
    let base_field = non_negative_index(cursor.read_i32()?);
    cursor.finish()?;
    let mut grouping = PivotCacheFieldGrouping {
        parent_field,
        base_field,
        range: None,
        discrete: None,
        items: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDF_GROUP => return Ok(grouping),
            rt::BEGIN_PCDFG_RANGE => {
                grouping.range = Some(parse_range_grouping(record.payload())?);
                walker.expect_end(rt::END_PCDFG_RANGE, "BrtBeginPCDFGRange")?;
            },
            rt::BEGIN_PCDFG_DISCRETE => {
                grouping.discrete = Some(parse_discrete_grouping(walker, record.payload())?);
            },
            rt::BEGIN_PCDFG_ITEMS => {
                parse_grouping_items(walker, record.payload(), &mut grouping.items)?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFGroup collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDFGroup".to_string()))
}

/// Interpret a spec `-1`-or-index field as an optional unsigned index.
fn non_negative_index(value: i32) -> Option<u32> {
    u32::try_from(value).ok()
}

/// `BrtBeginPCDFGRange` payload (MS-XLSB 2.4.134).
fn parse_range_grouping(data: &[u8]) -> Result<PivotCacheRangeGrouping> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDFGRange");
    let group_by = PivotCacheGroupBy::try_from(cursor.read_u8()?)?;
    let flags = cursor.read_u8()?;
    let start = cursor.read_f64()?;
    let end = cursor.read_f64()?;
    let interval = cursor.read_f64()?;
    cursor.finish()?;
    Ok(PivotCacheRangeGrouping {
        group_by,
        auto_start: flags & GROUP_RANGE_AUTO_START != 0,
        auto_end: flags & GROUP_RANGE_AUTO_END != 0,
        dates: flags & GROUP_RANGE_DATES != 0,
        start,
        end,
        interval,
    })
}

/// `BrtBeginPCDFGDiscrete` collection (MS-XLSB 2.4.132).
fn parse_discrete_grouping(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
) -> Result<PivotCacheDiscreteGrouping> {
    let declared = parse_collection_count(data, "BrtBeginPCDFGDiscrete")?;
    let mut grouping = PivotCacheDiscreteGrouping {
        item_indexes: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDFG_DISCRETE => {
                validate_collection_count(
                    declared,
                    grouping.item_indexes.len(),
                    "BrtBeginPCDFGDiscrete",
                )?;
                return Ok(grouping);
            },
            rt::PCDI_INDEX => {
                let mut cursor = Cursor::new(record.payload(), "BrtPCDIIndex");
                grouping.item_indexes.push(cursor.read_u32()?);
                cursor.finish()?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFGDiscrete collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDFGDiscrete".to_string(),
    ))
}

/// `BrtBeginPCDFGItems` collection (MS-XLSB 2.4.133): grouping cache items.
fn parse_grouping_items(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    items: &mut Vec<PivotCacheItem>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDFGItems")?;
    let first_item = items.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDFG_ITEMS => {
                validate_collection_count(
                    declared,
                    items.len() - first_item,
                    "BrtBeginPCDFGItems",
                )?;
                return Ok(());
            },
            rt::BEGIN_PCDI_RUN => {
                parse_item_run(record.payload(), items)?;
                walker.expect_end(rt::END_PCDI_RUN, "BrtBeginPCDIRun")?;
            },
            item_type @ (rt::PCDI_MISSING
            | rt::PCDI_NUMBER
            | rt::PCDI_BOOLEAN
            | rt::PCDI_ERROR
            | rt::PCDI_STRING
            | rt::PCDI_DATETIME
            | rt::PCDIA_MISSING
            | rt::PCDIA_NUMBER
            | rt::PCDIA_BOOLEAN
            | rt::PCDIA_ERROR
            | rt::PCDIA_STRING
            | rt::PCDIA_DATETIME) => {
                items.push(parse_cache_item(item_type, record.payload(), true)?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFGItems collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDFGItems".to_string()))
}

/// `BrtBeginPCDHierarchies` collection (MS-XLSB 2.4.145).
fn parse_hierarchies(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    definition: &mut PivotCacheDefinition,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDHierarchies")?;
    let first_hierarchy = definition.hierarchies.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_HIERARCHIES => {
                return validate_collection_count(
                    declared,
                    definition.hierarchies.len() - first_hierarchy,
                    "BrtBeginPCDHierarchies",
                );
            },
            rt::BEGIN_PCD_HIERARCHY => {
                let hierarchy = parse_hierarchy(walker, record.payload())?;
                definition.hierarchies.push(hierarchy);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHierarchies collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDHierarchies".to_string(),
    ))
}

/// `BrtBeginPCDHierarchy` collection (MS-XLSB 2.4.146).
fn parse_hierarchy(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<PivotCacheHierarchy> {
    let mut hierarchy = parse_hierarchy_payload(data)?;
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_HIERARCHY => return Ok(hierarchy),
            rt::BEGIN_PCDH_FIELDS_USAGE => {
                hierarchy.field_usage = parse_fields_usage(record.payload())?;
                walker.expect_end(rt::END_PCDH_FIELDS_USAGE, "BrtBeginPCDHFieldsUsage")?;
            },
            rt::BEGIN_PCDHG_LEVELS => {
                parse_grouping_levels(walker, record.payload(), &mut hierarchy.grouping_levels)?;
            },
            rt::BEGIN_PCDHGL_GROUPS => {
                parse_grouping_groups(walker, record.payload(), &mut hierarchy.grouping_groups)?;
            },
            PCD_H14 => {
                hierarchy.ext14 = Some(parse_hierarchy_ext14(record.payload())?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHierarchy collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDHierarchy".to_string(),
    ))
}

/// `BrtBeginPCDHierarchy` payload (MS-XLSB 2.4.146).
fn parse_hierarchy_payload(data: &[u8]) -> Result<PivotCacheHierarchy> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDHierarchy");
    let flags1 = cursor.read_u16()?;
    let level_count = cursor.read_u32()?;
    let set_parent_index = non_negative_index(cursor.read_i32()?);
    let icon_set = cursor.read_i32()?;
    let flags2 = cursor.read_u8()?;
    let attribute_member_value_type = cursor.read_u16()?;
    let unique_name = cursor.read_wide_string()?;
    let caption = cursor.read_wide_string()?;
    let dimension_unique_name = conditional_string(&mut cursor, flags2 & HIER_LOAD_DIM_UNQ != 0)?;
    let default_member_unique_name =
        conditional_string(&mut cursor, flags2 & HIER_LOAD_DEFAULT_UNQ != 0)?;
    let all_member_unique_name = conditional_string(&mut cursor, flags2 & HIER_LOAD_ALL_UNQ != 0)?;
    let all_member_display = conditional_string(&mut cursor, flags2 & HIER_LOAD_ALL_DISP != 0)?;
    let display_folder = conditional_string(&mut cursor, flags2 & HIER_LOAD_DISP_FLD != 0)?;
    let measure_group = conditional_string(&mut cursor, flags2 & HIER_LOAD_MEAS_GRP != 0)?;
    cursor.finish()?;
    Ok(PivotCacheHierarchy {
        unique_name,
        caption,
        dimension_unique_name,
        default_member_unique_name,
        all_member_unique_name,
        all_member_display,
        display_folder,
        measure_group,
        measure: flags1 & HIER_MEASURE != 0,
        set: flags1 & HIER_SET != 0,
        attribute_hierarchy: flags1 & HIER_ATTRIBUTE != 0,
        measure_hierarchy: flags1 & HIER_MEASURE_HIERARCHY != 0,
        only_one_field: flags1 & HIER_ONLY_ONE_FIELD != 0,
        time_hierarchy: flags1 & HIER_TIME != 0,
        key_attribute_hierarchy: flags1 & HIER_KEY_ATTRIBUTE != 0,
        hidden: flags1 & HIER_HIDDEN != 0,
        unbalanced_real: (flags1 & HIER_UNBALANCED_REAL_KNOWN != 0)
            .then_some(flags1 & HIER_UNBALANCED_REAL != 0),
        unbalanced_group: (flags1 & HIER_UNBALANCED_GROUP_KNOWN != 0)
            .then_some(flags1 & HIER_UNBALANCED_GROUP != 0),
        attribute_member_value_type: (flags1 & HIER_VALUE_TYPE_KNOWN != 0)
            .then_some(attribute_member_value_type),
        level_count,
        set_parent_index,
        icon_set,
        field_usage: Vec::new(),
        grouping_levels: Vec::new(),
        grouping_groups: Vec::new(),
        ext14: None,
    })
}

fn conditional_string(cursor: &mut Cursor<'_>, present: bool) -> Result<Option<String>> {
    if present {
        Ok(Some(cursor.read_wide_string()?))
    } else {
        Ok(None)
    }
}

/// `BrtBeginPCDHFieldsUsage` payload (MS-XLSB 2.4.138).
fn parse_fields_usage(data: &[u8]) -> Result<Vec<i32>> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDHFieldsUsage");
    let count = cursor.read_u32()?;
    let mut usage = Vec::new();
    for _ in 0..count {
        usage.push(cursor.read_i32()?);
    }
    cursor.finish()?;
    Ok(usage)
}

/// `BrtBeginPCDHGLevels` collection (MS-XLSB 2.4.140).
fn parse_grouping_levels(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    levels: &mut Vec<PivotCacheGroupingLevel>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDHGLevels")?;
    let first_level = levels.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDHG_LEVELS => {
                return validate_collection_count(
                    declared,
                    levels.len() - first_level,
                    "BrtBeginPCDHGLevels",
                );
            },
            rt::BEGIN_PCDHG_LEVEL => {
                let mut cursor = Cursor::new(record.payload(), "BrtBeginPCDHGLevel");
                let flags = cursor.read_u8()?;
                let unique_name = cursor.read_wide_string()?;
                let caption = cursor.read_wide_string()?;
                cursor.finish()?;
                levels.push(PivotCacheGroupingLevel {
                    group_level: flags & GROUPING_LEVEL_GROUP != 0,
                    unique_name,
                    caption,
                });
                walker.expect_end(rt::END_PCDHG_LEVEL, "BrtBeginPCDHGLevel")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHGLevels collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDHGLevels".to_string(),
    ))
}

/// `BrtBeginPCDHGLGroups` collection (MS-XLSB 2.4.144).
fn parse_grouping_groups(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    groups: &mut Vec<PivotCacheGroupingGroup>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDHGLGroups")?;
    let first_group = groups.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDHGL_GROUPS => {
                return validate_collection_count(
                    declared,
                    groups.len() - first_group,
                    "BrtBeginPCDHGLGroups",
                );
            },
            rt::BEGIN_PCDHGL_GROUP => {
                let group = parse_grouping_group(walker, record.payload())?;
                groups.push(group);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHGLGroups collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDHGLGroups".to_string(),
    ))
}

/// `BrtBeginPCDHGLGroup` collection (MS-XLSB 2.4.143).
fn parse_grouping_group(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
) -> Result<PivotCacheGroupingGroup> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDHGLGroup");
    let group_number = cursor.read_i32()?;
    let flags = cursor.read_u8()?;
    let name = cursor.read_wide_string()?;
    let unique_name = cursor.read_wide_string()?;
    let caption = cursor.read_wide_string()?;
    let parent_unique_name =
        conditional_string(&mut cursor, flags & GROUPING_GROUP_LOAD_PARENT != 0)?;
    cursor.finish()?;
    let mut group = PivotCacheGroupingGroup {
        group_number,
        name,
        unique_name,
        caption,
        parent_unique_name,
        members: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDHGL_GROUP => return Ok(group),
            rt::BEGIN_PCDHGLG_MEMBERS => {
                parse_grouping_group_members(walker, record.payload(), &mut group.members)?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHGLGroup collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDHGLGroup".to_string(),
    ))
}

/// `BrtBeginPCDHGLGMembers` collection (MS-XLSB 2.4.142).
fn parse_grouping_group_members(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    members: &mut Vec<PivotCacheGroupingGroupMember>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDHGLGMembers")?;
    let first_member = members.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDHGLG_MEMBERS => {
                return validate_collection_count(
                    declared,
                    members.len() - first_member,
                    "BrtBeginPCDHGLGMembers",
                );
            },
            rt::BEGIN_PCDHGLG_MEMBER => {
                let mut cursor = Cursor::new(record.payload(), "BrtBeginPCDHGLGMember");
                let is_group = cursor.read_u32()? != 0;
                let unique_name = cursor.read_wide_string()?;
                cursor.finish()?;
                members.push(PivotCacheGroupingGroupMember {
                    is_group,
                    unique_name,
                });
                walker.expect_end(rt::END_PCDHGLG_MEMBER, "BrtBeginPCDHGLGMember")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHGLGMembers collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDHGLGMembers".to_string(),
    ))
}

/// `BrtPCDH14` payload (MS-XLSB 2.4.726).
fn parse_hierarchy_ext14(data: &[u8]) -> Result<PivotCacheHierarchyExt14> {
    let mut cursor = Cursor::new(data, "BrtPCDH14");
    // FRTBlank header (4 bytes, MS-XLSB 2.5.55).
    cursor.guard(4)?;
    cursor.skip(4)?;
    let flags = cursor.read_u8()?;
    let hierarchy_count = cursor.read_u32()?;
    let mut hierarchy_indexes = Vec::new();
    for _ in 0..hierarchy_count {
        hierarchy_indexes.push(cursor.read_i32()?);
    }
    cursor.finish()?;
    Ok(PivotCacheHierarchyExt14 {
        flatten_hierarchies: flags & HIER14_FLATTEN != 0,
        measure_set: flags & HIER14_MEASURE_SET != 0,
        hierarchize_distinct: flags & HIER14_HIERARCHIZE_DISTINCT != 0,
        ignorable: flags & HIER14_IGNORABLE != 0,
        hierarchy_indexes,
    })
}

/// `BrtBeginPCDSDTupleCache` collection (MS-XLSB 2.4.164).
fn parse_tuple_cache(walker: &mut RecordWalker<'_>) -> Result<PivotCacheTupleCache> {
    let mut cache = PivotCacheTupleCache::default();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDSD_TUPLE_CACHE => return Ok(cache),
            rt::BEGIN_PCDSDTC_ENTRIES => {
                parse_tuple_cache_entries(walker, record.payload(), &mut cache.entries)?;
            },
            rt::BEGIN_PCDSDTC_QUERIES => {
                parse_tuple_cache_queries(walker, record.payload(), &mut cache.queries)?;
            },
            rt::BEGIN_PCDSDTC_SETS => {
                parse_tuple_cache_sets(walker, record.payload(), &mut cache.sets)?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSDTupleCache collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDSDTupleCache".to_string(),
    ))
}

/// `BrtBeginPCDSDTCEntries` collection (MS-XLSB 2.4.159).
fn parse_tuple_cache_entries(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    entries: &mut Vec<PivotCacheItemValue>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDSDTCEntries")?;
    let first_entry = entries.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDSDTC_ENTRIES => {
                return validate_collection_count(
                    declared,
                    entries.len() - first_entry,
                    "BrtBeginPCDSDTCEntries",
                );
            },
            // Tuple-cache entries may trail a PCDISrvFmt (sxvcellextra), so
            // parse non-strictly and drop the additional formatting bytes.
            item_type @ (rt::PCDI_MISSING
            | rt::PCDI_NUMBER
            | rt::PCDI_BOOLEAN
            | rt::PCDI_ERROR
            | rt::PCDI_STRING
            | rt::PCDI_DATETIME) => {
                entries.push(parse_cache_item(item_type, record.payload(), false)?.value);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSDTCEntries collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDSDTCEntries".to_string(),
    ))
}

/// `BrtBeginPCDSDTCQueries` collection (MS-XLSB 2.4.160).
fn parse_tuple_cache_queries(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    queries: &mut Vec<String>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDSDTCQueries")?;
    let first_query = queries.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDSDTC_QUERIES => {
                return validate_collection_count(
                    declared,
                    queries.len() - first_query,
                    "BrtBeginPCDSDTCQueries",
                );
            },
            rt::BEGIN_PCDSDTC_QUERY => {
                let mut cursor = Cursor::new(record.payload(), "BrtBeginPCDSDTCQuery");
                queries.push(cursor.read_wide_string()?);
                cursor.finish()?;
                walker.expect_end(rt::END_PCDSDTC_QUERY, "BrtBeginPCDSDTCQuery")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSDTCQueries collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDSDTCQueries".to_string(),
    ))
}

/// `BrtBeginPCDSDTCSets` collection (MS-XLSB 2.4.163).
fn parse_tuple_cache_sets(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    sets: &mut Vec<PivotCacheTupleCacheSet>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDSDTC Sets")?;
    let first_set = sets.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCDSDTC_SETS => {
                return validate_collection_count(
                    declared,
                    sets.len() - first_set,
                    "BrtBeginPCDSDTCSets",
                );
            },
            rt::BEGIN_PCDSDTC_SET => {
                sets.push(parse_tuple_cache_set(record.payload())?);
                walker.expect_end(rt::END_PCDSDTC_SET, "BrtBeginPCDSDTCSet")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSDTCSets collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDSDTCSets".to_string(),
    ))
}

/// `BrtBeginPCDSDTCSet` payload (MS-XLSB 2.4.162).
fn parse_tuple_cache_set(data: &[u8]) -> Result<PivotCacheTupleCacheSet> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDSDTCSet");
    let tuple_count = match cursor.read_u32()? {
        TUPLE_COUNT_UNKNOWN => None,
        count => Some(count),
    };
    let max_rank = cursor.read_u32()?;
    let sort_order = cursor.read_u32()?;
    let flags = cursor.read_u8()?;
    let definition = cursor.read_wide_string()?;
    cursor.finish()?;
    Ok(PivotCacheTupleCacheSet {
        tuple_count,
        max_rank,
        sort_order,
        query_failed: flags & TUPLE_SET_QUERY_FAILED != 0,
        definition,
    })
}

/// `BrtBeginPCDCalcItems` collection (MS-XLSB 2.4.125).
fn parse_calculated_items(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    definition: &mut PivotCacheDefinition,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDCalcItems")?;
    let first_item = definition.calculated_items.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_CALC_ITEMS => {
                return validate_collection_count(
                    declared,
                    definition.calculated_items.len() - first_item,
                    "BrtBeginPCDCalcItems",
                );
            },
            rt::BEGIN_PCD_CALC_ITEM => {
                let item = parse_calculated_item(walker, record.payload())?;
                definition.calculated_items.push(item);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDCalcItems collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDCalcItems".to_string(),
    ))
}

/// `BrtBeginPCDCalcItem` collection (MS-XLSB 2.4.124).
fn parse_calculated_item(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<CalculatedItem> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDCalcItem");
    // reserved (4 bytes): MUST be -1 and is ignored.
    cursor.guard(4)?;
    cursor.skip(4)?;
    let formula = parse_pivot_formula(&mut cursor)?;
    cursor.finish()?;
    let mut item = CalculatedItem {
        formula,
        names: Vec::new(),
        filters: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_CALC_ITEM => return Ok(item),
            rt::BEGIN_P_NAMES => parse_names(walker, record.payload(), &mut item.names)?,
            rt::BEGIN_PR_FILTERS => {
                parse_rule_filters(walker, record.payload(), &mut item.filters)?
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDCalcItem collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDCalcItem".to_string(),
    ))
}

/// `BrtBeginPNames` collection (MS-XLSB 2.4.177).
fn parse_names(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    names: &mut Vec<PivotName>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPNames")?;
    let first_name = names.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_P_NAMES => {
                return validate_collection_count(
                    declared,
                    names.len() - first_name,
                    "BrtBeginPNames",
                );
            },
            rt::BEGIN_P_NAME => {
                names.push(parse_name(walker, record.payload())?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPNames collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPNames".to_string()))
}

/// `BrtBeginPName` collection (MS-XLSB 2.4.176).
fn parse_name(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<PivotName> {
    let mut cursor = Cursor::new(data, "BrtBeginPName");
    let field_index = cursor.read_u32()?;
    let function = PivotNameFunction::try_from(cursor.read_u8()?)?;
    let flags = cursor.read_u8()?;
    // Two unnamed padding bytes.
    cursor.guard(2)?;
    cursor.skip(2)?;
    cursor.finish()?;
    let mut name = PivotName {
        field_index,
        function,
        err_name: flags & PNAME_ERR_NAME != 0,
        pairs: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_P_NAME => return Ok(name),
            rt::BEGIN_PN_PAIRS => parse_name_pairs(walker, record.payload(), &mut name.pairs)?,
            other => walker.skip_unhandled(other, "BrtBeginPName collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPName".to_string()))
}

/// `BrtBeginPNPairs` collection (MS-XLSB 2.4.179).
fn parse_name_pairs(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    pairs: &mut Vec<PivotNamePair>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPNPairs")?;
    let first_pair = pairs.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PN_PAIRS => {
                return validate_collection_count(
                    declared,
                    pairs.len() - first_pair,
                    "BrtBeginPNPairs",
                );
            },
            rt::BEGIN_PN_PAIR => {
                let mut cursor = Cursor::new(record.payload(), "BrtBeginPNPair");
                let flags = cursor.read_u8()?;
                let field_index = cursor.read_u32()?;
                let item_index = cursor.read_i32()?;
                // Three unnamed padding bytes.
                cursor.guard(3)?;
                cursor.skip(3)?;
                cursor.finish()?;
                pairs.push(PivotNamePair {
                    physical: flags & PNPAIR_PHYSICAL != 0,
                    relative: flags & PNPAIR_RELATIVE != 0,
                    field_index,
                    item_index,
                });
                walker.expect_end(rt::END_PN_PAIR, "BrtBeginPNPair")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPNPairs collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPNPairs".to_string()))
}

/// `BrtBeginPRFilters` collection (MS-XLSB 2.4.182).
fn parse_rule_filters(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    filters: &mut Vec<PivotRuleFilter>,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPRFilters")?;
    let first_filter = filters.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PR_FILTERS => {
                return validate_collection_count(
                    declared,
                    filters.len() - first_filter,
                    "BrtBeginPRFilters",
                );
            },
            rt::BEGIN_PR_FILTER => {
                filters.push(parse_rule_filter(walker, record.payload())?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPRFilters collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPRFilters".to_string()))
}

/// `BrtBeginPRFilter` collection (MS-XLSB 2.4.180; `PRFilter` structure).
fn parse_rule_filter(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<PivotRuleFilter> {
    let mut cursor = Cursor::new(data, "BrtBeginPRFilter");
    let field = cursor.read_i32()?;
    let declared_items = cursor.read_u32()?;
    let flags = u32::from(cursor.read_u8()?)
        | u32::from(cursor.read_u8()?) << 8
        | u32::from(cursor.read_u8()?) << 16;
    cursor.finish()?;
    let mut filter = PivotRuleFilter {
        field,
        item_types: flags & PR_FILTER_ITEM_TYPES_MASK,
        selected: flags & PR_FILTER_SELECTED != 0,
        items: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PR_FILTER => {
                validate_collection_count(declared_items, filter.items.len(), "BrtBeginPRFilter")?;
                return Ok(filter);
            },
            rt::BEGIN_PRF_ITEM => {
                let mut cursor = Cursor::new(record.payload(), "BrtBeginPRFItem");
                filter.items.push(cursor.read_u32()?);
                cursor.finish()?;
                walker.expect_end(rt::END_PRF_ITEM, "BrtBeginPRFItem")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPRFilter collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPRFilter".to_string()))
}

/// `BrtBeginPCDCalcMems` collection (MS-XLSB 2.4.129).
fn parse_calculated_members(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    definition: &mut PivotCacheDefinition,
) -> Result<()> {
    let declared = parse_collection_count(data, "BrtBeginPCDCalcMems")?;
    let first_member = definition.calculated_members.len();
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_CALC_MEMS => {
                return validate_collection_count(
                    declared,
                    definition.calculated_members.len() - first_member,
                    "BrtBeginPCDCalcMems",
                );
            },
            rt::BEGIN_PCD_CALC_MEM => {
                let member = parse_calculated_member(walker, record.payload())?;
                definition.calculated_members.push(member);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDCalcMems collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream(
        "BrtEndPCDCalcMems".to_string(),
    ))
}

/// `BrtBeginPCDCalcMem` collection (MS-XLSB 2.4.126; `PCDCalcMemCommon`, MS-XLSB 2.5.99).
fn parse_calculated_member(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<CalculatedMember> {
    let mut member = parse_calculated_member_payload(data)?;
    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PCD_CALC_MEM => return Ok(member),
            rt::BEGIN_PCD_CALC_MEM14 => {
                member.ext14 = Some(parse_calculated_member_ext14(record.payload())?);
                walker.expect_end(rt::END_PCD_CALC_MEM14, "BrtBeginPCDCalcMem14")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDCalcMem collection")?,
        }
    }
    Err(Error::UnexpectedEndOfStream("BrtEndPCDCalcMem".to_string()))
}

/// `PCDCalcMemCommon` (MS-XLSB 2.5.99).
fn parse_calculated_member_payload(data: &[u8]) -> Result<CalculatedMember> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDCalcMem");
    let flags = cursor.read_u32()?;
    let solve_order = cursor.read_i32()?;
    let is_set = cursor.read_u32()? != 0;
    let name = cursor.read_wide_string()?;
    let mdx = cursor.read_wide_string()?;
    let member_name = conditional_string(&mut cursor, flags & CALC_MEM_LOAD_MEMBER_NAME != 0)?;
    let source_hierarchy = conditional_string(&mut cursor, flags & CALC_MEM_LOAD_SOURCE_HIER != 0)?;
    let parent_unique = conditional_string(&mut cursor, flags & CALC_MEM_LOAD_PARENT_UNIQUE != 0)?;
    cursor.finish()?;
    Ok(CalculatedMember {
        name,
        mdx,
        solve_order,
        is_set,
        member_name,
        source_hierarchy,
        parent_unique,
        ext14: None,
    })
}

/// `BrtBeginPCDCalcMem14` payload (MS-XLSB 2.4.127).
fn parse_calculated_member_ext14(data: &[u8]) -> Result<CalculatedMemberExt14> {
    let mut cursor = Cursor::new(data, "BrtBeginPCDCalcMem14");
    // FRTBlank header (4 bytes, MS-XLSB 2.5.55).
    cursor.guard(4)?;
    cursor.skip(4)?;
    let flags = cursor.read_u8()?;
    let display_folder = cursor.read_wide_string()?;
    // The long MDX overflow string is present iff bytes remain.
    let long_mdx = if cursor.remaining() > 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    cursor.finish()?;
    Ok(CalculatedMemberExt14 {
        flatten_hierarchies: flags & CALC_MEM14_FLATTEN != 0,
        dynamic_set: flags & CALC_MEM14_DYNAMIC_SET != 0,
        hierarchize_distinct: flags & CALC_MEM14_HIERARCHIZE_DISTINCT != 0,
        display_folder,
        long_mdx,
    })
}

/// `BrtBeginPCD14` payload (MS-XLSB 2.4.123).
fn parse_pcd14(data: &[u8]) -> Result<PivotCacheDefinitionExt14> {
    let mut cursor = Cursor::new(data, "BrtBeginPCD14");
    // FRTBlank header (4 bytes, MS-XLSB 2.5.55).
    cursor.guard(4)?;
    cursor.skip(4)?;
    let flags = cursor.read_u8()?;
    let cache_id = cursor.read_i32()?;
    cursor.finish()?;
    Ok(PivotCacheDefinitionExt14 {
        slicer_data: flags & PCD14_SLICER_DATA != 0,
        server_support_subquery_calc_mem: flags & PCD14_SUBQUERY_CALC_MEM != 0,
        server_support_subquery_non_visual: flags & PCD14_SUBQUERY_NON_VISUAL != 0,
        server_support_add_calc_mems: flags & PCD14_ADD_CALC_MEMS != 0,
        cache_id,
    })
}
