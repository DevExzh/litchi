//! Record-walking parser for the XLSB PivotCache definition stream
//! (MS-XLSB 2.1.7.38).
//!
//! The parser is strict about record payloads it fully understands and
//! tolerant about everything else: unknown record types are ignored, and
//! known begin/end record pairs that carry no modelled data (KPIs, FRT
//! wrappers, server-format collections, ...) are skipped as balanced
//! collections.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::pivot::model::*;
use crate::xlsb::records::{XlsbRecord, XlsbRecordIter, record_types as rt, wide_str_with_len};
use litchi_core::binary;

// Record types used by this stream that have no constant in `records::record_types` yet.
/// `BrtPCDField14` (MS-XLSB 2.4.725): marks the preceding cache field as ignorable.
const PCD_FIELD14: u16 = 1141;
/// `BrtPCDH14` (MS-XLSB 2.4.726): named-set extension of a cache hierarchy.
const PCD_H14: u16 = 1037;
/// `BrtBeginPRule` / `BrtEndPRule` (MS-XLSB 2.4.186): unmodelled pivot rule.
const BEGIN_PRULE: u16 = 247;
const END_PRULE: u16 = 248;
/// `BrtBeginPCDKPIs` / `BrtEndPCDKPIs` (MS-XLSB 2.4.149): unmodelled KPI collection.
const BEGIN_PCD_KPIS: u16 = 269;
const END_PCD_KPIS: u16 = 270;
/// `BrtBeginPCDKPI` / `BrtEndPCDKPI` (MS-XLSB 2.4.148): unmodelled KPI.
const BEGIN_PCD_KPI: u16 = 271;
const END_PCD_KPI: u16 = 272;

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
pub fn parse_pivot_cache_definition(data: &[u8]) -> XlsbResult<PivotCacheDefinition> {
    let mut walker = RecordWalker::new(data);
    let first = walker.required("BrtBeginPivotCacheDef")?;
    if first.header.record_type != rt::BEGIN_PIVOT_CACHE_DEF {
        return Err(XlsbError::UnexpectedRecord {
            expected: rt::BEGIN_PIVOT_CACHE_DEF,
            found: first.header.record_type,
        });
    }
    let mut definition = parse_definition_payload(&first.data)?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PIVOT_CACHE_DEF => return Ok(definition),
            rt::BEGIN_PCD_SOURCE => {
                definition.source = Some(parse_source(&mut walker, &record.data)?);
            },
            rt::BEGIN_PCD_FIELDS => {
                parse_fields(&mut walker, &record.data, &mut definition)?;
            },
            rt::BEGIN_PCD_HIERARCHIES => {
                parse_hierarchies(&mut walker, &record.data, &mut definition)?;
            },
            rt::BEGIN_PCDSD_TUPLE_CACHE => {
                definition.tuple_cache = Some(parse_tuple_cache(&mut walker)?);
            },
            rt::BEGIN_PCD_CALC_ITEMS => {
                parse_calculated_items(&mut walker, &record.data, &mut definition)?;
            },
            rt::BEGIN_PCD_CALC_MEMS => {
                parse_calculated_members(&mut walker, &record.data, &mut definition)?;
            },
            rt::BEGIN_PCD14 => {
                definition.ext14 = Some(parse_pcd14(&record.data)?);
                walker.expect_end(rt::END_PCD14, "BrtBeginPCD14")?;
            },
            other => walker.skip_unhandled(other, "PivotCache definition stream")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPivotCacheDef".to_string(),
    ))
}

/// Wraps the shared record iterator with the collection helpers this parser needs.
struct RecordWalker<'a> {
    iter: XlsbRecordIter<&'a [u8]>,
}

impl<'a> RecordWalker<'a> {
    fn new(data: &'a [u8]) -> Self {
        RecordWalker {
            iter: XlsbRecordIter::new(data),
        }
    }

    fn next(&mut self) -> XlsbResult<Option<XlsbRecord>> {
        self.iter.next().transpose()
    }

    fn required(&mut self, context: &'static str) -> XlsbResult<XlsbRecord> {
        self.next()?
            .ok_or_else(|| XlsbError::UnexpectedEndOfStream(context.to_string()))
    }

    /// Consume records up to and including `end_type`, tolerating nested
    /// collections of the same record pair.
    fn skip_collection(
        &mut self,
        begin_type: u16,
        end_type: u16,
        context: &'static str,
    ) -> XlsbResult<()> {
        let mut depth = 1u32;
        while let Some(record) = self.next()? {
            if record.header.record_type == begin_type {
                depth += 1;
            } else if record.header.record_type == end_type {
                depth -= 1;
                if depth == 0 {
                    return Ok(());
                }
            }
        }
        Err(XlsbError::UnexpectedEndOfStream(context.to_string()))
    }

    /// Skip a record the parser does not handle: a balanced collection when
    /// the type is a known begin record, a single record otherwise.
    fn skip_unhandled(&mut self, record_type: u16, context: &'static str) -> XlsbResult<()> {
        if let Some(end_type) = paired_end(record_type) {
            self.skip_collection(record_type, end_type, context)?;
        }
        Ok(())
    }

    /// Consume everything up to the matching end record of a collection that
    /// is expected to contain no modelled children.
    fn expect_end(&mut self, end_type: u16, context: &'static str) -> XlsbResult<()> {
        while let Some(record) = self.next()? {
            let record_type = record.header.record_type;
            if record_type == end_type {
                return Ok(());
            }
            self.skip_unhandled(record_type, context)?;
        }
        Err(XlsbError::UnexpectedEndOfStream(context.to_string()))
    }
}

/// Map a known begin record type to its matching end record type.
///
/// Returns `None` for standalone records and unknown types, which the parser
/// then skips as single records.
fn paired_end(record_type: u16) -> Option<u16> {
    Some(match record_type {
        rt::BEGIN_PCD_SOURCE => rt::END_PCD_SOURCE,
        rt::BEGIN_PCDS_RANGE => rt::END_PCDS_RANGE,
        rt::BEGIN_PCDS_CONSOL => rt::END_PCDS_CONSOL,
        rt::BEGIN_PCDSC_PAGES => rt::END_PCDSC_PAGES,
        rt::BEGIN_PCDSC_PAGE => rt::END_PCDSC_PAGE,
        rt::BEGIN_PCDSCP_ITEM => rt::END_PCDSCP_ITEM,
        rt::BEGIN_PCDSC_SETS => rt::END_PCDSC_SETS,
        rt::BEGIN_PCDSC_SET => rt::END_PCDSC_SET,
        rt::BEGIN_PCD_FIELDS => rt::END_PCD_FIELDS,
        rt::BEGIN_PCD_FIELD => rt::END_PCD_FIELD,
        rt::BEGIN_PCDF_ATBL => rt::END_PCDF_ATBL,
        rt::BEGIN_PCDI_RUN => rt::END_PCDI_RUN,
        rt::BEGIN_PCDF_GROUP => rt::END_PCDF_GROUP,
        rt::BEGIN_PCDFG_ITEMS => rt::END_PCDFG_ITEMS,
        rt::BEGIN_PCDFG_RANGE => rt::END_PCDFG_RANGE,
        rt::BEGIN_PCDFG_DISCRETE => rt::END_PCDFG_DISCRETE,
        rt::BEGIN_PCD_HIERARCHIES => rt::END_PCD_HIERARCHIES,
        rt::BEGIN_PCD_HIERARCHY => rt::END_PCD_HIERARCHY,
        rt::BEGIN_PCDH_FIELDS_USAGE => rt::END_PCDH_FIELDS_USAGE,
        rt::BEGIN_PCDHG_LEVELS => rt::END_PCDHG_LEVELS,
        rt::BEGIN_PCDHG_LEVEL => rt::END_PCDHG_LEVEL,
        rt::BEGIN_PCDHGL_GROUPS => rt::END_PCDHGL_GROUPS,
        rt::BEGIN_PCDHGL_GROUP => rt::END_PCDHGL_GROUP,
        rt::BEGIN_PCDHGLG_MEMBERS => rt::END_PCDHGLG_MEMBERS,
        rt::BEGIN_PCDHGLG_MEMBER => rt::END_PCDHGLG_MEMBER,
        rt::BEGIN_PCDSD_TUPLE_CACHE => rt::END_PCDSD_TUPLE_CACHE,
        rt::BEGIN_PCDSDTC_ENTRIES => rt::END_PCDSDTC_ENTRIES,
        rt::BEGIN_PCDSDTC_MEMBERS => rt::END_PCDSDTC_MEMBERS,
        rt::BEGIN_PCDSDTC_MEMBER => rt::END_PCDSDTC_MEMBER,
        rt::BEGIN_PCDSDTC_QUERIES => rt::END_PCDSDTC_QUERIES,
        rt::BEGIN_PCDSDTC_QUERY => rt::END_PCDSDTC_QUERY,
        rt::BEGIN_PCDSDTC_SETS => rt::END_PCDSDTC_SETS,
        rt::BEGIN_PCDSDTC_SET => rt::END_PCDSDTC_SET,
        rt::BEGIN_PCDSDTC_MEMBERS_SORT_BY => rt::END_PCDSDTC_MEMBERS_SORT_BY,
        rt::BEGIN_PCD_SFCI_ENTRIES => rt::END_PCD_SFCI_ENTRIES,
        rt::BEGIN_PCD_CALC_ITEMS => rt::END_PCD_CALC_ITEMS,
        rt::BEGIN_PCD_CALC_ITEM => rt::END_PCD_CALC_ITEM,
        rt::BEGIN_PCD_CALC_MEMS => rt::END_PCD_CALC_MEMS,
        rt::BEGIN_PCD_CALC_MEM => rt::END_PCD_CALC_MEM,
        rt::BEGIN_PCD_CALC_MEM14 => rt::END_PCD_CALC_MEM14,
        rt::BEGIN_PCD_CALC_MEM_EXT => rt::END_PCD_CALC_MEM_EXT,
        rt::BEGIN_PCD_CALC_MEMS_EXT => rt::END_PCD_CALC_MEMS_EXT,
        rt::BEGIN_PCD14 => rt::END_PCD14,
        rt::BEGIN_PR_FILTERS => rt::END_PR_FILTERS,
        rt::BEGIN_PR_FILTER => rt::END_PR_FILTER,
        rt::BEGIN_PRF_ITEM => rt::END_PRF_ITEM,
        rt::BEGIN_PR_FILTERS14 => rt::END_PR_FILTERS14,
        rt::BEGIN_PR_FILTER14 => rt::END_PR_FILTER14,
        rt::BEGIN_PRF_ITEM14 => rt::END_PRF_ITEM14,
        rt::BEGIN_P_NAMES => rt::END_P_NAMES,
        rt::BEGIN_P_NAME => rt::END_P_NAME,
        rt::BEGIN_PN_PAIRS => rt::END_PN_PAIRS,
        rt::BEGIN_PN_PAIR => rt::END_PN_PAIR,
        rt::BEGIN_ITEM_UNIQUE_NAMES => rt::END_ITEM_UNIQUE_NAMES,
        rt::FRT_BEGIN => rt::FRT_END,
        rt::AC_BEGIN => rt::AC_END,
        BEGIN_PRULE => END_PRULE,
        BEGIN_PCD_KPIS => END_PCD_KPIS,
        BEGIN_PCD_KPI => END_PCD_KPI,
        _ => return None,
    })
}

/// Bounds-checked cursor over one record payload.
struct PayloadCursor<'a> {
    data: &'a [u8],
    offset: usize,
    context: &'static str,
}

impl<'a> PayloadCursor<'a> {
    fn new(data: &'a [u8], context: &'static str) -> Self {
        PayloadCursor {
            data,
            offset: 0,
            context,
        }
    }

    fn remaining(&self) -> usize {
        self.data.len() - self.offset
    }

    fn guard(&self, needed: usize) -> XlsbResult<()> {
        if self.remaining() < needed {
            return Err(XlsbError::InvalidLength {
                expected: self.offset + needed,
                found: self.data.len(),
            });
        }
        Ok(())
    }

    fn read_u8(&mut self) -> XlsbResult<u8> {
        self.guard(1)?;
        let value = self.data[self.offset];
        self.offset += 1;
        Ok(value)
    }

    fn read_u16(&mut self) -> XlsbResult<u16> {
        self.guard(2)?;
        let value = binary::read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;
        Ok(value)
    }

    fn read_u32(&mut self) -> XlsbResult<u32> {
        self.guard(4)?;
        let value = binary::read_u32_le_at(self.data, self.offset)?;
        self.offset += 4;
        Ok(value)
    }

    fn read_i32(&mut self) -> XlsbResult<i32> {
        Ok(self.read_u32()? as i32)
    }

    fn read_f64(&mut self) -> XlsbResult<f64> {
        self.guard(8)?;
        let value = binary::read_f64_le_at(self.data, self.offset)?;
        self.offset += 8;
        Ok(value)
    }

    /// Read a full-width Boolean (any nonzero value is `true`).
    fn read_bool8(&mut self) -> XlsbResult<bool> {
        Ok(self.read_u8()? != 0)
    }

    fn read_range(&mut self) -> XlsbResult<PivotCacheRange> {
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

    /// Read an `XLWideString` (MS-XLSB 2.5.169).
    fn read_wide_string(&mut self) -> XlsbResult<String> {
        let (value, consumed) = wide_str_with_len(&self.data[self.offset..])?;
        self.offset += consumed;
        Ok(value)
    }

    /// Read an `XLNullableWideString` (MS-XLSB 2.5.167).
    fn read_nullable_wide_string(&mut self) -> XlsbResult<Option<String>> {
        self.guard(4)?;
        if binary::read_u32_le_at(self.data, self.offset)? == u32::MAX {
            self.offset += 4;
            return Ok(None);
        }
        self.read_wide_string().map(Some)
    }

    /// Read a length-prefixed byte blob (`cce`/`cb` prefixed formula parts).
    fn read_blob(&mut self) -> XlsbResult<Vec<u8>> {
        let len = usize::try_from(self.read_u32()?)
            .map_err(|_| malformed(self.context, "byte blob length overflow"))?;
        self.guard(len)?;
        let blob = self.data[self.offset..self.offset + len].to_vec();
        self.offset += len;
        Ok(blob)
    }

    /// Reject payloads with unparsed trailing bytes.
    fn finish(&self) -> XlsbResult<()> {
        if self.remaining() != 0 {
            return Err(malformed(
                self.context,
                format!("{} trailing bytes", self.remaining()),
            ));
        }
        Ok(())
    }
}

fn malformed(context: &str, detail: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: context.to_string(),
        val: detail.into(),
    }
}

/// `BrtBeginPivotCacheDef` payload (MS-XLSB 2.4.168).
fn parse_definition_payload(data: &[u8]) -> XlsbResult<PivotCacheDefinition> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPivotCacheDef");
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
        cursor.offset += 4;
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
fn parse_source(walker: &mut RecordWalker<'_>, data: &[u8]) -> XlsbResult<PivotCacheSource> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDSource");
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
        match record.header.record_type {
            rt::END_PCD_SOURCE => return Ok(source),
            rt::BEGIN_PCDS_RANGE => {
                source.worksheet = Some(parse_worksheet_range(&record.data)?);
                walker.expect_end(rt::END_PCDS_RANGE, "BrtBeginPCDSRange")?;
            },
            rt::BEGIN_PCDS_CONSOL => {
                source.consolidation = Some(parse_consolidation(walker, &record.data)?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSource collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSource".to_string(),
    ))
}

/// `BrtBeginPCDSRange` payload (MS-XLSB 2.4.167).
fn parse_worksheet_range(data: &[u8]) -> XlsbResult<PivotCacheWorksheetSource> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDSRange");
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
) -> XlsbResult<PivotCacheConsolidationSource> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDSConsol");
    let flags = cursor.read_u16()?;
    cursor.finish()?;
    let mut consolidation = PivotCacheConsolidationSource {
        auto_page: flags & CONSOL_AUTO_PAGE != 0,
        sets: Vec::new(),
        pages: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDS_CONSOL => return Ok(consolidation),
            rt::BEGIN_PCDSC_SETS => parse_consolidation_sets(walker, &mut consolidation)?,
            rt::BEGIN_PCDSC_PAGES => parse_consolidation_pages(walker, &mut consolidation)?,
            other => walker.skip_unhandled(other, "BrtBeginPCDSConsol collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSConsol".to_string(),
    ))
}

/// `BrtBeginPCDSCSets` collection (MS-XLSB 2.4.155).
fn parse_consolidation_sets(
    walker: &mut RecordWalker<'_>,
    consolidation: &mut PivotCacheConsolidationSource,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDSC_SETS => return Ok(()),
            rt::BEGIN_PCDSC_SET => {
                consolidation
                    .sets
                    .push(parse_consolidation_set(&record.data)?);
                walker.expect_end(rt::END_PCDSC_SET, "BrtBeginPCDSCSet")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSCSets collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSCSets".to_string(),
    ))
}

/// `BrtBeginPCDSCSet` payload (MS-XLSB 2.4.154).
fn parse_consolidation_set(data: &[u8]) -> XlsbResult<PivotCacheConsolidationSet> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDSCSet");
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
    consolidation: &mut PivotCacheConsolidationSource,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDSC_PAGES => return Ok(()),
            rt::BEGIN_PCDSC_PAGE => {
                let page = parse_consolidation_page(walker)?;
                consolidation.pages.push(page);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSCPages collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSCPages".to_string(),
    ))
}

/// `BrtBeginPCDSCPage` collection (MS-XLSB 2.4.151).
fn parse_consolidation_page(
    walker: &mut RecordWalker<'_>,
) -> XlsbResult<PivotCacheConsolidationPage> {
    let mut page = PivotCacheConsolidationPage {
        item_names: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDSC_PAGE => return Ok(page),
            rt::BEGIN_PCDSCP_ITEM => {
                let mut cursor = PayloadCursor::new(&record.data, "BrtBeginPCDSCPItem");
                page.item_names.push(cursor.read_wide_string()?);
                cursor.finish()?;
                walker.expect_end(rt::END_PCDSCP_ITEM, "BrtBeginPCDSCPItem")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSCPage collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSCPage".to_string(),
    ))
}

/// `BrtBeginPCDFields` collection (MS-XLSB 2.4.137).
fn parse_fields(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    definition: &mut PivotCacheDefinition,
) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDFields");
    // `cFields` declares the field count; the actual records define the model.
    let _declared_fields = cursor.read_u32()?;
    cursor.finish()?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCD_FIELDS => return Ok(()),
            rt::BEGIN_PCD_FIELD => {
                let field = parse_field(walker, &record.data)?;
                definition.fields.push(field);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFields collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDFields".to_string(),
    ))
}

/// `BrtBeginPCDField` collection (MS-XLSB 2.4.136).
fn parse_field(walker: &mut RecordWalker<'_>, data: &[u8]) -> XlsbResult<PivotCacheField> {
    let mut field = parse_field_payload(data)?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCD_FIELD => return Ok(field),
            rt::BEGIN_PCDF_ATBL => {
                parse_shared_items(walker, &record.data, &mut field.shared_items)?
            },
            rt::BEGIN_PCDF_GROUP => {
                field.grouping = Some(parse_grouping(walker, &record.data)?);
            },
            PCD_FIELD14 => field.ignore = true,
            other => walker.skip_unhandled(other, "BrtBeginPCDField collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDField".to_string(),
    ))
}

/// `BrtBeginPCDField` payload (MS-XLSB 2.4.136).
fn parse_field_payload(data: &[u8]) -> XlsbResult<PivotCacheField> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDField");
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
fn parse_pivot_formula(cursor: &mut PayloadCursor<'_>) -> XlsbResult<PivotParsedFormulaData> {
    let tokens = cursor.read_blob()?;
    let extra = cursor.read_blob()?;
    Ok(PivotParsedFormulaData { tokens, extra })
}

/// `BrtBeginPCDFAtbl` collection (MS-XLSB 2.4.131): shared item statistics
/// followed by the raw cache items.
fn parse_shared_items(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    shared_items: &mut PivotCacheSharedItems,
) -> XlsbResult<()> {
    shared_items.stats = Some(parse_shared_items_stats(data)?);
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDF_ATBL => return Ok(()),
            rt::BEGIN_PCDI_RUN => {
                parse_item_run(&record.data, &mut shared_items.items)?;
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
                let item = parse_cache_item(item_type, &record.data, true)?;
                shared_items.items.push(item);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFAtbl collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDFAtbl".to_string(),
    ))
}

/// `BrtBeginPCDFAtbl` payload (MS-XLSB 2.4.131).
fn parse_shared_items_stats(data: &[u8]) -> XlsbResult<PivotCacheSharedItemsStats> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDFAtbl");
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
fn parse_cache_item(record_type: u16, data: &[u8], strict: bool) -> XlsbResult<PivotCacheItem> {
    let mut cursor = PayloadCursor::new(data, "BrtPCDI cache item");
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
            return Err(XlsbError::UnexpectedRecord {
                expected: rt::PCDI_MISSING,
                found: record_type,
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
fn read_date_time(cursor: &mut PayloadCursor<'_>) -> XlsbResult<PivotCacheDateTime> {
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
fn parse_item_info(cursor: &mut PayloadCursor<'_>) -> XlsbResult<PivotCacheItemInfo> {
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
fn parse_item_run(data: &[u8], items: &mut Vec<PivotCacheItem>) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDIRun");
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
    cursor.finish()
}

/// `BrtBeginPCDFGroup` collection (MS-XLSB 2.4.135).
fn parse_grouping(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
) -> XlsbResult<PivotCacheFieldGrouping> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDFGroup");
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
        match record.header.record_type {
            rt::END_PCDF_GROUP => return Ok(grouping),
            rt::BEGIN_PCDFG_RANGE => {
                grouping.range = Some(parse_range_grouping(&record.data)?);
                walker.expect_end(rt::END_PCDFG_RANGE, "BrtBeginPCDFGRange")?;
            },
            rt::BEGIN_PCDFG_DISCRETE => {
                grouping.discrete = Some(parse_discrete_grouping(walker)?);
            },
            rt::BEGIN_PCDFG_ITEMS => {
                parse_grouping_items(walker, &mut grouping.items)?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFGroup collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDFGroup".to_string(),
    ))
}

/// Interpret a spec `-1`-or-index field as an optional unsigned index.
fn non_negative_index(value: i32) -> Option<u32> {
    u32::try_from(value).ok()
}

/// `BrtBeginPCDFGRange` payload (MS-XLSB 2.4.134).
fn parse_range_grouping(data: &[u8]) -> XlsbResult<PivotCacheRangeGrouping> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDFGRange");
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
) -> XlsbResult<PivotCacheDiscreteGrouping> {
    let mut grouping = PivotCacheDiscreteGrouping {
        item_indexes: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDFG_DISCRETE => return Ok(grouping),
            rt::PCDI_INDEX => {
                let mut cursor = PayloadCursor::new(&record.data, "BrtPCDIIndex");
                grouping.item_indexes.push(cursor.read_u32()?);
                cursor.finish()?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFGDiscrete collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDFGDiscrete".to_string(),
    ))
}

/// `BrtBeginPCDFGItems` collection (MS-XLSB 2.4.133): grouping cache items.
fn parse_grouping_items(
    walker: &mut RecordWalker<'_>,
    items: &mut Vec<PivotCacheItem>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDFG_ITEMS => return Ok(()),
            rt::BEGIN_PCDI_RUN => {
                parse_item_run(&record.data, items)?;
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
                items.push(parse_cache_item(item_type, &record.data, true)?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDFGItems collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDFGItems".to_string(),
    ))
}

/// `BrtBeginPCDHierarchies` collection (MS-XLSB 2.4.145).
fn parse_hierarchies(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    definition: &mut PivotCacheDefinition,
) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDHierarchies");
    let _declared_hierarchies = cursor.read_u32()?;
    cursor.finish()?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCD_HIERARCHIES => return Ok(()),
            rt::BEGIN_PCD_HIERARCHY => {
                let hierarchy = parse_hierarchy(walker, &record.data)?;
                definition.hierarchies.push(hierarchy);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHierarchies collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDHierarchies".to_string(),
    ))
}

/// `BrtBeginPCDHierarchy` collection (MS-XLSB 2.4.146).
fn parse_hierarchy(walker: &mut RecordWalker<'_>, data: &[u8]) -> XlsbResult<PivotCacheHierarchy> {
    let mut hierarchy = parse_hierarchy_payload(data)?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCD_HIERARCHY => return Ok(hierarchy),
            rt::BEGIN_PCDH_FIELDS_USAGE => {
                hierarchy.field_usage = parse_fields_usage(&record.data)?;
                walker.expect_end(rt::END_PCDH_FIELDS_USAGE, "BrtBeginPCDHFieldsUsage")?;
            },
            rt::BEGIN_PCDHG_LEVELS => {
                parse_grouping_levels(walker, &mut hierarchy.grouping_levels)?;
            },
            rt::BEGIN_PCDHGL_GROUPS => {
                parse_grouping_groups(walker, &mut hierarchy.grouping_groups)?;
            },
            PCD_H14 => {
                hierarchy.ext14 = Some(parse_hierarchy_ext14(&record.data)?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHierarchy collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDHierarchy".to_string(),
    ))
}

/// `BrtBeginPCDHierarchy` payload (MS-XLSB 2.4.146).
fn parse_hierarchy_payload(data: &[u8]) -> XlsbResult<PivotCacheHierarchy> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDHierarchy");
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

fn conditional_string(cursor: &mut PayloadCursor<'_>, present: bool) -> XlsbResult<Option<String>> {
    if present {
        cursor.read_wide_string().map(Some)
    } else {
        Ok(None)
    }
}

/// `BrtBeginPCDHFieldsUsage` payload (MS-XLSB 2.4.138).
fn parse_fields_usage(data: &[u8]) -> XlsbResult<Vec<i32>> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDHFieldsUsage");
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
    levels: &mut Vec<PivotCacheGroupingLevel>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDHG_LEVELS => return Ok(()),
            rt::BEGIN_PCDHG_LEVEL => {
                let mut cursor = PayloadCursor::new(&record.data, "BrtBeginPCDHGLevel");
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
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDHGLevels".to_string(),
    ))
}

/// `BrtBeginPCDHGLGroups` collection (MS-XLSB 2.4.144).
fn parse_grouping_groups(
    walker: &mut RecordWalker<'_>,
    groups: &mut Vec<PivotCacheGroupingGroup>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDHGL_GROUPS => return Ok(()),
            rt::BEGIN_PCDHGL_GROUP => {
                let group = parse_grouping_group(walker, &record.data)?;
                groups.push(group);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHGLGroups collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDHGLGroups".to_string(),
    ))
}

/// `BrtBeginPCDHGLGroup` collection (MS-XLSB 2.4.143).
fn parse_grouping_group(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
) -> XlsbResult<PivotCacheGroupingGroup> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDHGLGroup");
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
        match record.header.record_type {
            rt::END_PCDHGL_GROUP => return Ok(group),
            rt::BEGIN_PCDHGLG_MEMBERS => {
                parse_grouping_group_members(walker, &mut group.members)?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDHGLGroup collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDHGLGroup".to_string(),
    ))
}

/// `BrtBeginPCDHGLGMembers` collection (MS-XLSB 2.4.142).
fn parse_grouping_group_members(
    walker: &mut RecordWalker<'_>,
    members: &mut Vec<PivotCacheGroupingGroupMember>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDHGLG_MEMBERS => return Ok(()),
            rt::BEGIN_PCDHGLG_MEMBER => {
                let mut cursor = PayloadCursor::new(&record.data, "BrtBeginPCDHGLGMember");
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
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDHGLGMembers".to_string(),
    ))
}

/// `BrtPCDH14` payload (MS-XLSB 2.4.726).
fn parse_hierarchy_ext14(data: &[u8]) -> XlsbResult<PivotCacheHierarchyExt14> {
    let mut cursor = PayloadCursor::new(data, "BrtPCDH14");
    // FRTBlank header (4 bytes, MS-XLSB 2.5.55).
    cursor.guard(4)?;
    cursor.offset += 4;
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
fn parse_tuple_cache(walker: &mut RecordWalker<'_>) -> XlsbResult<PivotCacheTupleCache> {
    let mut cache = PivotCacheTupleCache::default();
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDSD_TUPLE_CACHE => return Ok(cache),
            rt::BEGIN_PCDSDTC_ENTRIES => {
                parse_tuple_cache_entries(walker, &mut cache.entries)?;
            },
            rt::BEGIN_PCDSDTC_QUERIES => {
                parse_tuple_cache_queries(walker, &mut cache.queries)?;
            },
            rt::BEGIN_PCDSDTC_SETS => {
                parse_tuple_cache_sets(walker, &mut cache.sets)?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSDTupleCache collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSDTupleCache".to_string(),
    ))
}

/// `BrtBeginPCDSDTCEntries` collection (MS-XLSB 2.4.159).
fn parse_tuple_cache_entries(
    walker: &mut RecordWalker<'_>,
    entries: &mut Vec<PivotCacheItemValue>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDSDTC_ENTRIES => return Ok(()),
            // Tuple-cache entries may trail a PCDISrvFmt (sxvcellextra), so
            // parse non-strictly and drop the additional formatting bytes.
            item_type @ (rt::PCDI_MISSING
            | rt::PCDI_NUMBER
            | rt::PCDI_BOOLEAN
            | rt::PCDI_ERROR
            | rt::PCDI_STRING
            | rt::PCDI_DATETIME) => {
                entries.push(parse_cache_item(item_type, &record.data, false)?.value);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSDTCEntries collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSDTCEntries".to_string(),
    ))
}

/// `BrtBeginPCDSDTCQueries` collection (MS-XLSB 2.4.160).
fn parse_tuple_cache_queries(
    walker: &mut RecordWalker<'_>,
    queries: &mut Vec<String>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDSDTC_QUERIES => return Ok(()),
            rt::BEGIN_PCDSDTC_QUERY => {
                let mut cursor = PayloadCursor::new(&record.data, "BrtBeginPCDSDTCQuery");
                queries.push(cursor.read_wide_string()?);
                cursor.finish()?;
                walker.expect_end(rt::END_PCDSDTC_QUERY, "BrtBeginPCDSDTCQuery")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSDTCQueries collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSDTCQueries".to_string(),
    ))
}

/// `BrtBeginPCDSDTCSets` collection (MS-XLSB 2.4.163).
fn parse_tuple_cache_sets(
    walker: &mut RecordWalker<'_>,
    sets: &mut Vec<PivotCacheTupleCacheSet>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCDSDTC_SETS => return Ok(()),
            rt::BEGIN_PCDSDTC_SET => {
                sets.push(parse_tuple_cache_set(&record.data)?);
                walker.expect_end(rt::END_PCDSDTC_SET, "BrtBeginPCDSDTCSet")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDSDTCSets collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDSDTCSets".to_string(),
    ))
}

/// `BrtBeginPCDSDTCSet` payload (MS-XLSB 2.4.162).
fn parse_tuple_cache_set(data: &[u8]) -> XlsbResult<PivotCacheTupleCacheSet> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDSDTCSet");
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
) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDCalcItems");
    let _declared_items = cursor.read_u32()?;
    cursor.finish()?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCD_CALC_ITEMS => return Ok(()),
            rt::BEGIN_PCD_CALC_ITEM => {
                let item = parse_calculated_item(walker, &record.data)?;
                definition.calculated_items.push(item);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDCalcItems collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDCalcItems".to_string(),
    ))
}

/// `BrtBeginPCDCalcItem` collection (MS-XLSB 2.4.124).
fn parse_calculated_item(walker: &mut RecordWalker<'_>, data: &[u8]) -> XlsbResult<CalculatedItem> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDCalcItem");
    // reserved (4 bytes): MUST be -1 and is ignored.
    cursor.guard(4)?;
    cursor.offset += 4;
    let formula = parse_pivot_formula(&mut cursor)?;
    cursor.finish()?;
    let mut item = CalculatedItem {
        formula,
        names: Vec::new(),
        filters: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCD_CALC_ITEM => return Ok(item),
            rt::BEGIN_P_NAMES => parse_names(walker, &mut item.names)?,
            rt::BEGIN_PR_FILTERS => parse_rule_filters(walker, &mut item.filters)?,
            other => walker.skip_unhandled(other, "BrtBeginPCDCalcItem collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDCalcItem".to_string(),
    ))
}

/// `BrtBeginPNames` collection (MS-XLSB 2.4.177).
fn parse_names(walker: &mut RecordWalker<'_>, names: &mut Vec<PivotName>) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_P_NAMES => return Ok(()),
            rt::BEGIN_P_NAME => {
                names.push(parse_name(walker, &record.data)?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPNames collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream("BrtEndPNames".to_string()))
}

/// `BrtBeginPName` collection (MS-XLSB 2.4.176).
fn parse_name(walker: &mut RecordWalker<'_>, data: &[u8]) -> XlsbResult<PivotName> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPName");
    let field_index = cursor.read_u32()?;
    let function = PivotNameFunction::try_from(cursor.read_u8()?)?;
    let flags = cursor.read_u8()?;
    // Two unnamed padding bytes.
    cursor.guard(2)?;
    cursor.offset += 2;
    cursor.finish()?;
    let mut name = PivotName {
        field_index,
        function,
        err_name: flags & PNAME_ERR_NAME != 0,
        pairs: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_P_NAME => return Ok(name),
            rt::BEGIN_PN_PAIRS => parse_name_pairs(walker, &mut name.pairs)?,
            other => walker.skip_unhandled(other, "BrtBeginPName collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream("BrtEndPName".to_string()))
}

/// `BrtBeginPNPairs` collection (MS-XLSB 2.4.179).
fn parse_name_pairs(
    walker: &mut RecordWalker<'_>,
    pairs: &mut Vec<PivotNamePair>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PN_PAIRS => return Ok(()),
            rt::BEGIN_PN_PAIR => {
                let mut cursor = PayloadCursor::new(&record.data, "BrtBeginPNPair");
                let flags = cursor.read_u8()?;
                let field_index = cursor.read_u32()?;
                let item_index = cursor.read_i32()?;
                // Three unnamed padding bytes.
                cursor.guard(3)?;
                cursor.offset += 3;
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
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPNPairs".to_string(),
    ))
}

/// `BrtBeginPRFilters` collection (MS-XLSB 2.4.182).
fn parse_rule_filters(
    walker: &mut RecordWalker<'_>,
    filters: &mut Vec<PivotRuleFilter>,
) -> XlsbResult<()> {
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PR_FILTERS => return Ok(()),
            rt::BEGIN_PR_FILTER => {
                filters.push(parse_rule_filter(walker, &record.data)?);
            },
            other => walker.skip_unhandled(other, "BrtBeginPRFilters collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPRFilters".to_string(),
    ))
}

/// `BrtBeginPRFilter` collection (MS-XLSB 2.4.180; `PRFilter` structure).
fn parse_rule_filter(walker: &mut RecordWalker<'_>, data: &[u8]) -> XlsbResult<PivotRuleFilter> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPRFilter");
    let field = cursor.read_i32()?;
    let _declared_items = cursor.read_u32()?;
    let flags = cursor.read_u8()? as u32
        | (cursor.read_u8()? as u32) << 8
        | (cursor.read_u8()? as u32) << 16;
    cursor.finish()?;
    let mut filter = PivotRuleFilter {
        field,
        item_types: flags & PR_FILTER_ITEM_TYPES_MASK,
        selected: flags & PR_FILTER_SELECTED != 0,
        items: Vec::new(),
    };
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PR_FILTER => return Ok(filter),
            rt::BEGIN_PRF_ITEM => {
                let mut cursor = PayloadCursor::new(&record.data, "BrtBeginPRFItem");
                filter.items.push(cursor.read_u32()?);
                cursor.finish()?;
                walker.expect_end(rt::END_PRF_ITEM, "BrtBeginPRFItem")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPRFilter collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPRFilter".to_string(),
    ))
}

/// `BrtBeginPCDCalcMems` collection (MS-XLSB 2.4.129).
fn parse_calculated_members(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    definition: &mut PivotCacheDefinition,
) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDCalcMems");
    let _declared_members = cursor.read_u32()?;
    cursor.finish()?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCD_CALC_MEMS => return Ok(()),
            rt::BEGIN_PCD_CALC_MEM => {
                let member = parse_calculated_member(walker, &record.data)?;
                definition.calculated_members.push(member);
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDCalcMems collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDCalcMems".to_string(),
    ))
}

/// `BrtBeginPCDCalcMem` collection (MS-XLSB 2.4.126; `PCDCalcMemCommon`, MS-XLSB 2.5.99).
fn parse_calculated_member(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
) -> XlsbResult<CalculatedMember> {
    let mut member = parse_calculated_member_payload(data)?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_PCD_CALC_MEM => return Ok(member),
            rt::BEGIN_PCD_CALC_MEM14 => {
                member.ext14 = Some(parse_calculated_member_ext14(&record.data)?);
                walker.expect_end(rt::END_PCD_CALC_MEM14, "BrtBeginPCDCalcMem14")?;
            },
            other => walker.skip_unhandled(other, "BrtBeginPCDCalcMem collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndPCDCalcMem".to_string(),
    ))
}

/// `PCDCalcMemCommon` (MS-XLSB 2.5.99).
fn parse_calculated_member_payload(data: &[u8]) -> XlsbResult<CalculatedMember> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDCalcMem");
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
fn parse_calculated_member_ext14(data: &[u8]) -> XlsbResult<CalculatedMemberExt14> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCDCalcMem14");
    // FRTBlank header (4 bytes, MS-XLSB 2.5.55).
    cursor.guard(4)?;
    cursor.offset += 4;
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
fn parse_pcd14(data: &[u8]) -> XlsbResult<PivotCacheDefinitionExt14> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginPCD14");
    // FRTBlank header (4 bytes, MS-XLSB 2.5.55).
    cursor.guard(4)?;
    cursor.offset += 4;
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
