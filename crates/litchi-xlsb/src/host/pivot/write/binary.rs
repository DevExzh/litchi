//! Binary serializer for the XLSB PivotCache definition stream (MS-XLSB 2.1.7.38).
//!
//! This is the exact inverse of `parse.rs`: record order, payload layouts,
//! flag bits, optional-field presence flags, and collection counts all
//! mirror the reader so authored cache definitions round-trip through
//! `parse_pivot_cache_definition` and, at the package level, through
//! `Workbook::pivot_cache_definitions`.
//!
//! The serializer is lossless-or-refuse: model content that cannot be
//! represented such that the reader recovers it verbatim is rejected with a
//! clear error instead of being silently dropped. Concretely, this refuses:
//!
//! - [`PivotCacheItemValue::Index`] items inside shared items, grouping
//!   items, or tuple cache entries (the reader only accepts `BrtPCDIIndex`
//!   inside discrete groupings),
//! - shared items without statistics (the `BrtBeginPCDFAtbl` record carries
//!   both the statistics and the item collection, so emitting the items
//!   would fabricate statistics the model does not have),
//! - statistics with only one of minimum/maximum set,
//! - sources and consolidation sets that carry both a named range and a
//!   cell range, or neither.

use crate::package::error::Result;
use crate::package::pivot::model::*;
use crate::package::walker::malformed;
use crate::raw::Writer;
use crate::raw::kind as rt;

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

// `BrtBeginPCDSConsol` flags word (MS-XLSB 2.4.150).
const CONSOL_AUTO_PAGE: u16 = 1 << 0;

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

/// Byte length of an `FRTBlank` header (MS-XLSB 2.5.55).
const FRT_BLANK_LEN: usize = 4;
/// `cpairs` value required by `BrtBeginPNPairs` (MS-XLSB 2.4.179).
const PN_PAIRS_DECLARED: u32 = 1;
/// Maximum number of consolidation page fields (MS-XLSB 2.4.152).
const MAX_CONSOLIDATION_PAGES: usize = 4;

fn write_wide_string(data: &mut Vec<u8>, value: &str) {
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
}

fn write_nullable_wide_string(data: &mut Vec<u8>, value: &Option<String>) {
    match value {
        Some(value) => write_wide_string(data, value),
        None => data.extend_from_slice(&u32::MAX.to_le_bytes()),
    }
}

fn write_blob(data: &mut Vec<u8>, blob: &[u8]) {
    data.extend_from_slice(&(blob.len() as u32).to_le_bytes());
    data.extend_from_slice(blob);
}

fn write_count(data: &mut Vec<u8>, count: usize, context: &'static str) -> Result<()> {
    let count =
        u32::try_from(count).map_err(|_| malformed(context, "collection count overflow"))?;
    data.extend_from_slice(&count.to_le_bytes());
    Ok(())
}

/// Inverse of the reader's `non_negative_index`: `None` becomes `-1`.
fn optional_index(value: Option<u32>, context: &'static str) -> Result<i32> {
    match value {
        Some(value) => i32::try_from(value).map_err(|_| malformed(context, "index overflow")),
        None => Ok(-1),
    }
}

/// Serialize a PivotCache definition into its complete
/// `pivotCacheDefinition` part stream.
///
/// Everything the model can hold is serialized: refresh metadata, the cache
/// source (worksheet range, consolidation, or external), cache fields with
/// shared items of every value type, range and discrete grouping, OLAP
/// hierarchies, the tuple cache, calculated items and members, and the
/// Excel 2013 extensions. Content that cannot round-trip through the reader
/// is rejected; see the module documentation for the exact refusal rules.
pub(super) fn write_pivot_cache_definition(definition: &PivotCacheDefinition) -> Result<Vec<u8>> {
    let mut data = Vec::with_capacity(512);
    let mut writer = Writer::new(&mut data);
    writer.write_record(rt::BEGIN_PIVOT_CACHE_DEF, &definition_payload(definition))?;

    if let Some(source) = &definition.source {
        write_source(&mut writer, source)?;
    }
    if !definition.fields.is_empty() {
        let mut payload = Vec::with_capacity(4);
        write_count(&mut payload, definition.fields.len(), "BrtBeginPCDFields")?;
        writer.write_record(rt::BEGIN_PCD_FIELDS, &payload)?;
        for field in &definition.fields {
            write_field(&mut writer, field)?;
        }
        writer.write_record(rt::END_PCD_FIELDS, &[])?;
    }
    if !definition.hierarchies.is_empty() {
        let mut payload = Vec::with_capacity(4);
        write_count(
            &mut payload,
            definition.hierarchies.len(),
            "BrtBeginPCDHierarchies",
        )?;
        writer.write_record(rt::BEGIN_PCD_HIERARCHIES, &payload)?;
        for hierarchy in &definition.hierarchies {
            write_hierarchy(&mut writer, hierarchy)?;
        }
        writer.write_record(rt::END_PCD_HIERARCHIES, &[])?;
    }
    if let Some(tuple_cache) = &definition.tuple_cache {
        write_tuple_cache(&mut writer, tuple_cache)?;
    }
    if !definition.calculated_items.is_empty() {
        let mut payload = Vec::with_capacity(4);
        write_count(
            &mut payload,
            definition.calculated_items.len(),
            "BrtBeginPCDCalcItems",
        )?;
        writer.write_record(rt::BEGIN_PCD_CALC_ITEMS, &payload)?;
        for item in &definition.calculated_items {
            write_calculated_item(&mut writer, item)?;
        }
        writer.write_record(rt::END_PCD_CALC_ITEMS, &[])?;
    }
    if !definition.calculated_members.is_empty() {
        let mut payload = Vec::with_capacity(4);
        write_count(
            &mut payload,
            definition.calculated_members.len(),
            "BrtBeginPCDCalcMems",
        )?;
        writer.write_record(rt::BEGIN_PCD_CALC_MEMS, &payload)?;
        for member in &definition.calculated_members {
            write_calculated_member(&mut writer, member)?;
        }
        writer.write_record(rt::END_PCD_CALC_MEMS, &[])?;
    }
    if let Some(ext14) = &definition.ext14 {
        writer.write_record(rt::BEGIN_PCD14, &pcd14_payload(ext14))?;
        writer.write_record(rt::END_PCD14, &[])?;
    }

    writer.write_record(rt::END_PIVOT_CACHE_DEF, &[])?;
    Ok(data)
}

/// `BrtBeginPivotCacheDef` payload (MS-XLSB 2.4.168).
fn definition_payload(definition: &PivotCacheDefinition) -> Vec<u8> {
    let mut data = Vec::with_capacity(32);
    data.push(definition.version_last_refresh);
    data.push(definition.version_refreshable_min);
    data.push(definition.version_created);
    let mut flags1 = 0u8;
    if definition.save_data {
        flags1 |= DEF_SAVE_DATA;
    }
    if definition.invalid {
        flags1 |= DEF_INVALID;
    }
    if definition.refresh_on_load {
        flags1 |= DEF_REFRESH_ON_LOAD;
    }
    if definition.optimize_cache {
        flags1 |= DEF_OPTIMIZE_CACHE;
    }
    if definition.enable_refresh {
        flags1 |= DEF_ENABLE_REFRESH;
    }
    if definition.background_query {
        flags1 |= DEF_BACKGROUND_QUERY;
    }
    if definition.upgrade_on_refresh {
        flags1 |= DEF_UPGRADE_ON_REFRESH;
    }
    if definition.cube_functions {
        flags1 |= DEF_CUBE_FUNCTIONS;
    }
    data.push(flags1);
    data.extend_from_slice(&definition.ghost_items_max.to_le_bytes());
    data.extend_from_slice(&definition.refreshed_date_serial.to_le_bytes());
    let mut flags2 = 0u8;
    if definition.refreshed_by.is_some() {
        flags2 |= DEF_LOAD_REFRESHED_WHO;
    }
    if definition.records_rel_id.is_some() {
        flags2 |= DEF_LOAD_REL_ID_RECORDS;
    }
    if definition.support_subquery {
        flags2 |= DEF_SUPPORT_SUBQUERY;
    }
    if definition.support_attrib_drill {
        flags2 |= DEF_SUPPORT_ATTRIB_DRILL;
    }
    data.push(flags2);
    data.extend_from_slice(&definition.record_count.to_le_bytes());
    if let Some(refreshed_by) = &definition.refreshed_by {
        write_wide_string(&mut data, refreshed_by);
    }
    if definition.records_rel_id.is_some() {
        write_nullable_wide_string(&mut data, &definition.records_rel_id);
    }
    if definition.refreshed_by.is_none() {
        // `unused` (4 bytes) exists iff fLoadRefreshedWho is 0.
        data.extend_from_slice(&[0; 4]);
    }
    data
}

/// `BrtBeginPCDSource` collection (MS-XLSB 2.4.166).
fn write_source<W: std::io::Write>(
    writer: &mut Writer<W>,
    source: &PivotCacheSource,
) -> Result<()> {
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(&(source.source_type as u32).to_le_bytes());
    payload.extend_from_slice(&source.connection_id.unwrap_or(0).to_le_bytes());
    writer.write_record(rt::BEGIN_PCD_SOURCE, &payload)?;
    if let Some(worksheet) = &source.worksheet {
        writer.write_record(rt::BEGIN_PCDS_RANGE, &worksheet_range_payload(worksheet)?)?;
        writer.write_record(rt::END_PCDS_RANGE, &[])?;
    }
    if let Some(consolidation) = &source.consolidation {
        write_consolidation(writer, consolidation)?;
    }
    writer.write_record(rt::END_PCD_SOURCE, &[])?;
    Ok(())
}

fn write_range(data: &mut Vec<u8>, range: &PivotCacheRange) {
    data.extend_from_slice(&range.first_row.to_le_bytes());
    data.extend_from_slice(&range.last_row.to_le_bytes());
    data.extend_from_slice(&range.first_column.to_le_bytes());
    data.extend_from_slice(&range.last_column.to_le_bytes());
}

/// Reject a source located both by name and by range, or by neither.
fn check_name_or_range(
    named_range: &Option<String>,
    range: &Option<PivotCacheRange>,
    context: &'static str,
) -> Result<()> {
    if named_range.is_some() == range.is_some() {
        return Err(malformed(
            context,
            "source must carry exactly one of a named range or a cell range",
        ));
    }
    Ok(())
}

/// `BrtBeginPCDSRange` payload (MS-XLSB 2.4.167).
fn worksheet_range_payload(source: &PivotCacheWorksheetSource) -> Result<Vec<u8>> {
    check_name_or_range(&source.named_range, &source.range, "BrtBeginPCDSRange")?;
    let mut data = Vec::with_capacity(24);
    data.push(if source.named_range.is_some() {
        RANGE_NAME
    } else {
        0
    });
    data.push(if source.built_in_name {
        RANGE_BUILT_IN
    } else {
        0
    });
    let mut flags2 = 0u8;
    if source.sheet_name.is_some() {
        flags2 |= RANGE_LOAD_SHEET;
    }
    if source.external_rel_id.is_some() {
        flags2 |= RANGE_LOAD_REL_ID;
    }
    data.push(flags2);
    if let Some(sheet_name) = &source.sheet_name {
        write_wide_string(&mut data, sheet_name);
    }
    if source.external_rel_id.is_some() {
        write_nullable_wide_string(&mut data, &source.external_rel_id);
    }
    match (&source.named_range, &source.range) {
        (Some(named_range), None) => write_wide_string(&mut data, named_range),
        (None, Some(range)) => write_range(&mut data, range),
        _ => unreachable!("name-or-range invariant checked above"),
    }
    Ok(data)
}

/// `BrtBeginPCDSConsol` collection (MS-XLSB 2.4.150).
fn write_consolidation<W: std::io::Write>(
    writer: &mut Writer<W>,
    consolidation: &PivotCacheConsolidationSource,
) -> Result<()> {
    if consolidation.pages.len() > MAX_CONSOLIDATION_PAGES {
        return Err(malformed(
            "BrtBeginPCDSCPages",
            format!(
                "{} consolidation pages exceed the maximum of {MAX_CONSOLIDATION_PAGES}",
                consolidation.pages.len()
            ),
        ));
    }
    let mut payload = Vec::with_capacity(2);
    payload.extend_from_slice(
        &(if consolidation.auto_page {
            CONSOL_AUTO_PAGE
        } else {
            0
        })
        .to_le_bytes(),
    );
    writer.write_record(rt::BEGIN_PCDS_CONSOL, &payload)?;

    let mut payload = Vec::with_capacity(4);
    write_count(&mut payload, consolidation.sets.len(), "BrtBeginPCDSCSets")?;
    writer.write_record(rt::BEGIN_PCDSC_SETS, &payload)?;
    for set in &consolidation.sets {
        writer.write_record(rt::BEGIN_PCDSC_SET, &consolidation_set_payload(set)?)?;
        writer.write_record(rt::END_PCDSC_SET, &[])?;
    }
    writer.write_record(rt::END_PCDSC_SETS, &[])?;

    let mut payload = Vec::with_capacity(4);
    write_count(
        &mut payload,
        consolidation.pages.len(),
        "BrtBeginPCDSCPages",
    )?;
    writer.write_record(rt::BEGIN_PCDSC_PAGES, &payload)?;
    for page in &consolidation.pages {
        writer.write_record(rt::BEGIN_PCDSC_PAGE, &[])?;
        for item_name in &page.item_names {
            let mut payload = Vec::with_capacity(item_name.len() * 2 + 4);
            write_wide_string(&mut payload, item_name);
            writer.write_record(rt::BEGIN_PCDSCP_ITEM, &payload)?;
            writer.write_record(rt::END_PCDSCP_ITEM, &[])?;
        }
        writer.write_record(rt::END_PCDSC_PAGE, &[])?;
    }
    writer.write_record(rt::END_PCDSC_PAGES, &[])?;

    writer.write_record(rt::END_PCDS_CONSOL, &[])?;
    Ok(())
}

/// `BrtBeginPCDSCSet` payload (MS-XLSB 2.4.154).
fn consolidation_set_payload(set: &PivotCacheConsolidationSet) -> Result<Vec<u8>> {
    check_name_or_range(&set.named_range, &set.range, "BrtBeginPCDSCSet")?;
    let mut data = Vec::with_capacity(32);
    for index in set.item_indexes {
        data.extend_from_slice(&index.to_le_bytes());
    }
    data.push(u8::from(set.named_range.is_some()));
    data.push(u8::from(set.built_in_name));
    let mut flags = 0u8;
    if set.sheet_name.is_some() {
        flags |= CONSOL_SET_LOAD_SHEET;
    }
    if set.external_rel_id.is_some() {
        flags |= CONSOL_SET_LOAD_REL_ID;
    }
    data.push(flags);
    if let Some(sheet_name) = &set.sheet_name {
        write_wide_string(&mut data, sheet_name);
    }
    if set.external_rel_id.is_some() {
        write_nullable_wide_string(&mut data, &set.external_rel_id);
    }
    match (&set.named_range, &set.range) {
        (Some(named_range), None) => write_wide_string(&mut data, named_range),
        (None, Some(range)) => write_range(&mut data, range),
        _ => unreachable!("name-or-range invariant checked above"),
    }
    Ok(data)
}

/// `BrtBeginPCDField` collection (MS-XLSB 2.4.136).
fn write_field<W: std::io::Write>(writer: &mut Writer<W>, field: &PivotCacheField) -> Result<()> {
    writer.write_record(rt::BEGIN_PCD_FIELD, &field_payload(field)?)?;
    if !field.shared_items.items.is_empty() || field.shared_items.stats.is_some() {
        write_shared_items(writer, &field.shared_items)?;
    }
    if let Some(grouping) = &field.grouping {
        write_grouping(writer, grouping)?;
    }
    if field.ignore {
        // FRTBlank header; the reader only notes the record's presence.
        writer.write_record(PCD_FIELD14, &[0; FRT_BLANK_LEN])?;
    }
    writer.write_record(rt::END_PCD_FIELD, &[])?;
    Ok(())
}

/// `BrtBeginPCDField` payload (MS-XLSB 2.4.136).
fn field_payload(field: &PivotCacheField) -> Result<Vec<u8>> {
    let mut data = Vec::with_capacity(32);
    let mut flags = 0u16;
    if field.server_based {
        flags |= FIELD_SERVER_BASED;
    }
    if field.cant_get_unique_items {
        flags |= FIELD_CANT_GET_UNIQUE_ITEMS;
    }
    if field.source_field {
        flags |= FIELD_SRC_FIELD;
    }
    if field.caption.is_some() {
        flags |= FIELD_CAPTION;
    }
    if field.olap_member_property_field {
        flags |= FIELD_OLAP_MEM_PROP;
    }
    if field.formula.is_some() {
        flags |= FIELD_LOAD_FMLA;
    }
    if field.member_property_name.is_some() {
        flags |= FIELD_LOAD_PROP_NAME;
    }
    data.extend_from_slice(&flags.to_le_bytes());
    data.extend_from_slice(
        &field
            .number_format
            .unwrap_or(DEFAULT_NUMBER_FORMAT)
            .to_le_bytes(),
    );
    data.extend_from_slice(&field.sql_type.to_le_bytes());
    data.extend_from_slice(&field.hierarchy_index.to_le_bytes());
    data.extend_from_slice(&field.level.to_le_bytes());
    write_count(
        &mut data,
        field.member_property_fields.len(),
        "BrtBeginPCDField",
    )?;
    write_wide_string(&mut data, &field.name);
    if let Some(caption) = &field.caption {
        write_wide_string(&mut data, caption);
    }
    if let Some(formula) = &field.formula {
        write_blob(&mut data, &formula.tokens);
        write_blob(&mut data, &formula.extra);
    }
    if !field.member_property_fields.is_empty() {
        let byte_count = u32::try_from(field.member_property_fields.len())
            .ok()
            .and_then(|count| count.checked_mul(4))
            .ok_or_else(|| malformed("BrtBeginPCDField", "member property byte count overflow"))?;
        data.extend_from_slice(&byte_count.to_le_bytes());
        for index in &field.member_property_fields {
            data.extend_from_slice(&index.to_le_bytes());
        }
    }
    if let Some(member_property_name) = &field.member_property_name {
        write_wide_string(&mut data, member_property_name);
    }
    Ok(data)
}

/// `BrtBeginPCDFAtbl` collection (MS-XLSB 2.4.131).
fn write_shared_items<W: std::io::Write>(
    writer: &mut Writer<W>,
    shared_items: &PivotCacheSharedItems,
) -> Result<()> {
    let stats = shared_items.stats.as_ref().ok_or_else(|| {
        malformed(
            "BrtBeginPCDFAtbl",
            "shared items without statistics cannot be emitted losslessly",
        )
    })?;
    writer.write_record(rt::BEGIN_PCDF_ATBL, &shared_items_stats_payload(stats)?)?;
    for item in &shared_items.items {
        write_cache_item(
            writer,
            &item.value,
            item.additional.as_ref(),
            "BrtBeginPCDFAtbl",
        )?;
    }
    writer.write_record(rt::END_PCDF_ATBL, &[])?;
    Ok(())
}

/// `BrtBeginPCDFAtbl` payload (MS-XLSB 2.4.131).
fn shared_items_stats_payload(stats: &PivotCacheSharedItemsStats) -> Result<Vec<u8>> {
    if stats.minimum.is_some() != stats.maximum.is_some() {
        return Err(malformed(
            "BrtBeginPCDFAtbl",
            "minimum and maximum must both be set or both be absent",
        ));
    }
    let mut data = Vec::with_capacity(24);
    let mut flags = 0u16;
    if stats.text_field {
        flags |= ATBL_TEXT_FIELD;
    }
    if stats.non_dates {
        flags |= ATBL_NON_DATES;
    }
    if stats.date_in_field {
        flags |= ATBL_DATE_IN_FIELD;
    }
    if stats.has_text_item {
        flags |= ATBL_HAS_TEXT_ITEM;
    }
    if stats.has_blank_item {
        flags |= ATBL_HAS_BLANK_ITEM;
    }
    if stats.mixed_types_ignoring_blanks {
        flags |= ATBL_MIXED_TYPES;
    }
    if stats.numeric_field {
        flags |= ATBL_NUM_FIELD;
    }
    if stats.integer_field {
        flags |= ATBL_INT_FIELD;
    }
    if stats.minimum.is_some() {
        flags |= ATBL_NUM_MIN_MAX_VALID;
    }
    if stats.has_long_text_item {
        flags |= ATBL_HAS_LONG_TEXT_ITEM;
    }
    data.extend_from_slice(&flags.to_le_bytes());
    data.extend_from_slice(&stats.item_count.to_le_bytes());
    if let (Some(minimum), Some(maximum)) = (stats.minimum, stats.maximum) {
        data.extend_from_slice(&minimum.to_le_bytes());
        data.extend_from_slice(&maximum.to_le_bytes());
    }
    Ok(data)
}

/// Write one cache item as a `BrtPCDI*` (no additional info) or `BrtPCDIA*`
/// (with additional info) record.
///
/// [`PivotCacheItemValue::Index`] is refused: outside a discrete grouping
/// the reader skips `BrtPCDIIndex`, so writing it would silently drop the
/// item on re-read.
fn write_cache_item<W: std::io::Write>(
    writer: &mut Writer<W>,
    value: &PivotCacheItemValue,
    additional: Option<&PivotCacheItemInfo>,
    context: &'static str,
) -> Result<()> {
    let with_info = additional.is_some();
    let record_type = match (value, with_info) {
        (PivotCacheItemValue::Missing, false) => rt::PCDI_MISSING,
        (PivotCacheItemValue::Missing, true) => rt::PCDIA_MISSING,
        (PivotCacheItemValue::Number(_), false) => rt::PCDI_NUMBER,
        (PivotCacheItemValue::Number(_), true) => rt::PCDIA_NUMBER,
        (PivotCacheItemValue::Boolean(_), false) => rt::PCDI_BOOLEAN,
        (PivotCacheItemValue::Boolean(_), true) => rt::PCDIA_BOOLEAN,
        (PivotCacheItemValue::Error(_), false) => rt::PCDI_ERROR,
        (PivotCacheItemValue::Error(_), true) => rt::PCDIA_ERROR,
        (PivotCacheItemValue::String(_), false) => rt::PCDI_STRING,
        (PivotCacheItemValue::String(_), true) => rt::PCDIA_STRING,
        (PivotCacheItemValue::DateTime(_), false) => rt::PCDI_DATETIME,
        (PivotCacheItemValue::DateTime(_), true) => rt::PCDIA_DATETIME,
        (PivotCacheItemValue::Index(_), _) => {
            return Err(malformed(
                context,
                "index items are only valid inside a discrete grouping",
            ));
        },
    };
    let mut payload = Vec::with_capacity(16);
    match value {
        PivotCacheItemValue::Missing => {},
        PivotCacheItemValue::Number(value) => payload.extend_from_slice(&value.to_le_bytes()),
        PivotCacheItemValue::Boolean(value) => payload.push(u8::from(*value)),
        PivotCacheItemValue::Error(code) => payload.push(*code as u8),
        PivotCacheItemValue::String(value) => write_wide_string(&mut payload, value),
        PivotCacheItemValue::DateTime(value) => write_date_time(&mut payload, value),
        PivotCacheItemValue::Index(_) => unreachable!("index items refused above"),
    }
    if let Some(additional) = additional {
        let mut flags = 0u16;
        if additional.ghost {
            flags |= ADDL_GHOST;
        }
        if additional.calculated {
            flags |= ADDL_CALCULATED;
        }
        if additional.caption.is_some() {
            flags |= ADDL_CAPTION;
        }
        payload.extend_from_slice(&flags.to_le_bytes());
        if additional.caption.is_some() {
            write_nullable_wide_string(&mut payload, &additional.caption);
        }
        write_count(
            &mut payload,
            additional.member_property_items.len(),
            "PCDIAddlInfo",
        )?;
        for index in &additional.member_property_items {
            payload.extend_from_slice(&index.to_le_bytes());
        }
    }
    Ok(writer.write_record(record_type, &payload)?)
}

/// `PCDIDateTime` (MS-XLSB 2.5.101).
fn write_date_time(data: &mut Vec<u8>, value: &PivotCacheDateTime) {
    data.extend_from_slice(&value.year.to_le_bytes());
    data.extend_from_slice(&value.month.to_le_bytes());
    data.push(value.day);
    data.push(value.hour);
    data.push(value.minute);
    data.push(value.second);
}

/// `BrtBeginPCDFGroup` collection (MS-XLSB 2.4.135).
fn write_grouping<W: std::io::Write>(
    writer: &mut Writer<W>,
    grouping: &PivotCacheFieldGrouping,
) -> Result<()> {
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(
        &optional_index(grouping.parent_field, "BrtBeginPCDFGroup")?.to_le_bytes(),
    );
    payload.extend_from_slice(
        &optional_index(grouping.base_field, "BrtBeginPCDFGroup")?.to_le_bytes(),
    );
    writer.write_record(rt::BEGIN_PCDF_GROUP, &payload)?;

    if let Some(range) = &grouping.range {
        let mut payload = Vec::with_capacity(26);
        payload.push(range.group_by as u8);
        let mut flags = 0u8;
        if range.auto_start {
            flags |= GROUP_RANGE_AUTO_START;
        }
        if range.auto_end {
            flags |= GROUP_RANGE_AUTO_END;
        }
        if range.dates {
            flags |= GROUP_RANGE_DATES;
        }
        payload.push(flags);
        payload.extend_from_slice(&range.start.to_le_bytes());
        payload.extend_from_slice(&range.end.to_le_bytes());
        payload.extend_from_slice(&range.interval.to_le_bytes());
        writer.write_record(rt::BEGIN_PCDFG_RANGE, &payload)?;
        writer.write_record(rt::END_PCDFG_RANGE, &[])?;
    }
    if let Some(discrete) = &grouping.discrete {
        writer.write_record(rt::BEGIN_PCDFG_DISCRETE, &[])?;
        for index in &discrete.item_indexes {
            writer.write_record(rt::PCDI_INDEX, &index.to_le_bytes())?;
        }
        writer.write_record(rt::END_PCDFG_DISCRETE, &[])?;
    }
    if !grouping.items.is_empty() {
        writer.write_record(rt::BEGIN_PCDFG_ITEMS, &[])?;
        for item in &grouping.items {
            write_cache_item(
                writer,
                &item.value,
                item.additional.as_ref(),
                "BrtBeginPCDFGItems",
            )?;
        }
        writer.write_record(rt::END_PCDFG_ITEMS, &[])?;
    }
    writer.write_record(rt::END_PCDF_GROUP, &[])?;
    Ok(())
}

/// `BrtBeginPCDHierarchy` collection (MS-XLSB 2.4.146).
fn write_hierarchy<W: std::io::Write>(
    writer: &mut Writer<W>,
    hierarchy: &PivotCacheHierarchy,
) -> Result<()> {
    writer.write_record(rt::BEGIN_PCD_HIERARCHY, &hierarchy_payload(hierarchy)?)?;

    if !hierarchy.field_usage.is_empty() {
        let mut payload = Vec::with_capacity(4 + hierarchy.field_usage.len() * 4);
        write_count(
            &mut payload,
            hierarchy.field_usage.len(),
            "BrtBeginPCDHFieldsUsage",
        )?;
        for index in &hierarchy.field_usage {
            payload.extend_from_slice(&index.to_le_bytes());
        }
        writer.write_record(rt::BEGIN_PCDH_FIELDS_USAGE, &payload)?;
        writer.write_record(rt::END_PCDH_FIELDS_USAGE, &[])?;
    }
    if !hierarchy.grouping_levels.is_empty() {
        let mut payload = Vec::with_capacity(4);
        write_count(
            &mut payload,
            hierarchy.grouping_levels.len(),
            "BrtBeginPCDHGLevels",
        )?;
        writer.write_record(rt::BEGIN_PCDHG_LEVELS, &payload)?;
        for level in &hierarchy.grouping_levels {
            let mut payload = Vec::with_capacity(16);
            payload.push(if level.group_level {
                GROUPING_LEVEL_GROUP
            } else {
                0
            });
            write_wide_string(&mut payload, &level.unique_name);
            write_wide_string(&mut payload, &level.caption);
            writer.write_record(rt::BEGIN_PCDHG_LEVEL, &payload)?;
            writer.write_record(rt::END_PCDHG_LEVEL, &[])?;
        }
        writer.write_record(rt::END_PCDHG_LEVELS, &[])?;
    }
    if !hierarchy.grouping_groups.is_empty() {
        let mut payload = Vec::with_capacity(4);
        write_count(
            &mut payload,
            hierarchy.grouping_groups.len(),
            "BrtBeginPCDHGLGroups",
        )?;
        writer.write_record(rt::BEGIN_PCDHGL_GROUPS, &payload)?;
        for group in &hierarchy.grouping_groups {
            write_grouping_group(writer, group)?;
        }
        writer.write_record(rt::END_PCDHGL_GROUPS, &[])?;
    }
    if let Some(ext14) = &hierarchy.ext14 {
        let mut payload = vec![0; FRT_BLANK_LEN];
        let mut flags = 0u8;
        if ext14.flatten_hierarchies {
            flags |= HIER14_FLATTEN;
        }
        if ext14.measure_set {
            flags |= HIER14_MEASURE_SET;
        }
        if ext14.hierarchize_distinct {
            flags |= HIER14_HIERARCHIZE_DISTINCT;
        }
        if ext14.ignorable {
            flags |= HIER14_IGNORABLE;
        }
        payload.push(flags);
        write_count(&mut payload, ext14.hierarchy_indexes.len(), "BrtPCDH14")?;
        for index in &ext14.hierarchy_indexes {
            payload.extend_from_slice(&index.to_le_bytes());
        }
        writer.write_record(PCD_H14, &payload)?;
    }
    writer.write_record(rt::END_PCD_HIERARCHY, &[])?;
    Ok(())
}

/// `BrtBeginPCDHierarchy` payload (MS-XLSB 2.4.146).
fn hierarchy_payload(hierarchy: &PivotCacheHierarchy) -> Result<Vec<u8>> {
    let mut data = Vec::with_capacity(32);
    let mut flags1 = 0u16;
    if hierarchy.measure {
        flags1 |= HIER_MEASURE;
    }
    if hierarchy.set {
        flags1 |= HIER_SET;
    }
    if hierarchy.attribute_hierarchy {
        flags1 |= HIER_ATTRIBUTE;
    }
    if hierarchy.measure_hierarchy {
        flags1 |= HIER_MEASURE_HIERARCHY;
    }
    if hierarchy.only_one_field {
        flags1 |= HIER_ONLY_ONE_FIELD;
    }
    if hierarchy.time_hierarchy {
        flags1 |= HIER_TIME;
    }
    if hierarchy.key_attribute_hierarchy {
        flags1 |= HIER_KEY_ATTRIBUTE;
    }
    if hierarchy.attribute_member_value_type.is_some() {
        flags1 |= HIER_VALUE_TYPE_KNOWN;
    }
    if let Some(unbalanced_real) = hierarchy.unbalanced_real {
        flags1 |= HIER_UNBALANCED_REAL_KNOWN;
        if unbalanced_real {
            flags1 |= HIER_UNBALANCED_REAL;
        }
    }
    if let Some(unbalanced_group) = hierarchy.unbalanced_group {
        flags1 |= HIER_UNBALANCED_GROUP_KNOWN;
        if unbalanced_group {
            flags1 |= HIER_UNBALANCED_GROUP;
        }
    }
    if hierarchy.hidden {
        flags1 |= HIER_HIDDEN;
    }
    data.extend_from_slice(&flags1.to_le_bytes());
    data.extend_from_slice(&hierarchy.level_count.to_le_bytes());
    data.extend_from_slice(
        &optional_index(hierarchy.set_parent_index, "BrtBeginPCDHierarchy")?.to_le_bytes(),
    );
    data.extend_from_slice(&hierarchy.icon_set.to_le_bytes());
    let mut flags2 = 0u8;
    if hierarchy.dimension_unique_name.is_some() {
        flags2 |= HIER_LOAD_DIM_UNQ;
    }
    if hierarchy.default_member_unique_name.is_some() {
        flags2 |= HIER_LOAD_DEFAULT_UNQ;
    }
    if hierarchy.all_member_unique_name.is_some() {
        flags2 |= HIER_LOAD_ALL_UNQ;
    }
    if hierarchy.all_member_display.is_some() {
        flags2 |= HIER_LOAD_ALL_DISP;
    }
    if hierarchy.display_folder.is_some() {
        flags2 |= HIER_LOAD_DISP_FLD;
    }
    if hierarchy.measure_group.is_some() {
        flags2 |= HIER_LOAD_MEAS_GRP;
    }
    data.push(flags2);
    data.extend_from_slice(
        &hierarchy
            .attribute_member_value_type
            .unwrap_or(0)
            .to_le_bytes(),
    );
    write_wide_string(&mut data, &hierarchy.unique_name);
    write_wide_string(&mut data, &hierarchy.caption);
    if let Some(value) = &hierarchy.dimension_unique_name {
        write_wide_string(&mut data, value);
    }
    if let Some(value) = &hierarchy.default_member_unique_name {
        write_wide_string(&mut data, value);
    }
    if let Some(value) = &hierarchy.all_member_unique_name {
        write_wide_string(&mut data, value);
    }
    if let Some(value) = &hierarchy.all_member_display {
        write_wide_string(&mut data, value);
    }
    if let Some(value) = &hierarchy.display_folder {
        write_wide_string(&mut data, value);
    }
    if let Some(value) = &hierarchy.measure_group {
        write_wide_string(&mut data, value);
    }
    Ok(data)
}

/// `BrtBeginPCDHGLGroup` collection (MS-XLSB 2.4.143).
fn write_grouping_group<W: std::io::Write>(
    writer: &mut Writer<W>,
    group: &PivotCacheGroupingGroup,
) -> Result<()> {
    let mut payload = Vec::with_capacity(24);
    payload.extend_from_slice(&group.group_number.to_le_bytes());
    payload.push(if group.parent_unique_name.is_some() {
        GROUPING_GROUP_LOAD_PARENT
    } else {
        0
    });
    write_wide_string(&mut payload, &group.name);
    write_wide_string(&mut payload, &group.unique_name);
    write_wide_string(&mut payload, &group.caption);
    if let Some(parent_unique_name) = &group.parent_unique_name {
        write_wide_string(&mut payload, parent_unique_name);
    }
    writer.write_record(rt::BEGIN_PCDHGL_GROUP, &payload)?;

    let mut payload = Vec::with_capacity(4);
    write_count(&mut payload, group.members.len(), "BrtBeginPCDHGLGMembers")?;
    writer.write_record(rt::BEGIN_PCDHGLG_MEMBERS, &payload)?;
    for member in &group.members {
        let mut payload = Vec::with_capacity(8);
        payload.extend_from_slice(&u32::from(member.is_group).to_le_bytes());
        write_wide_string(&mut payload, &member.unique_name);
        writer.write_record(rt::BEGIN_PCDHGLG_MEMBER, &payload)?;
        writer.write_record(rt::END_PCDHGLG_MEMBER, &[])?;
    }
    writer.write_record(rt::END_PCDHGLG_MEMBERS, &[])?;

    writer.write_record(rt::END_PCDHGL_GROUP, &[])?;
    Ok(())
}

/// `BrtBeginPCDSDTupleCache` collection (MS-XLSB 2.4.164).
fn write_tuple_cache<W: std::io::Write>(
    writer: &mut Writer<W>,
    cache: &PivotCacheTupleCache,
) -> Result<()> {
    writer.write_record(rt::BEGIN_PCDSD_TUPLE_CACHE, &[])?;

    let mut payload = Vec::with_capacity(4);
    write_count(&mut payload, cache.entries.len(), "BrtBeginPCDSDTCEntries")?;
    writer.write_record(rt::BEGIN_PCDSDTC_ENTRIES, &payload)?;
    for value in &cache.entries {
        // Tuple cache entries carry no additional info; the reader parses
        // them non-strictly from plain `BrtPCDI*` records.
        write_cache_item(writer, value, None, "BrtBeginPCDSDTCEntries")?;
    }
    writer.write_record(rt::END_PCDSDTC_ENTRIES, &[])?;

    let mut payload = Vec::with_capacity(4);
    write_count(&mut payload, cache.queries.len(), "BrtBeginPCDSDTCQueries")?;
    writer.write_record(rt::BEGIN_PCDSDTC_QUERIES, &payload)?;
    for query in &cache.queries {
        let mut payload = Vec::with_capacity(query.len() * 2 + 4);
        write_wide_string(&mut payload, query);
        writer.write_record(rt::BEGIN_PCDSDTC_QUERY, &payload)?;
        writer.write_record(rt::END_PCDSDTC_QUERY, &[])?;
    }
    writer.write_record(rt::END_PCDSDTC_QUERIES, &[])?;

    let mut payload = Vec::with_capacity(4);
    write_count(&mut payload, cache.sets.len(), "BrtBeginPCDSDTCSets")?;
    writer.write_record(rt::BEGIN_PCDSDTC_SETS, &payload)?;
    for set in &cache.sets {
        let mut payload = Vec::with_capacity(16);
        payload.extend_from_slice(&set.tuple_count.unwrap_or(TUPLE_COUNT_UNKNOWN).to_le_bytes());
        payload.extend_from_slice(&set.max_rank.to_le_bytes());
        payload.extend_from_slice(&set.sort_order.to_le_bytes());
        payload.push(if set.query_failed {
            TUPLE_SET_QUERY_FAILED
        } else {
            0
        });
        write_wide_string(&mut payload, &set.definition);
        writer.write_record(rt::BEGIN_PCDSDTC_SET, &payload)?;
        writer.write_record(rt::END_PCDSDTC_SET, &[])?;
    }
    writer.write_record(rt::END_PCDSDTC_SETS, &[])?;

    writer.write_record(rt::END_PCDSD_TUPLE_CACHE, &[])?;
    Ok(())
}

/// `BrtBeginPCDCalcItem` collection (MS-XLSB 2.4.124).
fn write_calculated_item<W: std::io::Write>(
    writer: &mut Writer<W>,
    item: &CalculatedItem,
) -> Result<()> {
    let mut payload = Vec::with_capacity(16);
    // reserved (4 bytes): MUST be -1.
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    write_blob(&mut payload, &item.formula.tokens);
    write_blob(&mut payload, &item.formula.extra);
    writer.write_record(rt::BEGIN_PCD_CALC_ITEM, &payload)?;

    if !item.names.is_empty() {
        let mut payload = Vec::with_capacity(4);
        write_count(&mut payload, item.names.len(), "BrtBeginPNames")?;
        writer.write_record(rt::BEGIN_P_NAMES, &payload)?;
        for name in &item.names {
            write_name(writer, name)?;
        }
        writer.write_record(rt::END_P_NAMES, &[])?;
    }
    if !item.filters.is_empty() {
        let mut payload = Vec::with_capacity(4);
        write_count(&mut payload, item.filters.len(), "BrtBeginPRFilters")?;
        writer.write_record(rt::BEGIN_PR_FILTERS, &payload)?;
        for filter in &item.filters {
            write_rule_filter(writer, filter)?;
        }
        writer.write_record(rt::END_PR_FILTERS, &[])?;
    }
    writer.write_record(rt::END_PCD_CALC_ITEM, &[])?;
    Ok(())
}

/// `BrtBeginPName` collection (MS-XLSB 2.4.176).
fn write_name<W: std::io::Write>(writer: &mut Writer<W>, name: &PivotName) -> Result<()> {
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(&name.field_index.to_le_bytes());
    payload.push(name.function as u8);
    payload.push(if name.err_name { PNAME_ERR_NAME } else { 0 });
    payload.extend_from_slice(&[0; 2]); // padding
    writer.write_record(rt::BEGIN_P_NAME, &payload)?;

    let mut payload = Vec::with_capacity(4);
    payload.extend_from_slice(&PN_PAIRS_DECLARED.to_le_bytes());
    writer.write_record(rt::BEGIN_PN_PAIRS, &payload)?;
    for pair in &name.pairs {
        let mut payload = Vec::with_capacity(12);
        let mut flags = 0u8;
        if pair.physical {
            flags |= PNPAIR_PHYSICAL;
        }
        if pair.relative {
            flags |= PNPAIR_RELATIVE;
        }
        payload.push(flags);
        payload.extend_from_slice(&pair.field_index.to_le_bytes());
        payload.extend_from_slice(&pair.item_index.to_le_bytes());
        payload.extend_from_slice(&[0; 3]); // padding
        writer.write_record(rt::BEGIN_PN_PAIR, &payload)?;
        writer.write_record(rt::END_PN_PAIR, &[])?;
    }
    writer.write_record(rt::END_PN_PAIRS, &[])?;

    writer.write_record(rt::END_P_NAME, &[])?;
    Ok(())
}

/// `BrtBeginPRFilter` collection (MS-XLSB 2.4.180; `PRFilter` structure).
fn write_rule_filter<W: std::io::Write>(
    writer: &mut Writer<W>,
    filter: &PivotRuleFilter,
) -> Result<()> {
    if filter.item_types & !PR_FILTER_ITEM_TYPES_MASK != 0 {
        return Err(malformed(
            "BrtBeginPRFilter",
            format!(
                "item types 0x{:X} exceed the 13-bit mask",
                filter.item_types
            ),
        ));
    }
    let mut payload = Vec::with_capacity(11);
    payload.extend_from_slice(&filter.field.to_le_bytes());
    write_count(&mut payload, filter.items.len(), "BrtBeginPRFilter")?;
    let mut flags = filter.item_types & PR_FILTER_ITEM_TYPES_MASK;
    if filter.selected {
        flags |= PR_FILTER_SELECTED;
    }
    payload.push(flags as u8);
    payload.push((flags >> 8) as u8);
    payload.push((flags >> 16) as u8);
    writer.write_record(rt::BEGIN_PR_FILTER, &payload)?;
    for item in &filter.items {
        writer.write_record(rt::BEGIN_PRF_ITEM, &item.to_le_bytes())?;
        writer.write_record(rt::END_PRF_ITEM, &[])?;
    }
    writer.write_record(rt::END_PR_FILTER, &[])?;
    Ok(())
}

/// `BrtBeginPCDCalcMem` collection (MS-XLSB 2.4.126; `PCDCalcMemCommon`,
/// MS-XLSB 2.5.99).
fn write_calculated_member<W: std::io::Write>(
    writer: &mut Writer<W>,
    member: &CalculatedMember,
) -> Result<()> {
    let mut payload = Vec::with_capacity(24);
    let mut flags = 0u32;
    if member.member_name.is_some() {
        flags |= CALC_MEM_LOAD_MEMBER_NAME;
    }
    if member.source_hierarchy.is_some() {
        flags |= CALC_MEM_LOAD_SOURCE_HIER;
    }
    if member.parent_unique.is_some() {
        flags |= CALC_MEM_LOAD_PARENT_UNIQUE;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&member.solve_order.to_le_bytes());
    payload.extend_from_slice(&u32::from(member.is_set).to_le_bytes());
    write_wide_string(&mut payload, &member.name);
    write_wide_string(&mut payload, &member.mdx);
    if let Some(value) = &member.member_name {
        write_wide_string(&mut payload, value);
    }
    if let Some(value) = &member.source_hierarchy {
        write_wide_string(&mut payload, value);
    }
    if let Some(value) = &member.parent_unique {
        write_wide_string(&mut payload, value);
    }
    writer.write_record(rt::BEGIN_PCD_CALC_MEM, &payload)?;

    if let Some(ext14) = &member.ext14 {
        let mut payload = vec![0; FRT_BLANK_LEN];
        let mut flags = 0u8;
        if ext14.flatten_hierarchies {
            flags |= CALC_MEM14_FLATTEN;
        }
        if ext14.dynamic_set {
            flags |= CALC_MEM14_DYNAMIC_SET;
        }
        if ext14.hierarchize_distinct {
            flags |= CALC_MEM14_HIERARCHIZE_DISTINCT;
        }
        payload.push(flags);
        write_wide_string(&mut payload, &ext14.display_folder);
        if let Some(long_mdx) = &ext14.long_mdx {
            write_wide_string(&mut payload, long_mdx);
        }
        writer.write_record(rt::BEGIN_PCD_CALC_MEM14, &payload)?;
        writer.write_record(rt::END_PCD_CALC_MEM14, &[])?;
    }
    writer.write_record(rt::END_PCD_CALC_MEM, &[])?;
    Ok(())
}

/// `BrtBeginPCD14` payload (MS-XLSB 2.4.123).
fn pcd14_payload(ext14: &PivotCacheDefinitionExt14) -> Vec<u8> {
    let mut payload = vec![0; FRT_BLANK_LEN];
    let mut flags = 0u8;
    if ext14.slicer_data {
        flags |= PCD14_SLICER_DATA;
    }
    if ext14.server_support_subquery_calc_mem {
        flags |= PCD14_SUBQUERY_CALC_MEM;
    }
    if ext14.server_support_subquery_non_visual {
        flags |= PCD14_SUBQUERY_NON_VISUAL;
    }
    if ext14.server_support_add_calc_mems {
        flags |= PCD14_ADD_CALC_MEMS;
    }
    payload.push(flags);
    payload.extend_from_slice(&ext14.cache_id.to_le_bytes());
    payload
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::error::Error;
    use crate::package::pivot::parse_pivot_cache_definition;
    use crate::writer::{MutableWorksheet, WorkbookWriter};
    use std::io::Cursor;

    fn field(name: &str) -> PivotCacheField {
        PivotCacheField {
            name: name.to_string(),
            caption: None,
            number_format: None,
            sql_type: 0,
            hierarchy_index: 0,
            level: 0x7FFF,
            member_property_fields: Vec::new(),
            member_property_name: None,
            formula: None,
            server_based: false,
            cant_get_unique_items: false,
            source_field: true,
            olap_member_property_field: false,
            ignore: false,
            shared_items: PivotCacheSharedItems::default(),
            grouping: None,
        }
    }

    fn stats() -> PivotCacheSharedItemsStats {
        PivotCacheSharedItemsStats {
            text_field: true,
            non_dates: false,
            date_in_field: false,
            has_text_item: true,
            has_blank_item: true,
            mixed_types_ignoring_blanks: true,
            numeric_field: false,
            integer_field: false,
            has_long_text_item: false,
            item_count: 7,
            minimum: None,
            maximum: None,
        }
    }

    /// A field whose shared items cover every value type, both plain and
    /// with additional info.
    fn full_item_field() -> PivotCacheField {
        let mut region = field("Region");
        region.caption = Some("Sales Region".to_string());
        region.cant_get_unique_items = true;
        region.shared_items = PivotCacheSharedItems {
            stats: Some(stats()),
            items: vec![
                PivotCacheItem {
                    value: PivotCacheItemValue::Missing,
                    additional: None,
                },
                PivotCacheItem {
                    value: PivotCacheItemValue::String("North".into()),
                    additional: None,
                },
                PivotCacheItem {
                    value: PivotCacheItemValue::String("South".into()),
                    additional: Some(PivotCacheItemInfo {
                        ghost: true,
                        calculated: false,
                        caption: Some("SOUTH!".to_string()),
                        member_property_items: vec![1, -1],
                    }),
                },
                PivotCacheItem {
                    value: PivotCacheItemValue::Number(42.5),
                    additional: None,
                },
                PivotCacheItem {
                    value: PivotCacheItemValue::Boolean(true),
                    additional: Some(PivotCacheItemInfo {
                        ghost: false,
                        calculated: true,
                        caption: None,
                        member_property_items: Vec::new(),
                    }),
                },
                PivotCacheItem {
                    value: PivotCacheItemValue::Error(PivotCacheErrorCode::NA),
                    additional: None,
                },
                PivotCacheItem {
                    value: PivotCacheItemValue::DateTime(PivotCacheDateTime {
                        year: 2024,
                        month: 3,
                        day: 14,
                        hour: 9,
                        minute: 30,
                        second: 0,
                    }),
                    additional: None,
                },
            ],
        };
        region
    }

    fn range_grouped_field() -> PivotCacheField {
        let mut amount = field("Amount");
        amount.number_format = Some(44);
        amount.grouping = Some(PivotCacheFieldGrouping {
            parent_field: None,
            base_field: Some(1),
            range: Some(PivotCacheRangeGrouping {
                group_by: PivotCacheGroupBy::Days,
                auto_start: true,
                auto_end: false,
                dates: true,
                start: 45_000.0,
                end: 46_000.0,
                interval: 7.0,
            }),
            discrete: None,
            items: Vec::new(),
        });
        amount
    }

    fn discrete_grouped_field() -> PivotCacheField {
        let mut group = field("RegionGroup");
        group.formula = Some(PivotParsedFormulaData {
            tokens: vec![0x1E, 0x02],
            extra: vec![0xAA],
        });
        group.member_property_fields = vec![2, 3];
        group.member_property_name = Some("Prop".to_string());
        group.ignore = true;
        group.grouping = Some(PivotCacheFieldGrouping {
            parent_field: Some(0),
            base_field: Some(0),
            range: None,
            discrete: Some(PivotCacheDiscreteGrouping {
                item_indexes: vec![1, 3, 5],
            }),
            items: vec![PivotCacheItem {
                value: PivotCacheItemValue::String("Grouped".into()),
                additional: None,
            }],
        });
        group
    }

    fn hierarchy() -> PivotCacheHierarchy {
        PivotCacheHierarchy {
            unique_name: "[Region]".to_string(),
            caption: "Region".to_string(),
            dimension_unique_name: Some("[RegionDim]".to_string()),
            default_member_unique_name: Some("[Region].[All]".to_string()),
            all_member_unique_name: None,
            all_member_display: Some("All Regions".to_string()),
            display_folder: None,
            measure_group: Some("MG".to_string()),
            measure: false,
            set: false,
            attribute_hierarchy: true,
            measure_hierarchy: false,
            only_one_field: true,
            time_hierarchy: false,
            key_attribute_hierarchy: true,
            hidden: false,
            unbalanced_real: Some(false),
            unbalanced_group: None,
            attribute_member_value_type: Some(0x0007),
            level_count: 2,
            set_parent_index: None,
            icon_set: -1,
            field_usage: vec![0, 1],
            grouping_levels: vec![PivotCacheGroupingLevel {
                group_level: true,
                unique_name: "[Region].[Custom]".to_string(),
                caption: "Custom".to_string(),
            }],
            grouping_groups: vec![PivotCacheGroupingGroup {
                group_number: 1,
                name: "Group1".to_string(),
                unique_name: "[Region].[Group1]".to_string(),
                caption: "Group 1".to_string(),
                parent_unique_name: Some("[Region].[All]".to_string()),
                members: vec![
                    PivotCacheGroupingGroupMember {
                        is_group: false,
                        unique_name: "[Region].[North]".to_string(),
                    },
                    PivotCacheGroupingGroupMember {
                        is_group: true,
                        unique_name: "[Region].[South]".to_string(),
                    },
                ],
            }],
            ext14: Some(PivotCacheHierarchyExt14 {
                flatten_hierarchies: true,
                measure_set: false,
                hierarchize_distinct: true,
                ignorable: false,
                hierarchy_indexes: vec![0, -2],
            }),
        }
    }

    fn full_definition() -> PivotCacheDefinition {
        PivotCacheDefinition {
            version_last_refresh: 3,
            version_refreshable_min: 0,
            version_created: 2,
            save_data: true,
            invalid: false,
            refresh_on_load: true,
            optimize_cache: false,
            enable_refresh: true,
            background_query: true,
            upgrade_on_refresh: false,
            cube_functions: true,
            support_subquery: true,
            support_attrib_drill: true,
            ghost_items_max: -1,
            refreshed_date_serial: 44_000.5,
            record_count: 5,
            refreshed_by: Some("analyst".to_string()),
            records_rel_id: Some("rIdRecords".to_string()),
            source: Some(PivotCacheSource {
                source_type: PivotCacheSourceType::Worksheet,
                connection_id: None,
                worksheet: Some(PivotCacheWorksheetSource {
                    named_range: None,
                    built_in_name: false,
                    sheet_name: Some("Data Sheet".to_string()),
                    external_rel_id: None,
                    range: Some(PivotCacheRange {
                        first_row: 0,
                        last_row: 99,
                        first_column: 1,
                        last_column: 7,
                    }),
                }),
                consolidation: None,
            }),
            fields: vec![
                full_item_field(),
                range_grouped_field(),
                discrete_grouped_field(),
            ],
            hierarchies: vec![hierarchy()],
            tuple_cache: Some(PivotCacheTupleCache {
                entries: vec![
                    PivotCacheItemValue::Missing,
                    PivotCacheItemValue::Number(1.5),
                    PivotCacheItemValue::String("cube".into()),
                    PivotCacheItemValue::Boolean(false),
                    PivotCacheItemValue::Error(PivotCacheErrorCode::Div0),
                    PivotCacheItemValue::DateTime(PivotCacheDateTime {
                        year: 2023,
                        month: 12,
                        day: 31,
                        hour: 23,
                        minute: 59,
                        second: 59,
                    }),
                ],
                queries: vec!["SELECT {} ON 0".to_string()],
                sets: vec![
                    PivotCacheTupleCacheSet {
                        tuple_count: Some(4),
                        max_rank: 2,
                        sort_order: 1,
                        query_failed: false,
                        definition: "{[Region].Members}".to_string(),
                    },
                    PivotCacheTupleCacheSet {
                        tuple_count: None,
                        max_rank: 0,
                        sort_order: 0,
                        query_failed: true,
                        definition: "{}".to_string(),
                    },
                ],
            }),
            calculated_items: vec![CalculatedItem {
                formula: PivotParsedFormulaData {
                    tokens: vec![0x03, 0x04],
                    extra: Vec::new(),
                },
                names: vec![PivotName {
                    field_index: 0,
                    function: PivotNameFunction::Sum,
                    err_name: false,
                    pairs: vec![PivotNamePair {
                        physical: true,
                        relative: false,
                        field_index: 0,
                        item_index: 2,
                    }],
                }],
                filters: vec![PivotRuleFilter {
                    field: -2,
                    item_types: 0x1F,
                    selected: true,
                    items: vec![0, 2],
                }],
            }],
            calculated_members: vec![CalculatedMember {
                name: "[Measures].[Calc]".to_string(),
                mdx: "1+1".to_string(),
                solve_order: 5,
                is_set: true,
                member_name: Some("Calc".to_string()),
                source_hierarchy: Some("[Measures]".to_string()),
                parent_unique: Some("[Measures].[All]".to_string()),
                ext14: Some(CalculatedMemberExt14 {
                    flatten_hierarchies: false,
                    dynamic_set: true,
                    hierarchize_distinct: true,
                    display_folder: "Folder".to_string(),
                    long_mdx: Some("1+1 /* long */".to_string()),
                }),
            }],
            ext14: Some(PivotCacheDefinitionExt14 {
                slicer_data: true,
                server_support_subquery_calc_mem: true,
                server_support_subquery_non_visual: false,
                server_support_add_calc_mems: true,
                cache_id: 12,
            }),
        }
    }

    fn round_trip(definition: &PivotCacheDefinition) -> PivotCacheDefinition {
        let bytes = write_pivot_cache_definition(definition).unwrap();
        parse_pivot_cache_definition(&bytes).unwrap()
    }

    #[test]
    fn serialized_full_definition_round_trips_through_the_reader() {
        let definition = full_definition();
        assert_eq!(round_trip(&definition), definition);
    }

    #[test]
    fn serialized_minimal_definition_round_trips() {
        let definition = PivotCacheDefinition::default();
        assert_eq!(round_trip(&definition), definition);
    }

    #[test]
    fn serialized_consolidation_source_round_trips() {
        let definition = PivotCacheDefinition {
            source: Some(PivotCacheSource {
                source_type: PivotCacheSourceType::Consolidation,
                connection_id: None,
                worksheet: None,
                consolidation: Some(PivotCacheConsolidationSource {
                    auto_page: true,
                    sets: vec![
                        PivotCacheConsolidationSet {
                            item_indexes: [1, u32::MAX, u32::MAX, u32::MAX],
                            named_range: Some("MyRange".to_string()),
                            built_in_name: true,
                            sheet_name: None,
                            external_rel_id: None,
                            range: None,
                        },
                        PivotCacheConsolidationSet {
                            item_indexes: [0, 1, u32::MAX, u32::MAX],
                            named_range: None,
                            built_in_name: false,
                            sheet_name: Some("Q1".to_string()),
                            external_rel_id: Some("rIdExt".to_string()),
                            range: Some(PivotCacheRange {
                                first_row: 4,
                                last_row: 20,
                                first_column: 0,
                                last_column: 3,
                            }),
                        },
                    ],
                    pages: vec![
                        PivotCacheConsolidationPage {
                            item_names: vec!["Region1".to_string(), "Region2".to_string()],
                        },
                        PivotCacheConsolidationPage {
                            item_names: Vec::new(),
                        },
                    ],
                }),
            }),
            ..PivotCacheDefinition::default()
        };
        assert_eq!(round_trip(&definition), definition);
    }

    #[test]
    fn serialized_external_source_and_named_range_round_trips() {
        let definition = PivotCacheDefinition {
            source: Some(PivotCacheSource {
                source_type: PivotCacheSourceType::External,
                connection_id: Some(17),
                worksheet: Some(PivotCacheWorksheetSource {
                    named_range: Some("ExternalData".to_string()),
                    built_in_name: false,
                    sheet_name: None,
                    external_rel_id: Some("rIdBook".to_string()),
                    range: None,
                }),
                consolidation: None,
            }),
            ..PivotCacheDefinition::default()
        };
        assert_eq!(round_trip(&definition), definition);
    }

    #[test]
    fn refuses_content_that_cannot_round_trip() {
        // Index items inside shared items are skipped by the reader.
        let mut definition = PivotCacheDefinition::default();
        let mut broken = field("F");
        broken.shared_items = PivotCacheSharedItems {
            stats: Some(stats()),
            items: vec![PivotCacheItem {
                value: PivotCacheItemValue::Index(3),
                additional: None,
            }],
        };
        definition.fields = vec![broken];
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));

        // Shared items without statistics would fabricate an ATBL payload.
        let mut definition = PivotCacheDefinition::default();
        let mut broken = field("F");
        broken.shared_items = PivotCacheSharedItems {
            stats: None,
            items: vec![PivotCacheItem {
                value: PivotCacheItemValue::Number(1.0),
                additional: None,
            }],
        };
        definition.fields = vec![broken];
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));

        // Statistics with only one bound set.
        let mut definition = PivotCacheDefinition::default();
        let mut broken = field("F");
        broken.shared_items = PivotCacheSharedItems {
            stats: Some(PivotCacheSharedItemsStats {
                minimum: Some(1.0),
                ..stats()
            }),
            items: Vec::new(),
        };
        definition.fields = vec![broken];
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));

        // Index items inside grouping items.
        let mut definition = PivotCacheDefinition::default();
        let mut broken = field("F");
        broken.grouping = Some(PivotCacheFieldGrouping {
            parent_field: None,
            base_field: None,
            range: None,
            discrete: None,
            items: vec![PivotCacheItem {
                value: PivotCacheItemValue::Index(0),
                additional: None,
            }],
        });
        definition.fields = vec![broken];
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));

        // Index values inside tuple cache entries.
        let definition = PivotCacheDefinition {
            tuple_cache: Some(PivotCacheTupleCache {
                entries: vec![PivotCacheItemValue::Index(1)],
                ..PivotCacheTupleCache::default()
            }),
            ..PivotCacheDefinition::default()
        };
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));
    }

    #[test]
    fn refuses_ambiguous_or_empty_sources() {
        // Both a named range and a cell range.
        let definition = PivotCacheDefinition {
            source: Some(PivotCacheSource {
                source_type: PivotCacheSourceType::Worksheet,
                connection_id: None,
                worksheet: Some(PivotCacheWorksheetSource {
                    named_range: Some("R".to_string()),
                    built_in_name: false,
                    sheet_name: None,
                    external_rel_id: None,
                    range: Some(PivotCacheRange {
                        first_row: 0,
                        last_row: 1,
                        first_column: 0,
                        last_column: 0,
                    }),
                }),
                consolidation: None,
            }),
            ..PivotCacheDefinition::default()
        };
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));

        // Neither a named range nor a cell range.
        let definition = PivotCacheDefinition {
            source: Some(PivotCacheSource {
                source_type: PivotCacheSourceType::Worksheet,
                connection_id: None,
                worksheet: Some(PivotCacheWorksheetSource {
                    named_range: None,
                    built_in_name: false,
                    sheet_name: None,
                    external_rel_id: None,
                    range: None,
                }),
                consolidation: None,
            }),
            ..PivotCacheDefinition::default()
        };
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));

        // A consolidation set with no locator.
        let definition = PivotCacheDefinition {
            source: Some(PivotCacheSource {
                source_type: PivotCacheSourceType::Consolidation,
                connection_id: None,
                worksheet: None,
                consolidation: Some(PivotCacheConsolidationSource {
                    auto_page: false,
                    sets: vec![PivotCacheConsolidationSet {
                        item_indexes: [u32::MAX; 4],
                        named_range: None,
                        built_in_name: false,
                        sheet_name: None,
                        external_rel_id: None,
                        range: None,
                    }],
                    pages: Vec::new(),
                }),
            }),
            ..PivotCacheDefinition::default()
        };
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));
    }

    #[test]
    fn refuses_out_of_range_model_values() {
        // More than four consolidation pages.
        let definition = PivotCacheDefinition {
            source: Some(PivotCacheSource {
                source_type: PivotCacheSourceType::Consolidation,
                connection_id: None,
                worksheet: None,
                consolidation: Some(PivotCacheConsolidationSource {
                    auto_page: false,
                    sets: Vec::new(),
                    pages: vec![
                        PivotCacheConsolidationPage {
                            item_names: Vec::new(),
                        };
                        5
                    ],
                }),
            }),
            ..PivotCacheDefinition::default()
        };
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));

        // Rule-filter item types exceeding the 13-bit mask.
        let definition = PivotCacheDefinition {
            calculated_items: vec![CalculatedItem {
                formula: PivotParsedFormulaData::default(),
                names: Vec::new(),
                filters: vec![PivotRuleFilter {
                    field: 0,
                    item_types: 1 << 13,
                    selected: false,
                    items: Vec::new(),
                }],
            }],
            ..PivotCacheDefinition::default()
        };
        assert!(matches!(
            write_pivot_cache_definition(&definition),
            Err(Error::Unrecognized { .. })
        ));
    }

    #[test]
    fn written_caches_round_trip_through_the_package() {
        let first = full_definition();
        let second = PivotCacheDefinition {
            record_count: 2,
            source: Some(PivotCacheSource {
                source_type: PivotCacheSourceType::Worksheet,
                connection_id: None,
                worksheet: Some(PivotCacheWorksheetSource {
                    named_range: Some("Data".to_string()),
                    built_in_name: false,
                    sheet_name: Some("Sheet1".to_string()),
                    external_rel_id: None,
                    range: None,
                }),
                consolidation: None,
            }),
            ..PivotCacheDefinition::default()
        };

        let mut workbook = WorkbookWriter::new();
        workbook.add_worksheet(MutableWorksheet::new("Sheet1"));
        let first_id = workbook.add_pivot_cache(&first).unwrap();
        let second_id = workbook.add_pivot_cache(&second).unwrap();
        assert_eq!(first_id, 1);
        assert_eq!(second_id, 2);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::package::Workbook::new(Cursor::new(output.into_inner())).unwrap();

        let definitions = reader.pivot_cache_definitions();
        assert_eq!(definitions.len(), 2);
        assert_eq!(definitions[0].0, first_id);
        assert_eq!(definitions[1].0, second_id);
        assert_eq!(definitions[0].1, first);
        assert_eq!(definitions[1].1, second);
        assert_eq!(reader.pivot_cache_definition(first_id), Some(&first));
        assert!(reader.pivot_cache_definition(99).is_none());
    }

    #[test]
    fn add_pivot_cache_surfaces_serializer_refusals() {
        let mut broken = PivotCacheDefinition::default();
        let mut broken_field = field("F");
        broken_field.shared_items = PivotCacheSharedItems {
            stats: None,
            items: vec![PivotCacheItem {
                value: PivotCacheItemValue::Number(1.0),
                additional: None,
            }],
        };
        broken.fields = vec![broken_field];

        let mut workbook = WorkbookWriter::new();
        workbook.add_worksheet(MutableWorksheet::new("Sheet1"));
        assert!(matches!(
            workbook.add_pivot_cache(&broken),
            Err(Error::Unrecognized { .. })
        ));
        // The refused cache was not attached; a valid cache gets id 1.
        let valid = PivotCacheDefinition::default();
        assert_eq!(workbook.add_pivot_cache(&valid).unwrap(), 1);
    }
}
