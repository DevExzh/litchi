//! BIFF12 wire payloads and record primitives for PivotCache writing.

use crate::package::error::Result;
use crate::package::pivot::model::*;
use crate::package::walker::malformed;
use crate::raw::kind as rt;

use super::validation;

/// `BrtPCDField14` (MS-XLSB 2.4.725): marks the preceding cache field as ignorable.
pub(super) const PCD_FIELD14: crate::raw::Kind = rt::PCD_FIELD14;
/// `BrtPCDH14` (MS-XLSB 2.4.726): named-set extension of a cache hierarchy.
pub(super) const PCD_H14: crate::raw::Kind = rt::PCD_H14;

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
pub(super) const CONSOL_AUTO_PAGE: u16 = 1 << 0;

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
pub(super) const GROUP_RANGE_AUTO_START: u8 = 1 << 0;
pub(super) const GROUP_RANGE_AUTO_END: u8 = 1 << 1;
pub(super) const GROUP_RANGE_DATES: u8 = 1 << 2;

// `PCDIAddlInfo` flags word (MS-XLSB 2.5.100).
pub(super) const ADDL_GHOST: u16 = 1 << 0;
pub(super) const ADDL_CALCULATED: u16 = 1 << 1;
pub(super) const ADDL_CAPTION: u16 = 1 << 2;

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
pub(super) const GROUPING_LEVEL_GROUP: u8 = 1 << 0;
// `BrtBeginPCDHGLGroup` flags byte (MS-XLSB 2.4.143).
pub(super) const GROUPING_GROUP_LOAD_PARENT: u8 = 1 << 0;

// `BrtPCDH14` flags byte (MS-XLSB 2.4.726).
pub(super) const HIER14_FLATTEN: u8 = 1 << 0;
pub(super) const HIER14_MEASURE_SET: u8 = 1 << 1;
pub(super) const HIER14_HIERARCHIZE_DISTINCT: u8 = 1 << 2;
pub(super) const HIER14_IGNORABLE: u8 = 1 << 3;

// `PCDCalcMemCommon` flags word (MS-XLSB 2.5.99).
pub(super) const CALC_MEM_LOAD_MEMBER_NAME: u32 = 1 << 0;
pub(super) const CALC_MEM_LOAD_SOURCE_HIER: u32 = 1 << 1;
pub(super) const CALC_MEM_LOAD_PARENT_UNIQUE: u32 = 1 << 2;

// `BrtBeginPCDCalcMem14` flags byte (MS-XLSB 2.4.127).
pub(super) const CALC_MEM14_FLATTEN: u8 = 1 << 0;
pub(super) const CALC_MEM14_DYNAMIC_SET: u8 = 1 << 1;
pub(super) const CALC_MEM14_HIERARCHIZE_DISTINCT: u8 = 1 << 2;

// `BrtBeginPName` flags byte (MS-XLSB 2.4.176).
pub(super) const PNAME_ERR_NAME: u8 = 1 << 0;
// `BrtBeginPNPair` flags byte (MS-XLSB 2.4.178).
pub(super) const PNPAIR_PHYSICAL: u8 = 1 << 0;
pub(super) const PNPAIR_RELATIVE: u8 = 1 << 1;

// `PRFilter` 24-bit flags (MS-XLSB 2.4.180).
pub(super) const PR_FILTER_ITEM_TYPES_MASK: u32 = 0x1FFF;
pub(super) const PR_FILTER_SELECTED: u32 = 1 << 16;

// `BrtBeginPCDSDTCSet` flags byte (MS-XLSB 2.4.162).
pub(super) const TUPLE_SET_QUERY_FAILED: u8 = 1 << 0;
/// `cTuples` value meaning the tuple count is unknown (MS-XLSB 2.4.162).
pub(super) const TUPLE_COUNT_UNKNOWN: u32 = u32::MAX;

// `BrtBeginPCD14` flags byte (MS-XLSB 2.4.123).
const PCD14_SLICER_DATA: u8 = 1 << 0;
const PCD14_SUBQUERY_CALC_MEM: u8 = 1 << 1;
const PCD14_SUBQUERY_NON_VISUAL: u8 = 1 << 2;
const PCD14_ADD_CALC_MEMS: u8 = 1 << 3;

/// Byte length of an `FRTBlank` header (MS-XLSB 2.5.55).
pub(super) const FRT_BLANK_LEN: usize = 4;
/// `cpairs` value required by `BrtBeginPNPairs` (MS-XLSB 2.4.179).
pub(super) const PN_PAIRS_DECLARED: u32 = 1;

pub(super) fn write_wide_string(data: &mut Vec<u8>, value: &str) {
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
}

pub(super) fn write_nullable_wide_string(data: &mut Vec<u8>, value: &Option<String>) {
    match value {
        Some(value) => write_wide_string(data, value),
        None => data.extend_from_slice(&u32::MAX.to_le_bytes()),
    }
}

pub(super) fn write_blob(data: &mut Vec<u8>, blob: &[u8]) {
    data.extend_from_slice(&(blob.len() as u32).to_le_bytes());
    data.extend_from_slice(blob);
}

/// `BrtBeginPivotCacheDef` payload (MS-XLSB 2.4.168).
pub(super) fn definition_payload(definition: &PivotCacheDefinition) -> Vec<u8> {
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

pub(super) fn write_range(data: &mut Vec<u8>, range: &PivotCacheRange) {
    data.extend_from_slice(&range.first_row.to_le_bytes());
    data.extend_from_slice(&range.last_row.to_le_bytes());
    data.extend_from_slice(&range.first_column.to_le_bytes());
    data.extend_from_slice(&range.last_column.to_le_bytes());
}

/// `BrtBeginPCDSRange` payload (MS-XLSB 2.4.167).
pub(super) fn worksheet_range_payload(source: &PivotCacheWorksheetSource) -> Result<Vec<u8>> {
    validation::check_name_or_range(&source.named_range, &source.range, "BrtBeginPCDSRange")?;
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

/// `BrtBeginPCDSCSet` payload (MS-XLSB 2.4.154).
pub(super) fn consolidation_set_payload(set: &PivotCacheConsolidationSet) -> Result<Vec<u8>> {
    validation::check_name_or_range(&set.named_range, &set.range, "BrtBeginPCDSCSet")?;
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

/// `BrtBeginPCDField` payload (MS-XLSB 2.4.136).
pub(super) fn field_payload(field: &PivotCacheField) -> Result<Vec<u8>> {
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
    validation::write_count(
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

/// `BrtBeginPCDFAtbl` payload (MS-XLSB 2.4.131).
pub(super) fn shared_items_stats_payload(stats: &PivotCacheSharedItemsStats) -> Result<Vec<u8>> {
    validation::validate_shared_items_stats(stats)?;
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

/// `PCDIDateTime` (MS-XLSB 2.5.101).
pub(super) fn write_date_time(data: &mut Vec<u8>, value: &PivotCacheDateTime) {
    data.extend_from_slice(&value.year.to_le_bytes());
    data.extend_from_slice(&value.month.to_le_bytes());
    data.push(value.day);
    data.push(value.hour);
    data.push(value.minute);
    data.push(value.second);
}

/// `BrtBeginPCDHierarchy` payload (MS-XLSB 2.4.146).
pub(super) fn hierarchy_payload(hierarchy: &PivotCacheHierarchy) -> Result<Vec<u8>> {
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
        &validation::optional_index(hierarchy.set_parent_index, "BrtBeginPCDHierarchy")?
            .to_le_bytes(),
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

/// `BrtBeginPCD14` payload (MS-XLSB 2.4.123).
pub(super) fn pcd14_payload(ext14: &PivotCacheDefinitionExt14) -> Vec<u8> {
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
