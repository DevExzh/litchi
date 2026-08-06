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
use crate::raw::Writer;
use crate::raw::kind as rt;

use super::{validation, wire};
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
    writer.write_record(
        rt::BEGIN_PIVOT_CACHE_DEF,
        &wire::definition_payload(definition),
    )?;

    if let Some(source) = &definition.source {
        write_source(&mut writer, source)?;
    }
    if !definition.fields.is_empty() {
        let mut payload = Vec::with_capacity(4);
        validation::write_count(&mut payload, definition.fields.len(), "BrtBeginPCDFields")?;
        writer.write_record(rt::BEGIN_PCD_FIELDS, &payload)?;
        for field in &definition.fields {
            write_field(&mut writer, field)?;
        }
        writer.write_record(rt::END_PCD_FIELDS, &[])?;
    }
    if !definition.hierarchies.is_empty() {
        let mut payload = Vec::with_capacity(4);
        validation::write_count(
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
        validation::write_count(
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
        validation::write_count(
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
        writer.write_record(rt::BEGIN_PCD14, &wire::pcd14_payload(ext14))?;
        writer.write_record(rt::END_PCD14, &[])?;
    }

    writer.write_record(rt::END_PIVOT_CACHE_DEF, &[])?;
    Ok(data)
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
        writer.write_record(
            rt::BEGIN_PCDS_RANGE,
            &wire::worksheet_range_payload(worksheet)?,
        )?;
        writer.write_record(rt::END_PCDS_RANGE, &[])?;
    }
    if let Some(consolidation) = &source.consolidation {
        write_consolidation(writer, consolidation)?;
    }
    writer.write_record(rt::END_PCD_SOURCE, &[])?;
    Ok(())
}

/// `BrtBeginPCDSConsol` collection (MS-XLSB 2.4.150).
fn write_consolidation<W: std::io::Write>(
    writer: &mut Writer<W>,
    consolidation: &PivotCacheConsolidationSource,
) -> Result<()> {
    validation::validate_consolidation(consolidation)?;
    let mut payload = Vec::with_capacity(2);
    payload.extend_from_slice(
        &(if consolidation.auto_page {
            wire::CONSOL_AUTO_PAGE
        } else {
            0
        })
        .to_le_bytes(),
    );
    writer.write_record(rt::BEGIN_PCDS_CONSOL, &payload)?;

    let mut payload = Vec::with_capacity(4);
    validation::write_count(&mut payload, consolidation.sets.len(), "BrtBeginPCDSCSets")?;
    writer.write_record(rt::BEGIN_PCDSC_SETS, &payload)?;
    for set in &consolidation.sets {
        writer.write_record(rt::BEGIN_PCDSC_SET, &wire::consolidation_set_payload(set)?)?;
        writer.write_record(rt::END_PCDSC_SET, &[])?;
    }
    writer.write_record(rt::END_PCDSC_SETS, &[])?;

    let mut payload = Vec::with_capacity(4);
    validation::write_count(
        &mut payload,
        consolidation.pages.len(),
        "BrtBeginPCDSCPages",
    )?;
    writer.write_record(rt::BEGIN_PCDSC_PAGES, &payload)?;
    for page in &consolidation.pages {
        writer.write_record(rt::BEGIN_PCDSC_PAGE, &[])?;
        for item_name in &page.item_names {
            let mut payload = Vec::with_capacity(item_name.len() * 2 + 4);
            wire::write_wide_string(&mut payload, item_name);
            writer.write_record(rt::BEGIN_PCDSCP_ITEM, &payload)?;
            writer.write_record(rt::END_PCDSCP_ITEM, &[])?;
        }
        writer.write_record(rt::END_PCDSC_PAGE, &[])?;
    }
    writer.write_record(rt::END_PCDSC_PAGES, &[])?;

    writer.write_record(rt::END_PCDS_CONSOL, &[])?;
    Ok(())
}

/// `BrtBeginPCDField` collection (MS-XLSB 2.4.136).
fn write_field<W: std::io::Write>(writer: &mut Writer<W>, field: &PivotCacheField) -> Result<()> {
    writer.write_record(rt::BEGIN_PCD_FIELD, &wire::field_payload(field)?)?;
    if !field.shared_items.items.is_empty() || field.shared_items.stats.is_some() {
        write_shared_items(writer, &field.shared_items)?;
    }
    if let Some(grouping) = &field.grouping {
        write_grouping(writer, grouping)?;
    }
    if field.ignore {
        // FRTBlank header; the reader only notes the record's presence.
        writer.write_record(wire::PCD_FIELD14, &[0; wire::FRT_BLANK_LEN])?;
    }
    writer.write_record(rt::END_PCD_FIELD, &[])?;
    Ok(())
}

/// `BrtBeginPCDFAtbl` collection (MS-XLSB 2.4.131).
fn write_shared_items<W: std::io::Write>(
    writer: &mut Writer<W>,
    shared_items: &PivotCacheSharedItems,
) -> Result<()> {
    validation::validate_shared_items(shared_items)?;
    let stats = shared_items
        .stats
        .as_ref()
        .expect("validated shared-item statistics");
    writer.write_record(
        rt::BEGIN_PCDF_ATBL,
        &wire::shared_items_stats_payload(stats)?,
    )?;
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
    validation::validate_cache_item(value, context)?;
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
        (PivotCacheItemValue::Index(_), _) => unreachable!("index items refused above"),
    };
    let mut payload = Vec::with_capacity(16);
    match value {
        PivotCacheItemValue::Missing => {},
        PivotCacheItemValue::Number(value) => payload.extend_from_slice(&value.to_le_bytes()),
        PivotCacheItemValue::Boolean(value) => payload.push(u8::from(*value)),
        PivotCacheItemValue::Error(code) => payload.push(*code as u8),
        PivotCacheItemValue::String(value) => wire::write_wide_string(&mut payload, value),
        PivotCacheItemValue::DateTime(value) => wire::write_date_time(&mut payload, value),
        PivotCacheItemValue::Index(_) => unreachable!("index items refused above"),
    }
    if let Some(additional) = additional {
        let mut flags = 0u16;
        if additional.ghost {
            flags |= wire::ADDL_GHOST;
        }
        if additional.calculated {
            flags |= wire::ADDL_CALCULATED;
        }
        if additional.caption.is_some() {
            flags |= wire::ADDL_CAPTION;
        }
        payload.extend_from_slice(&flags.to_le_bytes());
        if additional.caption.is_some() {
            wire::write_nullable_wide_string(&mut payload, &additional.caption);
        }
        validation::write_count(
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

/// `BrtBeginPCDFGroup` collection (MS-XLSB 2.4.135).
fn write_grouping<W: std::io::Write>(
    writer: &mut Writer<W>,
    grouping: &PivotCacheFieldGrouping,
) -> Result<()> {
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(
        &validation::optional_index(grouping.parent_field, "BrtBeginPCDFGroup")?.to_le_bytes(),
    );
    payload.extend_from_slice(
        &validation::optional_index(grouping.base_field, "BrtBeginPCDFGroup")?.to_le_bytes(),
    );
    writer.write_record(rt::BEGIN_PCDF_GROUP, &payload)?;

    if let Some(range) = &grouping.range {
        let mut payload = Vec::with_capacity(26);
        payload.push(range.group_by as u8);
        let mut flags = 0u8;
        if range.auto_start {
            flags |= wire::GROUP_RANGE_AUTO_START;
        }
        if range.auto_end {
            flags |= wire::GROUP_RANGE_AUTO_END;
        }
        if range.dates {
            flags |= wire::GROUP_RANGE_DATES;
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
    writer.write_record(
        rt::BEGIN_PCD_HIERARCHY,
        &wire::hierarchy_payload(hierarchy)?,
    )?;

    if !hierarchy.field_usage.is_empty() {
        let mut payload = Vec::with_capacity(4 + hierarchy.field_usage.len() * 4);
        validation::write_count(
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
        validation::write_count(
            &mut payload,
            hierarchy.grouping_levels.len(),
            "BrtBeginPCDHGLevels",
        )?;
        writer.write_record(rt::BEGIN_PCDHG_LEVELS, &payload)?;
        for level in &hierarchy.grouping_levels {
            let mut payload = Vec::with_capacity(16);
            payload.push(if level.group_level {
                wire::GROUPING_LEVEL_GROUP
            } else {
                0
            });
            wire::write_wide_string(&mut payload, &level.unique_name);
            wire::write_wide_string(&mut payload, &level.caption);
            writer.write_record(rt::BEGIN_PCDHG_LEVEL, &payload)?;
            writer.write_record(rt::END_PCDHG_LEVEL, &[])?;
        }
        writer.write_record(rt::END_PCDHG_LEVELS, &[])?;
    }
    if !hierarchy.grouping_groups.is_empty() {
        let mut payload = Vec::with_capacity(4);
        validation::write_count(
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
        let mut payload = vec![0; wire::FRT_BLANK_LEN];
        let mut flags = 0u8;
        if ext14.flatten_hierarchies {
            flags |= wire::HIER14_FLATTEN;
        }
        if ext14.measure_set {
            flags |= wire::HIER14_MEASURE_SET;
        }
        if ext14.hierarchize_distinct {
            flags |= wire::HIER14_HIERARCHIZE_DISTINCT;
        }
        if ext14.ignorable {
            flags |= wire::HIER14_IGNORABLE;
        }
        payload.push(flags);
        validation::write_count(&mut payload, ext14.hierarchy_indexes.len(), "BrtPCDH14")?;
        for index in &ext14.hierarchy_indexes {
            payload.extend_from_slice(&index.to_le_bytes());
        }
        writer.write_record(wire::PCD_H14, &payload)?;
    }
    writer.write_record(rt::END_PCD_HIERARCHY, &[])?;
    Ok(())
}

/// `BrtBeginPCDHGLGroup` collection (MS-XLSB 2.4.143).
fn write_grouping_group<W: std::io::Write>(
    writer: &mut Writer<W>,
    group: &PivotCacheGroupingGroup,
) -> Result<()> {
    let mut payload = Vec::with_capacity(24);
    payload.extend_from_slice(&group.group_number.to_le_bytes());
    payload.push(if group.parent_unique_name.is_some() {
        wire::GROUPING_GROUP_LOAD_PARENT
    } else {
        0
    });
    wire::write_wide_string(&mut payload, &group.name);
    wire::write_wide_string(&mut payload, &group.unique_name);
    wire::write_wide_string(&mut payload, &group.caption);
    if let Some(parent_unique_name) = &group.parent_unique_name {
        wire::write_wide_string(&mut payload, parent_unique_name);
    }
    writer.write_record(rt::BEGIN_PCDHGL_GROUP, &payload)?;

    let mut payload = Vec::with_capacity(4);
    validation::write_count(&mut payload, group.members.len(), "BrtBeginPCDHGLGMembers")?;
    writer.write_record(rt::BEGIN_PCDHGLG_MEMBERS, &payload)?;
    for member in &group.members {
        let mut payload = Vec::with_capacity(8);
        payload.extend_from_slice(&u32::from(member.is_group).to_le_bytes());
        wire::write_wide_string(&mut payload, &member.unique_name);
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
    validation::write_count(&mut payload, cache.entries.len(), "BrtBeginPCDSDTCEntries")?;
    writer.write_record(rt::BEGIN_PCDSDTC_ENTRIES, &payload)?;
    for value in &cache.entries {
        // Tuple cache entries carry no additional info; the reader parses
        // them non-strictly from plain `BrtPCDI*` records.
        write_cache_item(writer, value, None, "BrtBeginPCDSDTCEntries")?;
    }
    writer.write_record(rt::END_PCDSDTC_ENTRIES, &[])?;

    let mut payload = Vec::with_capacity(4);
    validation::write_count(&mut payload, cache.queries.len(), "BrtBeginPCDSDTCQueries")?;
    writer.write_record(rt::BEGIN_PCDSDTC_QUERIES, &payload)?;
    for query in &cache.queries {
        let mut payload = Vec::with_capacity(query.len() * 2 + 4);
        wire::write_wide_string(&mut payload, query);
        writer.write_record(rt::BEGIN_PCDSDTC_QUERY, &payload)?;
        writer.write_record(rt::END_PCDSDTC_QUERY, &[])?;
    }
    writer.write_record(rt::END_PCDSDTC_QUERIES, &[])?;

    let mut payload = Vec::with_capacity(4);
    validation::write_count(&mut payload, cache.sets.len(), "BrtBeginPCDSDTCSets")?;
    writer.write_record(rt::BEGIN_PCDSDTC_SETS, &payload)?;
    for set in &cache.sets {
        let mut payload = Vec::with_capacity(16);
        payload.extend_from_slice(
            &set.tuple_count
                .unwrap_or(wire::TUPLE_COUNT_UNKNOWN)
                .to_le_bytes(),
        );
        payload.extend_from_slice(&set.max_rank.to_le_bytes());
        payload.extend_from_slice(&set.sort_order.to_le_bytes());
        payload.push(if set.query_failed {
            wire::TUPLE_SET_QUERY_FAILED
        } else {
            0
        });
        wire::write_wide_string(&mut payload, &set.definition);
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
    wire::write_blob(&mut payload, &item.formula.tokens);
    wire::write_blob(&mut payload, &item.formula.extra);
    writer.write_record(rt::BEGIN_PCD_CALC_ITEM, &payload)?;

    if !item.names.is_empty() {
        let mut payload = Vec::with_capacity(4);
        validation::write_count(&mut payload, item.names.len(), "BrtBeginPNames")?;
        writer.write_record(rt::BEGIN_P_NAMES, &payload)?;
        for name in &item.names {
            write_name(writer, name)?;
        }
        writer.write_record(rt::END_P_NAMES, &[])?;
    }
    if !item.filters.is_empty() {
        let mut payload = Vec::with_capacity(4);
        validation::write_count(&mut payload, item.filters.len(), "BrtBeginPRFilters")?;
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
    payload.push(if name.err_name {
        wire::PNAME_ERR_NAME
    } else {
        0
    });
    payload.extend_from_slice(&[0; 2]); // padding
    writer.write_record(rt::BEGIN_P_NAME, &payload)?;

    let mut payload = Vec::with_capacity(4);
    payload.extend_from_slice(&wire::PN_PAIRS_DECLARED.to_le_bytes());
    writer.write_record(rt::BEGIN_PN_PAIRS, &payload)?;
    for pair in &name.pairs {
        let mut payload = Vec::with_capacity(12);
        let mut flags = 0u8;
        if pair.physical {
            flags |= wire::PNPAIR_PHYSICAL;
        }
        if pair.relative {
            flags |= wire::PNPAIR_RELATIVE;
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
    validation::validate_rule_filter(filter)?;
    let mut payload = Vec::with_capacity(11);
    payload.extend_from_slice(&filter.field.to_le_bytes());
    validation::write_count(&mut payload, filter.items.len(), "BrtBeginPRFilter")?;
    let mut flags = filter.item_types & wire::PR_FILTER_ITEM_TYPES_MASK;
    if filter.selected {
        flags |= wire::PR_FILTER_SELECTED;
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
        flags |= wire::CALC_MEM_LOAD_MEMBER_NAME;
    }
    if member.source_hierarchy.is_some() {
        flags |= wire::CALC_MEM_LOAD_SOURCE_HIER;
    }
    if member.parent_unique.is_some() {
        flags |= wire::CALC_MEM_LOAD_PARENT_UNIQUE;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&member.solve_order.to_le_bytes());
    payload.extend_from_slice(&u32::from(member.is_set).to_le_bytes());
    wire::write_wide_string(&mut payload, &member.name);
    wire::write_wide_string(&mut payload, &member.mdx);
    if let Some(value) = &member.member_name {
        wire::write_wide_string(&mut payload, value);
    }
    if let Some(value) = &member.source_hierarchy {
        wire::write_wide_string(&mut payload, value);
    }
    if let Some(value) = &member.parent_unique {
        wire::write_wide_string(&mut payload, value);
    }
    writer.write_record(rt::BEGIN_PCD_CALC_MEM, &payload)?;

    if let Some(ext14) = &member.ext14 {
        let mut payload = vec![0; wire::FRT_BLANK_LEN];
        let mut flags = 0u8;
        if ext14.flatten_hierarchies {
            flags |= wire::CALC_MEM14_FLATTEN;
        }
        if ext14.dynamic_set {
            flags |= wire::CALC_MEM14_DYNAMIC_SET;
        }
        if ext14.hierarchize_distinct {
            flags |= wire::CALC_MEM14_HIERARCHIZE_DISTINCT;
        }
        payload.push(flags);
        wire::write_wide_string(&mut payload, &ext14.display_folder);
        if let Some(long_mdx) = &ext14.long_mdx {
            wire::write_wide_string(&mut payload, long_mdx);
        }
        writer.write_record(rt::BEGIN_PCD_CALC_MEM14, &payload)?;
        writer.write_record(rt::END_PCD_CALC_MEM14, &[])?;
    }
    writer.write_record(rt::END_PCD_CALC_MEM, &[])?;
    Ok(())
}
