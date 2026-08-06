//! Representability checks for the XLSB PivotCache binary writer.

use crate::package::error::Result;
use crate::package::pivot::model::*;
use crate::package::walker::malformed;

pub(super) fn write_count(data: &mut Vec<u8>, count: usize, context: &'static str) -> Result<()> {
    let count =
        u32::try_from(count).map_err(|_| malformed(context, "collection count overflow"))?;
    data.extend_from_slice(&count.to_le_bytes());
    Ok(())
}

pub(super) fn optional_index(value: Option<u32>, context: &'static str) -> Result<i32> {
    match value {
        Some(value) => i32::try_from(value).map_err(|_| malformed(context, "index overflow")),
        None => Ok(-1),
    }
}

pub(super) fn check_name_or_range(
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

pub(super) fn validate_consolidation(source: &PivotCacheConsolidationSource) -> Result<()> {
    const MAX_CONSOLIDATION_PAGES: usize = 4;
    if source.pages.len() > MAX_CONSOLIDATION_PAGES {
        return Err(malformed(
            "BrtBeginPCDSCPages",
            format!(
                "{} consolidation pages exceed the maximum of {MAX_CONSOLIDATION_PAGES}",
                source.pages.len()
            ),
        ));
    }
    Ok(())
}

pub(super) fn validate_shared_items(shared_items: &PivotCacheSharedItems) -> Result<()> {
    if shared_items.stats.is_none() {
        return Err(malformed(
            "BrtBeginPCDFAtbl",
            "shared items without statistics cannot be emitted losslessly",
        ));
    }
    Ok(())
}

pub(super) fn validate_shared_items_stats(stats: &PivotCacheSharedItemsStats) -> Result<()> {
    if stats.minimum.is_some() != stats.maximum.is_some() {
        return Err(malformed(
            "BrtBeginPCDFAtbl",
            "minimum and maximum must both be set or both be absent",
        ));
    }
    Ok(())
}

pub(super) fn validate_cache_item(
    value: &PivotCacheItemValue,
    context: &'static str,
) -> Result<()> {
    if matches!(value, PivotCacheItemValue::Index(_)) {
        return Err(malformed(
            context,
            "index items are only valid inside a discrete grouping",
        ));
    }
    Ok(())
}

pub(super) fn validate_rule_filter(filter: &PivotRuleFilter) -> Result<()> {
    const ITEM_TYPES_MASK: u32 = 0x1FFF;
    if filter.item_types & !ITEM_TYPES_MASK != 0 {
        return Err(malformed(
            "BrtBeginPRFilter",
            format!(
                "item types 0x{:X} exceed the 13-bit mask",
                filter.item_types
            ),
        ));
    }
    Ok(())
}
