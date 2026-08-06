//! Cross-record semantic validation for typed PivotCache snapshots.

use crate::package::error::Result;
use crate::package::pivot::model::{PivotCacheDefinition, PivotCacheItemValue, PivotCacheRecords};
use crate::package::walker::malformed;

/// Validate definition invariants that cannot be checked while walking one
/// BIFF12 record at a time.
pub(super) fn validate_definition(definition: &PivotCacheDefinition) -> Result<()> {
    validate_count(definition.fields.len(), "BrtBeginPCDFields")?;
    validate_count(definition.hierarchies.len(), "BrtBeginPCDHierarchies")?;
    validate_count(definition.calculated_items.len(), "BrtBeginPCDCalcItems")?;
    validate_count(definition.calculated_members.len(), "BrtBeginPCDCalcMems")?;

    for (field_index, field) in definition.fields.iter().enumerate() {
        for (item_index, item) in field.shared_items.items.iter().enumerate() {
            if let PivotCacheItemValue::Index(index) = &item.value
                && usize::try_from(*index)
                    .ok()
                    .is_some_and(|index| index >= field.shared_items.items.len())
            {
                return Err(malformed(
                    "BrtPCDIIndex",
                    format!("field {field_index} item {item_index} references {index}"),
                ));
            }
        }
    }

    if let Some(source) = &definition.source {
        if source.source_type == crate::package::pivot::model::PivotCacheSourceType::External
            && source.connection_id.is_none()
        {
            return Err(malformed(
                "BrtBeginPCDSource",
                "external source is missing its connection identifier",
            ));
        }
        if source.source_type == crate::package::pivot::model::PivotCacheSourceType::Worksheet
            && source.worksheet.is_none()
        {
            return Err(malformed(
                "BrtBeginPCDSource",
                "worksheet source is missing its range payload",
            ));
        }
        if source.source_type == crate::package::pivot::model::PivotCacheSourceType::Consolidation
            && source.consolidation.is_none()
        {
            return Err(malformed(
                "BrtBeginPCDSource",
                "consolidation source is missing its range payload",
            ));
        }
    }
    Ok(())
}

/// Check definition metadata before parsing a records part.
pub(super) fn validate_records_precondition(definition: &PivotCacheDefinition) -> Result<()> {
    validate_definition(definition)?;
    if definition
        .fields
        .iter()
        .filter(|field| field.source_field)
        .count()
        > u32::MAX as usize
    {
        return Err(malformed(
            "BrtPCRRecord",
            "source field count overflows the bounded record model",
        ));
    }
    Ok(())
}

/// Validate row width and shared-item indexes after records are decoded.
pub(super) fn validate_records(
    definition: &PivotCacheDefinition,
    records: &PivotCacheRecords,
) -> Result<()> {
    if u64::from(records.record_count) != records.records.len() as u64 {
        return Err(malformed(
            "BrtBeginPivotCacheRecords",
            format!(
                "declared {} records, found {}",
                records.record_count,
                records.records.len()
            ),
        ));
    }
    let source_field_count = definition
        .fields
        .iter()
        .filter(|field| field.source_field)
        .count();
    for (row_index, record) in records.records.iter().enumerate() {
        if record.values.len() != source_field_count {
            return Err(malformed(
                "BrtPCRRecord",
                format!(
                    "row {row_index} has {} values, expected {}",
                    record.values.len(),
                    source_field_count
                ),
            ));
        }
        let mut source_fields = definition.fields.iter().filter(|field| field.source_field);
        for (value, field) in record.values.iter().zip(&mut source_fields) {
            if let PivotCacheItemValue::Index(index) = value
                && usize::try_from(*index)
                    .ok()
                    .is_some_and(|index| index >= field.shared_items.items.len())
            {
                return Err(malformed(
                    "BrtPCDIIndex",
                    format!("record row {row_index} references shared item {index}"),
                ));
            }
        }
    }
    Ok(())
}

fn validate_count(count: usize, context: &'static str) -> Result<()> {
    if count > u32::MAX as usize {
        return Err(malformed(context, "collection count overflows u32"));
    }
    Ok(())
}
