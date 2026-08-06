//! Cheap semantic guards shared by the PivotCache writer facade.

use crate::package::error::Result;
use crate::package::pivot::model::PivotCacheDefinition;
use crate::package::walker::malformed;

/// Validate collection sizes before entering the binary encoder.
///
/// The binary layer repeats the same bounded conversions at the record that
/// carries each count. Keeping this inexpensive facade check explicit makes
/// the semantic/binary boundary visible without adding allocations or
/// changing successful BIFF12 output.
pub(super) fn validate_definition(definition: &PivotCacheDefinition) -> Result<()> {
    validate_count(definition.fields.len(), "BrtBeginPCDFields")?;
    validate_count(definition.hierarchies.len(), "BrtBeginPCDHierarchies")?;
    validate_count(definition.calculated_items.len(), "BrtBeginPCDCalcItems")?;
    validate_count(definition.calculated_members.len(), "BrtBeginPCDCalcMems")?;
    Ok(())
}

fn validate_count(count: usize, context: &'static str) -> Result<()> {
    if count > u32::MAX as usize {
        return Err(malformed(context, "collection count overflow"));
    }
    Ok(())
}
