//! Semantic facades for PivotCache definition and records snapshots.
//!
//! The binary readers in [`super::parse`] and [`super::records`] only decode
//! the record grammar. This layer applies cross-record rules that need the
//! complete cache definition, then exposes the typed operations to callers.

use crate::package::error::Result;

use super::model::{PivotCacheDefinition, PivotCacheRecords};

/// Parse and semantically validate one `pivotCacheDefinition*.bin` part.
pub fn parse_pivot_cache_definition(data: &[u8]) -> Result<PivotCacheDefinition> {
    let definition = super::parse::parse_pivot_cache_definition_binary(data)?;
    super::validation::validate_definition(&definition)?;
    Ok(definition)
}

/// Parse and semantically validate one `pivotCacheRecords*.bin` part.
///
/// The definition supplies the field order and shared-item/type metadata
/// needed to turn the two XLSB record encodings into typed source rows.
pub fn parse_pivot_cache_records(
    data: &[u8],
    definition: &PivotCacheDefinition,
) -> Result<PivotCacheRecords> {
    super::validation::validate_records_precondition(definition)?;
    let records = super::records::parse_pivot_cache_records_binary(data, definition)?;
    super::validation::validate_records(definition, &records)?;
    Ok(records)
}
