//! Semantic orchestration for PivotCache snapshot serialization.

use crate::package::error::Result;
use crate::package::pivot::model::PivotCacheDefinition;

/// Serialize a typed PivotCache snapshot after its top-level semantic shape
/// has been checked.
pub(crate) fn write_pivot_cache_definition(definition: &PivotCacheDefinition) -> Result<Vec<u8>> {
    super::validation::validate_definition(definition)?;
    super::binary::write_pivot_cache_definition(definition)
}
