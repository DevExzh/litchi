//! Layered MS-XLSB PivotCache definition writer.

mod semantic;
mod validation;
mod wire;

#[cfg(test)]
mod tests;

pub(super) fn write_pivot_cache_definition(
    definition: &crate::package::pivot::model::PivotCacheDefinition,
) -> crate::package::error::Result<Vec<u8>> {
    semantic::write_pivot_cache_definition(definition)
}
