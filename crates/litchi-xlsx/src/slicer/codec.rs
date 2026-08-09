//! Bounded `SpreadsheetML` slicer XML conversion.

use super::model::{Definition, Slicers};
use crate::error::Result;

/// Parse one `slicerCacheDefinition` part.
pub fn read(xml: &[u8]) -> Result<Definition> {
    crate::slicer_cache::parse(xml)
}

/// Serialize one `slicerCacheDefinition` part deterministically.
pub fn write(value: &Definition) -> Result<Vec<u8>> {
    crate::slicer_cache::write(value)
}

/// Parse one worksheet `slicers` part.
pub fn read_views(xml: &[u8]) -> Result<Slicers> {
    crate::slicer_cache::views::parse_slicers(xml)
}

/// Serialize one worksheet `slicers` part deterministically.
pub fn write_views(value: &Slicers) -> Result<Vec<u8>> {
    crate::slicer_cache::views::write_slicers(value)
}
