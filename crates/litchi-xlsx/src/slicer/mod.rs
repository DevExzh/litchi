//! Typed, bounded `SpreadsheetML` slicer ownership.
//!
//! The owner is deliberately split into semantic [`model`], XML [`codec`],
//! OPC [`package`] reads, read-only [`validation`], and the staged
//! [`Edit`] family (`Package::edit_slicers`, `Package::apply_slicers`). Slicers are inert controls: this owner never evaluates formulas,
//! refreshes a `PivotCache`, renders a UI, or applies a filter to cell data.

pub mod codec;
mod edit;
pub mod model;
pub mod package;
pub mod raw;
pub mod validation;

pub use edit::{Commit, Diagnostics, Edit, EditOptions, Error, LimitKind, Patch};
pub use model::{
    Cache, Data, DataKind, Definition, ExtensionList, Part, PivotTable, Slicer,
    SlicerExtensionList, Slicers,
};
pub use package::{load_caches, load_parts};

pub const CACHE_CONTENT_TYPE: &str = crate::slicer_cache::SLICER_CACHE_CONTENT_TYPE;
pub const CACHE_RELATIONSHIP_TYPE: &str = crate::slicer_cache::SLICER_CACHE_RELATIONSHIP_TYPE;

/// Read one slicer cache definition from its XML part.
pub fn read(xml: &[u8]) -> crate::error::Result<Definition> {
    codec::read(xml)
}

/// Read one worksheet slicers part from its XML.
pub fn read_views(xml: &[u8]) -> crate::error::Result<Slicers> {
    codec::read_views(xml)
}

/// Validate a slicer cache definition without mutating it.
pub fn validate(value: &Definition) -> crate::error::Result<()> {
    validation::definition(value)
}

/// Explicitly refuses runtime/UI behavior that this inert owner does not
/// implement.
pub fn unsupported_ui() -> crate::error::Result<()> {
    Err(Error::Unsupported {
        feature: "slicer UI, refresh, and filter application",
    })
}
