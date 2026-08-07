//! Media discovery, extraction, and transactional replacement.
//!
//! iWork packages store materialized assets as `Data/*` ZIP members and
//! describe them with `TSP.DataInfo` records in `Index/Metadata.iwa`.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::{EmbeddedMediaAsset, MediaAsset, MediaAssetId, MediaLimits, MediaStats, MediaType};
pub use package::{IWorkMediaEditor, MediaManager};

pub(crate) use codec::{embedded_assets, reachable_embedded_assets};

#[cfg(test)]
pub(crate) use codec::{PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE, field_payload};
#[cfg(test)]
pub(crate) use litchi_iwa_common::wire::parse_wire_fields;
#[cfg(test)]
pub(crate) use model::format_bytes;
