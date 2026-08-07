//! Semantic, bounded and inert PowerPoint font ownership.

mod codec;
mod model;
mod package;
mod patch;
#[cfg(feature = "fonts")]
pub(crate) mod prepared;
mod snapshot;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    EmbeddedFont, EotMetadata, Facet, Font, FontCollection, FontCollections, FontEmbeddingFlags,
    Limits, Scope, SharedFontData, validate_eot_facet,
};
pub use package::{PackageLimits, PackageOptions};
pub use patch::{Patch, Revision};
pub use snapshot::Snapshot;
pub use transaction::{Change, ChangeKind, Commit, Transaction};

#[cfg(feature = "fonts")]
pub use litchi_fonts::Prepared as PreparedFont;
#[cfg(feature = "fonts")]
pub use litchi_fonts::embedding::powerpoint::{Intent as EotIntent, Limits as EotLimits};
