//! Semantic, bounded and inert `PowerPoint` font ownership.

mod codec;
mod edit;
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

pub use edit::Edit;
#[allow(
    clippy::module_name_repetitions,
    reason = "`EmbeddedFont`, `FontCollection`, `FontCollections`, and `FontEmbeddingFlags` are established public API names re-exported through the `font` module; renaming them would break downstream crates"
)]
pub use model::{
    EmbeddedFont, EotMetadata, Facet, Font, FontCollection, FontCollections, FontEmbeddingFlags,
    Limits, Scope, SharedFontData, validate_eot_facet,
};
pub use package::{PackageLimits, PackageOptions};
pub(crate) use package::{require_stream_only_cfb, validate_unrelated_streams};
pub use patch::Revision;
pub use snapshot::Snapshot;
pub use transaction::{Change, ChangeKind};

crate::scoped::family!("font");

#[cfg(feature = "fonts")]
#[allow(
    clippy::module_name_repetitions,
    reason = "`PreparedFont` is the established public alias for `litchi_fonts::Prepared`; renaming it would break downstream crates"
)]
pub use litchi_fonts::Prepared as PreparedFont;
#[cfg(feature = "fonts")]
pub use litchi_fonts::embedding::powerpoint::{Intent as EotIntent, Limits as EotLimits};
