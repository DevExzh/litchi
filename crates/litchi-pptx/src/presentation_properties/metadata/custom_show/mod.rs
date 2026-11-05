//! Custom-show values and `PresentationML` fragment codec.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
mod wire;

#[cfg(test)]
mod tests;

pub use super::Limits;
pub use model::{List, Show};
pub use package::{load, load_snapshot};
pub use transaction::{Draft, Revision, Snapshot};

/// Rewrite a detached custom-show value against a presentation source while
/// retaining opaque XML owned by the custom-show list.
pub(crate) fn rewrite_source(
    source: &[u8],
    before: &List,
    after: &List,
    limits: Limits,
) -> crate::Result<Vec<u8>> {
    wire::rewrite_source(source, before, after, limits)
}
