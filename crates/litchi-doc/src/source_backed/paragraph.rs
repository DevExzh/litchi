//! The source-backed paragraph family: one same-length Unicode main-story
//! paragraph replacement published as a checked equal-length splice.

pub use super::Error;
pub use super::owner::{Diagnostics, Edit, Patch};

/// A published source-backed paragraph edit: the committed package, its
/// exact-source patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<super::Package, Patch, Diagnostics>;

/// The source-backed paragraph family has no caller-selected limit
/// dimension: the committed package is opened under the source package's
/// [`super::Limits`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {}
