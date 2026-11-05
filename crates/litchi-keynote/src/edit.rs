//! The primary slide-playback-state transaction family.
//!
//! [`crate::Package::edit`] starts an [`Edit`]; [`Edit::commit`] publishes a
//! [`Commit`] carrying the new immutable package, an exact-source [`Patch`]
//! and [`Diagnostics`]. [`crate::Package::apply`] replays a patch.

pub use crate::package::edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal};
