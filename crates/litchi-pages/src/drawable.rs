//! Drawables anchored in the Pages body and their layer order.

pub use crate::drawable_order::{BodyDrawableHandle, BodyDrawableSelector, DrawableLayerMove};

/// Exact-source transactions for the body drawables' layer order.
pub mod order {
    pub use crate::package::drawable_order::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
}
