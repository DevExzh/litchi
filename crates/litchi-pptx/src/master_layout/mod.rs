//! Semantic slide-master and slide-layout authoring.
//!
//! [`crate::Package::edit_master_layout`] stages master and layout authoring
//! through the [`Edit`] family; [`placeholder`] owns the placeholder
//! type-extension family.

mod codec;
mod edit;
mod model;
pub(crate) mod package;
pub mod placeholder;

#[cfg(test)]
mod tests;

pub use crate::shape::{PLACEHOLDER_TYPE_EXTENSION_URI, PlaceholderTypeExtension};
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub use model::{
    AuthoredSlideLayout, AuthoredSlideMaster, MIN_MASTER_OR_LAYOUT_ID, PlaceholderKind,
    PlaceholderSpec, SlideLayoutKind,
};
pub use package::{
    add_slide_layout, add_slide_master, remove_slide_layout, store_placeholder_shape,
    validate_master_layout_graph,
};

pub(crate) use package::{preflight_add_slide_layout, resolve_owner};
pub(crate) use placeholder::preflight_store_placeholder_shape;
