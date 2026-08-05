//! Semantic slide-master and slide-layout authoring facade.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{
    AuthoredSlideLayout, AuthoredSlideMaster, MIN_MASTER_OR_LAYOUT_ID, PlaceholderKind,
    PlaceholderSpec, SlideLayoutKind,
};
pub use package::{
    add_slide_layout, add_slide_master, remove_slide_layout, store_placeholder_shape,
    validate_master_layout_graph,
};
