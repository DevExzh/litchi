//! Borrowed semantic view of one `PresentationML` package graph.
//!
//! The typed facade delegates XML scanning to `codec` and package
//! relationship traversal to `package`. Embedded resources and media
//! authoring remain available through their existing child modules.

mod codec;
mod model;
mod package;
mod source;
mod transition;

#[cfg(test)]
mod tests;

pub mod embedded;
pub mod media;

pub use model::Presentation;
pub use source::{
    MAX_SOURCE_BACKED_SLIDE_BATCH, SourceBackedPresentation, SourceBackedPresentationEditor,
    SourceBackedSlideBatchCommit, SourceBackedSlideBatchEdit, SourceBackedSlideBatchPatch,
    SourceBackedSlideBatchSnapshot, SourceBackedSlideCommit, SourceBackedSlideEdit,
    SourceBackedSlidePatch, SourceBackedSlideSnapshot, SourceSlide,
};
