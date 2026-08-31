//! Borrowed semantic view of one `PresentationML` package graph.
//!
//! The typed facade delegates XML scanning to `codec` and package
//! relationship traversal to `package`. Embedded resources and media
//! authoring remain available through their existing child modules.

mod codec;
mod model;
mod order;
mod package;
mod source;
mod source_cross_copy;
mod transition;

#[cfg(test)]
mod tests;

pub mod embedded;
pub mod media;

pub use model::Presentation;
pub use order::{
    SourceBackedSlideOrderCommit, SourceBackedSlideOrderEdit, SourceBackedSlideOrderPatch,
    SourceBackedSlideOrderSnapshot,
};
pub use source::{
    MAX_SOURCE_BACKED_SLIDE_BATCH, SourceBackedPresentation, SourceBackedPresentationEditor,
    SourceBackedSlideBatchCommit, SourceBackedSlideBatchEdit, SourceBackedSlideBatchPatch,
    SourceBackedSlideBatchSnapshot, SourceBackedSlideCommit, SourceBackedSlideEdit,
    SourceBackedSlidePatch, SourceBackedSlideSnapshot, SourceImage, SourceImageDescriptor,
    SourceImageTarget, SourceSlide,
};
pub use source_cross_copy::{SourceBackedCrossSlideCopyPlan, SourceBackedCrossSlideCopySnapshot};
