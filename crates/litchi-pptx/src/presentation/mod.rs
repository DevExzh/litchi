//! Borrowed semantic view of one `PresentationML` package graph.
//!
//! The typed facade delegates XML scanning to `codec` and package
//! relationship traversal to `package`. Embedded resources and media
//! authoring remain available through their existing child modules.

mod codec;
mod model;
mod order;
mod package;
pub(crate) mod source;
mod source_cross_copy;
mod source_removal;
pub(crate) mod transition;

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
    MAX_SOURCE_BACKED_SLIDE_BATCH, SourceBackedIntoEditorError, SourceBackedPresentation,
    SourceBackedPresentationEditor, SourceBackedSlideBatchCommit, SourceBackedSlideBatchEdit,
    SourceBackedSlideBatchPatch, SourceBackedSlideBatchSnapshot, SourceBackedSlideCommit,
    SourceBackedSlideEdit, SourceBackedSlidePatch, SourceBackedSlideSnapshot,
    SourceBackedSvgAttachmentCommit, SourceBackedSvgAttachmentEdit, SourceBackedSvgAttachmentPatch,
    SourceBackedSvgAttachmentSnapshot, SourceBackedSvgCommit, SourceBackedSvgEdit,
    SourceBackedSvgPatch, SourceBackedSvgSnapshot, SourceImage, SourceImageDescriptor,
    SourceImageTarget, SourceSlide, SourceSvgAttachment, SourceSvgAttachmentReplacement,
    SourceSvgDescriptor, SourceSvgImage, SourceSvgReplacement,
};
pub use source_cross_copy::{SourceBackedCrossSlideCopyPlan, SourceBackedCrossSlideCopySnapshot};
pub use source_removal::SourceBackedSlideRemovalPlan;

/// Source-backed per-slide modern-comment snapshots.
pub use source::modern_comments;
/// Source-backed slide related-part snapshots.
pub use source::related;
/// Source-backed scene retention diagnostics.
pub use source::scenes as scene;

/// Explicit access to the retained presentation root.
pub use source::raw;
/// Detached presentation section membership.
pub use source::sections;
