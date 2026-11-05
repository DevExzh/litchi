//! Streaming `PresentationML` authoring.
//!
//! The in-memory authoring drafts live in [`crate::builder`].

pub mod streaming;

pub use streaming::{
    StreamingPresentationLimits, StreamingPresentationOptions, StreamingPresentationScratchLimits,
    StreamingPresentationWriter, StreamingSlideWriter, TextBoxSpec,
};
