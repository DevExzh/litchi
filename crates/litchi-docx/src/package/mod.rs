#![expect(
    clippy::module_inception,
    reason = "the module path preserves the established public facade"
)]
//! `WordprocessingML` package owner.
//!
//! The model stores typed package state, the codec owns OPC I/O, and the
//! package layer coordinates relationship-backed graph edits.

mod codec;
mod model;
mod package;
pub mod story;
#[cfg(test)]
mod tests;

pub use model::Package;
pub use story::{
    StoryDialect, StoryHyperlinkTextReplacement, StoryInventory, StoryKind, StoryLimits,
    StoryOwner, StoryPart, StoryTopology,
};
