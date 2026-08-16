//! Atomic semantic transactions over an opened presentation.
//!
//! One immutable root composes slide/text/notes edits with common, grouped, and
//! endpoint-closed connector shape transfer; picture-part creation/removal;
//! ordinary charts, media, table styles, legacy/modern comments and typed
//! extensions; master/layout authoring; and dependency relationship transfer.
//! Publication uses an exact finite OPC resource patch and classifies shape
//! transfer refusals before mutation.

mod copy_plan;
pub(crate) mod cross_copy_plan;
mod model;
mod patch;
mod remove_plan;
mod transaction;
mod xml;

#[cfg(test)]
mod tests;

pub use copy_plan::{SlideCopyPart, SlideCopyPlan};
pub use cross_copy_plan::{CrossSlideCopyPatch, CrossSlideCopyPlan};
pub use model::{Limits, MAX_SHAPE_TEXT_REPLACEMENTS, ShapeTextReplacement, Slide, Snapshot};
pub use patch::{Conflict, History, Patch, Resolution, ThreeWayPlan};
pub use remove_plan::{SlideRemovalPatch, SlideRemovalPlan};
pub use transaction::{Commit, Transaction};

pub(crate) use model::{capture_with_provenance, package_fingerprint};
pub(crate) use patch::apply;
pub(crate) use remove_plan::apply_patch as apply_removal_patch;
pub(crate) use xml::{insert_slide_binding, stage_shape_texts};
