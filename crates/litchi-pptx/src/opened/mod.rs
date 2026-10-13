//! Atomic semantic transactions over an opened presentation.
//!
//! One immutable root composes slide/text/notes edits with common, grouped, and
//! endpoint-closed connector shape transfer; picture-part creation/removal;
//! ordinary charts, media, table styles, legacy/modern comments and typed
//! extensions; master/layout authoring; and dependency relationship transfer.
//! Publication uses an exact finite OPC resource patch and classifies shape
//! transfer refusals before mutation.

mod chunks;
mod copy_plan;
pub(crate) mod cross_copy_plan;
mod geometry;
mod insert;
mod limit;
mod model;
mod patch;
mod properties;
mod removal_proof;
mod remove_plan;
mod slide_properties;
mod text_format;
mod transaction;
mod xml;

#[cfg(test)]
mod mce_retention_tests;
#[cfg(test)]
mod slide_root_memo_tests;
#[cfg(test)]
mod tests;
#[cfg(test)]
mod text_format_tests;
#[cfg(test)]
mod text_run_tests;

/// The failure type of this transaction family.
///
/// The family reports the crate's [`crate::Error`]; resource refusals of the
/// family's own [`Limits`] use [`crate::Error::EditLimit`] with a
/// [`LimitKind`]. Splitting a focused error out of the crate error is
/// deferred (record 0252).
pub use crate::Error;
pub use copy_plan::{SlideCopyPart, SlideCopyPlan};
pub use cross_copy_plan::{CrossSlideCopyPatch, CrossSlideCopyPlan};
pub use limit::LimitKind;
pub use model::{
    Limits, MAX_SHAPE_PROPERTIES_REPLACEMENTS, MAX_SHAPE_TEXT_REPLACEMENTS,
    ShapePropertiesReplacement, ShapeTextReplacement, Slide, Snapshot,
};
pub use patch::{Conflict, History, Patch, Resolution, ThreeWayPlan};
pub use remove_plan::{SlideRemovalPatch, SlideRemovalPlan};
pub(crate) use text_format::stage_text_properties;
pub use text_format::{
    MAX_FONT_SIZE, MAX_TYPEFACE_BYTES, MIN_FONT_SIZE, Rgb, TextEdit, TextProperties, TextTarget,
};
pub use transaction::{Commit, Diagnostics, Edit};

pub(crate) use chunks::{
    ChunkAdmission, admit_shape_property_chunks, admit_shape_text_chunks,
    stage_shape_property_chunks, stage_shape_property_chunks_admitted, stage_shape_text_chunks,
    stage_shape_text_chunks_admitted,
};
pub(crate) use geometry::geometry_skeleton;
pub(crate) use model::{
    PartDigests, capture_with_parent_digests, capture_with_provenance, package_fingerprint,
    package_fingerprint_with_memo,
};
pub(crate) use patch::{apply, apply_to_new};
pub(crate) use properties::stage_shape_properties;
pub(crate) use removal_proof::normalize_shape_removal_pair;
pub(crate) use remove_plan::apply_patch as apply_removal_patch;
pub(crate) use remove_plan::{
    validate_insertion_owner, validate_removal_presentation_owner, validate_removal_slide_surface,
};
pub(crate) use slide_properties::{
    normalize_slide_properties, slide_properties_skeleton, stage_slide_properties,
};
pub(crate) use transaction::{
    next_shape_id, stage_text_box, text_box_fragment, text_box_output_bound,
    validate_text_box_input,
};
pub(crate) use xml::{
    insert_slide_binding, reorder_slide_bindings, shape_text_skeleton, stage_shape_removal,
    stage_shape_texts, strip_appended_text_boxes, validate_slide_bindings,
};
pub(crate) use xml::{remove_slide_binding, validate_slide_catalog_delta};

#[cfg(test)]
mod owner_commit_tests;
