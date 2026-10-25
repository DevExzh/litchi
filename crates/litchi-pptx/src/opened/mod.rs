//! Atomic semantic transactions over an opened presentation.
//!
//! One immutable root composes slide/text/notes edits with common, grouped, and
//! endpoint-closed connector shape transfer; picture-part creation/removal;
//! ordinary charts, media, table styles, legacy/modern comments and typed
//! extensions; master/layout authoring; and dependency relationship transfer.
//! Publication uses an exact finite OPC resource patch and classifies shape
//! transfer refusals before mutation.

mod chunks;
pub(crate) mod copy_plan;
pub(crate) mod cross_copy_plan;
mod enclosed_removal;
mod geometry;
mod insert;
mod limit;
mod model;
mod new_slide;
mod patch;
mod properties;
mod removal_proof;
pub(crate) mod remove_plan;
pub(crate) mod sections;
mod slide_properties;
mod text_format;
mod transaction;
mod xml;

#[cfg(test)]
mod mce_retention_tests;
#[cfg(test)]
mod picture_edit_tests;
#[cfg(test)]
mod slide_root_memo_tests;
#[cfg(test)]
mod tests;
#[cfg(test)]
mod text_format_tests;
#[cfg(test)]
mod text_run_tests;
#[cfg(test)]
mod text_structure_tests;
#[cfg(test)]
mod transition_edit_tests;

/// The failure type of this transaction family.
///
/// The family reports the crate's [`crate::Error`]; resource refusals of the
/// family's own [`Limits`] use [`crate::Error::Limit`] with a
/// [`LimitKind`]. Splitting a focused error out of the crate error is
/// deferred (record 0252).
pub use crate::Error;
pub use limit::{CodecLimitKind, LimitKind};
pub use model::{
    Limits, MAX_SHAPE_PROPERTIES_REPLACEMENTS, MAX_SHAPE_TEXT_REPLACEMENTS,
    ShapePropertiesReplacement, ShapeTextReplacement, Slide, Snapshot,
};
pub use patch::{Conflict, History, Patch, Resolution, ThreeWayPlan};

/// Configuration of a primary edit (record 0351 §2.5): the bounds of its
/// [`Limits`]. The defaults are the finite [`Limits::default`].
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Limits,
}

impl EditOptions {
    /// Select the edit's bounds.
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    /// The edit's bounds.
    #[must_use]
    pub const fn limits_value(&self) -> Limits {
        self.limits
    }
}
pub use text_format::{
    CellFill, CellMargin, MAX_FONT_SIZE, MAX_TYPEFACE_BYTES, MIN_FONT_SIZE, Rgb, ShapeFill,
    TextAssignment, TextBody, TextProperties, TextTarget,
};
pub(crate) use text_format::{
    TextChange, stage_text_properties_with_limits, stage_text_property_chunks_with_limits,
    stage_text_structure_with_limits,
};
pub use transaction::{Commit, Diagnostics, Edit, MediaFile, ModernCommentChange, SmartArtParts};

pub(crate) use chunks::{
    ChunkAdmission, admit_shape_property_chunks, admit_shape_text_chunks,
    stage_shape_property_chunks_admitted_with_limits, stage_shape_property_chunks_with_limits,
    stage_shape_text_chunks_admitted_with_limits, stage_shape_text_chunks_with_limits,
};
pub(crate) use enclosed_removal::{
    stage_enclosed_shape_removal_with_limits, validate_enclosed_removal_pair_with_limits,
};
pub(crate) use geometry::geometry_skeleton_with_limits;
#[cfg(test)]
pub(crate) use model::fingerprint_log;
pub(crate) use model::{
    PartDigests, capture_with_parent_digests_and_semantic_limits,
    capture_with_provenance_and_semantic_limits, package_fingerprint,
    package_fingerprint_with_memo, relationships_equal,
};
pub(crate) use new_slide::{
    Planned as NewSlidePlaceholder, plan as plan_new_slide_placeholders,
    read_back_with_limits as read_back_new_slide, slide_xml_with_limits as new_slide_xml,
};
pub(crate) use patch::{apply_to_new_with_semantic_limits, apply_with_semantic_limits};
pub(crate) use properties::stage_shape_properties_with_limits;
pub(crate) use removal_proof::normalize_shape_removal_pair_with_limits;
pub(crate) use remove_plan::apply_patch_with_semantic_limits as apply_removal_patch_with_semantic_limits;
pub(crate) use remove_plan::{
    NON_POSITIONAL_EXTENSIONS, removal_layout_edge, validate_insertion_owner,
    validate_removal_presentation_owner, validate_removal_slide_surface,
};
pub(crate) use sections::SECTION_EXTENSION_URI;
pub(crate) use slide_properties::{
    normalize_slide_properties, slide_properties_skeleton, stage_common_slide_name,
    stage_slide_properties,
};
pub(crate) use transaction::{
    next_shape_id, stage_text_box, text_box_fragment, text_box_output_bound,
    validate_text_box_input,
};
pub(crate) use xml::{
    insert_slide_binding, reorder_slide_bindings, shape_text_skeleton_with_limits,
    stage_shape_removal_with_limits, stage_shape_texts_with_limits,
    strip_appended_text_boxes_with_limits, validate_slide_bindings,
};
pub(crate) use xml::{
    remove_slide_binding, slide_binding_removal_spans, validate_slide_catalog_delta,
};

#[cfg(test)]
mod owner_commit_tests;
