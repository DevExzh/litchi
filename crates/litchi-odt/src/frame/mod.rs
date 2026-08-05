//! Contextual ODT frame facade.
//!
//! Geometry, anchors, and bounded image resources are owned by
//! `litchi-odf-common`; this module retains only ODT element construction.

mod xml;

pub use litchi_odf_common::drawing::authoring::{Anchor, Frame, Length};
pub use litchi_odf_common::media::authoring::Format;

pub(crate) use litchi_odf_common::media::authoring::{
    Part, allocate_picture_path, validate_payload,
};
pub(crate) use xml::{image_element, text_box_element};
