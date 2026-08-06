//! PresentationML package-level programmable-tag ownership.
//!
//! The facade keeps package discovery and snapshot edits ergonomic while the
//! implementation is split into typed state, XML codecs, and bounded package
//! validation.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{discover, load, put, remove};

#[cfg(test)]
pub(crate) use codec::process_owner_ooxml;
pub(crate) use codec::{process_pptx_ooxml, replace_xml, staged_xml};
pub(crate) use model::Anchor;
pub(crate) use validation::{
    available_part_name, available_relationship_id, has_other_inbound, relationship_namespace,
    resolve_anchor, validate_relative_target, validate_selected_relationship,
};
