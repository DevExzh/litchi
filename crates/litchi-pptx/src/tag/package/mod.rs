//! `PresentationML` package-level programmable-tag ownership.
//!
//! The facade keeps package discovery and snapshot edits ergonomic while the
//! implementation is split into typed state, XML codecs, and bounded package
//! validation.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{discover, load};

#[cfg(test)]
pub(crate) use codec::process_owner_ooxml;
pub(crate) use codec::{
    discover_with_limits, load_with_limits, process_pptx_ooxml, put_with_limits,
    remove_with_limits, replace_xml, staged_xml, validate_replay_part,
};
#[cfg(test)]
pub(crate) use codec::{put, remove};
pub(crate) use model::Anchor;
pub(crate) use validation::{
    available_part_name, available_relationship_id, has_other_inbound, relationship_namespace,
    resolve_anchor_with_limits, validate_relative_target, validate_selected_relationship,
};
