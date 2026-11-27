//! `ActiveX` XML and relationship codecs.
//!
//! XML mutation stays byte-preserving for the worksheet owner; relationship
//! traversal is kept separate so package transactions can consume a compact,
//! deterministic dependency list.

mod relationships;
mod xml;

pub(crate) use relationships::descriptor_relationship_ids;
pub use xml::replace_controls_xml;
pub(crate) use xml::{controls_span, relationship_ids_in_xml};
