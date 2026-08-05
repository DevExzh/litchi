//! ODF inline-element grammar for ruby annotations.
//!
//! The generated allowed-child tables live in the semantic model layer. The
//! parent ruby parser imports this narrow facade so namespace and parser state
//! remain owned by the ruby-family module.

mod model;

pub(super) use model::{is_hyperlink_child, is_ruby_base_child};
