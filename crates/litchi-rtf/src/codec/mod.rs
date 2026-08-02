//! Bounded transport, parsing, and serialization mechanics.

pub(crate) mod compressed;
pub(crate) mod error;
pub(crate) mod lexer;
pub(crate) mod limits;
pub(crate) mod parser;
pub(crate) mod writer;

// Parser code predates the responsibility split and uses parent-relative
// names for shared retained-model vocabularies. Keep those dependencies
// explicit at this boundary while the codec remains private.
pub(crate) use crate::native::*;
pub(crate) use crate::{
    annotation, bookmark, border, document_variable, field, form_field, info, list,
    navigation_entry, object, picture, section, shape, stylesheet, table, types, user_property,
};
