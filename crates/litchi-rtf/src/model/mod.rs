//! Retained native RTF snapshot storage.

pub(crate) mod document;
pub(crate) mod types;

pub(crate) use crate::{
    annotation, bookmark, border, compressed, document_variable, error, field, form_field, info,
    lexer, limits, list, navigation_entry, object, parser, picture, section, shape, stylesheet,
    table, user_property,
};
pub(crate) use types::DocumentElement;
