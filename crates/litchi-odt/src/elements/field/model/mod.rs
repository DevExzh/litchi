//! Stable facade for the contextual ODF field semantic model.

#![allow(
    clippy::wildcard_imports,
    reason = "the model facade intentionally gathers contextual field owners"
)]

use super::*;
use crate::elements::element::{Element, ElementBase};
use litchi_core::{Error, Result};

mod common;
mod database;
mod document;
mod dynamic;
mod metadata;
mod reference;
mod values;

pub(super) use common::{
    is_xml_1_0_char, parse_drop_down_boolean, push_xml_attribute, push_xml_text, set_data_style,
    validate_double, validate_dynamic_value,
};
pub(super) use document::{
    validate_xml_schema_date, validate_xml_schema_date_time, validate_xml_schema_time,
};
pub(super) use metadata::{
    MetaContentGrammar, add_meta_size, is_allowed_meta_namespace, meta_child_grammar,
    validate_document_metadata_value, validate_meta_element_parts, validate_xml_id,
};

pub use database::*;
pub use document::*;
pub use dynamic::*;
pub use metadata::*;
pub use reference::*;
pub use values::*;
