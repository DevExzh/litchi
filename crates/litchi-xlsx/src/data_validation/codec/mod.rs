//! Layered SpreadsheetML data-validation codec facade.
//!
//! Wire/XML boundaries live in `wire`, parsing in `parser`, semantic
//! invariants in `validation`, and canonical serialization in `writer`.
#![allow(
    clippy::module_inception,
    reason = "the nested codec module distinguishes the codec facade from its implementation"
)]

mod parser;
mod validation;
mod wire;
mod writer;

#[cfg(test)]
mod tests;

pub use parser::parse_data_validation_collections;
pub use validation::validate_data_validation_collections;
pub use writer::{
    write_data_validation_collections, write_data_validation_core, write_data_validation_extensions,
};

pub(crate) use validation::{
    validate_collection, validate_optional_text, validate_rule, validate_text,
};
pub(crate) use wire::{
    BoundedXml, append_bounded_bytes, exact, invalid, optional_attr, parse_sqref, reserve_vec,
    spreadsheet, xml_error,
};
