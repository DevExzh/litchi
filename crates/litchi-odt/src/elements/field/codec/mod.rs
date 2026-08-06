//! Namespace-aware ODF field XML codec facade.

mod semantic;
#[cfg(test)]
mod tests;
mod validation;
mod wire;

pub(super) use validation::checked_field_depth;
pub(super) use wire::{
    parse_database_fields, parse_drop_down_fields, parse_meta_fields, parse_note_body_contents,
};
