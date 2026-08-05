//! Shared external-workbook relationship vocabulary.

mod codec;
mod model;

pub use codec::is_external_workbook_relationship;
pub use model::EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES;

#[cfg(test)]
mod tests;
