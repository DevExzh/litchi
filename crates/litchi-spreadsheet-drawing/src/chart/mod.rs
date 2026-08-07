//! Shared `SpreadsheetML` worksheet-chart integration.

pub mod anchor;
mod codec;
mod model;
mod pivot;
pub mod relationship;

pub use anchor::Anchor;
pub use codec::{decode, read, write, write_with_external_data_id};
pub use model::{Chart, Series};
pub use relationship::{
    ExternalDataPart, ExternalDataTarget, Relationship, Target, UserShapesPart,
};

#[cfg(test)]
mod tests;
