//! Shared `SpreadsheetML` worksheet-chart integration.

pub mod anchor;
pub mod builder;
mod codec;
mod model;
mod pivot;
pub mod relationship;

pub use codec::{decode, read, write, write_with_external_data_id};

#[cfg(test)]
mod tests;
