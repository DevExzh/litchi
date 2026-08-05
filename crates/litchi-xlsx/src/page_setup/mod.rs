//! Typed worksheet page-setup metadata and authoring.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{parse_worksheet_page_setup, parse_worksheet_page_setup_relationship_id};
pub use model::*;
