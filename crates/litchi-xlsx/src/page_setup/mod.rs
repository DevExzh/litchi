//! Typed worksheet page-setup metadata and authoring.

mod codec;
mod model;
mod patch;
mod snapshot;
pub mod source;

#[cfg(test)]
mod tests;

pub use codec::{
    parse_worksheet_page_setup, parse_worksheet_page_setup_relationship_id,
    replace_worksheet_page_setup, write_page_setup,
};
pub use model::*;
pub(crate) use patch::Patch;
pub use snapshot::Snapshot;
