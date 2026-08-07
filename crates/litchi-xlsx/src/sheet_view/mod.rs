//! Semantic worksheet-view facade.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::parse_worksheet_views;
pub use model::{
    Collection, Entry, Extension, PivotArea, PivotAreaType, PivotSelection, PivotSelectionAxis,
};

use crate::error::Error;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
