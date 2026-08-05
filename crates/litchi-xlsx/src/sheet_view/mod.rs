//! Semantic worksheet-view facade.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::parse_worksheet_views;
pub use model::{
    CellReference, Extension, Pane, PanePosition, PaneState, PivotArea, PivotAreaType,
    PivotSelection, PivotSelectionAxis, RangeReference, Selection, Sqref, View, ViewType, Views,
};

use crate::error::Error;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
