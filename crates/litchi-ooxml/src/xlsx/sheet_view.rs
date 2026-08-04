//! XLSX worksheet-view host adapter.

use crate::error::{OoxmlError, Result};

pub use litchi_xlsx::sheet_view::{
    CellReference, Extension, Pane, PanePosition, PaneState, PivotArea, PivotAreaType,
    PivotSelection, PivotSelectionAxis, RangeReference, Selection, Sqref, View, ViewType, Views,
};

/// Parse the worksheet's sheetViews collection using the canonical XLSX owner.
///
/// The host-facing wrapper retains the historical OoxmlError return type
/// while the model and codec live in litchi-xlsx.
pub fn parse_worksheet_views(xml: &[u8]) -> Result<Option<Views>> {
    litchi_xlsx::sheet_view::parse_worksheet_views(xml).map_err(OoxmlError::from)
}
