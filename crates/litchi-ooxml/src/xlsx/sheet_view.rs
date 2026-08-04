//! Compatibility exports for the canonical XLSX worksheet-view codec.

use crate::error::{OoxmlError, Result};

pub use litchi_xlsx::sheet_view::{
    PivotAreaType, PivotSelectionAxis, WorksheetCellReference, WorksheetPanePosition,
    WorksheetPaneState, WorksheetPivotArea, WorksheetPivotSelection, WorksheetRangeReference,
    WorksheetViewCollection, WorksheetViewDefinition, WorksheetViewExtension, WorksheetViewPane,
    WorksheetViewSelection, WorksheetViewSqref, WorksheetViewType,
};

/// Parse the worksheet's sheetViews collection using the canonical XLSX owner.
///
/// The host-facing wrapper retains the historical OoxmlError return type
/// while the model and codec live in litchi-xlsx.
pub fn parse_worksheet_views(xml: &[u8]) -> Result<Option<WorksheetViewCollection>> {
    litchi_xlsx::sheet_view::parse_worksheet_views(xml).map_err(OoxmlError::from)
}
