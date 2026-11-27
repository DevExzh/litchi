//! Chartsheet view children from `CT_ChartsheetView` and `CT_CustomChartsheetView`.

/// Visibility state used by workbook sheet entries and custom chartsheet views.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum State {
    Visible,
    Hidden,
    VeryHidden,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Default,
    Portrait,
    Landscape,
}

/// One `sheetView` child of the chartsheet's required `sheetViews` collection.
///
/// Chartsheet views have a smaller attribute set than worksheet views in
/// ISO/IEC 29500-1 `CT_ChartsheetView`: tab selection, zoom, the workbook-view
/// index, and zoom-to-fit are the complete typed surface here.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct View {
    pub tab_selected: Option<bool>,
    pub zoom_scale: Option<u32>,
    pub workbook_view_id: u32,
    pub zoom_to_fit: Option<bool>,
}

/// One saved chartsheet view from `CT_CustomChartsheetView`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomView {
    /// Braced UUID lexical form required by `SpreadsheetML` `ST_Guid`.
    pub guid: String,
    pub scale: Option<u32>,
    pub state: Option<State>,
    pub zoom_to_fit: Option<bool>,
}
