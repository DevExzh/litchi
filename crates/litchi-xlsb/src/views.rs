//! Typed worksheet-view values for SpreadsheetML.
//!
//! The values in this module mirror the worksheet `sheetView`, `pane`, and
//! `selection` model. XML parsing and writing remain in the format adapter;
//! these types deliberately contain no package or parser state.

/// Worksheet view type used by the `sheetView@view` attribute.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetViewType {
    /// Normal worksheet view.
    Normal,
    /// Page break preview view.
    PageBreakPreview,
    /// Page layout view.
    PageLayout,
}

impl SheetViewType {
    /// Returns the exact SpreadsheetML token for the `view` attribute.
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::PageBreakPreview => "pageBreakPreview",
            Self::PageLayout => "pageLayout",
        }
    }

    /// Parses an exact SpreadsheetML `view` attribute token.
    #[must_use]
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "normal" => Some(Self::Normal),
            "pageBreakPreview" => Some(Self::PageBreakPreview),
            "pageLayout" => Some(Self::PageLayout),
            _ => None,
        }
    }
}

/// Pane position within a split or frozen worksheet view.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetPanePosition {
    /// Lower-right pane.
    BottomRight,
    /// Upper-right pane.
    TopRight,
    /// Lower-left pane.
    BottomLeft,
    /// Upper-left pane.
    TopLeft,
}

impl SheetPanePosition {
    /// Returns the exact SpreadsheetML token for `pane@activePane` or
    /// `selection@pane`.
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::BottomRight => "bottomRight",
            Self::TopRight => "topRight",
            Self::BottomLeft => "bottomLeft",
            Self::TopLeft => "topLeft",
        }
    }

    /// Parses an exact SpreadsheetML pane-position token.
    #[must_use]
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "bottomRight" => Some(Self::BottomRight),
            "topRight" => Some(Self::TopRight),
            "bottomLeft" => Some(Self::BottomLeft),
            "topLeft" => Some(Self::TopLeft),
            _ => None,
        }
    }
}

/// Split/freeze state for a worksheet pane.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetPaneState {
    /// A movable split pane.
    Split,
    /// Frozen rows and/or columns.
    Frozen,
    /// A frozen pane that also contains a split.
    FrozenSplit,
}

impl SheetPaneState {
    /// Returns the exact SpreadsheetML token for the `pane@state` attribute.
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Split => "split",
            Self::Frozen => "frozen",
            Self::FrozenSplit => "frozenSplit",
        }
    }

    /// Parses an exact SpreadsheetML pane-state token.
    #[must_use]
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "split" => Some(Self::Split),
            "frozen" => Some(Self::Frozen),
            "frozenSplit" => Some(Self::FrozenSplit),
            _ => None,
        }
    }
}

/// Pane configuration within a worksheet view.
#[derive(Debug, Clone, Default)]
pub struct SheetPane {
    /// Horizontal split position.
    pub x_split: Option<f64>,
    /// Vertical split position.
    pub y_split: Option<f64>,
    /// First visible cell in the lower-right pane.
    pub top_left_cell: Option<String>,
    /// Pane that currently has focus.
    pub active_pane: Option<SheetPanePosition>,
    /// Split/freeze state.
    pub state: Option<SheetPaneState>,
}

/// Cell selection within a worksheet pane.
#[derive(Debug, Clone, Default)]
pub struct SheetSelection {
    /// Pane containing the selection.
    pub pane: Option<SheetPanePosition>,
    /// Active cell reference.
    pub active_cell: Option<String>,
    /// Index of the active cell within `sqref`.
    pub active_cell_id: Option<u32>,
    /// Selected cell references.
    pub sqref: Option<String>,
}

/// Worksheet view configuration.
#[derive(Debug, Clone, Default)]
pub struct SheetView {
    /// Workbook window this view belongs to.
    pub workbook_view_id: Option<u32>,
    /// Whether the view window is protected.
    pub window_protection: Option<bool>,
    /// Show formulas instead of values.
    pub show_formulas: Option<bool>,
    /// Show grid lines.
    pub show_grid_lines: Option<bool>,
    /// Show row/column headers.
    pub show_row_col_headers: Option<bool>,
    /// Show zero values.
    pub show_zeros: Option<bool>,
    /// Right-to-left display.
    pub right_to_left: Option<bool>,
    /// Whether this worksheet tab is selected in the workbook window.
    pub tab_selected: Option<bool>,
    /// Show the ruler in page-layout view.
    pub show_ruler: Option<bool>,
    /// Show worksheet outline symbols.
    pub show_outline_symbols: Option<bool>,
    /// Use the system default grid color.
    pub default_grid_color: Option<bool>,
    /// Show white space in page-layout view.
    pub show_white_space: Option<bool>,
    /// View type.
    pub view_type: Option<SheetViewType>,
    /// Top-left visible cell.
    pub top_left_cell: Option<String>,
    /// Indexed grid-line color.
    pub color_id: Option<u32>,
    /// Zoom scale (10-400).
    pub zoom_scale: Option<u16>,
    /// Zoom scale for normal view.
    pub zoom_scale_normal: Option<u16>,
    /// Zoom scale for page-break preview.
    pub zoom_scale_sheet_layout_view: Option<u16>,
    /// Zoom scale for page-layout view.
    pub zoom_scale_page_layout_view: Option<u16>,
    /// Optional split or frozen pane.
    pub pane: Option<SheetPane>,
    /// Selections associated with this view's panes.
    pub selections: Vec<SheetSelection>,
}

#[cfg(test)]
mod tests {
    use super::{SheetPanePosition, SheetPaneState, SheetViewType};

    #[test]
    fn view_type_tokens_round_trip() {
        for (value, token) in [
            (SheetViewType::Normal, "normal"),
            (SheetViewType::PageBreakPreview, "pageBreakPreview"),
            (SheetViewType::PageLayout, "pageLayout"),
        ] {
            assert_eq!(value.as_str(), token);
            assert_eq!(SheetViewType::parse(token), Some(value));
        }
        assert_eq!(SheetViewType::parse("pageLayoutView"), None);
    }

    #[test]
    fn pane_position_tokens_round_trip() {
        for (value, token) in [
            (SheetPanePosition::BottomRight, "bottomRight"),
            (SheetPanePosition::TopRight, "topRight"),
            (SheetPanePosition::BottomLeft, "bottomLeft"),
            (SheetPanePosition::TopLeft, "topLeft"),
        ] {
            assert_eq!(value.as_str(), token);
            assert_eq!(SheetPanePosition::parse(token), Some(value));
        }
        assert_eq!(SheetPanePosition::parse("center"), None);
    }

    #[test]
    fn pane_state_tokens_round_trip() {
        for (value, token) in [
            (SheetPaneState::Split, "split"),
            (SheetPaneState::Frozen, "frozen"),
            (SheetPaneState::FrozenSplit, "frozenSplit"),
        ] {
            assert_eq!(value.as_str(), token);
            assert_eq!(SheetPaneState::parse(token), Some(value));
        }
        assert_eq!(SheetPaneState::parse("freeze"), None);
    }
}
