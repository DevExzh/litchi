//! Worksheet view definitions for Excel.
//!
//! This module provides data structures for worksheet view settings such as
//! zoom, right-to-left mode, and the active view type.

/// Worksheet view type.
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
    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::PageBreakPreview => "pageBreakPreview",
            Self::PageLayout => "pageLayout",
        }
    }

    pub(crate) fn parse(value: &str) -> Option<Self> {
        match value {
            "normal" => Some(Self::Normal),
            "pageBreakPreview" => Some(Self::PageBreakPreview),
            "pageLayout" => Some(Self::PageLayout),
            _ => None,
        }
    }
}

/// Worksheet view configuration.
#[derive(Debug, Clone, Default)]
pub struct SheetView {
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
    /// View type.
    pub view_type: Option<SheetViewType>,
    /// Top-left visible cell.
    pub top_left_cell: Option<String>,
    /// Zoom scale (10-400).
    pub zoom_scale: Option<u16>,
    /// Zoom scale for normal view.
    pub zoom_scale_normal: Option<u16>,
}
