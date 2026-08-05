//! Semantic worksheet-window values for legacy BIFF8.

use crate::error::{XlsError, XlsResult};

use super::SELECTION_RECORD_TYPE;

pub(super) fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Return whether a pane exists for the supplied horizontal and vertical
/// split axes.
pub(crate) const fn pane_exists(horizontal: u16, vertical: u16, pane: PaneType) -> bool {
    match pane {
        PaneType::LowerRight => horizontal > 0 && vertical > 0,
        PaneType::UpperRight => horizontal > 0,
        PaneType::LowerLeft => vertical > 0,
        PaneType::UpperLeft => true,
    }
}

/// Logical pane containing an active cell or selection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PaneType {
    LowerRight,
    UpperRight,
    LowerLeft,
    UpperLeft,
}

impl PaneType {
    pub(super) fn parse(value: u8, record_type: u16) -> XlsResult<Self> {
        match value {
            0 => Ok(Self::LowerRight),
            1 => Ok(Self::UpperRight),
            2 => Ok(Self::LowerLeft),
            3 => Ok(Self::UpperLeft),
            _ => Err(invalid(record_type, "pane type must be between 0 and 3")),
        }
    }

    pub(crate) fn code(self) -> u8 {
        match self {
            Self::LowerRight => 0,
            Self::UpperRight => 1,
            Self::LowerLeft => 2,
            Self::UpperLeft => 3,
        }
    }
}

/// Inclusive selected cell range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) first_column: u8,
    pub(super) last_column: u8,
}

impl Range {
    /// Create an inclusive range, rejecting inverted endpoints.
    pub fn new(
        first_row: u16,
        last_row: u16,
        first_column: u8,
        last_column: u8,
    ) -> XlsResult<Self> {
        if first_row > last_row || first_column > last_column {
            return Err(XlsError::InvalidData(
                "selection range endpoints must be ordered".to_string(),
            ));
        }
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }

    /// Create a one-cell range.
    pub const fn cell(row: u16, column: u8) -> Self {
        Self {
            first_row: row,
            last_row: row,
            first_column: column,
            last_column: column,
        }
    }

    pub const fn first_row(&self) -> u16 {
        self.first_row
    }
    pub const fn last_row(&self) -> u16 {
        self.last_row
    }
    pub const fn first_column(&self) -> u8 {
        self.first_column
    }
    pub const fn last_column(&self) -> u8 {
        self.last_column
    }
}

/// One BIFF8 `SELECTION` record in a worksheet window.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Selection {
    pub(super) pane: PaneType,
    pub(super) active_row: u16,
    pub(super) active_column: u8,
    pub(super) active_range_index: u16,
    pub(super) ranges: Vec<Range>,
}

impl Selection {
    pub fn pane(&self) -> PaneType {
        self.pane
    }
    pub fn active_row(&self) -> u16 {
        self.active_row
    }
    pub fn active_column(&self) -> u8 {
        self.active_column
    }
    pub fn active_range_index(&self) -> u16 {
        self.active_range_index
    }
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }
}

/// Frozen or split pane configuration for a worksheet window.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Pane {
    pub(super) horizontal_split: u16,
    pub(super) vertical_split: u16,
    pub(super) bottom_pane_top_row: u16,
    pub(super) right_pane_left_column: u8,
    pub(super) active_pane: PaneType,
}

impl Pane {
    /// Horizontal split in columns when frozen, otherwise in twips.
    pub fn horizontal_split(&self) -> u16 {
        self.horizontal_split
    }
    /// Vertical split in rows when frozen, otherwise in twips.
    pub fn vertical_split(&self) -> u16 {
        self.vertical_split
    }
    pub fn bottom_pane_top_row(&self) -> u16 {
        self.bottom_pane_top_row
    }
    pub fn right_pane_left_column(&self) -> u8 {
        self.right_pane_left_column
    }
    pub fn active_pane(&self) -> PaneType {
        self.active_pane
    }
}

/// Display and navigation state for one worksheet window.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct View {
    pub(super) flags: u16,
    pub(super) first_visible_row: u16,
    pub(super) first_visible_column: u8,
    pub(super) gridline_color_index: u16,
    pub(super) page_break_zoom_percent: Option<u16>,
    pub(super) normal_zoom_percent: Option<u16>,
    pub(super) zoom_fraction: Option<(u16, u16)>,
    pub(super) pane: Option<Pane>,
    pub(super) selections: Vec<Selection>,
}

impl View {
    pub fn shows_formulas(&self) -> bool {
        self.flags & 0x0001 != 0
    }
    pub fn shows_gridlines(&self) -> bool {
        self.flags & 0x0002 != 0
    }
    pub fn shows_row_column_headers(&self) -> bool {
        self.flags & 0x0004 != 0
    }
    pub fn has_frozen_panes(&self) -> bool {
        self.flags & 0x0008 != 0
    }
    pub fn shows_zero_values(&self) -> bool {
        self.flags & 0x0010 != 0
    }
    pub fn uses_default_gridline_color(&self) -> bool {
        self.flags & 0x0020 != 0
    }
    pub fn is_right_to_left(&self) -> bool {
        self.flags & 0x0040 != 0
    }
    pub fn shows_outline_symbols(&self) -> bool {
        self.flags & 0x0080 != 0
    }
    pub fn is_frozen_without_split(&self) -> bool {
        self.flags & 0x0100 != 0
    }
    pub fn is_selected(&self) -> bool {
        self.flags & 0x0200 != 0
    }
    pub fn is_displayed(&self) -> bool {
        self.flags & 0x0400 != 0
    }
    pub fn is_page_break_preview(&self) -> bool {
        self.flags & 0x0800 != 0
    }
    pub fn first_visible_row(&self) -> u16 {
        self.first_visible_row
    }
    pub fn first_visible_column(&self) -> u8 {
        self.first_visible_column
    }
    pub fn gridline_color_index(&self) -> u16 {
        self.gridline_color_index
    }
    pub fn page_break_zoom_percent(&self) -> Option<u16> {
        self.page_break_zoom_percent
    }
    pub fn normal_zoom_percent(&self) -> Option<u16> {
        self.normal_zoom_percent
    }
    pub fn zoom_fraction(&self) -> Option<(u16, u16)> {
        self.zoom_fraction
    }
    pub fn pane(&self) -> Option<&Pane> {
        self.pane.as_ref()
    }
    pub fn selections(&self) -> &[Selection] {
        &self.selections
    }

    pub(super) fn validate_selection_groups(&self) -> XlsResult<()> {
        let mut start = 0;
        while start < self.selections.len() {
            let first = &self.selections[start];
            let mut end = start + 1;
            let mut range_count = first.ranges.len();
            while end < self.selections.len() && self.selections[end].pane == first.pane {
                range_count = range_count
                    .checked_add(self.selections[end].ranges.len())
                    .ok_or_else(|| {
                        invalid(
                            SELECTION_RECORD_TYPE,
                            "SELECTION range aggregation overflows",
                        )
                    })?;
                end += 1;
            }
            let active_index = usize::from(first.active_range_index);
            if active_index >= range_count {
                return Err(invalid(
                    SELECTION_RECORD_TYPE,
                    "SELECTION active range index is outside its contiguous range aggregation",
                ));
            }
            let mut remaining = active_index;
            let mut active_range = None;
            for selection in &self.selections[start..end] {
                if remaining < selection.ranges.len() {
                    active_range = selection.ranges.get(remaining);
                    break;
                }
                remaining -= selection.ranges.len();
            }
            let range = active_range.ok_or_else(|| {
                invalid(
                    SELECTION_RECORD_TYPE,
                    "SELECTION active range could not be resolved",
                )
            })?;
            if first.active_row < range.first_row
                || first.active_row > range.last_row
                || first.active_column < range.first_column
                || first.active_column > range.last_column
            {
                return Err(invalid(
                    SELECTION_RECORD_TYPE,
                    "SELECTION active range does not contain the active cell",
                ));
            }
            start = end;
        }
        Ok(())
    }
}
