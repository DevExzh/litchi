use crate::error::{XlsError, XlsResult};
use crate::{XlsPaneType, XlsSelectionRange};

const MAX_SCL_TERM: u16 = i16::MAX as u16;
pub(crate) const MAX_SELECTION_RANGES: usize = 1_369;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsViewScale {
    pub numerator: u16,
    pub denominator: u16,
}

impl XlsViewScale {
    pub fn new(numerator: u16, denominator: u16) -> XlsResult<Self> {
        let value = Self {
            numerator,
            denominator,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn validate(self) -> XlsResult<()> {
        if self.numerator == 0
            || self.denominator == 0
            || self.numerator > MAX_SCL_TERM
            || self.denominator > MAX_SCL_TERM
            || u32::from(self.numerator) * 10 < u32::from(self.denominator)
            || u32::from(self.numerator) > u32::from(self.denominator) * 4
        {
            return Err(XlsError::InvalidData(
                "view scale terms must be 1..=32767 and the fraction must be between 1/10 and 4"
                    .to_string(),
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsPaneMode {
    Frozen,
    Split,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsWorksheetPaneOptions {
    pub mode: XlsPaneMode,
    pub horizontal_split: u16,
    pub vertical_split: u16,
    pub bottom_pane_top_row: u16,
    pub right_pane_left_column: u8,
    pub active_pane: XlsPaneType,
}

impl XlsWorksheetPaneOptions {
    pub fn frozen(rows: u16, columns: u8) -> XlsResult<Self> {
        let value = Self {
            mode: XlsPaneMode::Frozen,
            horizontal_split: u16::from(columns),
            vertical_split: rows,
            bottom_pane_top_row: rows,
            right_pane_left_column: columns,
            active_pane: active_pane(columns > 0, rows > 0),
        };
        value.validate()?;
        Ok(value)
    }

    pub fn split(
        horizontal_twips: u16,
        vertical_twips: u16,
        bottom_pane_top_row: u16,
        right_pane_left_column: u8,
        active_pane: XlsPaneType,
    ) -> XlsResult<Self> {
        let value = Self {
            mode: XlsPaneMode::Split,
            horizontal_split: horizontal_twips,
            vertical_split: vertical_twips,
            bottom_pane_top_row,
            right_pane_left_column,
            active_pane,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn validate(self) -> XlsResult<()> {
        if self.horizontal_split == 0 && self.vertical_split == 0 {
            return Err(XlsError::InvalidData(
                "PANE must split at least one axis".to_string(),
            ));
        }
        if self.mode == XlsPaneMode::Split
            && (self.horizontal_split > i16::MAX as u16 || self.vertical_split > i16::MAX as u16)
        {
            return Err(XlsError::InvalidData(
                "split positions exceed 32767 twips".to_string(),
            ));
        }
        if self.mode == XlsPaneMode::Frozen && self.horizontal_split > 255 {
            return Err(XlsError::InvalidData(
                "frozen horizontal split exceeds 255 columns".to_string(),
            ));
        }
        if !pane_exists(self.horizontal_split, self.vertical_split, self.active_pane) {
            return Err(XlsError::InvalidData(
                "active pane does not exist for the split axes".to_string(),
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsWorksheetSelectionOptions {
    pub pane: XlsPaneType,
    pub active_row: u16,
    pub active_column: u8,
    pub active_range_index: u16,
    pub ranges: Vec<XlsSelectionRange>,
}

impl XlsWorksheetSelectionOptions {
    pub fn single_cell(pane: XlsPaneType, row: u16, column: u8) -> Self {
        Self {
            pane,
            active_row: row,
            active_column: column,
            active_range_index: 0,
            ranges: vec![XlsSelectionRange::new(row, row, column, column)],
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsWorksheetViewOptions {
    pub show_formulas: bool,
    pub show_gridlines: bool,
    pub show_row_column_headers: bool,
    pub show_zero_values: bool,
    /// `None` uses the window default color; custom palette indices are 0..=63.
    pub gridline_color_index: Option<u16>,
    pub right_to_left: bool,
    pub show_outline_symbols: bool,
    pub selected: bool,
    pub displayed: bool,
    pub page_break_preview: bool,
    pub first_visible_row: u16,
    pub first_visible_column: u8,
    pub page_break_zoom_percent: Option<u16>,
    pub normal_zoom_percent: Option<u16>,
    pub scale: Option<XlsViewScale>,
    pub pane: Option<XlsWorksheetPaneOptions>,
    pub selections: Vec<XlsWorksheetSelectionOptions>,
}

impl Default for XlsWorksheetViewOptions {
    fn default() -> Self {
        Self {
            show_formulas: false,
            show_gridlines: true,
            show_row_column_headers: true,
            show_zero_values: true,
            gridline_color_index: None,
            right_to_left: false,
            show_outline_symbols: true,
            selected: true,
            displayed: true,
            page_break_preview: false,
            first_visible_row: 0,
            first_visible_column: 0,
            page_break_zoom_percent: None,
            normal_zoom_percent: None,
            scale: None,
            pane: None,
            selections: vec![XlsWorksheetSelectionOptions::single_cell(
                XlsPaneType::UpperLeft,
                0,
                0,
            )],
        }
    }
}

impl XlsWorksheetViewOptions {
    pub(crate) fn validate(&self) -> XlsResult<()> {
        if self.gridline_color_index.is_some_and(|value| value > 63) {
            return Err(XlsError::InvalidData(
                "custom gridline color index must be <= 63".to_string(),
            ));
        }
        for zoom in [self.page_break_zoom_percent, self.normal_zoom_percent]
            .into_iter()
            .flatten()
        {
            if !(10..=400).contains(&zoom) {
                return Err(XlsError::InvalidData(
                    "view zoom percent must be between 10 and 400".to_string(),
                ));
            }
        }
        if let Some(scale) = self.scale {
            scale.validate()?;
        }
        if let Some(pane) = self.pane {
            pane.validate()?;
            if pane.mode == XlsPaneMode::Frozen
                && (self.first_visible_row == u16::MAX || self.first_visible_column == u8::MAX)
            {
                return Err(XlsError::InvalidData(
                    "frozen view cannot use sentinel origins".to_string(),
                ));
            }
        }
        validate_selections(self.pane.as_ref(), &self.selections)
    }

    pub(crate) fn set_frozen(&mut self, rows: u16, columns: u8) -> XlsResult<()> {
        let pane = XlsWorksheetPaneOptions::frozen(rows, columns)?;
        self.selections = vec![XlsWorksheetSelectionOptions::single_cell(
            pane.active_pane,
            rows,
            columns,
        )];
        self.pane = Some(pane);
        Ok(())
    }

    pub(crate) fn clear_pane(&mut self) {
        self.pane = None;
        self.selections = vec![XlsWorksheetSelectionOptions::single_cell(
            XlsPaneType::UpperLeft,
            0,
            0,
        )];
    }
}

fn validate_selections(
    pane: Option<&XlsWorksheetPaneOptions>,
    selections: &[XlsWorksheetSelectionOptions],
) -> XlsResult<()> {
    if selections.is_empty() {
        return Ok(());
    }
    let active_pane = pane.map_or(XlsPaneType::UpperLeft, |value| value.active_pane);
    let mut has_active = false;
    let mut seen = Vec::new();
    let mut start = 0usize;
    while start < selections.len() {
        let first = &selections[start];
        if first.active_range_index & 0x8000 != 0 {
            return Err(XlsError::InvalidData(
                "active range index must be a signed nonnegative integer".to_string(),
            ));
        }
        if seen.contains(&first.pane) {
            return Err(XlsError::InvalidData(
                "SELECTION pane groups must be contiguous".to_string(),
            ));
        }
        if !pane.map_or(first.pane == XlsPaneType::UpperLeft, |value| {
            pane_exists(value.horizontal_split, value.vertical_split, first.pane)
        }) {
            return Err(XlsError::InvalidData(
                "SELECTION references a nonexistent pane".to_string(),
            ));
        }
        has_active |= first.pane == active_pane;
        let mut end = start;
        let mut count = 0usize;
        while end < selections.len() && selections[end].pane == first.pane {
            let current = &selections[end];
            if current.ranges.is_empty() || current.ranges.len() > MAX_SELECTION_RANGES {
                return Err(XlsError::InvalidData(
                    "each SELECTION needs 1..=1369 ranges".to_string(),
                ));
            }
            if (
                current.active_row,
                current.active_column,
                current.active_range_index,
            ) != (
                first.active_row,
                first.active_column,
                first.active_range_index,
            ) {
                return Err(XlsError::InvalidData(
                    "contiguous SELECTION records disagree".to_string(),
                ));
            }
            if current.ranges.iter().any(|range| {
                range.first_row() > range.last_row() || range.first_column() > range.last_column()
            }) {
                return Err(XlsError::InvalidData(
                    "SELECTION contains an inverted range".to_string(),
                ));
            }
            count = count.checked_add(current.ranges.len()).ok_or_else(|| {
                XlsError::InvalidData("SELECTION range count overflow".to_string())
            })?;
            end += 1;
        }
        let active = selections[start..end]
            .iter()
            .flat_map(|selection| selection.ranges.iter())
            .nth(usize::from(first.active_range_index))
            .ok_or_else(|| {
                XlsError::InvalidData("active range index is out of bounds".to_string())
            })?;
        if first.active_row < active.first_row()
            || first.active_row > active.last_row()
            || first.active_column < active.first_column()
            || first.active_column > active.last_column()
        {
            return Err(XlsError::InvalidData(
                "active range does not contain the active cell".to_string(),
            ));
        }
        seen.push(first.pane);
        start = end;
    }
    if !has_active {
        return Err(XlsError::InvalidData(
            "active pane has no SELECTION".to_string(),
        ));
    }
    Ok(())
}

fn active_pane(has_columns: bool, has_rows: bool) -> XlsPaneType {
    match (has_columns, has_rows) {
        (true, true) => XlsPaneType::LowerRight,
        (true, false) => XlsPaneType::UpperRight,
        (false, true) => XlsPaneType::LowerLeft,
        (false, false) => XlsPaneType::UpperLeft,
    }
}

pub(crate) fn pane_exists(x: u16, y: u16, pane: XlsPaneType) -> bool {
    match pane {
        XlsPaneType::LowerRight => x > 0 && y > 0,
        XlsPaneType::UpperRight => x > 0,
        XlsPaneType::LowerLeft => y > 0,
        XlsPaneType::UpperLeft => true,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::writer::biff;

    #[test]
    fn serializes_exact_view_records() {
        let pane =
            XlsWorksheetPaneOptions::split(1_200, 800, 7, 4, XlsPaneType::LowerRight).unwrap();
        let options = XlsWorksheetViewOptions {
            show_formulas: true,
            show_gridlines: false,
            gridline_color_index: Some(8),
            first_visible_row: 2,
            first_visible_column: 1,
            pane: Some(pane),
            selections: vec![XlsWorksheetSelectionOptions::single_cell(
                XlsPaneType::LowerRight,
                7,
                4,
            )],
            ..XlsWorksheetViewOptions::default()
        };
        let mut window = Vec::new();
        let mut pane_bytes = Vec::new();
        let mut selection = Vec::new();
        biff::write_window2_options(&mut window, &options).unwrap();
        biff::write_pane_options(&mut pane_bytes, &pane).unwrap();
        biff::write_selection_options(&mut selection, &options.selections[0]).unwrap();
        assert_eq!(
            window,
            vec![
                0x3e, 0x02, 18, 0, 0x95, 0x06, 2, 0, 1, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
            ]
        );
        assert_eq!(
            pane_bytes,
            vec![0x41, 0, 10, 0, 0xb0, 4, 0x20, 3, 7, 0, 4, 0, 0, 0]
        );
        assert_eq!(
            selection,
            vec![0x1d, 0, 15, 0, 0, 7, 0, 4, 0, 0, 0, 1, 0, 7, 0, 7, 0, 4, 4]
        );
    }

    #[test]
    fn rejects_inconsistent_view_options() {
        let mut options = XlsWorksheetViewOptions::default();
        options.selections[0].pane = XlsPaneType::LowerRight;
        assert!(options.validate().is_err());
        assert!(XlsViewScale::new(32_768, 32_768).is_err());
        assert!(XlsWorksheetPaneOptions::split(32_768, 0, 0, 0, XlsPaneType::UpperRight).is_err());

        let mut no_selection = XlsWorksheetViewOptions::default();
        no_selection.selections.clear();
        assert!(no_selection.validate().is_ok());

        let invalid_frozen = XlsWorksheetPaneOptions {
            mode: XlsPaneMode::Frozen,
            horizontal_split: 256,
            vertical_split: 0,
            bottom_pane_top_row: 0,
            right_pane_left_column: 0,
            active_pane: XlsPaneType::UpperRight,
        };
        assert!(invalid_frozen.validate().is_err());

        let mut signed_index = XlsWorksheetViewOptions::default();
        signed_index.selections[0].active_range_index = 0x8000;
        assert!(signed_index.validate().is_err());

        let mut too_many_ranges = XlsWorksheetViewOptions::default();
        too_many_ranges.selections[0].ranges =
            vec![XlsSelectionRange::new(0, 0, 0, 0); MAX_SELECTION_RANGES + 1];
        assert!(too_many_ranges.validate().is_err());
    }
}
