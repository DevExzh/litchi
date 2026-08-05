use crate::error::{XlsError, XlsResult};
use crate::view::{PaneType, Range, pane_exists};

const MAX_SCL_TERM: u16 = i16::MAX as u16;
pub(crate) const MAX_SELECTION_RANGES: usize = 1_369;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// A checked BIFF8 worksheet zoom fraction between 10% and 400%.
pub struct Scale {
    numerator: u16,
    denominator: u16,
}

impl Scale {
    /// Create a scale from positive terms no larger than 32767.
    pub fn new(numerator: u16, denominator: u16) -> XlsResult<Self> {
        let value = Self {
            numerator,
            denominator,
        };
        value.validate()?;
        Ok(value)
    }

    /// Return the fraction numerator.
    pub const fn numerator(self) -> u16 {
        self.numerator
    }

    /// Return the fraction denominator.
    pub const fn denominator(self) -> u16 {
        self.denominator
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
/// How a worksheet pane divides its window.
pub enum Mode {
    /// Rows and columns remain visible while scrolling.
    Frozen,
    /// The window is divided at twip positions.
    Split,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// A checked frozen or split worksheet pane.
pub struct Pane {
    mode: Mode,
    horizontal_split: u16,
    vertical_split: u16,
    bottom_pane_top_row: u16,
    right_pane_left_column: u8,
    active_pane: PaneType,
}

impl Pane {
    /// Freeze `rows` rows and `columns` columns from the top-left corner.
    pub fn frozen(rows: u16, columns: u8) -> XlsResult<Self> {
        let value = Self {
            mode: Mode::Frozen,
            horizontal_split: u16::from(columns),
            vertical_split: rows,
            bottom_pane_top_row: rows,
            right_pane_left_column: columns,
            active_pane: active_pane(columns > 0, rows > 0),
        };
        value.validate()?;
        Ok(value)
    }

    /// Create a split pane with BIFF8 twip offsets and scroll origins.
    pub fn split(
        horizontal_twips: u16,
        vertical_twips: u16,
        bottom_pane_top_row: u16,
        right_pane_left_column: u8,
        active_pane: PaneType,
    ) -> XlsResult<Self> {
        let value = Self {
            mode: Mode::Split,
            horizontal_split: horizontal_twips,
            vertical_split: vertical_twips,
            bottom_pane_top_row,
            right_pane_left_column,
            active_pane,
        };
        value.validate()?;
        Ok(value)
    }

    /// Return whether this pane is frozen or split.
    pub const fn mode(self) -> Mode {
        self.mode
    }

    /// Return the horizontal split in columns when frozen, otherwise twips.
    pub const fn horizontal(self) -> u16 {
        self.horizontal_split
    }

    /// Return the vertical split in rows when frozen, otherwise twips.
    pub const fn vertical(self) -> u16 {
        self.vertical_split
    }

    /// Return the first row shown in the bottom pane.
    pub const fn row(self) -> u16 {
        self.bottom_pane_top_row
    }

    /// Return the first column shown in the right pane.
    pub const fn column(self) -> u8 {
        self.right_pane_left_column
    }

    /// Return the pane that owns the active selection.
    pub const fn active(self) -> PaneType {
        self.active_pane
    }

    pub(crate) fn validate(self) -> XlsResult<()> {
        if self.horizontal_split == 0 && self.vertical_split == 0 {
            return Err(XlsError::InvalidData(
                "PANE must split at least one axis".to_string(),
            ));
        }
        if self.mode == Mode::Split
            && (self.horizontal_split > i16::MAX as u16 || self.vertical_split > i16::MAX as u16)
        {
            return Err(XlsError::InvalidData(
                "split positions exceed 32767 twips".to_string(),
            ));
        }
        if self.mode == Mode::Frozen && self.horizontal_split > 255 {
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
/// A checked active cell and its selected ranges for one pane.
pub struct Selection {
    pane: PaneType,
    active_row: u16,
    active_column: u8,
    active_range_index: u16,
    ranges: Vec<Range>,
}

impl Selection {
    /// Create a selection record from ordered, nonempty ranges.
    pub fn new(
        pane: PaneType,
        active_row: u16,
        active_column: u8,
        active_range_index: u16,
        ranges: Vec<Range>,
    ) -> XlsResult<Self> {
        let value = Self {
            pane,
            active_row,
            active_column,
            active_range_index,
            ranges,
        };
        validate_selection(&value)?;
        Ok(value)
    }

    /// Select one cell in `pane`.
    pub fn cell(pane: PaneType, row: u16, column: u8) -> Self {
        Self {
            pane,
            active_row: row,
            active_column: column,
            active_range_index: 0,
            ranges: vec![Range::cell(row, column)],
        }
    }

    /// Return the pane containing this selection.
    pub const fn pane(&self) -> PaneType {
        self.pane
    }

    /// Return the active row.
    pub const fn row(&self) -> u16 {
        self.active_row
    }

    /// Return the active column.
    pub const fn column(&self) -> u8 {
        self.active_column
    }

    /// Return the zero-based active range index.
    pub const fn active(&self) -> u16 {
        self.active_range_index
    }

    /// Borrow the inclusive selected ranges.
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }
}

bitflags::bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    struct Flags: u16 {
        const FORMULAS = 0x0001;
        const GRIDLINES = 0x0002;
        const HEADERS = 0x0004;
        const ZEROS = 0x0010;
        const RIGHT_TO_LEFT = 0x0040;
        const OUTLINES = 0x0080;
        const SELECTED = 0x0200;
        const DISPLAYED = 0x0400;
        const PAGE_BREAKS = 0x0800;
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
/// Checked display, navigation, pane, and selection state for a worksheet.
pub struct View {
    flags: Flags,
    /// `None` uses the window default color; custom palette indices are 0..=63.
    gridline_color_index: Option<u16>,
    first_visible_row: u16,
    first_visible_column: u8,
    page_break_zoom_percent: Option<u16>,
    normal_zoom_percent: Option<u16>,
    scale: Option<Scale>,
    pane: Option<Pane>,
    selections: Vec<Selection>,
}

impl Default for View {
    fn default() -> Self {
        Self {
            flags: Flags::GRIDLINES
                | Flags::HEADERS
                | Flags::ZEROS
                | Flags::OUTLINES
                | Flags::SELECTED
                | Flags::DISPLAYED,
            gridline_color_index: None,
            first_visible_row: 0,
            first_visible_column: 0,
            page_break_zoom_percent: None,
            normal_zoom_percent: None,
            scale: None,
            pane: None,
            selections: vec![Selection::cell(PaneType::UpperLeft, 0, 0)],
        }
    }
}

impl View {
    /// Return whether formulas are shown instead of their results.
    pub const fn shows_formulas(&self) -> bool {
        self.flags.contains(Flags::FORMULAS)
    }

    /// Return whether gridlines are shown.
    pub const fn shows_gridlines(&self) -> bool {
        self.flags.contains(Flags::GRIDLINES)
    }

    /// Return whether row and column headers are shown.
    pub const fn shows_headers(&self) -> bool {
        self.flags.contains(Flags::HEADERS)
    }

    /// Return whether zero-valued cells are shown.
    pub const fn shows_zeros(&self) -> bool {
        self.flags.contains(Flags::ZEROS)
    }

    /// Return the custom palette index, or `None` for the default color.
    pub const fn grid_color_index(&self) -> Option<u16> {
        self.gridline_color_index
    }

    /// Return whether the worksheet is laid out right-to-left.
    pub const fn right_to_left(&self) -> bool {
        self.flags.contains(Flags::RIGHT_TO_LEFT)
    }

    /// Return whether outline symbols are shown.
    pub const fn shows_outlines(&self) -> bool {
        self.flags.contains(Flags::OUTLINES)
    }

    /// Return whether this worksheet is selected in the workbook window.
    pub const fn is_selected(&self) -> bool {
        self.flags.contains(Flags::SELECTED)
    }

    /// Return whether this worksheet window is displayed.
    pub const fn is_displayed(&self) -> bool {
        self.flags.contains(Flags::DISPLAYED)
    }

    /// Return whether page-break preview is active.
    pub const fn is_page_break_preview(&self) -> bool {
        self.flags.contains(Flags::PAGE_BREAKS)
    }

    /// Return the first visible row.
    pub const fn row(&self) -> u16 {
        self.first_visible_row
    }

    /// Return the first visible column.
    pub const fn column(&self) -> u8 {
        self.first_visible_column
    }

    /// Return the page-break preview zoom percentage.
    pub const fn page_zoom_percent(&self) -> Option<u16> {
        self.page_break_zoom_percent
    }

    /// Return the normal-view zoom percentage.
    pub const fn normal_zoom_percent(&self) -> Option<u16> {
        self.normal_zoom_percent
    }

    /// Return the optional exact zoom fraction.
    pub const fn scale(&self) -> Option<Scale> {
        self.scale
    }

    /// Borrow the current pane.
    pub const fn pane(&self) -> Option<&Pane> {
        self.pane.as_ref()
    }

    /// Borrow all pane selections.
    pub fn selections(&self) -> &[Selection] {
        &self.selections
    }

    /// Choose whether formulas are shown instead of their results.
    pub fn formulas(&mut self, show: bool) -> &mut Self {
        self.flags.set(Flags::FORMULAS, show);
        self
    }

    /// Choose whether gridlines are shown.
    pub fn gridlines(&mut self, show: bool) -> &mut Self {
        self.flags.set(Flags::GRIDLINES, show);
        self
    }

    /// Choose whether row and column headers are shown.
    pub fn headers(&mut self, show: bool) -> &mut Self {
        self.flags.set(Flags::HEADERS, show);
        self
    }

    /// Choose whether zero-valued cells are shown.
    pub fn zeros(&mut self, show: bool) -> &mut Self {
        self.flags.set(Flags::ZEROS, show);
        self
    }

    /// Choose right-to-left worksheet layout.
    pub fn rtl(&mut self, enabled: bool) -> &mut Self {
        self.flags.set(Flags::RIGHT_TO_LEFT, enabled);
        self
    }

    /// Choose whether outline symbols are shown.
    pub fn outlines(&mut self, show: bool) -> &mut Self {
        self.flags.set(Flags::OUTLINES, show);
        self
    }

    /// Choose whether this worksheet is selected in the workbook window.
    pub fn select(&mut self, selected: bool) -> &mut Self {
        self.flags.set(Flags::SELECTED, selected);
        self
    }

    /// Choose whether this worksheet window is displayed.
    pub fn display(&mut self, displayed: bool) -> &mut Self {
        self.flags.set(Flags::DISPLAYED, displayed);
        self
    }

    /// Choose whether page-break preview is active.
    pub fn page_breaks(&mut self, enabled: bool) -> &mut Self {
        self.flags.set(Flags::PAGE_BREAKS, enabled);
        self
    }

    /// Set the first visible cell after checking pane-specific sentinels.
    pub fn origin(&mut self, row: u16, column: u8) -> XlsResult<&mut Self> {
        validate_origin(row, column, self.pane.as_ref())?;
        self.first_visible_row = row;
        self.first_visible_column = column;
        Ok(self)
    }

    /// Set a palette index in `0..=63`, or use the default with `None`.
    pub fn grid_color(&mut self, index: Option<u16>) -> XlsResult<&mut Self> {
        if index.is_some_and(|value| value > 63) {
            return Err(XlsError::InvalidData(
                "custom gridline color index must be <= 63".to_string(),
            ));
        }
        self.gridline_color_index = index;
        Ok(self)
    }

    /// Set page-break preview zoom in `10..=400`, or clear it with `None`.
    pub fn page_zoom(&mut self, percent: Option<u16>) -> XlsResult<&mut Self> {
        validate_zoom(percent)?;
        self.page_break_zoom_percent = percent;
        Ok(self)
    }

    /// Set normal-view zoom in `10..=400`, or clear it with `None`.
    pub fn normal_zoom(&mut self, percent: Option<u16>) -> XlsResult<&mut Self> {
        validate_zoom(percent)?;
        self.normal_zoom_percent = percent;
        Ok(self)
    }

    /// Replace the exact zoom fraction and return the previous value.
    pub fn put_scale(&mut self, scale: Option<Scale>) -> Option<Scale> {
        std::mem::replace(&mut self.scale, scale)
    }

    /// Atomically replace the pane and its selections after preflight checks.
    pub fn put_pane(
        &mut self,
        pane: Pane,
        selections: Vec<Selection>,
    ) -> XlsResult<(Option<Pane>, Vec<Selection>)> {
        self.validate_with(Some(&pane), &selections)?;
        let previous_pane = self.pane.replace(pane);
        let previous_selections = std::mem::replace(&mut self.selections, selections);
        Ok((previous_pane, previous_selections))
    }

    /// Atomically replace selections after checking them against the pane.
    pub fn put_selections(&mut self, selections: Vec<Selection>) -> XlsResult<Vec<Selection>> {
        self.validate_with(self.pane.as_ref(), &selections)?;
        Ok(std::mem::replace(&mut self.selections, selections))
    }

    /// Remove the pane and return the previous owned pane and selections.
    pub fn clear_pane(&mut self) -> (Option<Pane>, Vec<Selection>) {
        let previous_pane = self.pane.take();
        let previous_selections = std::mem::replace(
            &mut self.selections,
            vec![Selection::cell(PaneType::UpperLeft, 0, 0)],
        );
        (previous_pane, previous_selections)
    }

    pub(crate) fn validate(&self) -> XlsResult<()> {
        self.validate_with(self.pane.as_ref(), &self.selections)
    }

    fn validate_with(&self, pane: Option<&Pane>, selections: &[Selection]) -> XlsResult<()> {
        if self.gridline_color_index.is_some_and(|value| value > 63) {
            return Err(XlsError::InvalidData(
                "custom gridline color index must be <= 63".to_string(),
            ));
        }
        validate_zoom(self.page_break_zoom_percent)?;
        validate_zoom(self.normal_zoom_percent)?;
        if let Some(scale) = self.scale {
            scale.validate()?;
        }
        if let Some(pane) = pane {
            pane.validate()?;
        }
        validate_origin(self.first_visible_row, self.first_visible_column, pane)?;
        validate_selections(pane, selections)
    }

    pub(crate) fn set_frozen(&mut self, rows: u16, columns: u8) -> XlsResult<()> {
        let pane = Pane::frozen(rows, columns)?;
        self.put_pane(pane, vec![Selection::cell(pane.active(), rows, columns)])?;
        Ok(())
    }
}

fn validate_origin(row: u16, column: u8, pane: Option<&Pane>) -> XlsResult<()> {
    if pane.is_some_and(|value| value.mode == Mode::Frozen)
        && (row == u16::MAX || column == u8::MAX)
    {
        return Err(XlsError::InvalidData(
            "frozen view cannot use sentinel origins".to_string(),
        ));
    }
    Ok(())
}

fn validate_zoom(percent: Option<u16>) -> XlsResult<()> {
    if percent.is_some_and(|value| !(10..=400).contains(&value)) {
        return Err(XlsError::InvalidData(
            "view zoom percent must be between 10 and 400".to_string(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_selection(selection: &Selection) -> XlsResult<()> {
    if selection.active_range_index & 0x8000 != 0 {
        return Err(XlsError::InvalidData(
            "active range index must be a signed nonnegative integer".to_string(),
        ));
    }
    if selection.ranges.is_empty() || selection.ranges.len() > MAX_SELECTION_RANGES {
        return Err(XlsError::InvalidData(
            "each SELECTION needs 1..=1369 ranges".to_string(),
        ));
    }
    if selection.ranges.iter().any(|range| {
        range.first_row() > range.last_row() || range.first_column() > range.last_column()
    }) {
        return Err(XlsError::InvalidData(
            "SELECTION contains an inverted range".to_string(),
        ));
    }
    Ok(())
}

fn validate_selections(pane: Option<&Pane>, selections: &[Selection]) -> XlsResult<()> {
    if selections.is_empty() {
        return if pane.is_none() {
            Ok(())
        } else {
            Err(XlsError::InvalidData(
                "a pane view needs at least one SELECTION".to_string(),
            ))
        };
    }
    let active_pane = pane.map_or(PaneType::UpperLeft, |value| value.active_pane);
    let mut has_active = false;
    let mut seen = 0u8;
    let mut start = 0usize;
    while start < selections.len() {
        let first = &selections[start];
        validate_selection(first)?;
        let pane_bit = 1u8 << first.pane.code();
        if seen & pane_bit != 0 {
            return Err(XlsError::InvalidData(
                "SELECTION pane groups must be contiguous".to_string(),
            ));
        }
        if !pane.map_or(first.pane == PaneType::UpperLeft, |value| {
            pane_exists(value.horizontal_split, value.vertical_split, first.pane)
        }) {
            return Err(XlsError::InvalidData(
                "SELECTION references a nonexistent pane".to_string(),
            ));
        }
        has_active |= first.pane == active_pane;
        let mut end = start;
        while end < selections.len() && selections[end].pane == first.pane {
            let current = &selections[end];
            validate_selection(current)?;
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
        seen |= pane_bit;
        start = end;
    }
    if !has_active {
        return Err(XlsError::InvalidData(
            "active pane has no SELECTION".to_string(),
        ));
    }
    Ok(())
}

fn active_pane(has_columns: bool, has_rows: bool) -> PaneType {
    match (has_columns, has_rows) {
        (true, true) => PaneType::LowerRight,
        (true, false) => PaneType::UpperRight,
        (false, true) => PaneType::LowerLeft,
        (false, false) => PaneType::UpperLeft,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::writer::biff;
    use std::panic::{AssertUnwindSafe, catch_unwind};

    #[test]
    fn serializes_exact_view_records() {
        let pane = Pane::split(1_200, 800, 7, 4, PaneType::LowerRight).unwrap();
        let mut view = View::default();
        view.formulas(true).gridlines(false);
        view.grid_color(Some(8)).unwrap();
        view.origin(2, 1).unwrap();
        view.put_pane(pane, vec![Selection::cell(PaneType::LowerRight, 7, 4)])
            .unwrap();
        let mut window = Vec::new();
        let mut pane_bytes = Vec::new();
        let mut selection = Vec::new();
        biff::write_window2_options(&mut window, &view).unwrap();
        biff::write_pane_options(&mut pane_bytes, &pane).unwrap();
        biff::write_selection_options(&mut selection, &view.selections()[0]).unwrap();
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
    fn checked_values_cover_biff8_boundaries() {
        assert!(Scale::new(1, 10).is_ok());
        assert!(Scale::new(4, 1).is_ok());
        assert!(Scale::new(32_767, 32_767).is_ok());
        assert!(Scale::new(32_768, 32_768).is_err());
        assert!(Scale::new(1, 11).is_err());
        assert!(Scale::new(5, 1).is_err());

        assert!(Pane::split(32_767, 0, 0, 0, PaneType::UpperRight).is_ok());
        assert!(Pane::split(32_768, 0, 0, 0, PaneType::UpperRight).is_err());
        assert!(Pane::split(0, 32_767, 0, 0, PaneType::LowerLeft).is_ok());
        assert!(Pane::split(0, 0, 0, 0, PaneType::UpperLeft).is_err());
        assert!(Pane::frozen(0, 0).is_err());

        assert!(Range::new(0, u16::MAX, 0, u8::MAX).is_ok());
        assert!(Range::new(1, 0, 0, 0).is_err());
        assert!(Range::new(0, 0, 1, 0).is_err());
        assert!(Selection::new(PaneType::UpperLeft, 0, 0, 0, Vec::new()).is_err());
        assert!(
            Selection::new(
                PaneType::UpperLeft,
                0,
                0,
                0,
                vec![Range::cell(0, 0); MAX_SELECTION_RANGES + 1],
            )
            .is_err()
        );
    }

    #[test]
    fn display_options_share_one_biff_word() {
        assert_eq!(std::mem::size_of::<Flags>(), std::mem::size_of::<u16>());
        let mut view = View::default();
        view.formulas(true)
            .gridlines(false)
            .headers(false)
            .zeros(false)
            .rtl(true)
            .outlines(false)
            .select(false)
            .display(false)
            .page_breaks(true);
        assert!(view.shows_formulas());
        assert!(!view.shows_gridlines());
        assert!(!view.shows_headers());
        assert!(!view.shows_zeros());
        assert!(view.right_to_left());
        assert!(!view.shows_outlines());
        assert!(!view.is_selected());
        assert!(!view.is_displayed());
        assert!(view.is_page_break_preview());
    }

    #[test]
    fn pane_preflight_does_not_unwind_or_mutate_on_error() {
        let mut view = View::default();
        view.formulas(true);
        let before = view.clone();
        let pane = Pane::split(1_200, 0, 0, 4, PaneType::UpperRight).unwrap();
        let invalid = vec![Selection::cell(PaneType::LowerLeft, 0, 0)];

        let outcome = catch_unwind(AssertUnwindSafe(|| view.put_pane(pane, invalid)));
        assert!(outcome.is_ok());
        assert!(outcome.unwrap().is_err());
        assert_eq!(view, before);
    }

    #[test]
    fn frozen_origin_preflight_is_failure_atomic() {
        let mut view = View::default();
        view.origin(u16::MAX, u8::MAX).unwrap();
        let before = view.clone();
        let pane = Pane::frozen(1, 1).unwrap();

        assert!(
            view.put_pane(pane, vec![Selection::cell(PaneType::LowerRight, 1, 1)],)
                .is_err()
        );
        assert_eq!(view, before);
    }

    #[test]
    fn pane_replacement_returns_owned_previous_state() {
        let mut view = View::default();
        let frozen = Pane::frozen(1, 1).unwrap();
        let (old_pane, old_selections) = view
            .put_pane(frozen, vec![Selection::cell(PaneType::LowerRight, 1, 1)])
            .unwrap();
        assert!(old_pane.is_none());
        assert_eq!(old_selections.len(), 1);

        let split = Pane::split(600, 0, 0, 2, PaneType::UpperRight).unwrap();
        let (old_pane, old_selections) = view
            .put_pane(split, vec![Selection::cell(PaneType::UpperRight, 0, 2)])
            .unwrap();
        assert_eq!(old_pane, Some(frozen));
        assert_eq!(old_selections[0].pane(), PaneType::LowerRight);
        assert_eq!(view.pane(), Some(&split));
    }
}
