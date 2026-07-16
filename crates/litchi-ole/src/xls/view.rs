//! BIFF8 worksheet window, zoom, pane, and selection records.

use crate::xls::error::{XlsError, XlsResult};

pub(crate) const WINDOW2_RECORD_TYPE: u16 = 0x023e;
pub(crate) const SCL_RECORD_TYPE: u16 = 0x00a0;
pub(crate) const PANE_RECORD_TYPE: u16 = 0x0041;
pub(crate) const SELECTION_RECORD_TYPE: u16 = 0x001d;
const PLV_RECORD_TYPE: u16 = 0x088b;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

/// Logical pane containing an active cell or selection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsPaneType {
    LowerRight,
    UpperRight,
    LowerLeft,
    UpperLeft,
}

impl XlsPaneType {
    fn parse(value: u8, record_type: u16) -> XlsResult<Self> {
        match value {
            0 => Ok(Self::LowerRight),
            1 => Ok(Self::UpperRight),
            2 => Ok(Self::LowerLeft),
            3 => Ok(Self::UpperLeft),
            _ => Err(invalid(record_type, "pane type must be between 0 and 3")),
        }
    }
}

/// Inclusive selected cell range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsSelectionRange {
    first_row: u16,
    last_row: u16,
    first_column: u8,
    last_column: u8,
}

impl XlsSelectionRange {
    pub fn first_row(&self) -> u16 { self.first_row }
    pub fn last_row(&self) -> u16 { self.last_row }
    pub fn first_column(&self) -> u8 { self.first_column }
    pub fn last_column(&self) -> u8 { self.last_column }
}

/// One BIFF8 `SELECTION` record in a worksheet window.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsSelection {
    pane: XlsPaneType,
    active_row: u16,
    active_column: u8,
    active_range_index: u16,
    ranges: Vec<XlsSelectionRange>,
}

impl XlsSelection {
    pub fn pane(&self) -> XlsPaneType { self.pane }
    pub fn active_row(&self) -> u16 { self.active_row }
    pub fn active_column(&self) -> u8 { self.active_column }
    pub fn active_range_index(&self) -> u16 { self.active_range_index }
    pub fn ranges(&self) -> &[XlsSelectionRange] { &self.ranges }
}

/// Frozen or split pane configuration for a worksheet window.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsPane {
    horizontal_split: u16,
    vertical_split: u16,
    bottom_pane_top_row: u16,
    right_pane_left_column: u8,
    active_pane: XlsPaneType,
}

impl XlsPane {
    /// Horizontal split in columns when frozen, otherwise in twips.
    pub fn horizontal_split(&self) -> u16 { self.horizontal_split }
    /// Vertical split in rows when frozen, otherwise in twips.
    pub fn vertical_split(&self) -> u16 { self.vertical_split }
    pub fn bottom_pane_top_row(&self) -> u16 { self.bottom_pane_top_row }
    pub fn right_pane_left_column(&self) -> u8 { self.right_pane_left_column }
    pub fn active_pane(&self) -> XlsPaneType { self.active_pane }
}

/// Display and navigation state for one worksheet window.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsWorksheetView {
    flags: u16,
    first_visible_row: u16,
    first_visible_column: u8,
    gridline_color_index: u16,
    page_break_zoom_percent: Option<u16>,
    normal_zoom_percent: Option<u16>,
    zoom_fraction: Option<(u16, u16)>,
    pane: Option<XlsPane>,
    selections: Vec<XlsSelection>,
}

impl XlsWorksheetView {
    pub fn shows_formulas(&self) -> bool { self.flags & 0x0001 != 0 }
    pub fn shows_gridlines(&self) -> bool { self.flags & 0x0002 != 0 }
    pub fn shows_row_column_headers(&self) -> bool { self.flags & 0x0004 != 0 }
    pub fn has_frozen_panes(&self) -> bool { self.flags & 0x0008 != 0 }
    pub fn shows_zero_values(&self) -> bool { self.flags & 0x0010 != 0 }
    pub fn uses_default_gridline_color(&self) -> bool { self.flags & 0x0020 != 0 }
    pub fn is_right_to_left(&self) -> bool { self.flags & 0x0040 != 0 }
    pub fn shows_outline_symbols(&self) -> bool { self.flags & 0x0080 != 0 }
    pub fn is_frozen_without_split(&self) -> bool { self.flags & 0x0100 != 0 }
    pub fn is_selected(&self) -> bool { self.flags & 0x0200 != 0 }
    pub fn is_displayed(&self) -> bool { self.flags & 0x0400 != 0 }
    pub fn is_page_break_preview(&self) -> bool { self.flags & 0x0800 != 0 }
    pub fn first_visible_row(&self) -> u16 { self.first_visible_row }
    pub fn first_visible_column(&self) -> u8 { self.first_visible_column }
    pub fn gridline_color_index(&self) -> u16 { self.gridline_color_index }
    pub fn page_break_zoom_percent(&self) -> Option<u16> { self.page_break_zoom_percent }
    pub fn normal_zoom_percent(&self) -> Option<u16> { self.normal_zoom_percent }
    pub fn zoom_fraction(&self) -> Option<(u16, u16)> { self.zoom_fraction }
    pub fn pane(&self) -> Option<&XlsPane> { self.pane.as_ref() }
    pub fn selections(&self) -> &[XlsSelection] { &self.selections }

    fn validate_selection_groups(&self) -> XlsResult<()> {
        let mut start = 0;
        while start < self.selections.len() {
            let first = &self.selections[start];
            let mut end = start + 1;
            let mut range_count = first.ranges.len();
            while end < self.selections.len() && self.selections[end].pane == first.pane {
                range_count = range_count
                    .checked_add(self.selections[end].ranges.len())
                    .ok_or_else(|| invalid(SELECTION_RECORD_TYPE, "SELECTION range aggregation overflows"))?;
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
            let range = active_range.expect("validated active selection index");
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

fn parse_zoom_percent(value: u16, record_type: u16, name: &str) -> XlsResult<Option<u16>> {
    if value == 0 {
        Ok(None)
    } else if (10..=400).contains(&value) {
        Ok(Some(value))
    } else {
        Err(invalid(record_type, format!("{name} must be zero or between 10 and 400")))
    }
}

fn parse_window2(data: &[u8]) -> XlsResult<XlsWorksheetView> {
    if data.len() != 18 {
        return Err(invalid(WINDOW2_RECORD_TYPE, format!("WINDOW2 payload must be 18 bytes, found {}", data.len())));
    }
    let flags = read_u16(data, 0);
    let first_visible_row = read_u16(data, 2);
    let first_visible_column = read_u16(data, 4);
    let gridline_color_index = read_u16(data, 6);
    if flags & 0xf000 != 0 {
        return Err(invalid(WINDOW2_RECORD_TYPE, "WINDOW2 reserved flag bits must be zero"));
    }
    if first_visible_column > 255 {
        return Err(invalid(WINDOW2_RECORD_TYPE, "WINDOW2 first visible column exceeds 255"));
    }
    if flags & 0x0100 != 0 && flags & 0x0008 == 0 {
        return Err(invalid(WINDOW2_RECORD_TYPE, "WINDOW2 frozen-without-split requires frozen panes"));
    }
    if flags & 0x0008 != 0 && (first_visible_row == u16::MAX || first_visible_column == 255) {
        return Err(invalid(WINDOW2_RECORD_TYPE, "WINDOW2 sentinel visible origins cannot be frozen"));
    }
    let uses_default_color = flags & 0x0020 != 0;
    if gridline_color_index > 64 || (gridline_color_index == 64) != uses_default_color {
        return Err(invalid(WINDOW2_RECORD_TYPE, "WINDOW2 gridline color and default-color flag disagree"));
    }
    if read_u16(data, 8) != 0 || read_u16(data, 16) != 0 {
        return Err(invalid(WINDOW2_RECORD_TYPE, "WINDOW2 reserved fields must be zero"));
    }

    Ok(XlsWorksheetView {
        flags,
        first_visible_row,
        first_visible_column: first_visible_column as u8,
        gridline_color_index,
        page_break_zoom_percent: parse_zoom_percent(read_u16(data, 10), WINDOW2_RECORD_TYPE, "page-break zoom")?,
        normal_zoom_percent: parse_zoom_percent(read_u16(data, 12), WINDOW2_RECORD_TYPE, "normal zoom")?,
        zoom_fraction: None,
        pane: None,
        selections: Vec::new(),
    })
}

fn parse_scl(data: &[u8]) -> XlsResult<(u16, u16)> {
    if data.len() != 4 {
        return Err(invalid(SCL_RECORD_TYPE, format!("SCL payload must be 4 bytes, found {}", data.len())));
    }
    let numerator = read_u16(data, 0);
    let denominator = read_u16(data, 2);
    if numerator == 0 || denominator == 0 {
        return Err(invalid(SCL_RECORD_TYPE, "SCL numerator and denominator must be positive"));
    }
    let numerator = u32::from(numerator);
    let denominator = u32::from(denominator);
    if numerator * 10 < denominator || numerator > denominator * 4 {
        return Err(invalid(SCL_RECORD_TYPE, "SCL zoom fraction must be between 1/10 and 4"));
    }
    Ok((numerator as u16, denominator as u16))
}

fn parse_pane(data: &[u8], frozen: bool) -> XlsResult<XlsPane> {
    if data.len() != 10 {
        return Err(invalid(PANE_RECORD_TYPE, format!("PANE payload must be 10 bytes, found {}", data.len())));
    }
    let horizontal_split = read_u16(data, 0);
    let vertical_split = read_u16(data, 2);
    let right_pane_left_column = read_u16(data, 6);
    if (frozen && horizontal_split > 255)
        || (!frozen && (horizontal_split > 32767 || vertical_split > 32767))
    {
        return Err(invalid(PANE_RECORD_TYPE, "PANE split position is outside its mode-specific bounds"));
    }
    if right_pane_left_column > 255 {
        return Err(invalid(PANE_RECORD_TYPE, "PANE right-pane column exceeds 255"));
    }
    if data[9] != 0 {
        return Err(invalid(PANE_RECORD_TYPE, "PANE reserved byte must be zero"));
    }
    Ok(XlsPane {
        horizontal_split,
        vertical_split,
        bottom_pane_top_row: read_u16(data, 4),
        right_pane_left_column: right_pane_left_column as u8,
        active_pane: XlsPaneType::parse(data[8], PANE_RECORD_TYPE)?,
    })
}

fn parse_selection(data: &[u8]) -> XlsResult<XlsSelection> {
    if data.len() < 9 {
        return Err(invalid(SELECTION_RECORD_TYPE, "SELECTION payload is shorter than 9 bytes"));
    }
    let active_column = read_u16(data, 3);
    let active_range_index = read_u16(data, 5);
    let range_count = usize::from(read_u16(data, 7));
    if active_column > 255 {
        return Err(invalid(SELECTION_RECORD_TYPE, "SELECTION active column exceeds 255"));
    }
    if active_range_index & 0x8000 != 0 {
        return Err(invalid(SELECTION_RECORD_TYPE, "SELECTION active range index must be nonnegative"));
    }
    if range_count > 1369 || data.len() != 9 + range_count * 6 {
        return Err(invalid(SELECTION_RECORD_TYPE, "SELECTION range count does not match its payload"));
    }
    let mut ranges = Vec::with_capacity(range_count);
    for chunk in data[9..].chunks_exact(6) {
        let range = XlsSelectionRange {
            first_row: read_u16(chunk, 0),
            last_row: read_u16(chunk, 2),
            first_column: chunk[4],
            last_column: chunk[5],
        };
        if range.first_row > range.last_row || range.first_column > range.last_column {
            return Err(invalid(SELECTION_RECORD_TYPE, "SELECTION contains an inverted range"));
        }
        ranges.push(range);
    }
    Ok(XlsSelection {
        pane: XlsPaneType::parse(data[0], SELECTION_RECORD_TYPE)?,
        active_row: read_u16(data, 1),
        active_column: active_column as u8,
        active_range_index,
        ranges,
    })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ViewPhase {
    Start,
    Zoom,
    Pane,
    Selections,
}

/// Collects only primary `WINDOW` productions, excluding custom-view selections.
pub(crate) struct ViewCollector {
    views: Vec<XlsWorksheetView>,
    current: Option<XlsWorksheetView>,
    phase: ViewPhase,
    saw_plv: bool,
}

impl ViewCollector {
    pub(crate) fn new() -> Self {
        Self {
            views: Vec::new(),
            current: None,
            phase: ViewPhase::Start,
            saw_plv: false,
        }
    }

    fn finish_current(&mut self) -> XlsResult<()> {
        if let Some(view) = self.current.take() {
            view.validate_selection_groups()?;
            self.views.push(view);
        }
        self.phase = ViewPhase::Start;
        self.saw_plv = false;
        Ok(())
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if record_type == WINDOW2_RECORD_TYPE {
            self.finish_current()?;
            self.current = Some(parse_window2(data)?);
            return Ok(());
        }
        let Some(view) = self.current.as_mut() else {
            return Ok(());
        };
        match record_type {
            PLV_RECORD_TYPE if self.phase == ViewPhase::Start && !self.saw_plv => {
                self.saw_plv = true;
            }
            SCL_RECORD_TYPE if self.phase == ViewPhase::Start => {
                view.zoom_fraction = Some(parse_scl(data)?);
                self.phase = ViewPhase::Zoom;
            }
            PANE_RECORD_TYPE if matches!(self.phase, ViewPhase::Start | ViewPhase::Zoom) => {
                view.pane = Some(parse_pane(data, view.has_frozen_panes())?);
                self.phase = ViewPhase::Pane;
            }
            SELECTION_RECORD_TYPE => {
                let selection = parse_selection(data)?;
                if let Some(previous) = view.selections.last()
                    && previous.pane == selection.pane
                    && (previous.active_row != selection.active_row
                        || previous.active_column != selection.active_column
                        || previous.active_range_index != selection.active_range_index)
                {
                    return Err(invalid(SELECTION_RECORD_TYPE, "contiguous selections for one pane disagree on the active cell"));
                }
                view.selections.push(selection);
                self.phase = ViewPhase::Selections;
            }
            PLV_RECORD_TYPE | SCL_RECORD_TYPE | PANE_RECORD_TYPE => {
                return Err(invalid(record_type, "record is out of order in the WINDOW production"));
            }
            _ => self.finish_current()?,
        }
        Ok(())
    }

    pub(crate) fn finish(mut self) -> XlsResult<Vec<XlsWorksheetView>> {
        self.finish_current()?;
        Ok(self.views)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn window2(flags: u16) -> [u8; 18] {
        let mut data = [0u8; 18];
        data[0..2].copy_from_slice(&flags.to_le_bytes());
        data[6..8].copy_from_slice(&64u16.to_le_bytes());
        data[10..12].copy_from_slice(&60u16.to_le_bytes());
        data[12..14].copy_from_slice(&75u16.to_le_bytes());
        data
    }

    fn pane() -> [u8; 10] {
        [5, 0, 7, 0, 7, 0, 34, 0, 0, 0]
    }

    fn selection(pane: u8) -> [u8; 15] {
        [pane, 7, 0, 34, 0, 0, 0, 1, 0, 7, 0, 7, 0, 34, 34]
    }

    #[test]
    fn parses_window_zoom_pane_and_selection() {
        let mut collector = ViewCollector::new();
        collector.feed_record(WINDOW2_RECORD_TYPE, &window2(0x07be)).unwrap();
        collector.feed_record(SCL_RECORD_TYPE, &[3, 0, 4, 0]).unwrap();
        collector.feed_record(PANE_RECORD_TYPE, &pane()).unwrap();
        collector.feed_record(SELECTION_RECORD_TYPE, &selection(0)).unwrap();
        let views = collector.finish().unwrap();
        let view = &views[0];
        assert!(view.has_frozen_panes());
        assert!(view.is_frozen_without_split());
        assert_eq!(view.zoom_fraction(), Some((3, 4)));
        assert_eq!(view.normal_zoom_percent(), Some(75));
        assert_eq!(view.pane().unwrap().right_pane_left_column(), 34);
        assert_eq!(view.selections()[0].ranges()[0].first_row(), 7);
    }

    #[test]
    fn rejects_malformed_and_out_of_order_view_records() {
        assert!(parse_window2(&[0; 17]).is_err());
        assert!(parse_scl(&[0, 0, 1, 0]).is_err());
        assert!(parse_pane(&[0; 9], false).is_err());
        assert!(parse_selection(&[0; 8]).is_err());

        let mut collector = ViewCollector::new();
        collector.feed_record(WINDOW2_RECORD_TYPE, &window2(0x002e)).unwrap();
        collector.feed_record(PANE_RECORD_TYPE, &pane()).unwrap();
        assert!(collector.feed_record(SCL_RECORD_TYPE, &[1, 0, 1, 0]).is_err());
    }

    #[test]
    fn ignores_custom_view_selections_after_window_closes() {
        let mut collector = ViewCollector::new();
        collector.feed_record(WINDOW2_RECORD_TYPE, &window2(0x0026)).unwrap();
        collector.feed_record(SELECTION_RECORD_TYPE, &selection(3)).unwrap();
        collector.feed_record(0x01aa, &[]).unwrap();
        collector.feed_record(SELECTION_RECORD_TYPE, &selection(0)).unwrap();
        let views = collector.finish().unwrap();
        assert_eq!(views[0].selections().len(), 1);
    }

    #[test]
    fn validates_active_range_across_contiguous_selection_records() {
        let mut first = selection(0);
        first[5..7].copy_from_slice(&1u16.to_le_bytes());
        let mut collector = ViewCollector::new();
        collector.feed_record(WINDOW2_RECORD_TYPE, &window2(0x002e)).unwrap();
        collector.feed_record(SELECTION_RECORD_TYPE, &first).unwrap();
        let mut second = selection(0);
        second[5..7].copy_from_slice(&1u16.to_le_bytes());
        collector.feed_record(SELECTION_RECORD_TYPE, &second).unwrap();
        assert!(collector.finish().is_ok());

        let mut invalid = selection(0);
        invalid[1..3].copy_from_slice(&8u16.to_le_bytes());
        let mut collector = ViewCollector::new();
        collector.feed_record(WINDOW2_RECORD_TYPE, &window2(0x002e)).unwrap();
        collector.feed_record(SELECTION_RECORD_TYPE, &invalid).unwrap();
        assert!(collector.finish().is_err());
    }

    #[test]
    fn reads_poi_pane_and_zoom_fixtures() {
        use crate::xls::XlsWorkbook;
        use std::fs::File;
        use std::path::Path;

        let fixture = |name: &str| {
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../3rdparty/poi/test-data/spreadsheet")
                .join(name)
        };

        let zoomed = XlsWorkbook::new(File::open(fixture("41139.xls")).unwrap()).unwrap();
        let view = zoomed.xls_worksheet(0).unwrap().worksheet_view().unwrap();
        assert_eq!(view.zoom_fraction(), Some((3, 4)));
        assert_eq!(view.normal_zoom_percent(), Some(75));
        assert_eq!(view.selections().len(), 4);
        let pane = view.pane().unwrap();
        assert_eq!((pane.horizontal_split(), pane.vertical_split()), (5, 7));
        assert_eq!((pane.bottom_pane_top_row(), pane.right_pane_left_column()), (7, 34));
        assert_eq!(pane.active_pane(), XlsPaneType::LowerRight);

        let split = XlsWorkbook::new(File::open(fixture("50939.xls")).unwrap()).unwrap();
        let view = split.xls_worksheet(0).unwrap().worksheet_view().unwrap();
        assert!(view.has_frozen_panes());
        assert!(!view.is_frozen_without_split());
        let pane = view.pane().unwrap();
        assert_eq!((pane.horizontal_split(), pane.vertical_split()), (8, 4));
        assert_eq!((pane.bottom_pane_top_row(), pane.right_pane_left_column()), (4, 26));
    }
}
