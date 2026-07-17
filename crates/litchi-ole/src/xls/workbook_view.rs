//! BIFF8 workbook-global window and sheet-tab identity metadata.

use std::collections::HashSet;

use super::{XlsError, XlsResult};

pub(crate) const WINDOW1_RECORD_TYPE: u16 = 0x003d;
pub(crate) const RR_TAB_ID_RECORD_TYPE: u16 = 0x013d;
const BOUND_SHEET8_RECORD_TYPE: u16 = 0x0085;
const MAX_RR_TAB_IDS: usize = 4112;

/// Display and sheet-tab navigation state for one workbook window.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsWorkbookWindow {
    horizontal_position_twips: i16,
    vertical_position_twips: i16,
    width_twips: i16,
    height_twips: i16,
    hidden: bool,
    minimized: bool,
    very_hidden: bool,
    shows_horizontal_scrollbar: bool,
    shows_vertical_scrollbar: bool,
    shows_sheet_tabs: bool,
    groups_dates_in_autofilter: bool,
    active_sheet_index: u16,
    first_visible_sheet_index: u16,
    selected_sheet_count: u16,
    sheet_tab_ratio_per_mille: u16,
}

impl XlsWorkbookWindow {
    pub fn horizontal_position_twips(&self) -> i16 { self.horizontal_position_twips }
    pub fn vertical_position_twips(&self) -> i16 { self.vertical_position_twips }
    pub fn width_twips(&self) -> i16 { self.width_twips }
    pub fn height_twips(&self) -> i16 { self.height_twips }
    pub fn hidden(&self) -> bool { self.hidden }
    pub fn minimized(&self) -> bool { self.minimized }
    pub fn very_hidden(&self) -> bool { self.very_hidden }
    pub fn shows_horizontal_scrollbar(&self) -> bool { self.shows_horizontal_scrollbar }
    pub fn shows_vertical_scrollbar(&self) -> bool { self.shows_vertical_scrollbar }
    pub fn shows_sheet_tabs(&self) -> bool { self.shows_sheet_tabs }
    pub fn groups_dates_in_autofilter(&self) -> bool { self.groups_dates_in_autofilter }
    pub fn active_sheet_index(&self) -> u16 { self.active_sheet_index }
    pub fn first_visible_sheet_index(&self) -> u16 { self.first_visible_sheet_index }
    pub fn selected_sheet_count(&self) -> u16 { self.selected_sheet_count }
    pub fn sheet_tab_ratio_per_mille(&self) -> u16 { self.sheet_tab_ratio_per_mille }

    fn validate_sheet_references(&self, sheet_count: usize) -> XlsResult<()> {
        if usize::from(self.active_sheet_index) >= sheet_count {
            return invalid(WINDOW1_RECORD_TYPE, "active sheet index is outside BoundSheet8 collection");
        }
        if usize::from(self.first_visible_sheet_index) >= sheet_count {
            return invalid(WINDOW1_RECORD_TYPE, "first visible sheet index is outside BoundSheet8 collection");
        }
        if usize::from(self.selected_sheet_count) > sheet_count {
            return invalid(WINDOW1_RECORD_TYPE, "selected sheet count exceeds BoundSheet8 count");
        }
        Ok(())
    }
}

/// Workbook window collection and stable sheet identifiers.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsWorkbookView {
    sheet_ids: Vec<u16>,
    windows: Vec<XlsWorkbookWindow>,
}

impl XlsWorkbookView {
    /// Stable sheet identifiers in `BoundSheet8` order.
    pub fn sheet_ids(&self) -> &[u16] { &self.sheet_ids }
    /// Workbook windows in `Window1` record order.
    pub fn windows(&self) -> &[XlsWorkbookWindow] { &self.windows }
    pub fn primary_window(&self) -> Option<&XlsWorkbookWindow> { self.windows.first() }

    pub(crate) fn validate_sheet_state(
        &self,
        visible_tabs: &[bool],
        selected_worksheet_tabs: &[Option<bool>],
    ) -> XlsResult<()> {
        if visible_tabs.len() != selected_worksheet_tabs.len() {
            return invalid(
                WINDOW1_RECORD_TYPE,
                "Window1 cross-record state has inconsistent sheet cardinality",
            );
        }

        for window in &self.windows {
            let active = usize::from(window.active_sheet_index);
            let first_visible = usize::from(window.first_visible_sheet_index);
            if !visible_tabs[active] {
                return invalid(
                    WINDOW1_RECORD_TYPE,
                    format!("active sheet index {active} refers to a hidden BoundSheet8 tab"),
                );
            }
            if !visible_tabs[first_visible] {
                return invalid(
                    WINDOW1_RECORD_TYPE,
                    format!(
                        "first visible sheet index {first_visible} refers to a hidden BoundSheet8 tab"
                    ),
                );
            }
        }

        let Some(window) = self.primary_window() else { return Ok(()); };
        let active = usize::from(window.active_sheet_index);
        if selected_worksheet_tabs[active] == Some(false) {
            return invalid(
                WINDOW1_RECORD_TYPE,
                format!("active sheet index {active} is not selected in Window2"),
            );
        }

        let selected_count = selected_worksheet_tabs.iter().flatten().filter(|selected| **selected).count();
        let declared_count = usize::from(window.selected_sheet_count);
        if selected_count > declared_count
            || (selected_worksheet_tabs.iter().all(Option::is_some) && selected_count != declared_count)
        {
            return invalid(
                WINDOW1_RECORD_TYPE,
                format!(
                    "selected sheet count {declared_count} disagrees with Window2 selected state ({selected_count})"
                ),
            );
        }
        Ok(())
    }
}

pub(crate) struct WorkbookViewCollector {
    sheet_ids: Option<Vec<u16>>,
    windows: Vec<XlsWorkbookWindow>,
    boundsheets_started: bool,
}

impl WorkbookViewCollector {
    pub(crate) fn new() -> Self {
        Self { sheet_ids: None, windows: Vec::new(), boundsheets_started: false }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        match record_type {
            RR_TAB_ID_RECORD_TYPE => {
                if self.sheet_ids.is_some() {
                    return invalid(record_type, "duplicate RRTabId record");
                }
                if !self.windows.is_empty() || self.boundsheets_started {
                    return invalid(record_type, "RRTabId record is out of BIFF8 order");
                }
                self.sheet_ids = Some(parse_rr_tab_id(data)?);
            },
            WINDOW1_RECORD_TYPE => {
                if self.boundsheets_started {
                    return invalid(record_type, "Window1 record must precede BoundSheet8 records");
                }
                self.windows.push(parse_window1(data)?);
            },
            BOUND_SHEET8_RECORD_TYPE => self.boundsheets_started = true,
            _ => {},
        }
        Ok(())
    }

    pub(crate) fn finish(self, sheet_count: usize) -> XlsResult<XlsWorkbookView> {
        if sheet_count == 0 {
            return invalid(BOUND_SHEET8_RECORD_TYPE, "workbook contains no sheets");
        }
        let sheet_ids = if sheet_count <= MAX_RR_TAB_IDS {
            self.sheet_ids.ok_or_else(|| XlsError::InvalidRecord {
                record_type: RR_TAB_ID_RECORD_TYPE,
                message: "RRTabId is required for workbooks with at most 4112 sheets".to_string(),
            })?
        } else {
            if self.sheet_ids.is_some() {
                return invalid(RR_TAB_ID_RECORD_TYPE, "RRTabId must be absent above 4112 sheets");
            }
            (1..=sheet_count)
                .map(|index| u16::try_from(index).unwrap_or(u16::MAX))
                .collect()
        };
        if sheet_count <= MAX_RR_TAB_IDS && sheet_ids.len() != sheet_count {
            return invalid(RR_TAB_ID_RECORD_TYPE, "RRTabId count does not match BoundSheet8 count");
        }
        if self.windows.is_empty() {
            return invalid(WINDOW1_RECORD_TYPE, "Globals Substream requires at least one Window1");
        }
        for window in &self.windows {
            window.validate_sheet_references(sheet_count)?;
        }
        Ok(XlsWorkbookView { sheet_ids, windows: self.windows })
    }
}

fn parse_rr_tab_id(data: &[u8]) -> XlsResult<Vec<u16>> {
    if data.is_empty() || data.len() % 2 != 0 || data.len() > MAX_RR_TAB_IDS * 2 {
        return invalid(RR_TAB_ID_RECORD_TYPE, "RRTabId payload must contain 1..=4112 identifiers");
    }
    let mut ids = Vec::with_capacity(data.len() / 2);
    let mut unique = HashSet::with_capacity(data.len() / 2);
    for chunk in data.chunks_exact(2) {
        let id = u16::from_le_bytes([chunk[0], chunk[1]]);
        // RRTabId stores producer-assigned unsigned identifiers. Although the
        // TabId reference structure uses 1..=0xFFFE, BIFF8 producers including
        // Apache POI emit zero-based RRTabId arrays. Preserve zero as an actual
        // unique identifier while excluding 0xFFFF, the no-sheet sentinel used
        // by sheet-reference structures.
        if id == 0xffff {
            return invalid(RR_TAB_ID_RECORD_TYPE, "sheet identifier must not be the 0xFFFF no-sheet sentinel");
        }
        if !unique.insert(id) {
            return invalid(RR_TAB_ID_RECORD_TYPE, "sheet identifiers must be unique");
        }
        ids.push(id);
    }
    Ok(ids)
}

fn parse_window1(data: &[u8]) -> XlsResult<XlsWorkbookWindow> {
    if data.len() != 18 {
        return Err(XlsError::InvalidLength { expected: 18, found: data.len() });
    }
    let width_twips = read_i16(data, 4);
    let height_twips = read_i16(data, 6);
    if width_twips < 1 || height_twips < 1 {
        return invalid(WINDOW1_RECORD_TYPE, "window width and height must be positive");
    }
    let flags = read_u16(data, 8);
    if flags & 0xff80 != 0 {
        return invalid(WINDOW1_RECORD_TYPE, "Window1 reserved flag bits must be zero");
    }
    if flags & 0x0004 != 0 && flags & 0x0001 == 0 {
        return invalid(WINDOW1_RECORD_TYPE, "Window1 fVeryHidden requires fHidden");
    }
    let ratio = read_u16(data, 16);
    if ratio > 1000 {
        return invalid(WINDOW1_RECORD_TYPE, "sheet tab ratio must be at most 1000");
    }
    Ok(XlsWorkbookWindow {
        horizontal_position_twips: read_i16(data, 0),
        vertical_position_twips: read_i16(data, 2),
        width_twips,
        height_twips,
        hidden: flags & 0x0001 != 0,
        minimized: flags & 0x0002 != 0,
        very_hidden: flags & 0x0004 != 0,
        shows_horizontal_scrollbar: flags & 0x0008 != 0,
        shows_vertical_scrollbar: flags & 0x0010 != 0,
        shows_sheet_tabs: flags & 0x0020 != 0,
        groups_dates_in_autofilter: flags & 0x0040 == 0,
        active_sheet_index: read_u16(data, 10),
        first_visible_sheet_index: read_u16(data, 12),
        selected_sheet_count: read_u16(data, 14),
        sheet_tab_ratio_per_mille: ratio,
    })
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_i16(data: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes([data[offset], data[offset + 1]])
}

fn invalid<T>(record_type: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidRecord { record_type, message: message.into() })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_window() -> [u8; 18] {
        [0, 0, 0, 0, 1, 0, 1, 0, 0x38, 0, 0, 0, 0, 0, 1, 0, 0x58, 2]
    }

    #[test]
    fn rejects_malformed_rr_tab_ids() {
        assert!(parse_rr_tab_id(&[1]).is_err());
        assert_eq!(parse_rr_tab_id(&[0, 0]).unwrap(), vec![0]);
        assert!(parse_rr_tab_id(&[0xff, 0xff]).is_err());
        assert!(parse_rr_tab_id(&[1, 0, 1, 0]).is_err());
    }

    #[test]
    fn rejects_malformed_window_and_order() {
        let mut reserved = valid_window();
        reserved[9] = 0x80;
        assert!(parse_window1(&reserved).is_err());
        let mut zero_width = valid_window();
        zero_width[4] = 0;
        zero_width[5] = 0;
        assert!(parse_window1(&zero_width).is_err());
        let mut very_hidden_without_hidden = valid_window();
        very_hidden_without_hidden[8] |= 0x04;
        assert!(parse_window1(&very_hidden_without_hidden).is_err());

        let mut collector = WorkbookViewCollector::new();
        collector.feed_record(WINDOW1_RECORD_TYPE, &valid_window()).unwrap();
        assert!(collector.feed_record(RR_TAB_ID_RECORD_TYPE, &[1, 0]).is_err());
    }

    #[test]
    fn validates_boundsheet_visibility_and_window2_selection() {
        let mut first_hidden = valid_window();
        first_hidden[12] = 1;
        let view = XlsWorkbookView {
            sheet_ids: vec![1, 2],
            windows: vec![parse_window1(&first_hidden).unwrap()],
        };
        assert!(view.validate_sheet_state(&[true, false], &[Some(true), Some(false)]).is_err());

        let view = XlsWorkbookView {
            sheet_ids: vec![1, 2],
            windows: vec![parse_window1(&valid_window()).unwrap()],
        };
        assert!(view.validate_sheet_state(&[true, true], &[Some(false), Some(true)]).is_err());
        assert!(view.validate_sheet_state(&[true, true], &[Some(true), Some(true)]).is_err());
        assert!(view.validate_sheet_state(&[true, true], &[Some(true), None]).is_ok());
    }

    #[test]
    fn rejects_cardinality_and_boundsheet_mismatch() {
        let mut collector = WorkbookViewCollector::new();
        collector.feed_record(RR_TAB_ID_RECORD_TYPE, &[1, 0]).unwrap();
        collector.feed_record(WINDOW1_RECORD_TYPE, &valid_window()).unwrap();
        assert!(collector.finish(2).is_err());
    }
}
