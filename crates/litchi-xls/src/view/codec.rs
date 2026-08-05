//! Bounded BIFF8 worksheet-window record codec.

use super::model::{Pane, PaneType, Range, Selection, View, invalid};
use super::{PANE_RECORD_TYPE, SCL_RECORD_TYPE, SELECTION_RECORD_TYPE, WINDOW2_RECORD_TYPE};
use crate::error::Result;

const PLV_RECORD_TYPE: u16 = 0x088b;

pub(super) fn read_u8(data: &[u8], offset: usize, record_type: u16, field: &str) -> Result<u8> {
    data.get(offset)
        .copied()
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

pub(super) fn read_u16(data: &[u8], offset: usize, record_type: u16, field: &str) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid(record_type, format!("{field} offset overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))?;
    let bytes: [u8; 2] = bytes
        .try_into()
        .map_err(|_| invalid(record_type, format!("truncated {field}")))?;
    Ok(u16::from_le_bytes(bytes))
}

pub(super) fn parse_zoom_percent(value: u16, record_type: u16, name: &str) -> Result<Option<u16>> {
    if value == 0 {
        Ok(None)
    } else if (10..=400).contains(&value) {
        Ok(Some(value))
    } else {
        Err(invalid(
            record_type,
            format!("{name} must be zero or between 10 and 400"),
        ))
    }
}

pub(super) fn parse_window2(data: &[u8]) -> Result<View> {
    if data.len() != 18 {
        return Err(invalid(
            WINDOW2_RECORD_TYPE,
            format!("WINDOW2 payload must be 18 bytes, found {}", data.len()),
        ));
    }
    let flags = read_u16(data, 0, WINDOW2_RECORD_TYPE, "WINDOW2.flags")?;
    let first_visible_row = read_u16(data, 2, WINDOW2_RECORD_TYPE, "WINDOW2.first_visible_row")?;
    let first_visible_column =
        read_u16(data, 4, WINDOW2_RECORD_TYPE, "WINDOW2.first_visible_column")?;
    let gridline_color_index =
        read_u16(data, 6, WINDOW2_RECORD_TYPE, "WINDOW2.gridline_color_index")?;
    if flags & 0xf000 != 0 {
        return Err(invalid(
            WINDOW2_RECORD_TYPE,
            "WINDOW2 reserved flag bits must be zero",
        ));
    }
    if first_visible_column > 255 {
        return Err(invalid(
            WINDOW2_RECORD_TYPE,
            "WINDOW2 first visible column exceeds 255",
        ));
    }
    if flags & 0x0100 != 0 && flags & 0x0008 == 0 {
        return Err(invalid(
            WINDOW2_RECORD_TYPE,
            "WINDOW2 frozen-without-split requires frozen panes",
        ));
    }
    if flags & 0x0008 != 0 && (first_visible_row == u16::MAX || first_visible_column == 255) {
        return Err(invalid(
            WINDOW2_RECORD_TYPE,
            "WINDOW2 sentinel visible origins cannot be frozen",
        ));
    }
    let uses_default_color = flags & 0x0020 != 0;
    if gridline_color_index > 64 || (gridline_color_index == 64) != uses_default_color {
        return Err(invalid(
            WINDOW2_RECORD_TYPE,
            "WINDOW2 gridline color and default-color flag disagree",
        ));
    }
    if read_u16(data, 8, WINDOW2_RECORD_TYPE, "WINDOW2.reserved1")? != 0
        || read_u16(data, 16, WINDOW2_RECORD_TYPE, "WINDOW2.reserved2")? != 0
    {
        return Err(invalid(
            WINDOW2_RECORD_TYPE,
            "WINDOW2 reserved fields must be zero",
        ));
    }

    Ok(View {
        flags,
        first_visible_row,
        first_visible_column: first_visible_column as u8,
        gridline_color_index,
        page_break_zoom_percent: parse_zoom_percent(
            read_u16(data, 10, WINDOW2_RECORD_TYPE, "WINDOW2.page_break_zoom")?,
            WINDOW2_RECORD_TYPE,
            "page-break zoom",
        )?,
        normal_zoom_percent: parse_zoom_percent(
            read_u16(data, 12, WINDOW2_RECORD_TYPE, "WINDOW2.normal_zoom")?,
            WINDOW2_RECORD_TYPE,
            "normal zoom",
        )?,
        zoom_fraction: None,
        pane: None,
        selections: Vec::new(),
    })
}

pub(super) fn parse_scl(data: &[u8]) -> Result<(u16, u16)> {
    if data.len() != 4 {
        return Err(invalid(
            SCL_RECORD_TYPE,
            format!("SCL payload must be 4 bytes, found {}", data.len()),
        ));
    }
    let numerator = read_u16(data, 0, SCL_RECORD_TYPE, "SCL.numerator")?;
    let denominator = read_u16(data, 2, SCL_RECORD_TYPE, "SCL.denominator")?;
    if numerator == 0
        || denominator == 0
        || numerator > i16::MAX as u16
        || denominator > i16::MAX as u16
    {
        return Err(invalid(
            SCL_RECORD_TYPE,
            "SCL numerator and denominator must be signed positive integers",
        ));
    }
    let numerator = u32::from(numerator);
    let denominator = u32::from(denominator);
    if numerator * 10 < denominator || numerator > denominator * 4 {
        return Err(invalid(
            SCL_RECORD_TYPE,
            "SCL zoom fraction must be between 1/10 and 4",
        ));
    }
    Ok((numerator as u16, denominator as u16))
}

pub(super) fn parse_pane(data: &[u8], frozen: bool) -> Result<Pane> {
    if data.len() != 10 {
        return Err(invalid(
            PANE_RECORD_TYPE,
            format!("PANE payload must be 10 bytes, found {}", data.len()),
        ));
    }
    let horizontal_split = read_u16(data, 0, PANE_RECORD_TYPE, "PANE.horizontal_split")?;
    let vertical_split = read_u16(data, 2, PANE_RECORD_TYPE, "PANE.vertical_split")?;
    let right_pane_left_column =
        read_u16(data, 6, PANE_RECORD_TYPE, "PANE.right_pane_left_column")?;
    if (frozen && horizontal_split > 255)
        || (!frozen && (horizontal_split > 32767 || vertical_split > 32767))
    {
        return Err(invalid(
            PANE_RECORD_TYPE,
            "PANE split position is outside its mode-specific bounds",
        ));
    }
    if right_pane_left_column > 255 {
        return Err(invalid(
            PANE_RECORD_TYPE,
            "PANE right-pane column exceeds 255",
        ));
    }
    if read_u8(data, 9, PANE_RECORD_TYPE, "PANE.reserved")? != 0 {
        return Err(invalid(PANE_RECORD_TYPE, "PANE reserved byte must be zero"));
    }
    Ok(Pane {
        horizontal_split,
        vertical_split,
        bottom_pane_top_row: read_u16(data, 4, PANE_RECORD_TYPE, "PANE.bottom_pane_top_row")?,
        right_pane_left_column: right_pane_left_column as u8,
        active_pane: PaneType::parse(
            read_u8(data, 8, PANE_RECORD_TYPE, "PANE.active_pane")?,
            PANE_RECORD_TYPE,
        )?,
    })
}

pub(super) fn parse_selection(data: &[u8]) -> Result<Selection> {
    if data.len() < 9 {
        return Err(invalid(
            SELECTION_RECORD_TYPE,
            "SELECTION payload is shorter than 9 bytes",
        ));
    }
    let active_column = read_u16(data, 3, SELECTION_RECORD_TYPE, "SELECTION.active_column")?;
    let active_range_index = read_u16(
        data,
        5,
        SELECTION_RECORD_TYPE,
        "SELECTION.active_range_index",
    )?;
    let range_count = usize::from(read_u16(
        data,
        7,
        SELECTION_RECORD_TYPE,
        "SELECTION.range_count",
    )?);
    if active_column > 255 {
        return Err(invalid(
            SELECTION_RECORD_TYPE,
            "SELECTION active column exceeds 255",
        ));
    }
    if active_range_index & 0x8000 != 0 {
        return Err(invalid(
            SELECTION_RECORD_TYPE,
            "SELECTION active range index must be nonnegative",
        ));
    }
    if range_count > 1369 || data.len() != 9 + range_count * 6 {
        return Err(invalid(
            SELECTION_RECORD_TYPE,
            "SELECTION range count does not match its payload",
        ));
    }
    let mut ranges = Vec::with_capacity(range_count);
    let range_data = data
        .get(9..)
        .ok_or_else(|| invalid(SELECTION_RECORD_TYPE, "truncated SELECTION ranges"))?;
    for chunk in range_data.chunks_exact(6) {
        let range = Range {
            first_row: read_u16(chunk, 0, SELECTION_RECORD_TYPE, "SELECTION.first_row")?,
            last_row: read_u16(chunk, 2, SELECTION_RECORD_TYPE, "SELECTION.last_row")?,
            first_column: read_u8(chunk, 4, SELECTION_RECORD_TYPE, "SELECTION.first_column")?,
            last_column: read_u8(chunk, 5, SELECTION_RECORD_TYPE, "SELECTION.last_column")?,
        };
        if range.first_row > range.last_row || range.first_column > range.last_column {
            return Err(invalid(
                SELECTION_RECORD_TYPE,
                "SELECTION contains an inverted range",
            ));
        }
        ranges.push(range);
    }
    Ok(Selection {
        pane: PaneType::parse(
            read_u8(data, 0, SELECTION_RECORD_TYPE, "SELECTION.pane")?,
            SELECTION_RECORD_TYPE,
        )?,
        active_row: read_u16(data, 1, SELECTION_RECORD_TYPE, "SELECTION.active_row")?,
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
    views: Vec<View>,
    current: Option<View>,
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

    fn finish_current(&mut self) -> Result<()> {
        if let Some(view) = self.current.take() {
            view.validate_selection_groups()?;
            if let Some(pane) = view.pane.as_ref() {
                if pane.horizontal_split == 0 && pane.vertical_split == 0 {
                    return Err(invalid(PANE_RECORD_TYPE, "PANE does not split either axis"));
                }
                if !super::model::pane_exists(
                    pane.horizontal_split,
                    pane.vertical_split,
                    pane.active_pane,
                ) {
                    return Err(invalid(PANE_RECORD_TYPE, "PANE active pane does not exist"));
                }
                if view.selections.iter().any(|selection| {
                    !super::model::pane_exists(
                        pane.horizontal_split,
                        pane.vertical_split,
                        selection.pane,
                    )
                }) {
                    return Err(invalid(
                        SELECTION_RECORD_TYPE,
                        "SELECTION references a pane that does not exist",
                    ));
                }
            }
            self.views.push(view);
        }
        self.phase = ViewPhase::Start;
        self.saw_plv = false;
        Ok(())
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
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
            },
            SCL_RECORD_TYPE if self.phase == ViewPhase::Start => {
                view.zoom_fraction = Some(parse_scl(data)?);
                self.phase = ViewPhase::Zoom;
            },
            PANE_RECORD_TYPE if matches!(self.phase, ViewPhase::Start | ViewPhase::Zoom) => {
                view.pane = Some(parse_pane(data, view.has_frozen_panes())?);
                self.phase = ViewPhase::Pane;
            },
            SELECTION_RECORD_TYPE => {
                let selection = parse_selection(data)?;
                if let Some(previous) = view.selections.last()
                    && previous.pane == selection.pane
                    && (previous.active_row != selection.active_row
                        || previous.active_column != selection.active_column
                        || previous.active_range_index != selection.active_range_index)
                {
                    return Err(invalid(
                        SELECTION_RECORD_TYPE,
                        "contiguous selections for one pane disagree on the active cell",
                    ));
                }
                view.selections.push(selection);
                self.phase = ViewPhase::Selections;
            },
            PLV_RECORD_TYPE | SCL_RECORD_TYPE | PANE_RECORD_TYPE => {
                return Err(invalid(
                    record_type,
                    "record is out of order in the WINDOW production",
                ));
            },
            _ => self.finish_current()?,
        }
        Ok(())
    }

    pub(crate) fn finish(mut self) -> Result<Vec<View>> {
        self.finish_current()?;
        Ok(self.views)
    }
}
