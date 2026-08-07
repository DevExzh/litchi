//! XLSB worksheet-view record adapter.
//!
//! This module owns only the BIFF12 wire contract. The format-neutral view
//! model lives in `litchi_sheet::view`.

use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Writer, kind};
use litchi_sheet::view::{
    Color, Display, Mode, Pane, Position, Scale, Selection, Split, State, View, Window, Zoom,
};
use litchi_sheet::{Cell, Rect};

/// Maximum number of `BrtBeginWsView` records in one collection.
pub const MAX_SHEET_VIEWS: usize = 1024;
/// Maximum number of `BrtSel` records attached to one worksheet view.
pub const MAX_SHEET_VIEW_SELECTIONS: usize = 4;
/// Maximum number of ranges in a `BrtSel` `sqrfx` collection.
pub const MAX_SHEET_VIEW_SELECTION_RANGES: usize = 32_767;

const FLAG_WINDOW_PROTECTION: u16 = 0x0001;
const FLAG_SHOW_FORMULAS: u16 = 0x0002;
const FLAG_SHOW_GRID_LINES: u16 = 0x0004;
const FLAG_SHOW_ROW_COLUMN_HEADERS: u16 = 0x0008;
const FLAG_SHOW_ZERO_VALUES: u16 = 0x0010;
const FLAG_RIGHT_TO_LEFT: u16 = 0x0020;
const FLAG_TAB_SELECTED: u16 = 0x0040;
const FLAG_SHOW_RULER: u16 = 0x0080;
const FLAG_SHOW_OUTLINE_SYMBOLS: u16 = 0x0100;
const FLAG_DEFAULT_GRID_COLOR: u16 = 0x0200;
const FLAG_WHITESPACE_HIDDEN: u16 = 0x0400;
const VIEW_FLAG_MASK: u16 = FLAG_WINDOW_PROTECTION
    | FLAG_SHOW_FORMULAS
    | FLAG_SHOW_GRID_LINES
    | FLAG_SHOW_ROW_COLUMN_HEADERS
    | FLAG_SHOW_ZERO_VALUES
    | FLAG_RIGHT_TO_LEFT
    | FLAG_TAB_SELECTED
    | FLAG_SHOW_RULER
    | FLAG_SHOW_OUTLINE_SYMBOLS
    | FLAG_DEFAULT_GRID_COLOR
    | FLAG_WHITESPACE_HIDDEN;

const PANE_FLAG_FROZEN: u8 = 0x01;
const PANE_FLAG_FROZEN_NO_SPLIT: u8 = 0x02;
const PANE_FLAG_MASK: u8 = PANE_FLAG_FROZEN | PANE_FLAG_FROZEN_NO_SPLIT;

const WS_VIEW_LEN: usize = 30;
const PANE_LEN: usize = 29;
const SELECTION_HEADER_LEN: usize = 20;
const RANGE_LEN: usize = 16;

fn malformed(context: &str, detail: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: context.to_string(),
        val: detail.into(),
    }
}

fn position_from_u32(value: u32, context: &str) -> Result<Position> {
    match value {
        0 => Ok(Position::BottomRight),
        1 => Ok(Position::TopRight),
        2 => Ok(Position::BottomLeft),
        3 => Ok(Position::TopLeft),
        _ => Err(malformed(context, format!("invalid pane position {value}"))),
    }
}

fn position_to_u32(position: Position) -> u32 {
    match position {
        Position::BottomRight => 0,
        Position::TopRight => 1,
        Position::BottomLeft => 2,
        Position::TopLeft => 3,
    }
}

fn mode_from_u32(value: u32, context: &str) -> Result<Mode> {
    match value {
        0 => Ok(Mode::Normal),
        1 => Ok(Mode::PageBreakPreview),
        2 => Ok(Mode::PageLayout),
        _ => Err(malformed(context, format!("invalid xlView value {value}"))),
    }
}

fn mode_to_u32(mode: Mode) -> u32 {
    match mode {
        Mode::Normal => 0,
        Mode::PageBreakPreview => 1,
        Mode::PageLayout => 2,
    }
}

fn cell(row: u32, column: u32, context: &str) -> Result<Cell> {
    Cell::at(row, column).map_err(|error| malformed(context, error.to_string()))
}

fn scale(value: u16, context: &str) -> Result<Scale> {
    Scale::new(value).map_err(|error| malformed(context, error.to_string()))
}

fn split(value: f64, context: &str) -> Result<Split> {
    Split::new(value).map_err(|error| malformed(context, error.to_string()))
}

fn optional_scale(value: u16, context: &str) -> Result<Option<Scale>> {
    (value != 0).then(|| scale(value, context)).transpose()
}

fn legacy_default_view() -> View {
    View {
        tab_selected: true,
        ..View::default()
    }
}

/// Parse a `BrtBeginWsView` payload.
pub fn parse_view(data: &[u8]) -> Result<View> {
    let context = "BrtBeginWsView";
    let mut cursor = Cursor::new(data, context);
    let flags = cursor.read_u16()?;
    if flags & !VIEW_FLAG_MASK != 0 {
        return Err(malformed(context, "reserved view flags are set"));
    }

    let mode = mode_from_u32(cursor.read_u32()?, context)?;
    let origin = cell(cursor.read_u32()?, cursor.read_u32()?, context)?;
    let color =
        Color::new(cursor.read_u8()?).map_err(|error| malformed(context, error.to_string()))?;
    let reserved2 = cursor.read_u8()?;
    let reserved3 = cursor.read_u16()?;
    if reserved2 != 0 || reserved3 != 0 {
        return Err(malformed(context, "reserved view fields are nonzero"));
    }

    let current = scale(cursor.read_u16()?, context)?;
    let normal = optional_scale(cursor.read_u16()?, context)?;
    let page_break_preview = optional_scale(cursor.read_u16()?, context)?;
    let page_layout = optional_scale(cursor.read_u16()?, context)?;
    let window = Window::new(cursor.read_u32()?);
    cursor.finish()?;
    debug_assert_eq!(data.len(), WS_VIEW_LEN);

    Ok(View {
        window,
        mode,
        color,
        display: Display {
            window_protection: flags & FLAG_WINDOW_PROTECTION != 0,
            show_formulas: flags & FLAG_SHOW_FORMULAS != 0,
            grid_lines: flags & FLAG_SHOW_GRID_LINES != 0,
            row_column_headers: flags & FLAG_SHOW_ROW_COLUMN_HEADERS != 0,
            zero_values: flags & FLAG_SHOW_ZERO_VALUES != 0,
            right_to_left: flags & FLAG_RIGHT_TO_LEFT != 0,
            ruler: flags & FLAG_SHOW_RULER != 0,
            outline_symbols: flags & FLAG_SHOW_OUTLINE_SYMBOLS != 0,
            default_grid_color: flags & FLAG_DEFAULT_GRID_COLOR != 0,
            // `fWhitespaceHidden` is the inverse of the shared semantic flag.
            white_space: flags & FLAG_WHITESPACE_HIDDEN == 0,
        },
        zoom: Zoom {
            current,
            normal,
            page_layout,
            page_break_preview,
        },
        origin,
        pane: None,
        selections: Vec::new(),
        tab_selected: flags & FLAG_TAB_SELECTED != 0,
    })
}

/// Serialize a `BrtBeginWsView` payload. `None` writes the legacy default.
pub fn write_view_payload(view: Option<&View>) -> Result<Vec<u8>> {
    let view = view.cloned().unwrap_or_else(legacy_default_view);
    let display = view.display;
    let mut flags = 0u16;

    for (bit, enabled) in [
        (FLAG_WINDOW_PROTECTION, display.window_protection),
        (FLAG_SHOW_FORMULAS, display.show_formulas),
        (FLAG_SHOW_GRID_LINES, display.grid_lines),
        (FLAG_SHOW_ROW_COLUMN_HEADERS, display.row_column_headers),
        (FLAG_SHOW_ZERO_VALUES, display.zero_values),
        (FLAG_RIGHT_TO_LEFT, display.right_to_left),
        (FLAG_SHOW_RULER, display.ruler),
        (FLAG_SHOW_OUTLINE_SYMBOLS, display.outline_symbols),
        (FLAG_DEFAULT_GRID_COLOR, display.default_grid_color),
        (FLAG_WHITESPACE_HIDDEN, !display.white_space),
    ] {
        if enabled {
            flags |= bit;
        }
    }
    if view.tab_selected {
        flags |= FLAG_TAB_SELECTED;
    }

    let mut payload = Vec::with_capacity(WS_VIEW_LEN);
    let mut writer = Writer::new(&mut payload);
    writer.write_u16(flags)?;
    writer.write_u32(mode_to_u32(view.mode))?;
    writer.write_u32(view.origin.row().get())?;
    writer.write_u32(view.origin.column().get())?;
    writer.write_u8(view.color.get())?;
    writer.write_u8(0)?;
    writer.write_u16(0)?;
    writer.write_u16(view.zoom.current.get())?;
    writer.write_u16(view.zoom.normal.map_or(0, Scale::get))?;
    writer.write_u16(view.zoom.page_break_preview.map_or(0, Scale::get))?;
    writer.write_u16(view.zoom.page_layout.map_or(0, Scale::get))?;
    writer.write_u32(view.window.get())?;
    debug_assert_eq!(payload.len(), WS_VIEW_LEN);
    Ok(payload)
}

/// Parse a `BrtPane` payload.
pub fn parse_pane(data: &[u8]) -> Result<Pane> {
    let context = "BrtPane";
    let mut cursor = Cursor::new(data, context);
    let horizontal = cursor.read_f64()?;
    let vertical = cursor.read_f64()?;
    let top_left = cell(cursor.read_u32()?, cursor.read_u32()?, context)?;
    let position = position_from_u32(cursor.read_u32()?, context)?;
    let flags = cursor.read_u8()?;
    cursor.finish()?;
    debug_assert_eq!(data.len(), PANE_LEN);

    if flags & !PANE_FLAG_MASK != 0 {
        return Err(malformed(context, "reserved pane flags are set"));
    }
    if flags & PANE_FLAG_MASK == PANE_FLAG_MASK {
        return Err(malformed(
            context,
            "fFrozen and fFrozenNoSplit are both set",
        ));
    }

    let state = if flags & PANE_FLAG_FROZEN != 0 {
        State::FrozenSplit
    } else if flags & PANE_FLAG_FROZEN_NO_SPLIT != 0 {
        State::Frozen
    } else {
        State::Split
    };

    Ok(Pane {
        position,
        state,
        horizontal: (horizontal != 0.0)
            .then(|| split(horizontal, context))
            .transpose()?,
        vertical: (vertical != 0.0)
            .then(|| split(vertical, context))
            .transpose()?,
        top_left,
    })
}

/// Serialize a `BrtPane` payload.
pub fn write_pane_payload(pane: &Pane) -> Result<Vec<u8>> {
    let mut payload = Vec::with_capacity(PANE_LEN);
    let mut writer = Writer::new(&mut payload);
    writer.write_f64(pane.horizontal.map_or(0.0, Split::get))?;
    writer.write_f64(pane.vertical.map_or(0.0, Split::get))?;
    writer.write_u32(pane.top_left.row().get())?;
    writer.write_u32(pane.top_left.column().get())?;
    writer.write_u32(position_to_u32(pane.position))?;
    writer.write_u8(match pane.state {
        State::Split => 0,
        State::Frozen => PANE_FLAG_FROZEN_NO_SPLIT,
        State::FrozenSplit => PANE_FLAG_FROZEN,
    })?;
    debug_assert_eq!(payload.len(), PANE_LEN);
    Ok(payload)
}

/// Parse a `BrtSel` payload.
pub fn parse_selection(data: &[u8]) -> Result<Selection> {
    let context = "BrtSel";
    let mut cursor = Cursor::new(data, context);
    let position = position_from_u32(cursor.read_u32()?, context)?;
    let active_cell = cell(cursor.read_u32()?, cursor.read_u32()?, context)?;
    let active_range = usize::try_from(cursor.read_u32()?)
        .map_err(|_| malformed(context, "active range index overflows usize"))?;
    let count = usize::try_from(cursor.read_u32()?)
        .map_err(|_| malformed(context, "range count overflows usize"))?;
    if count == 0 || count > MAX_SHEET_VIEW_SELECTION_RANGES {
        return Err(malformed(context, format!("invalid range count {count}")));
    }

    let ranges_len = count
        .checked_mul(RANGE_LEN)
        .ok_or_else(|| malformed(context, "range payload length overflows"))?;
    let expected_len = SELECTION_HEADER_LEN
        .checked_add(ranges_len)
        .ok_or_else(|| malformed(context, "selection payload length overflows"))?;
    if data.len() != expected_len || cursor.remaining() != ranges_len {
        return Err(Error::InvalidLength {
            expected: expected_len,
            found: data.len(),
        });
    }

    let mut ranges = Vec::with_capacity(count);
    for _ in 0..count {
        let first_row = cursor.read_u32()?;
        let last_row = cursor.read_u32()?;
        let first_column = cursor.read_u32()?;
        let last_column = cursor.read_u32()?;
        let first = cell(first_row, first_column, context)?;
        let _last = cell(last_row, last_column, context)?;
        let end_row = last_row
            .checked_add(1)
            .ok_or_else(|| malformed(context, "range row end overflows"))?;
        let end_column = last_column
            .checked_add(1)
            .ok_or_else(|| malformed(context, "range column end overflows"))?;
        ranges.push(
            Rect::new(first, end_row, end_column)
                .map_err(|error| malformed(context, error.to_string()))?,
        );
    }
    cursor.finish()?;
    Selection::new(position, active_cell, active_range, ranges)
        .map_err(|error| malformed(context, error.to_string()))
}

/// Serialize a `BrtSel` payload.
pub fn write_selection_payload(selection: &Selection) -> Result<Vec<u8>> {
    let context = "BrtSel";
    let range_count = selection.ranges().len();
    if range_count == 0 || range_count > MAX_SHEET_VIEW_SELECTION_RANGES {
        return Err(malformed(
            context,
            format!("invalid range count {range_count}"),
        ));
    }
    let range_count =
        u32::try_from(range_count).map_err(|_| malformed(context, "range count exceeds u32"))?;
    let active_range = u32::try_from(selection.active_range())
        .map_err(|_| malformed(context, "active range index exceeds u32"))?;

    let ranges_len = selection
        .ranges()
        .len()
        .checked_mul(RANGE_LEN)
        .ok_or_else(|| malformed(context, "range payload length overflows"))?;
    let capacity = SELECTION_HEADER_LEN
        .checked_add(ranges_len)
        .ok_or_else(|| malformed(context, "selection payload length overflows"))?;
    let mut payload = Vec::with_capacity(capacity);
    let mut writer = Writer::new(&mut payload);
    writer.write_u32(position_to_u32(selection.position()))?;
    writer.write_u32(selection.active_cell().row().get())?;
    writer.write_u32(selection.active_cell().column().get())?;
    writer.write_u32(active_range)?;
    writer.write_u32(range_count)?;
    for range in selection.ranges() {
        let start = range.start();
        let (end_row, end_column) = range.end();
        let last_row = end_row
            .checked_sub(1)
            .ok_or_else(|| malformed(context, "range row end underflows"))?;
        let last_column = end_column
            .checked_sub(1)
            .ok_or_else(|| malformed(context, "range column end underflows"))?;
        let _last = cell(last_row, last_column, context)?;
        writer.write_u32(start.row().get())?;
        writer.write_u32(last_row)?;
        writer.write_u32(start.column().get())?;
        writer.write_u32(last_column)?;
    }
    debug_assert_eq!(payload.len(), capacity);
    Ok(payload)
}

/// Read a `BrtBeginWsViews` collection through its matching end record.
pub(crate) fn read_views<RS: std::io::Read + std::io::Seek>(
    iter: &mut crate::package::records::Stream<RS>,
    buf: &mut Vec<u8>,
) -> Result<Vec<View>> {
    let context = "BrtBeginWsViews collection";
    let mut views = Vec::new();
    let mut current = None;

    loop {
        let record_kind = iter.read_type()?;
        iter.fill_buffer(buf)?;
        match record_kind {
            kind::BEGIN_WS_VIEW => {
                if current.is_some() {
                    return Err(malformed(context, "nested BrtBeginWsView"));
                }
                if views.len() >= MAX_SHEET_VIEWS {
                    return Err(malformed(context, "too many worksheet views"));
                }
                views.push(parse_view(buf)?);
                current = Some(views.len() - 1);
            },
            kind::END_WS_VIEW => {
                if current.take().is_none() {
                    return Err(malformed(context, "BrtEndWsView without BrtBeginWsView"));
                }
            },
            kind::PANE => {
                let index =
                    current.ok_or_else(|| malformed(context, "BrtPane outside BrtBeginWsView"))?;
                let view = &mut views[index];
                if view.pane.is_some() {
                    return Err(malformed(context, "duplicate BrtPane"));
                }
                view.pane = Some(parse_pane(buf)?);
            },
            kind::SEL => {
                let index =
                    current.ok_or_else(|| malformed(context, "BrtSel outside BrtBeginWsView"))?;
                let view = &mut views[index];
                if view.selections.len() >= MAX_SHEET_VIEW_SELECTIONS {
                    return Err(malformed(context, "worksheet view exceeds four selections"));
                }
                view.selections.push(parse_selection(buf)?);
            },
            kind::END_WS_VIEWS => {
                if current.is_some() {
                    return Err(malformed(context, "unterminated BrtBeginWsView"));
                }
                if views.is_empty() {
                    return Err(malformed(context, "worksheet view collection is empty"));
                }
                return Ok(views);
            },
            _ => {},
        }
    }
}
