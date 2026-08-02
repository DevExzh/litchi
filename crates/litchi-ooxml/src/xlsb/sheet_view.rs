//! Worksheet sheet-view support for XLSB.
//!
//! Parses and serializes the sheet-view record sequence of a Worksheet part
//! ([MS-XLSB] 2.1.7.62): `BrtBeginWsViews` (2.4.308), `BrtBeginWsView`
//! (2.4.307), `BrtPane` (2.4.723), `BrtSel` (2.4.790), `BrtEndWsView`
//! (2.4.659), and `BrtEndWsViews` (2.4.660).
//!
//! The typed model is shared with the XLSX implementation
//! ([`crate::xlsx::views`]), so sheet views behave identically across both
//! formats: pane/frozen-split state, zoom scales, selections, and the
//! tab-selected flag.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::utils::{cell_reference, parse_cell_reference};
use litchi_xlsb::raw::{Cursor, Writer, kind};

pub use crate::xlsx::views::{
    SheetPane, SheetPanePosition, SheetPaneState, SheetSelection, SheetView, SheetViewType,
};

/// Maximum number of `BrtBeginWsView` records in one `BrtBeginWsViews` collection.
pub const MAX_SHEET_VIEWS: usize = 1024;
/// Maximum number of `BrtSel` records attached to one sheet view.
pub const MAX_SHEET_VIEW_SELECTIONS: usize = 4;
/// Maximum number of ranges in one `BrtSel` `sqrfx` collection (MS-XLSB 2.4.790).
pub const MAX_SHEET_VIEW_SELECTION_RANGES: usize = 32_767;

// BrtBeginWsView flag bits ([MS-XLSB] 2.4.307).
const FLAG_WINDOW_PROTECTION: u16 = 0x0001; // A - fWnProt
const FLAG_SHOW_FORMULAS: u16 = 0x0002; // B - fDspFmla
const FLAG_SHOW_GRID_LINES: u16 = 0x0004; // C - fDspGrid
const FLAG_SHOW_ROW_COL_HEADERS: u16 = 0x0008; // D - fDspRwCol
const FLAG_SHOW_ZEROS: u16 = 0x0010; // E - fDspZeros
const FLAG_RIGHT_TO_LEFT: u16 = 0x0020; // F - fRightToLeft
const FLAG_TAB_SELECTED: u16 = 0x0040; // G - fSelected
const FLAG_SHOW_RULER: u16 = 0x0080; // H - fDspRuler
const FLAG_SHOW_OUTLINE_SYMBOLS: u16 = 0x0100; // I - fDspGuts
const FLAG_DEFAULT_GRID_COLOR: u16 = 0x0200; // J - fDefaultHdr
const FLAG_WHITESPACE_HIDDEN: u16 = 0x0400; // K - fWhitespaceHidden

// BrtPane flag bits ([MS-XLSB] 2.4.723).
const PANE_FLAG_FROZEN: u8 = 0x01; // A - fFrozen
const PANE_FLAG_FROZEN_NO_SPLIT: u8 = 0x02; // B - fFrozenNoSplit

/// Serialized length of a `BrtBeginWsView` record payload.
const WS_VIEW_LEN: usize = 30;
/// Serialized length of a `BrtPane` record payload.
const PANE_LEN: usize = 29;
/// Fixed portion of a `BrtSel` record before the `sqrfx` range collection.
const SEL_HEADER_LEN: usize = 20;
/// Serialized length of one `UncheckedRfX` range (MS-XLSB 2.5.154).
const RANGE_LEN: usize = 16;

const MAX_ZOOM_SCALE: u16 = 400;
const MIN_ZOOM_SCALE: u16 = 10;

fn malformed(context: &str, detail: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: context.to_string(),
        val: detail.into(),
    }
}

fn pane_position_from_u32(value: u32) -> Option<SheetPanePosition> {
    match value {
        0 => Some(SheetPanePosition::BottomRight),
        1 => Some(SheetPanePosition::TopRight),
        2 => Some(SheetPanePosition::BottomLeft),
        3 => Some(SheetPanePosition::TopLeft),
        _ => None,
    }
}

fn pane_position_to_u32(position: SheetPanePosition) -> u32 {
    match position {
        SheetPanePosition::BottomRight => 0,
        SheetPanePosition::TopRight => 1,
        SheetPanePosition::BottomLeft => 2,
        SheetPanePosition::TopLeft => 3,
    }
}

fn view_type_from_u32(value: u32) -> Option<SheetViewType> {
    match value {
        0 => Some(SheetViewType::Normal),
        1 => Some(SheetViewType::PageBreakPreview),
        2 => Some(SheetViewType::PageLayout),
        _ => None,
    }
}

fn view_type_to_u32(view_type: SheetViewType) -> u32 {
    match view_type {
        SheetViewType::Normal => 0,
        SheetViewType::PageBreakPreview => 1,
        SheetViewType::PageLayout => 2,
    }
}

/// Inclusive cell range from an `UncheckedRfX` structure ([MS-XLSB] 2.5.154).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ViewRange {
    row_first: u32,
    row_last: u32,
    col_first: u32,
    col_last: u32,
}

impl ViewRange {
    fn format(&self) -> String {
        let first = cell_reference(self.row_first, self.col_first);
        if self.row_first == self.row_last && self.col_first == self.col_last {
            first
        } else {
            format!("{}:{}", first, cell_reference(self.row_last, self.col_last))
        }
    }
}

/// Parse one A1-style range token (`A1` or `A1:B2`) for writer-side sqref input.
fn parse_range_token(token: &str) -> XlsbResult<ViewRange> {
    let invalid = || XlsbError::InvalidCellReference(token.to_string());
    let mut cells = token.split(':');
    let first = cells.next().ok_or_else(invalid)?;
    if first.is_empty() {
        return Err(invalid());
    }
    let (row_first, col_first) = parse_cell_reference(first)?;
    let (row_last, col_last) = match cells.next() {
        Some(second) => {
            if second.is_empty() {
                return Err(invalid());
            }
            parse_cell_reference(second)?
        },
        None => (row_first, col_first),
    };
    if cells.next().is_some() || row_first > row_last || col_first > col_last {
        return Err(invalid());
    }
    Ok(ViewRange {
        row_first,
        row_last,
        col_first,
        col_last,
    })
}

/// Parse an A1-style cell reference, rejecting out-of-sheet and overflowing input.
fn parse_view_cell(reference: &str, context: &str) -> XlsbResult<(u32, u32)> {
    // Guard against absurd input before the unchecked arithmetic in
    // `parse_cell_reference` (e.g. column names long enough to overflow u32).
    if reference.len() > 10 {
        return Err(malformed(context, format!("invalid cell '{reference}'")));
    }
    let (row, col) = parse_cell_reference(reference)?;
    if row > crate::xlsb::merged_cells::MAX_MERGED_CELL_ROW
        || col > crate::xlsb::merged_cells::MAX_MERGED_CELL_COLUMN
    {
        return Err(malformed(
            context,
            format!("cell '{reference}' is out of bounds"),
        ));
    }
    Ok((row, col))
}

fn zoom_or_default(value: Option<u16>, default: u16, context: &str) -> XlsbResult<u16> {
    let zoom = value.unwrap_or(default);
    if zoom != 0 && !(MIN_ZOOM_SCALE..=MAX_ZOOM_SCALE).contains(&zoom) {
        return Err(malformed(
            context,
            format!("zoom scale {zoom} is outside 10..=400"),
        ));
    }
    Ok(zoom)
}

/// Parse a `BrtBeginWsView` record payload ([MS-XLSB] 2.4.307).
///
/// Pane and selection records following the view are attached by
/// `read_sheet_views`, not by this function.
pub fn parse_ws_view(data: &[u8]) -> XlsbResult<SheetView> {
    let context = "BrtBeginWsView";
    let mut cursor = Cursor::new(data, context);
    let flags = cursor.read_u16()?;
    let xl_view = cursor.read_u32()?;
    let row_top = cursor.read_u32()?;
    let col_left = cursor.read_u32()?;
    let icv_hdr = cursor.read_u8()?;
    cursor.skip(1)?; // reserved2
    cursor.skip(2)?; // reserved3
    let w_scale = cursor.read_u16()?;
    let w_scale_normal = cursor.read_u16()?;
    let w_scale_slv = cursor.read_u16()?;
    let w_scale_plv = cursor.read_u16()?;
    let i_wbk_view = cursor.read_u32()?;
    cursor.finish()?;

    // fWhitespaceHidden is the inverse of the XLSX `showWhiteSpace` attribute.
    let show_white_space = flags & FLAG_WHITESPACE_HIDDEN == 0;
    Ok(SheetView {
        workbook_view_id: Some(i_wbk_view),
        window_protection: Some(flags & FLAG_WINDOW_PROTECTION != 0),
        show_formulas: Some(flags & FLAG_SHOW_FORMULAS != 0),
        show_grid_lines: Some(flags & FLAG_SHOW_GRID_LINES != 0),
        show_row_col_headers: Some(flags & FLAG_SHOW_ROW_COL_HEADERS != 0),
        show_zeros: Some(flags & FLAG_SHOW_ZEROS != 0),
        right_to_left: Some(flags & FLAG_RIGHT_TO_LEFT != 0),
        tab_selected: Some(flags & FLAG_TAB_SELECTED != 0),
        show_ruler: Some(flags & FLAG_SHOW_RULER != 0),
        show_outline_symbols: Some(flags & FLAG_SHOW_OUTLINE_SYMBOLS != 0),
        default_grid_color: Some(flags & FLAG_DEFAULT_GRID_COLOR != 0),
        show_white_space: Some(show_white_space),
        view_type: view_type_from_u32(xl_view),
        top_left_cell: Some(cell_reference(row_top, col_left)),
        color_id: Some(u32::from(icv_hdr)),
        zoom_scale: Some(w_scale),
        // A zero wScaleNormal/wScaleSLV/wScalePLV means "default 100" and is
        // reported as absent, matching the XLSX attribute defaults.
        zoom_scale_normal: (w_scale_normal != 0).then_some(w_scale_normal),
        zoom_scale_sheet_layout_view: (w_scale_slv != 0).then_some(w_scale_slv),
        zoom_scale_page_layout_view: (w_scale_plv != 0).then_some(w_scale_plv),
        pane: None,
        selections: Vec::new(),
    })
}

/// Serialize a `BrtBeginWsView` record payload ([MS-XLSB] 2.4.307).
///
/// `None` emits the same default view the crate has always written for
/// otherwise-unconfigured worksheets.
pub fn write_ws_view_payload(view: Option<&SheetView>) -> XlsbResult<Vec<u8>> {
    let context = "BrtBeginWsView";
    let mut flags = 0u16;
    let mut set_flag = |bit: u16, value: Option<bool>, default: bool| {
        if value.unwrap_or(default) {
            flags |= bit;
        }
    };
    set_flag(
        FLAG_WINDOW_PROTECTION,
        view.and_then(|v| v.window_protection),
        false,
    );
    set_flag(
        FLAG_SHOW_FORMULAS,
        view.and_then(|v| v.show_formulas),
        false,
    );
    set_flag(
        FLAG_SHOW_GRID_LINES,
        view.and_then(|v| v.show_grid_lines),
        true,
    );
    set_flag(
        FLAG_SHOW_ROW_COL_HEADERS,
        view.and_then(|v| v.show_row_col_headers),
        true,
    );
    set_flag(FLAG_SHOW_ZEROS, view.and_then(|v| v.show_zeros), true);
    set_flag(
        FLAG_RIGHT_TO_LEFT,
        view.and_then(|v| v.right_to_left),
        false,
    );
    set_flag(FLAG_TAB_SELECTED, view.and_then(|v| v.tab_selected), true);
    set_flag(FLAG_SHOW_RULER, view.and_then(|v| v.show_ruler), true);
    set_flag(
        FLAG_SHOW_OUTLINE_SYMBOLS,
        view.and_then(|v| v.show_outline_symbols),
        true,
    );
    set_flag(
        FLAG_DEFAULT_GRID_COLOR,
        view.and_then(|v| v.default_grid_color),
        true,
    );
    // show_white_space defaults to shown, so fWhitespaceHidden defaults to 0.
    if !view.and_then(|v| v.show_white_space).unwrap_or(true) {
        flags |= FLAG_WHITESPACE_HIDDEN;
    }

    let xl_view = view
        .and_then(|v| v.view_type)
        .map(view_type_to_u32)
        .unwrap_or(0);
    let (row_top, col_left) = match view.and_then(|v| v.top_left_cell.as_deref()) {
        Some(reference) => parse_view_cell(reference, context)?,
        None => (0, 0),
    };
    let color_id = view.and_then(|v| v.color_id).unwrap_or(0x40);
    if color_id > u32::from(u8::MAX) {
        return Err(malformed(
            context,
            format!("grid color index {color_id} exceeds one byte"),
        ));
    }

    let mut payload = Vec::with_capacity(WS_VIEW_LEN);
    let mut writer = Writer::new(&mut payload);
    writer.write_u16(flags)?;
    writer.write_u32(xl_view)?;
    writer.write_u32(row_top)?;
    writer.write_u32(col_left)?;
    writer.write_u8(color_id as u8)?;
    writer.write_u8(0)?; // reserved2
    writer.write_u16(0)?; // reserved3
    writer.write_u16(zoom_or_default(
        view.and_then(|v| v.zoom_scale),
        100,
        context,
    )?)?;
    writer.write_u16(zoom_or_default(
        view.and_then(|v| v.zoom_scale_normal),
        0,
        context,
    )?)?;
    writer.write_u16(zoom_or_default(
        view.and_then(|v| v.zoom_scale_sheet_layout_view),
        0,
        context,
    )?)?;
    writer.write_u16(zoom_or_default(
        view.and_then(|v| v.zoom_scale_page_layout_view),
        0,
        context,
    )?)?;
    writer.write_u32(view.and_then(|v| v.workbook_view_id).unwrap_or(0))?;
    debug_assert_eq!(payload.len(), WS_VIEW_LEN);
    Ok(payload)
}

/// Parse a `BrtPane` record payload ([MS-XLSB] 2.4.723).
pub fn parse_pane(data: &[u8]) -> XlsbResult<SheetPane> {
    let context = "BrtPane";
    let mut cursor = Cursor::new(data, context);
    let x_split = cursor.read_f64()?;
    let y_split = cursor.read_f64()?;
    let row_top = cursor.read_u32()?;
    let col_left = cursor.read_u32()?;
    let active = cursor.read_u32()?;
    let flags = cursor.read_u8()?;
    cursor.finish()?;
    if flags & PANE_FLAG_FROZEN != 0 && flags & PANE_FLAG_FROZEN_NO_SPLIT != 0 {
        return Err(malformed(
            context,
            "fFrozen and fFrozenNoSplit are both set",
        ));
    }
    let state = if flags & PANE_FLAG_FROZEN != 0 {
        SheetPaneState::FrozenSplit
    } else if flags & PANE_FLAG_FROZEN_NO_SPLIT != 0 {
        SheetPaneState::Frozen
    } else {
        SheetPaneState::Split
    };
    Ok(SheetPane {
        x_split: (x_split != 0.0).then_some(x_split),
        y_split: (y_split != 0.0).then_some(y_split),
        top_left_cell: Some(cell_reference(row_top, col_left)),
        active_pane: pane_position_from_u32(active),
        state: Some(state),
    })
}

/// Serialize a `BrtPane` record payload ([MS-XLSB] 2.4.723).
pub fn write_pane_payload(pane: &SheetPane) -> XlsbResult<Vec<u8>> {
    let context = "BrtPane";
    let x_split = pane.x_split.unwrap_or(0.0);
    let y_split = pane.y_split.unwrap_or(0.0);
    if !x_split.is_finite() || !y_split.is_finite() {
        return Err(malformed(context, "split positions must be finite"));
    }
    let (row_top, col_left) = match pane.top_left_cell.as_deref() {
        Some(reference) => parse_view_cell(reference, context)?,
        None => (0, 0),
    };
    let flags = match pane.state {
        Some(SheetPaneState::FrozenSplit) => PANE_FLAG_FROZEN,
        Some(SheetPaneState::Frozen) => PANE_FLAG_FROZEN_NO_SPLIT,
        Some(SheetPaneState::Split) | None => 0,
    };
    let mut payload = Vec::with_capacity(PANE_LEN);
    let mut writer = Writer::new(&mut payload);
    writer.write_f64(x_split)?;
    writer.write_f64(y_split)?;
    writer.write_u32(row_top)?;
    writer.write_u32(col_left)?;
    writer.write_u32(pane.active_pane.map(pane_position_to_u32).unwrap_or(0))?;
    writer.write_u8(flags)?;
    debug_assert_eq!(payload.len(), PANE_LEN);
    Ok(payload)
}

/// Parse a `BrtSel` record payload ([MS-XLSB] 2.4.790).
pub fn parse_selection(data: &[u8]) -> XlsbResult<SheetSelection> {
    let context = "BrtSel";
    let mut cursor = Cursor::new(data, context);
    let pnn = cursor.read_u32()?;
    let row_active = cursor.read_u32()?;
    let col_active = cursor.read_u32()?;
    let active_id = cursor.read_u32()?;
    let count = usize::try_from(cursor.read_u32()?)
        .map_err(|_| malformed(context, "sqrfx count overflow"))?;
    if count > MAX_SHEET_VIEW_SELECTION_RANGES {
        return Err(malformed(context, format!("sqrfx contains {count} ranges")));
    }
    if cursor.remaining() != count * RANGE_LEN {
        return Err(XlsbError::InvalidLength {
            expected: SEL_HEADER_LEN + count * RANGE_LEN,
            found: data.len(),
        });
    }
    let mut ranges = Vec::with_capacity(count);
    for _ in 0..count {
        ranges.push(ViewRange {
            row_first: cursor.read_u32()?,
            row_last: cursor.read_u32()?,
            col_first: cursor.read_u32()?,
            col_last: cursor.read_u32()?,
        });
    }
    let sqref = ranges
        .iter()
        .map(ViewRange::format)
        .collect::<Vec<_>>()
        .join(" ");
    Ok(SheetSelection {
        pane: pane_position_from_u32(pnn),
        active_cell: Some(cell_reference(row_active, col_active)),
        active_cell_id: Some(active_id),
        sqref: Some(sqref),
    })
}

/// Serialize a `BrtSel` record payload ([MS-XLSB] 2.4.790).
pub fn write_selection_payload(selection: &SheetSelection) -> XlsbResult<Vec<u8>> {
    let context = "BrtSel";
    let (row_active, col_active) = match selection.active_cell.as_deref() {
        Some(reference) => parse_view_cell(reference, context)?,
        None => (0, 0),
    };
    let ranges = match selection.sqref.as_deref() {
        Some(sqref) => {
            let mut ranges = Vec::new();
            for token in sqref.split_whitespace() {
                ranges.push(parse_range_token(token)?);
            }
            if ranges.is_empty() {
                return Err(malformed(context, "selection sqref cannot be empty"));
            }
            ranges
        },
        // Without an explicit sqref the active cell is the only selected range.
        None => vec![ViewRange {
            row_first: row_active,
            row_last: row_active,
            col_first: col_active,
            col_last: col_active,
        }],
    };
    if ranges.len() > MAX_SHEET_VIEW_SELECTION_RANGES {
        return Err(malformed(
            context,
            format!("selection contains {} ranges", ranges.len()),
        ));
    }
    let active_id = selection.active_cell_id.unwrap_or(0);
    if active_id as usize >= ranges.len() {
        return Err(malformed(
            context,
            format!("active cell index {active_id} is outside the selection ranges"),
        ));
    }

    let mut payload = Vec::with_capacity(SEL_HEADER_LEN + ranges.len() * RANGE_LEN);
    let mut writer = Writer::new(&mut payload);
    writer.write_u32(selection.pane.map(pane_position_to_u32).unwrap_or(0))?;
    writer.write_u32(row_active)?;
    writer.write_u32(col_active)?;
    writer.write_u32(active_id)?;
    writer.write_u32(ranges.len() as u32)?;
    for range in ranges {
        writer.write_u32(range.row_first)?;
        writer.write_u32(range.row_last)?;
        writer.write_u32(range.col_first)?;
        writer.write_u32(range.col_last)?;
    }
    Ok(payload)
}

/// Read one `BrtBeginWsViews` collection from a worksheet stream.
///
/// The iterator is positioned just after the `BrtBeginWsViews` record and is
/// consumed through the matching `BrtEndWsViews` record. Records other than
/// `BrtBeginWsView`, `BrtPane`, and `BrtSel` — including future records — are
/// skipped, matching the reader's tolerance for unmodelled content.
pub(crate) fn read_sheet_views<RS: std::io::Read + std::io::Seek>(
    iter: &mut crate::xlsb::records::Stream<RS>,
    buf: &mut Vec<u8>,
) -> XlsbResult<Vec<SheetView>> {
    let context = "BrtBeginWsViews collection";
    let mut views: Vec<SheetView> = Vec::new();
    // Index of the view opened by the innermost unmatched BrtBeginWsView.
    let mut current: Option<usize> = None;
    loop {
        let typ = iter.read_type()?;
        let _ = iter.fill_buffer(buf)?;
        match typ {
            kind::BEGIN_WS_VIEW => {
                if current.is_some() {
                    return Err(malformed(context, "nested BrtBeginWsView"));
                }
                if views.len() >= MAX_SHEET_VIEWS {
                    return Err(malformed(context, "too many sheet views"));
                }
                views.push(parse_ws_view(buf)?);
                current = Some(views.len() - 1);
            },
            kind::END_WS_VIEW if current.take().is_none() => {
                return Err(malformed(context, "BrtEndWsView without BrtBeginWsView"));
            },
            kind::END_WS_VIEW => {},
            kind::PANE => {
                if let Some(index) = current {
                    let view = &mut views[index];
                    if view.pane.is_some() {
                        return Err(malformed(context, "duplicate BrtPane in sheet view"));
                    }
                    view.pane = Some(parse_pane(buf)?);
                }
            },
            kind::SEL => {
                if let Some(index) = current {
                    let view = &mut views[index];
                    if view.selections.len() >= MAX_SHEET_VIEW_SELECTIONS {
                        return Err(malformed(context, "sheet view exceeds four selections"));
                    }
                    view.selections.push(parse_selection(buf)?);
                }
            },
            kind::END_WS_VIEWS => {
                if current.is_some() {
                    return Err(malformed(context, "unterminated BrtBeginWsView"));
                }
                return Ok(views);
            },
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_xlsb::raw::{Error as WireError, Stage};

    // BrtBeginWsView payload captured from a default Excel worksheet view.
    const EXCEL_WS_VIEW: [u8; 30] = [
        0xDC, 0x03, // flags
        0x00, 0x00, 0x00, 0x00, // xlView
        0x00, 0x00, 0x00, 0x00, // rwTop
        0x00, 0x00, 0x00, 0x00, // colLeft
        0x40, // icvHdr
        0x00, // reserved2
        0x00, 0x00, // reserved3
        0x64, 0x00, // wScale
        0x00, 0x00, // wScaleNormal
        0x00, 0x00, // wScaleSLV
        0x00, 0x00, // wScalePLV
        0x00, 0x00, 0x00, 0x00, // iWbkView
    ];

    // BrtSel payload captured from Excel: top-left pane, active cell A1, one
    // single-cell range A1.
    const EXCEL_SEL: [u8; 36] = [
        0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    ];

    #[test]
    fn parses_excel_default_ws_view() {
        let view = parse_ws_view(&EXCEL_WS_VIEW).unwrap();
        assert_eq!(view.tab_selected, Some(true));
        assert_eq!(view.show_grid_lines, Some(true));
        assert_eq!(view.show_row_col_headers, Some(true));
        assert_eq!(view.show_zeros, Some(true));
        assert_eq!(view.show_ruler, Some(true));
        assert_eq!(view.show_outline_symbols, Some(true));
        assert_eq!(view.default_grid_color, Some(true));
        assert_eq!(view.show_white_space, Some(true));
        assert_eq!(view.window_protection, Some(false));
        assert_eq!(view.show_formulas, Some(false));
        assert_eq!(view.right_to_left, Some(false));
        assert_eq!(view.view_type, Some(SheetViewType::Normal));
        assert_eq!(view.top_left_cell.as_deref(), Some("A1"));
        assert_eq!(view.color_id, Some(0x40));
        assert_eq!(view.zoom_scale, Some(100));
        assert_eq!(view.zoom_scale_normal, None);
        assert_eq!(view.zoom_scale_sheet_layout_view, None);
        assert_eq!(view.zoom_scale_page_layout_view, None);
        assert_eq!(view.workbook_view_id, Some(0));
    }

    #[test]
    fn default_ws_view_payload_matches_legacy_bytes() {
        assert_eq!(write_ws_view_payload(None).unwrap(), EXCEL_WS_VIEW);
    }

    #[test]
    fn ws_view_round_trip_preserves_flags() {
        let view = SheetView {
            window_protection: Some(true),
            show_formulas: Some(true),
            show_grid_lines: Some(false),
            right_to_left: Some(true),
            tab_selected: Some(false),
            show_white_space: Some(false),
            view_type: Some(SheetViewType::PageLayout),
            top_left_cell: Some("C7".to_string()),
            zoom_scale: Some(150),
            zoom_scale_normal: Some(75),
            zoom_scale_sheet_layout_view: Some(60),
            zoom_scale_page_layout_view: Some(200),
            workbook_view_id: Some(1),
            ..SheetView::default()
        };
        let payload = write_ws_view_payload(Some(&view)).unwrap();
        let parsed = parse_ws_view(&payload).unwrap();
        assert_eq!(parsed.window_protection, Some(true));
        assert_eq!(parsed.show_formulas, Some(true));
        assert_eq!(parsed.show_grid_lines, Some(false));
        assert_eq!(parsed.right_to_left, Some(true));
        assert_eq!(parsed.tab_selected, Some(false));
        assert_eq!(parsed.show_white_space, Some(false));
        assert_eq!(parsed.view_type, Some(SheetViewType::PageLayout));
        assert_eq!(parsed.top_left_cell.as_deref(), Some("C7"));
        assert_eq!(parsed.zoom_scale, Some(150));
        assert_eq!(parsed.zoom_scale_normal, Some(75));
        assert_eq!(parsed.zoom_scale_sheet_layout_view, Some(60));
        assert_eq!(parsed.zoom_scale_page_layout_view, Some(200));
        assert_eq!(parsed.workbook_view_id, Some(1));
    }

    #[test]
    fn ws_view_rejects_out_of_range_zoom() {
        let view = SheetView {
            zoom_scale: Some(500),
            ..SheetView::default()
        };
        assert!(write_ws_view_payload(Some(&view)).is_err());
    }

    #[test]
    fn ws_view_rejects_invalid_length() {
        assert!(matches!(
            parse_ws_view(&EXCEL_WS_VIEW[..29]),
            Err(XlsbError::Wire(WireError::Truncated {
                stage: Stage::Value,
                ..
            }))
        ));
    }

    #[test]
    fn frozen_pane_round_trip() {
        let pane = SheetPane {
            x_split: Some(1.0),
            y_split: Some(2.0),
            top_left_cell: Some("B3".to_string()),
            active_pane: Some(SheetPanePosition::BottomRight),
            state: Some(SheetPaneState::Frozen),
        };
        let payload = write_pane_payload(&pane).unwrap();
        assert_eq!(payload.len(), PANE_LEN);
        let parsed = parse_pane(&payload).unwrap();
        assert_eq!(parsed.x_split, Some(1.0));
        assert_eq!(parsed.y_split, Some(2.0));
        assert_eq!(parsed.top_left_cell.as_deref(), Some("B3"));
        assert_eq!(parsed.active_pane, Some(SheetPanePosition::BottomRight));
        assert_eq!(parsed.state, Some(SheetPaneState::Frozen));
    }

    #[test]
    fn pane_state_flag_mapping() {
        for (flags, expected) in [
            (0u8, SheetPaneState::Split),
            (PANE_FLAG_FROZEN, SheetPaneState::FrozenSplit),
            (PANE_FLAG_FROZEN_NO_SPLIT, SheetPaneState::Frozen),
        ] {
            let mut payload = vec![0u8; PANE_LEN - 1];
            payload.push(flags);
            assert_eq!(parse_pane(&payload).unwrap().state, Some(expected));
        }
        let mut payload = vec![0u8; PANE_LEN - 1];
        payload.push(PANE_FLAG_FROZEN | PANE_FLAG_FROZEN_NO_SPLIT);
        assert!(matches!(
            parse_pane(&payload),
            Err(XlsbError::Unrecognized { .. })
        ));
    }

    #[test]
    fn parses_excel_selection() {
        let selection = parse_selection(&EXCEL_SEL).unwrap();
        assert_eq!(selection.pane, Some(SheetPanePosition::TopLeft));
        assert_eq!(selection.active_cell.as_deref(), Some("A1"));
        assert_eq!(selection.active_cell_id, Some(0));
        assert_eq!(selection.sqref.as_deref(), Some("A1"));
    }

    #[test]
    fn selection_multi_range_round_trip() {
        let selection = SheetSelection {
            pane: Some(SheetPanePosition::BottomRight),
            active_cell: Some("D5".to_string()),
            active_cell_id: Some(1),
            sqref: Some("A1:B2 D5".to_string()),
        };
        let payload = write_selection_payload(&selection).unwrap();
        assert_eq!(payload.len(), SEL_HEADER_LEN + 2 * RANGE_LEN);
        let parsed = parse_selection(&payload).unwrap();
        assert_eq!(parsed.pane, selection.pane);
        assert_eq!(parsed.active_cell, selection.active_cell);
        assert_eq!(parsed.active_cell_id, selection.active_cell_id);
        assert_eq!(parsed.sqref, selection.sqref);
    }

    #[test]
    fn selection_rejects_bad_input() {
        // Empty sqref.
        let selection = SheetSelection {
            sqref: Some(String::new()),
            ..SheetSelection::default()
        };
        assert!(write_selection_payload(&selection).is_err());
        // Reversed range.
        let selection = SheetSelection {
            sqref: Some("B2:A1".to_string()),
            ..SheetSelection::default()
        };
        assert!(matches!(
            write_selection_payload(&selection),
            Err(XlsbError::InvalidCellReference(_))
        ));
        // active_cell_id outside the range collection.
        let selection = SheetSelection {
            active_cell_id: Some(1),
            sqref: Some("A1".to_string()),
            ..SheetSelection::default()
        };
        assert!(write_selection_payload(&selection).is_err());
        // Declared range count exceeding the payload.
        let mut payload = EXCEL_SEL.to_vec();
        payload[16] = 2;
        assert!(matches!(
            parse_selection(&payload),
            Err(XlsbError::InvalidLength { .. })
        ));
    }
}
