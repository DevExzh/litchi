//! BIFF8 custom-view records (MS-XLS 2.4.333–2.4.336): `UserBView`,
//! `UserSViewBegin` / `UserSViewBegin_Chart`, and `UserSViewEnd`.
//!
//! A custom view is a named snapshot of workbook and sheet display settings.
//! The `UserBView` record in the workbook globals substream carries the
//! workbook-wide settings and the GUID that ties the view together; each
//! covered sheet substream holds a `UserSViewBegin` … `UserSViewEnd` bracket
//! (the `UserSViewBegin_Chart` layout on chart sheets) whose inner records
//! duplicate ordinary sheet settings for the view. The records are inert:
//! this module never applies a view, shows a window, or prints a sheet.

use super::environment::XlsObjectDisplayMode;
use super::records::XlsEncoding;
use super::sheet_metadata::XlsSheetVisibility;
use super::utils::parse_string_record;
use super::view::XlsPaneType;
use super::{XlsError, XlsResult};

/// Record type of the `UserBView` record.
pub(crate) const USER_B_VIEW_RECORD_TYPE: u16 = 0x01A9;
/// Record type of the `UserSViewBegin` and `UserSViewBegin_Chart` records.
pub(crate) const USER_S_VIEW_BEGIN_RECORD_TYPE: u16 = 0x01AA;
/// Record type of the `UserSViewEnd` record.
pub(crate) const USER_S_VIEW_END_RECORD_TYPE: u16 = 0x01AB;

/// Fixed part of a `UserBView` record, before the `st` name string.
const USER_B_VIEW_HEADER_LEN: usize = 50;
/// Payload length of a `UserSViewBegin` record (sheet layout).
const USER_S_VIEW_BEGIN_LEN: usize = 64;
/// Payload length of a `UserSViewBegin_Chart` record (chart-sheet layout).
const USER_S_VIEW_BEGIN_CHART_LEN: usize = 64;
/// Payload length of a `UserSViewEnd` record.
const USER_S_VIEW_END_LEN: usize = 2;

/// Maximum value of `UserBView.wTabRatio`.
const MAX_TAB_RATIO: u16 = 1000;
/// Zoom bounds of `UserSViewBegin(Chart).wScale`.
const MIN_SCALE: u32 = 10;
const MAX_SCALE: u32 = 400;
/// Maximum legal value of the `icvHdr` gridline color index.
const MAX_GRIDLINE_COLOR: u16 = 64;

// `UserBView` display-flag word bit assignments (fields A–N).
const DSP_FMLA_BAR: u16 = 1 << 0;
const DSP_STATUS: u16 = 1 << 1;
const NOTE_DISP_SHIFT: u16 = 2;
const NOTE_DISP_MASK: u16 = 0x3;
const DSP_HSCROLL: u16 = 1 << 4;
const DSP_VSCROLL: u16 = 1 << 5;
const BOT_ADORNMENT: u16 = 1 << 6;
const ZOOM: u16 = 1 << 7;
const HIDE_OBJ_SHIFT: u16 = 8;
const HIDE_OBJ_MASK: u16 = 0x3;
const PRINT_INCL: u16 = 1 << 10;
const ROW_COL_INCL: u16 = 1 << 11;
const INVALID_TAB_ID: u16 = 1 << 12;
const TIMED_UPDATE: u16 = 1 << 13;
const ALL_MEM_CHANGES: u16 = 1 << 14;
const ONLY_SYNC: u16 = 1 << 15;

// `UserBView` window-flag word bit assignments (fields O–P); bits 2–15 are
// undefined and preserved verbatim for round-trip fidelity.
const PERSONAL_VIEW: u16 = 1 << 0;
const ICONIC: u16 = 1 << 1;

// `UserSViewBegin` flag double-word bit assignments (fields A–b); the
// undefined bits (V, Z) are preserved verbatim.
const SHOW_BRKS: u32 = 1 << 0;
const DSP_FMLA_SV: u32 = 1 << 1;
const DSP_GRID_SV: u32 = 1 << 2;
const DSP_RW_COL_SV: u32 = 1 << 3;
const DSP_GUTS_SV: u32 = 1 << 4;
const DSP_ZEROS_SV: u32 = 1 << 5;
const HORIZONTAL: u32 = 1 << 6;
const VERTICAL: u32 = 1 << 7;
const PRINT_RW_COL: u32 = 1 << 8;
const PRINT_GRID: u32 = 1 << 9;
const FIT_TO_PAGE: u32 = 1 << 10;
const PRINT_AREA: u32 = 1 << 11;
const ONE_PRINT_AREA: u32 = 1 << 12;
const FILTER_MODE: u32 = 1 << 13;
const EZ_FILTER: u32 = 1 << 14;
const FROZEN: u32 = 1 << 15;
const FROZEN_NO_SPLIT: u32 = 1 << 16;
const SPLIT_V: u32 = 1 << 17;
const SPLIT_H: u32 = 1 << 18;
const HIDDEN_RW_SHIFT: u32 = 19;
const HIDDEN_RW_MASK: u32 = 0x3;
const HIDDEN_COL: u32 = 1 << 21;
const FILTER_UNIQUE: u32 = 1 << 25;
const SHEET_LAYOUT_VIEW: u32 = 1 << 26;
const PAGE_LAYOUT_VIEW: u32 = 1 << 27;
const RULER: u32 = 1 << 29;

// `UserSViewBegin_Chart` flag double-word bit assignments.
const HS_STATE_SHIFT: u32 = 22;
const HS_STATE_MASK: u32 = 0x3;
const ZOOM_TO_FIT: u32 = 1 << 30;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn read_f64(data: &[u8], offset: usize) -> f64 {
    f64::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ])
}

/// How cell comments appear in a custom view (`UserBView.mdNoteDisp`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsCustomViewNoteDisplay {
    /// Comment and visual cue are off for each cell with a comment.
    Off,
    /// Only the visual cue that indicates the cell has a comment.
    VisualCue,
    /// Comment and visual cue are on for each cell with a comment.
    On,
}

impl XlsCustomViewNoteDisplay {
    fn from_bits(bits: u16) -> XlsResult<Self> {
        match bits {
            0 => Ok(Self::Off),
            1 => Ok(Self::VisualCue),
            2 => Ok(Self::On),
            _ => Err(invalid(
                USER_B_VIEW_RECORD_TYPE,
                "UserBView.mdNoteDisp must be 0, 1, or 2",
            )),
        }
    }
}

/// Whether hidden rows are present in a custom view (`UserSViewBegin.fHiddenRw`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsCustomViewHiddenRows {
    /// At least one hidden row is present (filtered rows excluded).
    Present,
    /// No hidden row is present.
    NotPresent,
}

impl XlsCustomViewHiddenRows {
    fn from_bits(bits: u32) -> XlsResult<Self> {
        match bits {
            0 => Ok(Self::Present),
            1 => Ok(Self::NotPresent),
            _ => Err(invalid(
                USER_S_VIEW_BEGIN_RECORD_TYPE,
                "UserSViewBegin.fHiddenRw must be 0 or 1",
            )),
        }
    }
}

/// Workbook-wide settings of one custom view (`UserBView`, MS-XLS 2.4.333).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsWorkbookCustomView {
    guid: [u8; 16],
    active_tab: Option<u16>,
    window_x: i32,
    window_y: i32,
    window_width: i32,
    window_height: i32,
    tab_ratio: u16,
    /// Display-flag word (fields A–N); typed accessors decode the bits.
    display_flags: u16,
    /// Window-flag word (fields O–P plus the undefined high bits).
    window_flags: u16,
    merge_interval: u16,
    name: String,
}

impl XlsWorkbookCustomView {
    /// Parse a `UserBView` record payload.
    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < USER_B_VIEW_HEADER_LEN {
            return Err(XlsError::InvalidLength {
                expected: USER_B_VIEW_HEADER_LEN,
                found: data.len(),
            });
        }
        let tab_ratio = read_u16(data, 40);
        if tab_ratio > MAX_TAB_RATIO {
            return Err(invalid(
                USER_B_VIEW_RECORD_TYPE,
                "UserBView.wTabRatio must not exceed 1000",
            ));
        }
        let display_flags = read_u16(data, 42);
        // Reject out-of-range enumerations up front so accessors stay total.
        XlsCustomViewNoteDisplay::from_bits((display_flags >> NOTE_DISP_SHIFT) & NOTE_DISP_MASK)?;
        let object_display = (display_flags >> HIDE_OBJ_SHIFT) & HIDE_OBJ_MASK;
        if object_display == HIDE_OBJ_MASK {
            return Err(invalid(
                USER_B_VIEW_RECORD_TYPE,
                "UserBView.fHideObj must be 0, 1, or 2",
            ));
        }
        let active_tab = if display_flags & INVALID_TAB_ID != 0 {
            None
        } else {
            Some(read_u16(data, 4))
        };
        let mut guid = [0u8; 16];
        guid.copy_from_slice(&data[8..24]);
        let name = parse_string_record(&data[USER_B_VIEW_HEADER_LEN..], &XlsEncoding::Utf16Le)?;
        Ok(Self {
            guid,
            active_tab,
            window_x: read_i32(data, 24),
            window_y: read_i32(data, 28),
            window_width: read_i32(data, 32),
            window_height: read_i32(data, 36),
            tab_ratio,
            display_flags,
            window_flags: read_u16(data, 46),
            merge_interval: read_u16(data, 48),
            name,
        })
    }

    /// The GUID that ties this workbook view to its per-sheet views.
    pub const fn guid(&self) -> &[u8; 16] {
        &self.guid
    }
    /// The active sheet in this view; `None` when `fInvalidTabId` is set.
    pub const fn active_tab(&self) -> Option<u16> {
        self.active_tab
    }
    /// Workbook window position in pixels.
    pub const fn window_position(&self) -> (i32, i32) {
        (self.window_x, self.window_y)
    }
    /// Workbook window size in pixels.
    pub const fn window_size(&self) -> (i32, i32) {
        (self.window_width, self.window_height)
    }
    /// Ratio of sheet-tab area to horizontal-scroll-bar area (0–1000).
    pub const fn tab_ratio(&self) -> u16 {
        self.tab_ratio
    }
    /// Whether a formula bar is displayed.
    pub const fn shows_formula_bar(&self) -> bool {
        self.display_flags & DSP_FMLA_BAR != 0
    }
    /// Whether a status bar is displayed.
    pub const fn shows_status_bar(&self) -> bool {
        self.display_flags & DSP_STATUS != 0
    }
    /// How cell comments appear in this view.
    pub fn note_display(&self) -> XlsCustomViewNoteDisplay {
        XlsCustomViewNoteDisplay::from_bits(
            (self.display_flags >> NOTE_DISP_SHIFT) & NOTE_DISP_MASK,
        )
        .expect("validated at parse")
    }
    /// Whether a horizontal scroll bar is displayed.
    pub const fn shows_horizontal_scroll_bar(&self) -> bool {
        self.display_flags & DSP_HSCROLL != 0
    }
    /// Whether a vertical scroll bar is displayed.
    pub const fn shows_vertical_scroll_bar(&self) -> bool {
        self.display_flags & DSP_VSCROLL != 0
    }
    /// Whether sheet tabs are displayed.
    pub const fn shows_sheet_tabs(&self) -> bool {
        self.display_flags & BOT_ADORNMENT != 0
    }
    /// Whether the workbook window is maximized.
    pub const fn is_maximized(&self) -> bool {
        self.display_flags & ZOOM != 0
    }
    /// How drawing and OLE objects appear in the workbook window.
    pub const fn object_display(&self) -> XlsObjectDisplayMode {
        match (self.display_flags >> HIDE_OBJ_SHIFT) & HIDE_OBJ_MASK {
            1 => XlsObjectDisplayMode::ShowPlaceholders,
            2 => XlsObjectDisplayMode::HideAll,
            _ => XlsObjectDisplayMode::ShowAll,
        }
    }
    /// Whether the view includes the workbook print settings.
    pub const fn includes_print_settings(&self) -> bool {
        self.display_flags & PRINT_INCL != 0
    }
    /// Whether the view includes hidden rows, hidden columns, and filters.
    pub const fn includes_hidden_rows_columns_and_filters(&self) -> bool {
        self.display_flags & ROW_COL_INCL != 0
    }
    /// Whether updates of linked or external data are coordinated.
    pub const fn timed_update(&self) -> bool {
        self.display_flags & TIMED_UPDATE != 0
    }
    /// Whether the changes being saved have priority in a merge conflict.
    pub const fn all_memory_changes(&self) -> bool {
        self.display_flags & ALL_MEM_CHANGES != 0
    }
    /// Whether the automatic update only merges, or merges and also saves.
    pub const fn only_sync(&self) -> bool {
        self.display_flags & ONLY_SYNC != 0
    }
    /// Whether this is the personal view of a shared workbook.
    pub const fn is_personal_view(&self) -> bool {
        self.window_flags & PERSONAL_VIEW != 0
    }
    /// Whether the workbook window is minimized.
    pub const fn is_minimized(&self) -> bool {
        self.window_flags & ICONIC != 0
    }
    /// Minutes between automatic merges of a shared workbook.
    pub const fn merge_interval(&self) -> u16 {
        self.merge_interval
    }
    /// The name of the custom view.
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// The `Ref8U` visible area of the logical top-left pane in a sheet view.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsCustomViewTopLeft {
    first_row: u16,
    last_row: u16,
    first_col: u16,
    last_col: u16,
}

impl XlsCustomViewTopLeft {
    pub const fn first_row(&self) -> u16 {
        self.first_row
    }
    pub const fn last_row(&self) -> u16 {
        self.last_row
    }
    pub const fn first_col(&self) -> u16 {
        self.first_col
    }
    pub const fn last_col(&self) -> u16 {
        self.last_col
    }
}

/// Sheet settings of one custom view (`UserSViewBegin`, MS-XLS 2.4.334).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XlsSheetCustomViewBegin {
    guid: [u8; 16],
    tab_id: u16,
    scale: u32,
    gridline_color: u16,
    active_pane: XlsPaneType,
    /// Flag double-word (fields A–b); typed accessors decode the bits.
    flags: u32,
    top_left: XlsCustomViewTopLeft,
    split_x: f64,
    split_y: f64,
    right_pane_col: u16,
    bottom_pane_row: u16,
}

impl XlsSheetCustomViewBegin {
    /// Parse a `UserSViewBegin` record payload (worksheet/dialog/macro
    /// layout; chart sheets use [`XlsChartSheetCustomViewBegin`]).
    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != USER_S_VIEW_BEGIN_LEN {
            return Err(XlsError::InvalidLength {
                expected: USER_S_VIEW_BEGIN_LEN,
                found: data.len(),
            });
        }
        let scale = read_u32(data, 20);
        if !(MIN_SCALE..=MAX_SCALE).contains(&scale) {
            return Err(invalid(
                USER_S_VIEW_BEGIN_RECORD_TYPE,
                "UserSViewBegin.wScale must be between 10 and 400",
            ));
        }
        let gridline_color = read_u16(data, 24);
        if gridline_color > MAX_GRIDLINE_COLOR {
            return Err(invalid(
                USER_S_VIEW_BEGIN_RECORD_TYPE,
                "UserSViewBegin.icvHdr must not exceed 64",
            ));
        }
        let active_pane = match data[28] {
            0 => XlsPaneType::LowerRight,
            1 => XlsPaneType::UpperRight,
            2 => XlsPaneType::LowerLeft,
            3 => XlsPaneType::UpperLeft,
            value => {
                return Err(invalid(
                    USER_S_VIEW_BEGIN_RECORD_TYPE,
                    format!("UserSViewBegin.pnnSel has an invalid pane type {value}"),
                ));
            },
        };
        let flags = read_u32(data, 32);
        XlsCustomViewHiddenRows::from_bits((flags >> HIDDEN_RW_SHIFT) & HIDDEN_RW_MASK)?;
        let mut guid = [0u8; 16];
        guid.copy_from_slice(&data[0..16]);
        Ok(Self {
            guid,
            tab_id: read_u16(data, 16),
            scale,
            gridline_color,
            active_pane,
            flags,
            top_left: XlsCustomViewTopLeft {
                first_row: read_u16(data, 36),
                last_row: read_u16(data, 38),
                first_col: read_u16(data, 40),
                last_col: read_u16(data, 42),
            },
            split_x: read_f64(data, 44),
            split_y: read_f64(data, 52),
            right_pane_col: read_u16(data, 60),
            bottom_pane_row: read_u16(data, 62),
        })
    }

    /// The GUID of the associated `UserBView` record.
    pub const fn guid(&self) -> &[u8; 16] {
        &self.guid
    }
    /// The sheet this custom view belongs to.
    pub const fn tab_id(&self) -> u16 {
        self.tab_id
    }
    /// Zoom level of the window used to display the sheet (10–400).
    pub const fn scale(&self) -> u32 {
        self.scale
    }
    /// Color index of the gridlines displayed in the view.
    pub const fn gridline_color(&self) -> u16 {
        self.gridline_color
    }
    /// The active pane.
    pub const fn active_pane(&self) -> XlsPaneType {
        self.active_pane
    }
    /// Whether page breaks are displayed.
    pub const fn shows_page_breaks(&self) -> bool {
        self.flags & SHOW_BRKS != 0
    }
    /// Whether the window displays formulas instead of values.
    pub const fn shows_formulas(&self) -> bool {
        self.flags & DSP_FMLA_SV != 0
    }
    /// Whether gridlines are displayed.
    pub const fn shows_gridlines(&self) -> bool {
        self.flags & DSP_GRID_SV != 0
    }
    /// Whether row and column headings are displayed.
    pub const fn shows_headings(&self) -> bool {
        self.flags & DSP_RW_COL_SV != 0
    }
    /// Whether outline symbols are displayed.
    pub const fn shows_outline_symbols(&self) -> bool {
        self.flags & DSP_GUTS_SV != 0
    }
    /// Whether zero values are suppressed.
    pub const fn suppresses_zeros(&self) -> bool {
        self.flags & DSP_ZEROS_SV != 0
    }
    /// Whether the sheet is centered between the horizontal margins when printed.
    pub const fn print_centered_horizontally(&self) -> bool {
        self.flags & HORIZONTAL != 0
    }
    /// Whether the sheet is centered between the vertical margins when printed.
    pub const fn print_centered_vertically(&self) -> bool {
        self.flags & VERTICAL != 0
    }
    /// Whether row and column headings are printed.
    pub const fn prints_headings(&self) -> bool {
        self.flags & PRINT_RW_COL != 0
    }
    /// Whether gridlines are printed.
    pub const fn prints_gridlines(&self) -> bool {
        self.flags & PRINT_GRID != 0
    }
    /// Whether the fit-to-page option is enabled.
    pub const fn fits_to_page(&self) -> bool {
        self.flags & FIT_TO_PAGE != 0
    }
    /// Whether the sheet has at least one print area.
    pub const fn has_print_area(&self) -> bool {
        self.flags & PRINT_AREA != 0
    }
    /// Whether the sheet has exactly one print area.
    pub const fn has_single_print_area(&self) -> bool {
        self.flags & ONE_PRINT_AREA != 0
    }
    /// Whether cells are hidden because of filtering.
    pub const fn is_filter_mode(&self) -> bool {
        self.flags & FILTER_MODE != 0
    }
    /// Whether the AutoFilter icon is shown on the sheet.
    pub const fn shows_autofilter_icon(&self) -> bool {
        self.flags & EZ_FILTER != 0
    }
    /// Whether the panes are frozen.
    pub const fn is_frozen(&self) -> bool {
        self.flags & FROZEN != 0
    }
    /// Whether the panes are frozen but not split.
    pub const fn is_frozen_without_split(&self) -> bool {
        self.flags & FROZEN_NO_SPLIT != 0
    }
    /// Whether the window is split vertically.
    pub const fn is_split_vertically(&self) -> bool {
        self.flags & SPLIT_V != 0
    }
    /// Whether the window is split horizontally.
    pub const fn is_split_horizontally(&self) -> bool {
        self.flags & SPLIT_H != 0
    }
    /// Whether hidden rows (filtered rows excluded) are present.
    pub fn hidden_rows(&self) -> XlsCustomViewHiddenRows {
        XlsCustomViewHiddenRows::from_bits((self.flags >> HIDDEN_RW_SHIFT) & HIDDEN_RW_MASK)
            .expect("validated at parse")
    }
    /// Whether at least one hidden column is present.
    pub const fn has_hidden_columns(&self) -> bool {
        self.flags & HIDDEN_COL != 0
    }
    /// Whether the advanced filter shows only unique rows.
    pub const fn filters_unique_rows(&self) -> bool {
        self.flags & FILTER_UNIQUE != 0
    }
    /// Whether the sheet is in Page Break Preview view.
    pub const fn is_page_break_preview(&self) -> bool {
        self.flags & SHEET_LAYOUT_VIEW != 0
    }
    /// Whether the sheet is in Page Layout view.
    pub const fn is_page_layout_view(&self) -> bool {
        self.flags & PAGE_LAYOUT_VIEW != 0
    }
    /// Whether the ruler is displayed.
    pub const fn shows_ruler(&self) -> bool {
        self.flags & RULER != 0
    }
    /// The visible area of the logical top-left pane.
    pub const fn top_left(&self) -> XlsCustomViewTopLeft {
        self.top_left
    }
    /// Left-to-right position of the split, expressed as a column number.
    pub const fn split_x(&self) -> f64 {
        self.split_x
    }
    /// Top-to-bottom position of the split, expressed as a row number.
    pub const fn split_y(&self) -> f64 {
        self.split_y
    }
    /// First visible column of the logical right pane (65535 when split).
    pub const fn right_pane_col(&self) -> u16 {
        self.right_pane_col
    }
    /// First visible row of the bottom pane (65535 when split).
    pub const fn bottom_pane_row(&self) -> u16 {
        self.bottom_pane_row
    }
}

/// Chart-sheet settings of one custom view (`UserSViewBegin_Chart`,
/// MS-XLS 2.4.335). Shares the `UserSViewBegin` record type but uses a
/// chart-sheet-specific layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsChartSheetCustomViewBegin {
    guid: [u8; 16],
    tab_id: u32,
    scale: u32,
    visibility: XlsSheetVisibility,
    zoom_to_fit: bool,
}

impl XlsChartSheetCustomViewBegin {
    /// Parse a `UserSViewBegin_Chart` record payload.
    ///
    /// This is public because the workbook reader never walks chart-sheet
    /// substreams; callers that walk one themselves use this for the
    /// chart-specific layout of the shared `UserSViewBegin` record type.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != USER_S_VIEW_BEGIN_CHART_LEN {
            return Err(XlsError::InvalidLength {
                expected: USER_S_VIEW_BEGIN_CHART_LEN,
                found: data.len(),
            });
        }
        let scale = read_u32(data, 20);
        if !(MIN_SCALE..=MAX_SCALE).contains(&scale) {
            return Err(invalid(
                USER_S_VIEW_BEGIN_RECORD_TYPE,
                "UserSViewBegin_Chart.wScale must be between 10 and 400",
            ));
        }
        let flags = read_u32(data, 32);
        let visibility = match (flags >> HS_STATE_SHIFT) & HS_STATE_MASK {
            0 => XlsSheetVisibility::Visible,
            1 => XlsSheetVisibility::Hidden,
            2 => XlsSheetVisibility::VeryHidden,
            _ => {
                return Err(invalid(
                    USER_S_VIEW_BEGIN_RECORD_TYPE,
                    "UserSViewBegin_Chart.hsState must be 0, 1, or 2",
                ));
            },
        };
        let mut guid = [0u8; 16];
        guid.copy_from_slice(&data[0..16]);
        Ok(Self {
            guid,
            tab_id: read_u32(data, 16),
            scale,
            visibility,
            zoom_to_fit: flags & ZOOM_TO_FIT != 0,
        })
    }

    /// The GUID of the associated `UserBView` record.
    pub const fn guid(&self) -> &[u8; 16] {
        &self.guid
    }
    /// The sheet this custom view belongs to.
    pub const fn tab_id(&self) -> u32 {
        self.tab_id
    }
    /// Zoom level of the window used to display the sheet (10–400).
    pub const fn scale(&self) -> u32 {
        self.scale
    }
    /// Hidden state of the chart sheet in this view.
    pub const fn visibility(&self) -> XlsSheetVisibility {
        self.visibility
    }
    /// Whether the zoom is set to "Zoom to Fit Selection".
    pub const fn zoom_to_fit(&self) -> bool {
        self.zoom_to_fit
    }
}

/// The end of a per-sheet custom-view record collection (`UserSViewEnd`,
/// MS-XLS 2.4.336). Its 2-byte payload is defined as 1 and ignored.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsSheetCustomViewEnd {
    reserved: u16,
}

impl XlsSheetCustomViewEnd {
    /// Parse a `UserSViewEnd` record payload.
    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != USER_S_VIEW_END_LEN {
            return Err(XlsError::InvalidLength {
                expected: USER_S_VIEW_END_LEN,
                found: data.len(),
            });
        }
        Ok(Self {
            reserved: read_u16(data, 0),
        })
    }

    /// The raw reserved value (1 in conforming files).
    pub const fn reserved(&self) -> u16 {
        self.reserved
    }
}

/// One sheet's custom view: the bracket of records starting at
/// `UserSViewBegin` (or `UserSViewBegin_Chart`) and ending at `UserSViewEnd`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XlsSheetCustomView {
    begin: XlsSheetCustomViewBegin,
    end: XlsSheetCustomViewEnd,
}

impl XlsSheetCustomView {
    pub(crate) fn new(begin: XlsSheetCustomViewBegin, end: XlsSheetCustomViewEnd) -> Self {
        Self { begin, end }
    }

    /// The sheet-view settings that opened the custom-view bracket.
    pub const fn begin(&self) -> &XlsSheetCustomViewBegin {
        &self.begin
    }
    /// The record that closed the custom-view bracket.
    pub const fn end(&self) -> XlsSheetCustomViewEnd {
        self.end
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn push_string(data: &mut Vec<u8>, text: &str) {
        data.extend_from_slice(&(text.len() as u16).to_le_bytes());
        data.push(0); // compressed Latin-1
        data.extend_from_slice(text.as_bytes());
    }

    fn user_b_view_payload() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&[0xEE; 4]); // unused1
        data.extend_from_slice(&2u16.to_le_bytes()); // tabId
        data.extend_from_slice(&0u16.to_le_bytes()); // reserved1
        data.extend_from_slice(&[0xAB; 16]); // guid
        data.extend_from_slice(&(-10i32).to_le_bytes()); // x
        data.extend_from_slice(&20i32.to_le_bytes()); // y
        data.extend_from_slice(&800i32.to_le_bytes()); // dx
        data.extend_from_slice(&600i32.to_le_bytes()); // dy
        data.extend_from_slice(&600u16.to_le_bytes()); // wTabRatio
        let flags = DSP_FMLA_BAR
            | DSP_STATUS
            | (2 << NOTE_DISP_SHIFT)
            | DSP_HSCROLL
            | BOT_ADORNMENT
            | (2 << HIDE_OBJ_SHIFT)
            | PRINT_INCL
            | ROW_COL_INCL;
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&[0xDD; 2]); // unused2
        data.extend_from_slice(&(PERSONAL_VIEW | 0x8000).to_le_bytes());
        data.extend_from_slice(&30u16.to_le_bytes()); // wMergeInterval
        push_string(&mut data, "Quarterly");
        data
    }

    #[test]
    fn parses_user_b_view() {
        let view = XlsWorkbookCustomView::parse(&user_b_view_payload()).unwrap();
        assert_eq!(view.guid(), &[0xAB; 16]);
        assert_eq!(view.active_tab(), Some(2));
        assert_eq!(view.window_position(), (-10, 20));
        assert_eq!(view.window_size(), (800, 600));
        assert_eq!(view.tab_ratio(), 600);
        assert!(view.shows_formula_bar());
        assert!(view.shows_status_bar());
        assert_eq!(view.note_display(), XlsCustomViewNoteDisplay::On);
        assert!(view.shows_horizontal_scroll_bar());
        assert!(!view.shows_vertical_scroll_bar());
        assert!(view.shows_sheet_tabs());
        assert!(!view.is_maximized());
        assert_eq!(view.object_display(), XlsObjectDisplayMode::HideAll);
        assert!(view.includes_print_settings());
        assert!(view.includes_hidden_rows_columns_and_filters());
        assert!(!view.timed_update());
        assert!(view.is_personal_view());
        assert!(!view.is_minimized());
        assert_eq!(view.merge_interval(), 30);
        assert_eq!(view.name(), "Quarterly");
    }

    #[test]
    fn user_b_view_honors_invalid_tab_id() {
        let mut payload = user_b_view_payload();
        payload[42..44].copy_from_slice(&(DSP_STATUS | INVALID_TAB_ID).to_le_bytes());
        let view = XlsWorkbookCustomView::parse(&payload).unwrap();
        assert_eq!(view.active_tab(), None);
        assert!(!view.shows_formula_bar());
    }

    #[test]
    fn rejects_bad_user_b_view() {
        // Truncated header.
        assert!(XlsWorkbookCustomView::parse(&[0; 49]).is_err());
        // Tab ratio above the legal maximum.
        let mut payload = user_b_view_payload();
        payload[40..42].copy_from_slice(&1001u16.to_le_bytes());
        assert!(XlsWorkbookCustomView::parse(&payload).is_err());
        // Out-of-range note-display enumeration.
        let mut payload = user_b_view_payload();
        payload[42..44].copy_from_slice(&(3u16 << NOTE_DISP_SHIFT).to_le_bytes());
        assert!(XlsWorkbookCustomView::parse(&payload).is_err());
        // Out-of-range object-display enumeration.
        let mut payload = user_b_view_payload();
        payload[42..44].copy_from_slice(&(3u16 << HIDE_OBJ_SHIFT).to_le_bytes());
        assert!(XlsWorkbookCustomView::parse(&payload).is_err());
        // Truncated name string.
        let mut payload = user_b_view_payload();
        payload.truncate(USER_B_VIEW_HEADER_LEN + 2);
        assert!(XlsWorkbookCustomView::parse(&payload).is_err());
    }

    fn user_s_view_begin_payload() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&[0xCD; 16]); // guid
        data.extend_from_slice(&1u16.to_le_bytes()); // iTabid
        data.extend_from_slice(&0u16.to_le_bytes()); // reserved1
        data.extend_from_slice(&75u32.to_le_bytes()); // wScale
        data.extend_from_slice(&64u16.to_le_bytes()); // icvHdr
        data.extend_from_slice(&0u16.to_le_bytes()); // reserved2
        data.push(3); // pnnSel = UpperLeft
        data.extend_from_slice(&[0; 2]); // reserved3
        data.push(0); // reserved4
        let flags = DSP_GRID_SV
            | DSP_RW_COL_SV
            | FROZEN
            | SPLIT_V
            | (1 << HIDDEN_RW_SHIFT)
            | HIDDEN_COL
            | PAGE_LAYOUT_VIEW;
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&5u16.to_le_bytes()); // rwFirst
        data.extend_from_slice(&30u16.to_le_bytes()); // rwLast
        data.extend_from_slice(&2u16.to_le_bytes()); // colFirst
        data.extend_from_slice(&9u16.to_le_bytes()); // colLast
        data.extend_from_slice(&3.5f64.to_le_bytes()); // operNumX
        data.extend_from_slice(&12.0f64.to_le_bytes()); // operNumY
        data.extend_from_slice(&u16::MAX.to_le_bytes()); // colRPane
        data.extend_from_slice(&40u16.to_le_bytes()); // rwBPane
        data
    }

    #[test]
    fn parses_user_s_view_begin() {
        let begin = XlsSheetCustomViewBegin::parse(&user_s_view_begin_payload()).unwrap();
        assert_eq!(begin.guid(), &[0xCD; 16]);
        assert_eq!(begin.tab_id(), 1);
        assert_eq!(begin.scale(), 75);
        assert_eq!(begin.gridline_color(), 64);
        assert_eq!(begin.active_pane(), XlsPaneType::UpperLeft);
        assert!(!begin.shows_page_breaks());
        assert!(!begin.shows_formulas());
        assert!(begin.shows_gridlines());
        assert!(begin.shows_headings());
        assert!(!begin.suppresses_zeros());
        assert!(begin.is_frozen());
        assert!(!begin.is_frozen_without_split());
        assert!(begin.is_split_vertically());
        assert!(!begin.is_split_horizontally());
        assert_eq!(begin.hidden_rows(), XlsCustomViewHiddenRows::NotPresent);
        assert!(begin.has_hidden_columns());
        assert!(begin.is_page_layout_view());
        assert!(!begin.is_page_break_preview());
        let top_left = begin.top_left();
        assert_eq!(top_left.first_row(), 5);
        assert_eq!(top_left.last_row(), 30);
        assert_eq!(top_left.first_col(), 2);
        assert_eq!(top_left.last_col(), 9);
        assert_eq!(begin.split_x(), 3.5);
        assert_eq!(begin.split_y(), 12.0);
        assert_eq!(begin.right_pane_col(), u16::MAX);
        assert_eq!(begin.bottom_pane_row(), 40);
    }

    #[test]
    fn rejects_bad_user_s_view_begin() {
        // Wrong length.
        assert!(XlsSheetCustomViewBegin::parse(&[0; 63]).is_err());
        // Zoom out of range.
        let mut payload = user_s_view_begin_payload();
        payload[20..24].copy_from_slice(&9u32.to_le_bytes());
        assert!(XlsSheetCustomViewBegin::parse(&payload).is_err());
        // Gridline color out of range.
        let mut payload = user_s_view_begin_payload();
        payload[24..26].copy_from_slice(&65u16.to_le_bytes());
        assert!(XlsSheetCustomViewBegin::parse(&payload).is_err());
        // Invalid pane type.
        let mut payload = user_s_view_begin_payload();
        payload[28] = 4;
        assert!(XlsSheetCustomViewBegin::parse(&payload).is_err());
        // Out-of-range hidden-row state.
        let mut payload = user_s_view_begin_payload();
        let flags = read_u32(&payload, 32) | (2 << HIDDEN_RW_SHIFT);
        payload[32..36].copy_from_slice(&flags.to_le_bytes());
        assert!(XlsSheetCustomViewBegin::parse(&payload).is_err());
    }

    #[test]
    fn parses_user_s_view_begin_chart() {
        let mut data = Vec::new();
        data.extend_from_slice(&[0xEF; 16]); // guid
        data.extend_from_slice(&3u32.to_le_bytes()); // iTabid (4 bytes)
        data.extend_from_slice(&100u32.to_le_bytes()); // wScale
        data.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        data.extend_from_slice(&[0x11; 4]); // unused1
        let flags = (2 << HS_STATE_SHIFT) | ZOOM_TO_FIT;
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&[0x22; 8]); // unused2
        data.extend_from_slice(&[0x33; 8]); // unused3
        data.extend_from_slice(&[0x44; 8]); // unused4
        data.extend_from_slice(&[0x55; 2]); // unused5
        data.extend_from_slice(&[0x66; 2]); // unused6
        assert_eq!(data.len(), USER_S_VIEW_BEGIN_CHART_LEN);
        let begin = XlsChartSheetCustomViewBegin::parse(&data).unwrap();
        assert_eq!(begin.guid(), &[0xEF; 16]);
        assert_eq!(begin.tab_id(), 3);
        assert_eq!(begin.scale(), 100);
        assert_eq!(begin.visibility(), XlsSheetVisibility::VeryHidden);
        assert!(begin.zoom_to_fit());

        // Zoom bounds and hidden-state enumeration are validated.
        data[20..24].copy_from_slice(&401u32.to_le_bytes());
        assert!(XlsChartSheetCustomViewBegin::parse(&data).is_err());
        data[20..24].copy_from_slice(&100u32.to_le_bytes());
        data[32..36].copy_from_slice(&(3u32 << HS_STATE_SHIFT).to_le_bytes());
        assert!(XlsChartSheetCustomViewBegin::parse(&data).is_err());
        assert!(XlsChartSheetCustomViewBegin::parse(&data[..63]).is_err());
    }

    #[test]
    fn parses_user_s_view_end() {
        let end = XlsSheetCustomViewEnd::parse(&1u16.to_le_bytes()).unwrap();
        assert_eq!(end.reserved(), 1);
        assert!(XlsSheetCustomViewEnd::parse(&[0; 1]).is_err());
        assert!(XlsSheetCustomViewEnd::parse(&[0; 3]).is_err());
    }
}
