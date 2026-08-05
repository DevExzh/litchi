//! BIFF8 worksheet default dimensions and outline workspace metadata.

use crate::error::{XlsError, XlsResult};

pub(crate) const GUTS_RECORD_TYPE: u16 = 0x0080;
pub(crate) const WSBOOL_RECORD_TYPE: u16 = 0x0081;
pub(crate) const DEFAULT_ROW_HEIGHT_RECORD_TYPE: u16 = 0x0225;
pub(crate) const DEF_COL_WIDTH_RECORD_TYPE: u16 = 0x0055;
const DIMENSIONS_RECORD_TYPE: u16 = 0x0200;

/// Default row/column dimensions and outline workspace state for a worksheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Layout {
    default_row_height_twips: u16,
    empty_rows_hidden: bool,
    default_row_height_unsynced: bool,
    thick_top_border: bool,
    thick_bottom_border: bool,
    default_column_width_chars: u16,
    max_row_outline_level: u8,
    max_column_outline_level: u8,
    row_gutter_width: u16,
    column_gutter_height: u16,
    show_automatic_page_breaks: bool,
    apply_styles_to_outlines: bool,
    summary_rows_below: bool,
    summary_columns_right: bool,
    fit_to_page: bool,
    synchronize_horizontal_scrolling: bool,
    synchronize_vertical_scrolling: bool,
    alternate_expression_evaluation: bool,
    alternate_formula_entry: bool,
}

impl Default for Layout {
    fn default() -> Self {
        Self {
            default_row_height_twips: 255,
            empty_rows_hidden: false,
            default_row_height_unsynced: false,
            thick_top_border: false,
            thick_bottom_border: false,
            default_column_width_chars: 8,
            max_row_outline_level: 0,
            max_column_outline_level: 0,
            row_gutter_width: 0,
            column_gutter_height: 0,
            show_automatic_page_breaks: true,
            apply_styles_to_outlines: false,
            summary_rows_below: true,
            summary_columns_right: true,
            fit_to_page: false,
            synchronize_horizontal_scrolling: false,
            synchronize_vertical_scrolling: false,
            alternate_expression_evaluation: false,
            alternate_formula_entry: false,
        }
    }
}

impl Layout {
    pub fn default_row_height_twips(&self) -> u16 {
        self.default_row_height_twips
    }
    pub fn empty_rows_hidden(&self) -> bool {
        self.empty_rows_hidden
    }
    pub fn default_row_height_unsynced(&self) -> bool {
        self.default_row_height_unsynced
    }
    pub fn thick_top_border(&self) -> bool {
        self.thick_top_border
    }
    pub fn thick_bottom_border(&self) -> bool {
        self.thick_bottom_border
    }
    pub fn default_column_width_chars(&self) -> u16 {
        self.default_column_width_chars
    }
    pub fn max_row_outline_level(&self) -> u8 {
        self.max_row_outline_level
    }
    pub fn max_column_outline_level(&self) -> u8 {
        self.max_column_outline_level
    }
    pub fn row_gutter_width(&self) -> u16 {
        self.row_gutter_width
    }
    pub fn column_gutter_height(&self) -> u16 {
        self.column_gutter_height
    }
    pub fn show_automatic_page_breaks(&self) -> bool {
        self.show_automatic_page_breaks
    }
    pub fn apply_styles_to_outlines(&self) -> bool {
        self.apply_styles_to_outlines
    }
    pub fn summary_rows_below(&self) -> bool {
        self.summary_rows_below
    }
    pub fn summary_columns_right(&self) -> bool {
        self.summary_columns_right
    }
    pub fn fit_to_page(&self) -> bool {
        self.fit_to_page
    }
    pub fn synchronize_horizontal_scrolling(&self) -> bool {
        self.synchronize_horizontal_scrolling
    }
    pub fn synchronize_vertical_scrolling(&self) -> bool {
        self.synchronize_vertical_scrolling
    }
    pub fn alternate_expression_evaluation(&self) -> bool {
        self.alternate_expression_evaluation
    }
    pub fn alternate_formula_entry(&self) -> bool {
        self.alternate_formula_entry
    }
}

pub(crate) struct Collector {
    layout: Layout,
    seen: [bool; 4],
    last_rank: Option<u8>,
    dimensions_seen: bool,
}

impl Collector {
    pub(crate) fn new() -> Self {
        Self {
            layout: Layout::default(),
            seen: [false; 4],
            last_rank: None,
            dimensions_seen: false,
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if record_type == DIMENSIONS_RECORD_TYPE {
            self.dimensions_seen = true;
            return Ok(());
        }
        let rank = match record_type {
            GUTS_RECORD_TYPE => 0,
            DEFAULT_ROW_HEIGHT_RECORD_TYPE => 1,
            WSBOOL_RECORD_TYPE => 2,
            DEF_COL_WIDTH_RECORD_TYPE => 3,
            _ => return Ok(()),
        };
        if self.seen[rank as usize] {
            return invalid(record_type, "duplicate worksheet layout record");
        }
        if self.last_rank.is_some_and(|previous| rank < previous) {
            return invalid(record_type, "worksheet layout record is out of BIFF8 order");
        }
        if record_type == DEF_COL_WIDTH_RECORD_TYPE && self.dimensions_seen {
            return invalid(record_type, "DefColWidth must precede Dimensions");
        }
        self.seen[rank as usize] = true;
        self.last_rank = Some(rank);

        match record_type {
            GUTS_RECORD_TYPE => self.parse_guts(data)?,
            DEFAULT_ROW_HEIGHT_RECORD_TYPE => self.parse_default_row_height(data)?,
            WSBOOL_RECORD_TYPE => self.parse_wsbool(data)?,
            DEF_COL_WIDTH_RECORD_TYPE => self.parse_def_col_width(data)?,
            _ => unreachable!(),
        }
        Ok(())
    }

    fn parse_guts(&mut self, data: &[u8]) -> XlsResult<()> {
        require_length(GUTS_RECORD_TYPE, data, 8)?;
        self.layout.row_gutter_width = read_u16(data, 0);
        self.layout.column_gutter_height = read_u16(data, 2);
        self.layout.max_row_outline_level = decode_outline_level(read_u16(data, 4))?;
        self.layout.max_column_outline_level = decode_outline_level(read_u16(data, 6))?;
        Ok(())
    }

    fn parse_default_row_height(&mut self, data: &[u8]) -> XlsResult<()> {
        require_length(DEFAULT_ROW_HEIGHT_RECORD_TYPE, data, 4)?;
        let flags = read_u16(data, 0);
        if flags & !0x000f != 0 {
            return invalid(
                DEFAULT_ROW_HEIGHT_RECORD_TYPE,
                "reserved flags must be zero",
            );
        }
        let height = read_u16(data, 2);
        let hidden = flags & 0x0002 != 0;
        if height > 8179 || (!hidden && height == 0) {
            return invalid(
                DEFAULT_ROW_HEIGHT_RECORD_TYPE,
                "row height is outside the BIFF8 range",
            );
        }
        self.layout.default_row_height_twips = height;
        self.layout.default_row_height_unsynced = flags & 0x0001 != 0;
        self.layout.empty_rows_hidden = hidden;
        self.layout.thick_top_border = flags & 0x0004 != 0;
        self.layout.thick_bottom_border = flags & 0x0008 != 0;
        Ok(())
    }

    fn parse_wsbool(&mut self, data: &[u8]) -> XlsResult<()> {
        require_length(WSBOOL_RECORD_TYPE, data, 2)?;
        let flags = read_u16(data, 0);
        if flags & 0x020e != 0 {
            return invalid(WSBOOL_RECORD_TYPE, "reserved flags must be zero");
        }
        if flags & 0x0010 != 0 {
            return invalid(
                WSBOOL_RECORD_TYPE,
                "dialog-sheet flag is invalid in a worksheet substream",
            );
        }
        self.layout.show_automatic_page_breaks = flags & 0x0001 != 0;
        self.layout.apply_styles_to_outlines = flags & 0x0020 != 0;
        self.layout.summary_rows_below = flags & 0x0040 != 0;
        self.layout.summary_columns_right = flags & 0x0080 != 0;
        self.layout.fit_to_page = flags & 0x0100 != 0;
        self.layout.synchronize_horizontal_scrolling = flags & 0x1000 != 0;
        self.layout.synchronize_vertical_scrolling = flags & 0x2000 != 0;
        self.layout.alternate_expression_evaluation = flags & 0x4000 != 0;
        self.layout.alternate_formula_entry = flags & 0x8000 != 0;
        Ok(())
    }

    fn parse_def_col_width(&mut self, data: &[u8]) -> XlsResult<()> {
        require_length(DEF_COL_WIDTH_RECORD_TYPE, data, 2)?;
        let width = read_u16(data, 0);
        if width > 255 {
            return invalid(
                DEF_COL_WIDTH_RECORD_TYPE,
                "default column width exceeds 255 characters",
            );
        }
        self.layout.default_column_width_chars = width;
        Ok(())
    }

    pub(crate) fn finish(self) -> Layout {
        self.layout
    }
}

fn decode_outline_level(encoded: u16) -> XlsResult<u8> {
    match encoded {
        0 => Ok(0),
        2..=8 => Ok((encoded - 1) as u8),
        _ => invalid(
            GUTS_RECORD_TYPE,
            "outline level encoding must be 0 or 2..=8",
        ),
    }
}

fn require_length(record_type: u16, data: &[u8], expected: usize) -> XlsResult<()> {
    if data.len() != expected {
        return Err(XlsError::InvalidLength {
            expected,
            found: data.len(),
        });
    }
    let _ = record_type;
    Ok(())
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn invalid<T>(record_type: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_malformed_layout_records() {
        let mut collector = Collector::new();
        assert!(
            collector
                .feed_record(DEFAULT_ROW_HEIGHT_RECORD_TYPE, &[0, 0, 1])
                .is_err()
        );

        let mut collector = Collector::new();
        assert!(
            collector
                .feed_record(GUTS_RECORD_TYPE, &[0, 0, 0, 0, 1, 0, 0, 0])
                .is_err()
        );

        let mut collector = Collector::new();
        assert!(
            collector
                .feed_record(WSBOOL_RECORD_TYPE, &0x0002u16.to_le_bytes())
                .is_err()
        );

        let mut collector = Collector::new();
        assert!(
            collector
                .feed_record(DEF_COL_WIDTH_RECORD_TYPE, &256u16.to_le_bytes())
                .is_err()
        );
    }

    #[test]
    fn rejects_duplicates_and_out_of_order_records() {
        let mut collector = Collector::new();
        collector
            .feed_record(WSBOOL_RECORD_TYPE, &0x00c1u16.to_le_bytes())
            .unwrap();
        assert!(collector.feed_record(GUTS_RECORD_TYPE, &[0; 8]).is_err());

        let mut collector = Collector::new();
        collector.feed_record(GUTS_RECORD_TYPE, &[0; 8]).unwrap();
        assert!(collector.feed_record(GUTS_RECORD_TYPE, &[0; 8]).is_err());
    }
}
