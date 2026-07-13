//! Worksheet implementation for XLSB files

use crate::xlsb::cell::XlsbCell;
use crate::xlsb::comments::Comment;
use crate::xlsb::hyperlinks::Hyperlink;
use crate::xlsb::merged_cells::MergedCell;
use litchi_core::sheet::{
    Cell as SheetCell, CellIterator, CellValue, Result, RowIterator, Worksheet,
};
use std::borrow::Cow;
use std::collections::BTreeMap;

/// Width, style, visibility, and outline metadata for an XLSB column range.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsbColumnInfo {
    /// First zero-based column covered by this record.
    pub first_column: u32,
    /// Last zero-based column covered by this record, inclusive.
    pub last_column: u32,
    /// Width in standard-digit character units.
    pub width: f64,
    /// Zero-based default cell-XF index for the covered columns.
    pub style_id: u32,
    /// Whether the columns are hidden.
    pub hidden: bool,
    /// Whether the width differs from the worksheet default.
    pub user_set_width: bool,
    /// Whether the width was adjusted to fit cell contents.
    pub best_fit: bool,
    /// Whether cells show phonetic information by default.
    pub show_phonetic: bool,
    /// Outline level from 0 through 7.
    pub outline_level: u8,
    /// Whether this outline group is collapsed.
    pub collapsed: bool,
}

/// Formatting, visibility, and occupied-column metadata for an XLSB row.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsbRowInfo {
    /// Zero-based row index.
    pub row: u32,
    /// Applied row style, absent when `fGhostDirty` is clear.
    pub style_id: Option<u32>,
    /// Custom height in points, absent when `fUnsynced` is clear.
    pub height: Option<f64>,
    /// Whether extra top-border ascender padding is allocated.
    pub extra_ascender: bool,
    /// Whether extra bottom-border descender padding is allocated.
    pub extra_descender: bool,
    /// Outline level from 0 through 7.
    pub outline_level: u8,
    /// Whether preceding child rows are collapsed.
    pub collapsed: bool,
    /// Whether this row is hidden.
    pub hidden: bool,
    /// Whether cells show phonetic information by default.
    pub show_phonetic: bool,
    /// Inclusive occupied-column spans, one per 1,024-column segment.
    pub column_spans: Vec<(u32, u32)>,
}

/// Cell range governed by a worksheet AutoFilter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbAutoFilter {
    /// First zero-based row, inclusive.
    pub first_row: u32,
    /// Last zero-based row, inclusive.
    pub last_row: u32,
    /// First zero-based column, inclusive.
    pub first_column: u32,
    /// Last zero-based column, inclusive.
    pub last_column: u32,
}

/// Worksheet protection state and permissions from `BrtSheetProtection`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbSheetProtection {
    /// Legacy password verifier, absent when the stored value is zero.
    pub password_hash: Option<u16>,
    /// Whether the worksheet and locked-cell contents are protected.
    pub locked: bool,
    pub allow_edit_objects: bool,
    pub allow_edit_scenarios: bool,
    pub allow_format_cells: bool,
    pub allow_format_columns: bool,
    pub allow_format_rows: bool,
    pub allow_insert_columns: bool,
    pub allow_insert_rows: bool,
    pub allow_insert_hyperlinks: bool,
    pub allow_delete_columns: bool,
    pub allow_delete_rows: bool,
    pub allow_select_locked_cells: bool,
    pub allow_sort: bool,
    pub allow_auto_filter: bool,
    pub allow_pivot_tables: bool,
    pub allow_select_unlocked_cells: bool,
}

/// XLSB worksheet implementation
#[derive(Debug, Clone)]
pub struct XlsbWorksheet {
    name: String,
    cells: BTreeMap<(u32, u32), XlsbCell>,
    max_row: u32,
    max_col: u32,
    merged_cells: Vec<MergedCell>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<Comment>,
    column_infos: Vec<XlsbColumnInfo>,
    row_infos: Vec<XlsbRowInfo>,
    auto_filter: Option<XlsbAutoFilter>,
    sheet_protection: Option<XlsbSheetProtection>,
}

impl XlsbWorksheet {
    /// Create a new worksheet
    pub fn new(name: String) -> Self {
        XlsbWorksheet {
            name,
            cells: BTreeMap::new(),
            max_row: 0,
            max_col: 0,
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            column_infos: Vec::new(),
            row_infos: Vec::new(),
            auto_filter: None,
            sheet_protection: None,
        }
    }

    /// Add a cell to the worksheet
    pub fn add_cell(&mut self, cell: XlsbCell) {
        let pos = (cell.row(), cell.column());
        self.max_row = self.max_row.max(cell.row());
        self.max_col = self.max_col.max(cell.column());
        self.cells.insert(pos, cell);
    }

    /// Get cell at position
    pub fn get_cell(&self, row: u32, col: u32) -> Option<&XlsbCell> {
        self.cells.get(&(row, col))
    }

    /// Add a merged cell range
    pub fn add_merged_cell(&mut self, merged: MergedCell) {
        self.merged_cells.push(merged);
    }

    /// Add a hyperlink
    pub fn add_hyperlink(&mut self, hyperlink: Hyperlink) {
        self.hyperlinks.push(hyperlink);
    }

    /// Add a comment
    pub fn add_comment(&mut self, comment: Comment) {
        self.comments.push(comment);
    }

    pub(crate) fn set_column_infos(&mut self, infos: Vec<XlsbColumnInfo>) {
        self.column_infos = infos;
    }

    pub(crate) fn set_row_infos(&mut self, infos: Vec<XlsbRowInfo>) {
        self.row_infos = infos;
    }

    pub(crate) fn set_auto_filter(&mut self, auto_filter: Option<XlsbAutoFilter>) {
        self.auto_filter = auto_filter;
    }

    pub(crate) fn set_sheet_protection(&mut self, sheet_protection: Option<XlsbSheetProtection>) {
        self.sheet_protection = sheet_protection;
    }

    /// Get all merged cells
    pub fn merged_cells(&self) -> &[MergedCell] {
        &self.merged_cells
    }

    /// Get all hyperlinks
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.hyperlinks
    }

    /// Get all comments
    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }

    /// Column-range formatting and visibility records in stream order.
    pub fn column_infos(&self) -> &[XlsbColumnInfo] {
        &self.column_infos
    }

    /// Row formatting, visibility, and occupied-column records in row order.
    pub fn row_infos(&self) -> &[XlsbRowInfo] {
        &self.row_infos
    }

    /// Worksheet AutoFilter range, if present.
    pub fn auto_filter(&self) -> Option<XlsbAutoFilter> {
        self.auto_filter
    }

    /// Worksheet protection state and allowed operations, if enabled.
    pub fn sheet_protection(&self) -> Option<XlsbSheetProtection> {
        self.sheet_protection
    }
}

impl Worksheet for XlsbWorksheet {
    fn name(&self) -> &str {
        &self.name
    }

    fn row_count(&self) -> usize {
        (self.max_row + 1) as usize
    }

    fn column_count(&self) -> usize {
        (self.max_col + 1) as usize
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        if self.cells.is_empty() {
            None
        } else {
            Some((0, 0, self.max_row, self.max_col))
        }
    }

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn SheetCell + '_>> {
        match self.cells.get(&(row, column)) {
            Some(cell) => Ok(Box::new(cell.clone())),
            None => {
                // Return empty cell for missing positions
                let empty_cell = XlsbCell::new(row, column, CellValue::Empty);
                Ok(Box::new(empty_cell))
            },
        }
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn SheetCell + '_>> {
        let (row, col) = crate::xlsb::utils::parse_cell_reference(coordinate)
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error + Send + Sync>)?;
        self.cell(row, col)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(XlsbCellIterator {
            cells: self.cells.values().collect(),
            index: 0,
        })
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        Box::new(XlsbRowIterator {
            worksheet: self,
            current_row: 0,
        })
    }

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        let row_idx = row_idx as u32;
        let mut row_data = Vec::new();

        for col in 0..=self.max_col {
            match self.cells.get(&(row_idx, col)) {
                Some(cell) => row_data.push(cell.value().clone()),
                None => row_data.push(CellValue::Empty),
            }
        }

        Ok(Cow::Owned(row_data))
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        match self.cells.get(&(row, column)) {
            Some(cell) => Ok(Cow::Borrowed(cell.value())),
            None => Ok(Cow::Borrowed(CellValue::EMPTY)),
        }
    }
}

/// Cell iterator for XLSB worksheets
struct XlsbCellIterator<'a> {
    cells: Vec<&'a XlsbCell>,
    index: usize,
}

impl<'a> CellIterator<'a> for XlsbCellIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetCell + 'a>>> {
        if self.index >= self.cells.len() {
            None
        } else {
            let cell = self.cells[self.index];
            self.index += 1;
            Some(Ok(Box::new(cell.clone())))
        }
    }
}

/// Row iterator for XLSB worksheets
struct XlsbRowIterator<'a> {
    worksheet: &'a XlsbWorksheet,
    current_row: usize,
}

impl<'a> RowIterator<'a> for XlsbRowIterator<'a> {
    fn next(&mut self) -> Option<Result<Cow<'a, [CellValue]>>> {
        if self.current_row >= self.worksheet.row_count() {
            None
        } else {
            let result = self.worksheet.row(self.current_row);
            self.current_row += 1;
            Some(result)
        }
    }
}
