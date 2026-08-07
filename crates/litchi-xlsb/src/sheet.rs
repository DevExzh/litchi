//! Worksheet implementation for XLSB files

use crate::comments::Record;
use crate::conditional_formatting::Formatting;
use crate::hyperlinks::Hyperlink;
use crate::merged_cells::MergedCell;
use crate::package::cell::Cell;
use crate::package::data_validation::{Settings, Validation};
use crate::package::scenarios::Manager;
use crate::package::web_extension_bindings::Binding;
use crate::slicer;
use litchi_core::sheet::{
    Cell as SheetCell, CellIterator, CellValue, Result, RowIterator, Worksheet as SheetWorksheet,
};
use litchi_sheet::view::View;
use std::borrow::Cow;
use std::collections::BTreeMap;

/// Width, style, visibility, and outline metadata for an XLSB column range.
#[derive(Debug, Clone, PartialEq)]
pub struct ColumnInfo {
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
pub struct RowInfo {
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
pub struct AutoFilter {
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
pub struct SheetProtection {
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

/// Strong password-verifier metadata from `BrtSheetProtectionIso`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StrongProtection {
    /// Number of password-hash iterations, at most 10,000,000.
    pub spin_count: u32,
    /// Calculated password hash bytes.
    pub hash: Vec<u8>,
    /// Salt bytes used to calculate the hash.
    pub salt: Vec<u8>,
    /// Hash algorithm name, such as `SHA-512`.
    pub algorithm: String,
}

/// XLSB worksheet implementation
#[derive(Debug, Clone)]
pub struct Worksheet {
    name: String,
    cells: BTreeMap<(u32, u32), Cell>,
    max_row: u32,
    max_col: u32,
    merged_cells: Vec<MergedCell>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<Record>,
    column_infos: Vec<ColumnInfo>,
    row_infos: Vec<RowInfo>,
    auto_filter: Option<AutoFilter>,
    sheet_protection: Option<SheetProtection>,
    strong_sheet_protection: Option<StrongProtection>,
    data_validations: Vec<Validation>,
    data_validation_settings: Option<Settings>,
    data_validation14_settings: Option<Settings>,
    conditional_formattings: Vec<Formatting>,
    web_extension_bindings: Vec<Binding>,
    views: Vec<View>,
    scenarios: Option<Manager>,
    slicers: Option<slicer::Views>,
    timelines: Option<crate::timeline::Views>,
}

impl Worksheet {
    /// Create a new worksheet
    pub fn new(name: String) -> Self {
        Worksheet {
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
            strong_sheet_protection: None,
            data_validations: Vec::new(),
            data_validation_settings: None,
            data_validation14_settings: None,
            conditional_formattings: Vec::new(),
            web_extension_bindings: Vec::new(),
            views: Vec::new(),
            scenarios: None,
            slicers: None,
            timelines: None,
        }
    }

    /// Add a cell to the worksheet
    pub fn add_cell(&mut self, cell: Cell) {
        let pos = (cell.row(), cell.column());
        self.max_row = self.max_row.max(cell.row());
        self.max_col = self.max_col.max(cell.column());
        self.cells.insert(pos, cell);
    }

    /// Get cell at position
    pub fn get_cell(&self, row: u32, col: u32) -> Option<&Cell> {
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
    pub fn add_comment(&mut self, comment: Record) {
        self.comments.push(comment);
    }

    pub(crate) fn set_column_infos(&mut self, infos: Vec<ColumnInfo>) {
        self.column_infos = infos;
    }

    pub(crate) fn set_row_infos(&mut self, infos: Vec<RowInfo>) {
        self.row_infos = infos;
    }

    pub(crate) fn set_auto_filter(&mut self, auto_filter: Option<AutoFilter>) {
        self.auto_filter = auto_filter;
    }

    pub(crate) fn set_sheet_protection(&mut self, sheet_protection: Option<SheetProtection>) {
        self.sheet_protection = sheet_protection;
    }

    pub(crate) fn set_strong_sheet_protection(&mut self, protection: Option<StrongProtection>) {
        self.strong_sheet_protection = protection;
    }

    pub(crate) fn set_data_validations(
        &mut self,
        settings: Option<Settings>,
        extension14_settings: Option<Settings>,
        validations: Vec<Validation>,
    ) {
        self.data_validation_settings = settings;
        self.data_validation14_settings = extension14_settings;
        self.data_validations = validations;
    }

    pub(crate) fn set_conditional_formattings(&mut self, conditional_formattings: Vec<Formatting>) {
        self.conditional_formattings = conditional_formattings;
    }

    pub(crate) fn set_web_extension_bindings(&mut self, bindings: Vec<Binding>) {
        self.web_extension_bindings = bindings;
    }

    pub(crate) fn set_views(&mut self, views: Vec<View>) {
        self.views = views;
    }

    pub(crate) fn set_scenarios(&mut self, scenarios: Option<Manager>) {
        self.scenarios = scenarios;
    }

    pub(crate) fn set_slicers(&mut self, slicers: Option<slicer::Views>) {
        self.slicers = slicers;
    }

    pub(crate) fn set_timelines(&mut self, timelines: Option<crate::timeline::Views>) {
        self.timelines = timelines;
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
    pub fn comments(&self) -> &[Record] {
        &self.comments
    }

    /// Column-range formatting and visibility records in stream order.
    pub fn column_infos(&self) -> &[ColumnInfo] {
        &self.column_infos
    }

    /// Row formatting, visibility, and occupied-column records in row order.
    pub fn row_infos(&self) -> &[RowInfo] {
        &self.row_infos
    }

    /// Worksheet AutoFilter range, if present.
    pub fn auto_filter(&self) -> Option<AutoFilter> {
        self.auto_filter
    }

    /// Worksheet protection state and allowed operations, if enabled.
    pub fn sheet_protection(&self) -> Option<SheetProtection> {
        self.sheet_protection
    }

    /// Strong password-verifier metadata, if the ISO protection record exists.
    pub fn strong_sheet_protection(&self) -> Option<&StrongProtection> {
        self.strong_sheet_protection.as_ref()
    }

    /// Worksheet data-validation rules in stream order.
    pub fn data_validations(&self) -> &[Validation] {
        &self.data_validations
    }

    /// Worksheet-level UI settings for classic data validation.
    pub fn data_validation_settings(&self) -> Option<Settings> {
        self.data_validation_settings
    }

    /// Worksheet-level UI settings for Office 2013 data validation.
    pub fn data_validation14_settings(&self) -> Option<Settings> {
        self.data_validation14_settings
    }

    /// Conditional-formatting blocks in worksheet stream order.
    pub fn conditional_formattings(&self) -> &[Formatting] {
        &self.conditional_formattings
    }

    /// Inert Office Add-in range bindings in worksheet stream order.
    pub fn web_extension_bindings(&self) -> &[Binding] {
        &self.web_extension_bindings
    }

    /// Sheet views (zoom, panes, selections) in worksheet stream order.
    ///
    /// The view model is shared with XLSX worksheets; see
    /// [`litchi_sheet::view::View`].
    pub fn views(&self) -> &[View] {
        &self.views
    }

    /// The worksheet's inert Scenario Manager snapshot, if present.
    ///
    /// Scenario values are metadata only; litchi never substitutes them into
    /// cells or recalculates the workbook.
    pub fn scenarios(&self) -> Option<&Manager> {
        self.scenarios.as_ref()
    }

    /// The worksheet's inert XLSB slicer views, if present.
    pub fn slicers(&self) -> Option<&slicer::Views> {
        self.slicers.as_ref()
    }

    /// The worksheet's inert XLSB timeline views, if present.
    pub fn timelines(&self) -> Option<&crate::timeline::Views> {
        self.timelines.as_ref()
    }
}

impl SheetWorksheet for Worksheet {
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
                let empty_cell = Cell::new(row, column, CellValue::Empty);
                Ok(Box::new(empty_cell))
            },
        }
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn SheetCell + '_>> {
        let (row, col) = crate::package::utils::parse_cell_reference(coordinate)
            .map_err(|e| -> Box<dyn std::error::Error + Send + Sync> { Box::new(e) })?;
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
    cells: Vec<&'a Cell>,
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
    worksheet: &'a Worksheet,
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
