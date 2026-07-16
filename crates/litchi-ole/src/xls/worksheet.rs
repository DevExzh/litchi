//! Worksheet implementation for XLS files

use crate::xls::autofilter::{AutoFilterColumn, AutoFilterInfo, SortInfo};
use crate::xls::cell::XlsCell;
use crate::xls::comments::XlsComment;
use crate::xls::error::XlsError;
use crate::xls::hyperlinks::XlsHyperlink;
use crate::xls::merged_cells::MergedCellRange;
use crate::xls::number_format::{XlsExtendedFormat, XlsFormatting, XlsNumberFormat};
use crate::xls::pivot_table::PivotTable;
use crate::xls::protection::SheetProtection;
use litchi_core::sheet::{
    Cell as SheetCell, CellIterator, CellValue, Result, RowIterator, Worksheet,
};
use std::borrow::Cow;
use std::collections::BTreeMap;
use std::sync::Arc;

use super::data_validation::{XlsDataValidationRule, XlsDataValidationSettings};

/// XLS worksheet implementation
#[derive(Debug, Clone)]
pub struct XlsWorksheet {
    name: String,
    cells: BTreeMap<(u32, u32), XlsCell>,
    max_row: u32,
    max_col: u32,
    /// Shared string table (Arc for zero-copy sharing across worksheets)
    shared_strings: Option<Arc<Vec<String>>>,
    /// Merged cell ranges (MERGECELLS records)
    merged_cells: Vec<MergedCellRange>,
    /// Hyperlinks (HLINK records)
    hyperlinks: Vec<XlsHyperlink>,
    /// Comments/notes (NOTE records)
    comments: Vec<XlsComment>,
    /// AutoFilter configuration (AUTOFILTERINFO + AUTOFILTER records)
    autofilter: Option<AutoFilterInfo>,
    /// Sort configuration (SORT record)
    sort_info: Option<SortInfo>,
    /// Pivot tables (aggregated SX* records)
    pivot_tables: Vec<PivotTable>,
    /// Sheet protection state (PROTECT/OBJECTPROTECT/SCENPROTECT/PASSWORD)
    protection: SheetProtection,
    formatting: Arc<XlsFormatting>,
    data_validation_settings: Option<XlsDataValidationSettings>,
    data_validations: Vec<XlsDataValidationRule>,
}

impl XlsWorksheet {
    /// Create a new worksheet
    pub fn new(name: String) -> Self {
        XlsWorksheet {
            name,
            cells: BTreeMap::new(),
            max_row: 0,
            max_col: 0,
            shared_strings: None,
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            autofilter: None,
            sort_info: None,
            pivot_tables: Vec::new(),
            protection: SheetProtection::default(),
            formatting: Arc::new(XlsFormatting::default()),
            data_validation_settings: None,
            data_validations: Vec::new(),
        }
    }

    /// Create a new worksheet with shared strings (Arc for zero-copy sharing)
    pub fn with_shared_strings(name: String, shared_strings: Arc<Vec<String>>) -> Self {
        XlsWorksheet {
            name,
            cells: BTreeMap::new(),
            max_row: 0,
            max_col: 0,
            shared_strings: Some(shared_strings),
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            autofilter: None,
            sort_info: None,
            pivot_tables: Vec::new(),
            protection: SheetProtection::default(),
            formatting: Arc::new(XlsFormatting::default()),
            data_validation_settings: None,
            data_validations: Vec::new(),
        }
    }

    /// Add a cell to the worksheet
    pub fn add_cell(&mut self, cell: XlsCell) {
        let pos = (cell.row(), cell.column());
        self.max_row = self.max_row.max(cell.row());
        self.max_col = self.max_col.max(cell.column());
        self.cells.insert(pos, cell);
    }

    /// Set worksheet dimensions
    pub fn set_dimensions(
        &mut self,
        _first_row: u32,
        last_row: u32,
        _first_col: u32,
        last_col: u32,
    ) {
        // Adjust max_row and max_col based on dimensions
        self.max_row = self.max_row.max(last_row.saturating_sub(1));
        self.max_col = self.max_col.max(last_col.saturating_sub(1));
    }

    /// Get shared strings reference
    pub fn shared_strings(&self) -> Option<&[String]> {
        self.shared_strings.as_ref().map(|arc| arc.as_slice())
    }

    /// Get cell at position
    pub fn get_cell(&self, row: u32, col: u32) -> Option<&XlsCell> {
        self.cells.get(&(row, col))
    }

    pub(crate) fn get_cell_mut(&mut self, row: u32, col: u32) -> Option<&mut XlsCell> {
        self.cells.get_mut(&(row, col))
    }

    // -- Merged cells --

    /// Add merged cell ranges parsed from a MERGECELLS record.
    pub fn add_merged_cells(&mut self, ranges: &[MergedCellRange]) {
        self.merged_cells.extend_from_slice(ranges);
    }

    /// All merged cell ranges in this worksheet.
    pub fn merged_cells(&self) -> &[MergedCellRange] {
        &self.merged_cells
    }

    // -- Hyperlinks --

    /// Add a parsed hyperlink.
    pub fn add_hyperlink(&mut self, link: XlsHyperlink) {
        self.hyperlinks.push(link);
    }

    /// All hyperlinks in this worksheet.
    pub fn hyperlinks(&self) -> &[XlsHyperlink] {
        &self.hyperlinks
    }

    // -- Comments --

    /// All comments/notes in this worksheet.
    pub fn comments(&self) -> &[XlsComment] {
        &self.comments
    }

    pub(crate) fn set_comments(&mut self, comments: Vec<XlsComment>) {
        self.comments = comments;
    }

    // -- AutoFilter --

    /// Initialize AutoFilter with the given column count (from AUTOFILTERINFO).
    pub fn set_autofilter_info(&mut self, column_count: u16) {
        self.autofilter = Some(AutoFilterInfo {
            column_count,
            columns: Vec::new(),
        });
    }

    /// Add an AutoFilter column definition (from AUTOFILTER record).
    pub fn add_autofilter_column(&mut self, col: AutoFilterColumn) {
        if let Some(ref mut af) = self.autofilter {
            af.columns.push(col);
        }
    }

    /// AutoFilter configuration, if any.
    pub fn autofilter(&self) -> Option<&AutoFilterInfo> {
        self.autofilter.as_ref()
    }

    // -- Sort --

    /// Set the sort configuration (from SORT record).
    pub fn set_sort_info(&mut self, info: SortInfo) {
        self.sort_info = Some(info);
    }

    /// Sort configuration, if any.
    pub fn sort_info(&self) -> Option<&SortInfo> {
        self.sort_info.as_ref()
    }

    // -- Pivot tables --

    /// Add a fully assembled pivot table.
    pub fn add_pivot_table(&mut self, pt: PivotTable) {
        self.pivot_tables.push(pt);
    }

    /// All pivot tables in this worksheet.
    pub fn pivot_tables(&self) -> &[PivotTable] {
        &self.pivot_tables
    }

    // -- Sheet Protection --

    /// Get the sheet protection state.
    pub fn protection(&self) -> &SheetProtection {
        &self.protection
    }

    /// Get a mutable reference to the sheet protection state.
    pub fn protection_mut(&mut self) -> &mut SheetProtection {
        &mut self.protection
    }

    pub(crate) fn set_formatting(&mut self, formatting: Arc<XlsFormatting>) {
        self.formatting = formatting;
    }

    pub fn formatting(&self) -> &XlsFormatting {
        &self.formatting
    }

    /// Worksheet-level BIFF8 data-validation settings, when present.
    pub fn data_validation_settings(&self) -> Option<&XlsDataValidationSettings> {
        self.data_validation_settings.as_ref()
    }

    /// Data-validation rules in worksheet record order.
    pub fn data_validations(&self) -> &[XlsDataValidationRule] {
        &self.data_validations
    }

    pub(crate) fn set_data_validation_settings(&mut self, settings: XlsDataValidationSettings) {
        self.data_validation_settings = Some(settings);
    }

    pub(crate) fn add_data_validation(&mut self, rule: XlsDataValidationRule) {
        self.data_validations.push(rule);
    }

    pub fn format_for_cell(&self, row: u32, col: u32) -> Option<&XlsExtendedFormat> {
        let cell = self.get_cell(row, col)?;
        self.formatting.cell_format(cell.xf_index())
    }

    pub fn number_format_for_cell(&self, row: u32, col: u32) -> Option<&XlsNumberFormat> {
        let format = self.format_for_cell(row, col)?;
        self.formatting.number_format(format.number_format_id())
    }

    pub fn is_date_time_formatted(&self, row: u32, col: u32) -> bool {
        self.format_for_cell(row, col)
            .map(|format| {
                self.formatting
                    .is_date_time_format(format.number_format_id())
            })
            .unwrap_or(false)
    }
}

impl Worksheet for XlsWorksheet {
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
            // Return reference instead of clone - zero-copy!
            Some(cell) => Ok(Box::new(cell)),
            None => {
                // Return empty cell for missing positions (owned, unavoidable)
                let empty_cell = XlsCell::new(row, column, CellValue::Empty);
                Ok(Box::new(empty_cell))
            },
        }
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn SheetCell + '_>> {
        let (row, col) = crate::xls::utils::parse_cell_reference(coordinate).ok_or_else(|| {
            Box::new(XlsError::InvalidCellReference(coordinate.to_string()))
                as Box<dyn std::error::Error + Send + Sync>
        })?;
        self.cell(row, col)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(XlsCellIterator {
            cells: self.cells.values().collect(),
            index: 0,
        })
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        Box::new(XlsRowIterator {
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

        // Return owned data wrapped in Cow
        Ok(Cow::Owned(row_data))
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        match self.cells.get(&(row, column)) {
            Some(cell) => Ok(Cow::Borrowed(cell.value())),
            None => Ok(Cow::Borrowed(CellValue::EMPTY)),
        }
    }
}

// Implement Worksheet for &XlsWorksheet to allow zero-copy reference returns
impl Worksheet for &XlsWorksheet {
    fn name(&self) -> &str {
        (*self).name()
    }

    fn row_count(&self) -> usize {
        (*self).row_count()
    }

    fn column_count(&self) -> usize {
        (*self).column_count()
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        (*self).dimensions()
    }

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn SheetCell + '_>> {
        (*self).cell(row, column)
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn SheetCell + '_>> {
        (*self).cell_by_coordinate(coordinate)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        (*self).cells()
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        (*self).rows()
    }

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        (*self).row(row_idx)
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        (*self).cell_value(row, column)
    }
}

/// Cell iterator for XLS worksheets
struct XlsCellIterator<'a> {
    cells: Vec<&'a XlsCell>,
    index: usize,
}

impl<'a> CellIterator<'a> for XlsCellIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetCell + 'a>>> {
        if self.index >= self.cells.len() {
            None
        } else {
            let cell = self.cells[self.index];
            self.index += 1;
            // Return reference instead of clone - zero-copy!
            Some(Ok(Box::new(cell)))
        }
    }
}

/// Row iterator for XLS worksheets
struct XlsRowIterator<'a> {
    worksheet: &'a XlsWorksheet,
    current_row: usize,
}

impl<'a> RowIterator<'a> for XlsRowIterator<'a> {
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
