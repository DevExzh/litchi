//! Worksheet implementation for XLS files

use crate::xls::autofilter::{AutoFilterColumn, AutoFilterInfo, SortInfo};
use crate::xls::cell::XlsCell;
use crate::xls::comments::XlsComment;
use crate::xls::error::XlsError;
use crate::xls::hyperlinks::XlsHyperlink;
use crate::xls::layout::{XlsColumnLayout, XlsRowLayout};
use crate::xls::view::XlsWorksheetView;
use crate::xls::page_setup::XlsPageSetup;
use crate::xls::conditional_format::XlsConditionalFormatting;
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
    /// Rich-text and phonetic metadata parallel to the shared string table.
    shared_string_properties:
        Option<Arc<Vec<Option<Box<crate::xls::records::SharedStringProperties>>>>>,
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
    row_layouts: BTreeMap<u16, XlsRowLayout>,
    column_layouts: Vec<XlsColumnLayout>,
    worksheet_views: Vec<XlsWorksheetView>,
    page_setup: Option<XlsPageSetup>,
    calculation: crate::xls::calculation::XlsWorksheetCalculation,
    scenario_manager: Option<crate::xls::scenario::XlsScenarioManager>,
    conditional_formattings: Vec<XlsConditionalFormatting>,
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
            shared_string_properties: None,
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
            row_layouts: BTreeMap::new(),
            column_layouts: Vec::new(),
            worksheet_views: Vec::new(),
            page_setup: None,
            calculation: crate::xls::calculation::XlsWorksheetCalculation::default(),
            scenario_manager: None,
            conditional_formattings: Vec::new(),
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
            shared_string_properties: None,
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
            row_layouts: BTreeMap::new(),
            column_layouts: Vec::new(),
            worksheet_views: Vec::new(),
            page_setup: None,
            calculation: crate::xls::calculation::XlsWorksheetCalculation::default(),
            scenario_manager: None,
            conditional_formattings: Vec::new(),
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

    pub(crate) fn with_shared_string_properties(
        name: String,
        shared_strings: Arc<Vec<String>>,
        shared_string_properties: Arc<
            Vec<Option<Box<crate::xls::records::SharedStringProperties>>>,
        >,
    ) -> Self {
        let mut worksheet = Self::with_shared_strings(name, shared_strings);
        worksheet.shared_string_properties = Some(shared_string_properties);
        worksheet
    }

    /// Rich-text and phonetic metadata for a zero-based shared-string index.
    pub fn shared_string_properties(
        &self,
        index: u32,
    ) -> Option<&crate::xls::records::SharedStringProperties> {
        self.shared_string_properties
            .as_ref()?
            .get(index as usize)?
            .as_deref()
    }

    /// Rich-text and phonetic metadata for a `LabelSst` cell.
    pub fn shared_string_properties_for_cell(
        &self,
        cell: &XlsCell,
    ) -> Option<&crate::xls::records::SharedStringProperties> {
        self.shared_string_properties(cell.shared_string_index()?)
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

    /// All hyperlinks in this worksheet.
    pub fn hyperlinks(&self) -> &[XlsHyperlink] {
        &self.hyperlinks
    }

    pub(crate) fn set_hyperlinks(&mut self, hyperlinks: Vec<XlsHyperlink>) {
        self.hyperlinks = hyperlinks;
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

    /// Layout metadata for a zero-based row index.
    pub fn row_layout(&self, row: u16) -> Option<&XlsRowLayout> {
        self.row_layouts.get(&row)
    }

    /// Row layout metadata in ascending row order.
    pub fn row_layouts(&self) -> impl ExactSizeIterator<Item = &XlsRowLayout> {
        self.row_layouts.values()
    }

    /// Column layout ranges in BIFF record order.
    pub fn column_layouts(&self) -> &[XlsColumnLayout] {
        &self.column_layouts
    }

    /// The first layout range containing a zero-based column index.
    pub fn column_layout(&self, column: u8) -> Option<&XlsColumnLayout> {
        let column = u16::from(column);
        self.column_layouts
            .iter()
            .find(|layout| (layout.first_column()..=layout.last_column()).contains(&column))
    }

    pub(crate) fn set_layouts(
        &mut self,
        rows: BTreeMap<u16, XlsRowLayout>,
        columns: Vec<XlsColumnLayout>,
    ) {
        self.row_layouts = rows;
        self.column_layouts = columns;
    }

    /// The first display window associated with this worksheet.
    pub fn worksheet_view(&self) -> Option<&XlsWorksheetView> {
        self.worksheet_views.first()
    }

    /// All display windows associated with this worksheet in record order.
    pub fn worksheet_views(&self) -> &[XlsWorksheetView] {
        &self.worksheet_views
    }

    pub(crate) fn set_worksheet_views(&mut self, views: Vec<XlsWorksheetView>) {
        self.worksheet_views = views;
    }

    /// Print and page setup for this worksheet.
    pub fn page_setup(&self) -> Option<&XlsPageSetup> {
        self.page_setup.as_ref()
    }

    pub(crate) fn set_page_setup(&mut self, page_setup: Option<XlsPageSetup>) {
        self.page_setup = page_setup;
    }

    pub fn calculation(&self) -> &crate::xls::calculation::XlsWorksheetCalculation {
        &self.calculation
    }

    pub(crate) fn set_calculation(
        &mut self,
        calculation: crate::xls::calculation::XlsWorksheetCalculation,
    ) {
        self.calculation = calculation;
    }

    pub fn scenario_manager(&self) -> Option<&crate::xls::scenario::XlsScenarioManager> {
        self.scenario_manager.as_ref()
    }

    pub(crate) fn set_scenario_manager(
        &mut self,
        scenario_manager: Option<crate::xls::scenario::XlsScenarioManager>,
    ) {
        self.scenario_manager = scenario_manager;
    }

    /// Legacy conditional formatting groups in worksheet record order.
    pub fn conditional_formattings(&self) -> &[XlsConditionalFormatting] {
        &self.conditional_formattings
    }

    pub(crate) fn set_conditional_formattings(
        &mut self,
        conditional_formattings: Vec<XlsConditionalFormatting>,
    ) {
        self.conditional_formattings = conditional_formattings;
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
