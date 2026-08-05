//! Semantic sheet assembly and bounded row/column materialisation.

use super::{
    ConditionalFormat, DdeSource, Error, MAX_CONDITIONAL_FORMATS_PER_SHEET,
    MAX_DEFERRED_BLANK_ROW_RUNS, MAX_EXPANDED_CELLS_PER_SHEET, MAX_EXPANDED_COLUMNS_PER_SHEET,
    MAX_EXPANDED_ROWS_PER_SHEET, MAX_SPARKLINE_GROUPS_PER_SHEET, MAX_TRAILING_EMPTY_ROWS, Row,
    Sheet, SheetPrintSettings, SheetProtection, SheetScenario, SheetStyle, SheetTableSource,
    SparklineGroup, StructureStack,
};

/// Builder for constructing a semantic [`Sheet`] during content traversal.
pub(crate) struct SheetBuilder {
    name: String,
    rows: Vec<Row>,
    columns: Vec<Column>,
    row_structure: StructureStack,
    column_structure: StructureStack,
    style: SheetStyle,
    print_settings: SheetPrintSettings,
    title: Option<String>,
    description: Option<String>,
    table_source: Option<SheetTableSource>,
    dde_source: Option<DdeSource>,
    scenario: Option<SheetScenario>,
    conditional_formats: Vec<ConditionalFormat>,
    sparkline_groups: Vec<SparklineGroup>,
    images: Vec<crate::Image>,
    cell_count: usize,
    /// Runs of empty rows read but not yet materialised, in document order.
    deferred_rows: Vec<(Row, usize)>,
    /// Total number of rows the deferred runs stand for.
    deferred_row_count: usize,
}

impl SheetBuilder {
    #[cfg(test)]
    pub fn new(name: String) -> Self {
        Self::with_formatting(name, SheetStyle::default(), SheetPrintSettings::default())
    }

    pub(crate) fn with_formatting(
        name: String,
        style: SheetStyle,
        print_settings: SheetPrintSettings,
    ) -> Self {
        Self {
            name,
            rows: Vec::new(),
            columns: Vec::new(),
            row_structure: StructureStack::new(),
            column_structure: StructureStack::new(),
            style,
            print_settings,
            title: None,
            description: None,
            table_source: None,
            dde_source: None,
            scenario: None,
            conditional_formats: Vec::new(),
            sparkline_groups: Vec::new(),
            images: Vec::new(),
            cell_count: 0,
            deferred_rows: Vec::new(),
            deferred_row_count: 0,
        }
    }

    pub(crate) fn set_scenario(&mut self, scenario: SheetScenario) -> Result<()> {
        if self.scenario.replace(scenario).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one scenario".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn add_conditional_format(&mut self, format: ConditionalFormat) -> Result<()> {
        if self.conditional_formats.len() >= MAX_CONDITIONAL_FORMATS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "sheet exceeds the {MAX_CONDITIONAL_FORMATS_PER_SHEET} conditional format safety limit"
            )));
        }
        self.conditional_formats.push(format);
        Ok(())
    }

    pub(crate) fn add_sparkline_group(&mut self, group: SparklineGroup) -> Result<()> {
        if self.sparkline_groups.len() >= MAX_SPARKLINE_GROUPS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "sheet exceeds the {MAX_SPARKLINE_GROUPS_PER_SHEET} sparkline group safety limit"
            )));
        }
        self.sparkline_groups.push(group);
        Ok(())
    }

    pub(crate) fn set_dde_source(&mut self, source: DdeSource) -> Result<()> {
        if self.scenario.is_some() {
            return Err(Error::InvalidFormat(
                "office:dde-source must precede table:scenario".to_string(),
            ));
        }
        if self.dde_source.replace(source).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one office:dde-source".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn set_table_source(&mut self, source: SheetTableSource) -> Result<()> {
        if self.dde_source.is_some() || self.scenario.is_some() {
            return Err(Error::InvalidFormat(
                "table:table-source must precede office:dde-source and table:scenario".to_string(),
            ));
        }
        if self.table_source.replace(source).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one table source".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn set_title(&mut self, title: String) -> Result<()> {
        if self.title.replace(title).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one title".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn set_description(&mut self, description: String) -> Result<()> {
        if self.description.replace(description).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one description".to_string(),
            ));
        }
        Ok(())
    }

    pub fn add_row(&mut self, mut row: Row) {
        let row_index = self.rows.len();
        row.index = row_index;
        // Update row index for all cells in this row
        for cell in &mut row.cells {
            cell.row = row_index;
        }
        self.rows.push(row);
    }

    /// Number of rows the sheet logically spans, including runs of empty rows
    /// that have been deferred and may never be materialised.
    fn logical_row_count(&self) -> usize {
        self.rows.len().saturating_add(self.deferred_row_count)
    }

    /// Open a row grouping. Pending empty rows belong to the enclosing context,
    /// so they are materialised before the boundary moves.
    pub(crate) fn begin_row_group(&mut self, display: bool) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.begin_group(display)
    }

    pub(crate) fn end_row_group(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.end_group()
    }

    pub(crate) fn begin_row_header(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.begin_header(self.rows.len())
    }

    pub(crate) fn end_row_header(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.end_header(self.rows.len())
    }

    pub(crate) fn add_repeated_row(&mut self, row: Row, repeated: usize) -> Result<()> {
        // The logical extent has to stay inside the grid a spreadsheet can
        // address, whether or not the rows are eventually materialised.
        let logical_end = self
            .logical_row_count()
            .checked_add(repeated)
            .ok_or_else(|| {
                Error::InvalidFormat("table row repetition overflows address space".to_string())
            })?;
        if logical_end > MAX_EXPANDED_ROWS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "expanded sheet exceeds the {MAX_EXPANDED_ROWS_PER_SHEET} row safety limit"
            )));
        }

        // Rows with no cell at all are the sheet-height padding producers append
        // after the used range. Defer them: an interior run is expanded again as
        // soon as a row with content follows, while a long trailing run is
        // discarded by `build` instead of costing a million allocations.
        if row.cells.is_empty() && self.deferred_rows.len() < MAX_DEFERRED_BLANK_ROW_RUNS {
            self.deferred_rows.push((row, repeated));
            self.deferred_row_count = self.deferred_row_count.saturating_add(repeated);
            return Ok(());
        }

        self.flush_deferred_rows()?;
        self.materialize_repeated_row(row, repeated)
    }

    /// Materialise every deferred empty-row run because a row with content or a
    /// structure boundary follows, so later rows keep their true index.
    fn flush_deferred_rows(&mut self) -> Result<()> {
        for (row, repeated) in std::mem::take(&mut self.deferred_rows) {
            self.materialize_repeated_row(row, repeated)?;
        }
        self.deferred_row_count = 0;
        Ok(())
    }

    /// Resolve the run of empty rows still pending at the end of a table.
    ///
    /// A short tail is kept, so an authored gap of blank rows survives the round
    /// trip. A long one is producer grid padding — every ODF spreadsheet is
    /// written out to its full addressable height — and is discarded, since it
    /// holds no cell, value, formula, annotation, or text. Discarding it records
    /// no structure range either, so the sheet's row groups never describe rows
    /// that are not there.
    fn finish_deferred_rows(&mut self) -> Result<()> {
        if self.deferred_row_count <= MAX_TRAILING_EMPTY_ROWS {
            return self.flush_deferred_rows();
        }
        self.deferred_rows.clear();
        self.deferred_row_count = 0;
        Ok(())
    }

    /// Expand one run of rows and record the physical range it occupies.
    fn materialize_repeated_row(&mut self, row: Row, repeated: usize) -> Result<()> {
        let start = self.rows.len();
        let added_cells = row.cells.len().checked_mul(repeated).ok_or_else(|| {
            Error::InvalidFormat("table row repetition overflows cell count".to_string())
        })?;
        let expanded_cells = self.cell_count.checked_add(added_cells).ok_or_else(|| {
            Error::InvalidFormat("expanded sheet cell count overflows address space".to_string())
        })?;
        if expanded_cells > MAX_EXPANDED_CELLS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "expanded sheet exceeds the {MAX_EXPANDED_CELLS_PER_SHEET} cell safety limit"
            )));
        }
        self.cell_count = expanded_cells;
        self.rows.reserve(repeated);
        for _ in 0..repeated {
            self.add_row(row.clone());
        }
        self.row_structure.add_range(start, self.rows.len())
    }

    pub(crate) fn add_repeated_column(&mut self, column: Column, repeated: usize) -> Result<()> {
        let start = self.columns.len();
        let expanded = self.columns.len().checked_add(repeated).ok_or_else(|| {
            Error::InvalidFormat("table column repetition overflows address space".to_string())
        })?;
        if expanded > MAX_EXPANDED_COLUMNS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "expanded sheet exceeds the {MAX_EXPANDED_COLUMNS_PER_SHEET} column safety limit"
            )));
        }
        for _ in 0..repeated {
            let mut item = column.clone();
            item.index = self.columns.len();
            self.columns.push(item);
        }
        self.column_structure.add_range(start, self.columns.len())?;
        Ok(())
    }

    pub(crate) fn begin_column_group(&mut self, display: bool) -> Result<()> {
        self.column_structure.begin_group(display)
    }

    pub(crate) fn end_column_group(&mut self) -> Result<()> {
        self.column_structure.end_group()
    }

    pub(crate) fn begin_column_header(&mut self) -> Result<()> {
        self.column_structure.begin_header(self.columns.len())
    }

    pub(crate) fn end_column_header(&mut self) -> Result<()> {
        self.column_structure.end_header(self.columns.len())
    }

    pub(crate) fn build(mut self) -> Result<Sheet> {
        self.finish_deferred_rows()?;
        Ok(Sheet {
            name: self.name,
            rows: self.rows,
            columns: self.columns,
            column_structure: self.column_structure.finish()?,
            row_structure: self.row_structure.finish()?,
            style: self.style,
            print_settings: self.print_settings,
            title: self.title,
            description: self.description,
            table_source: self.table_source,
            dde_source: self.dde_source,
            scenario: self.scenario,
            conditional_formats: self.conditional_formats,
            sparkline_groups: self.sparkline_groups,
            images: self.images,
            shapes: Vec::new(),
            protection: SheetProtection::default(),
        })
    }
}
