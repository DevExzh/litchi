//! Parser-side assembly state for ODS `content.xml`.

use super::super::{
    Cell, CellDetective, CellMatrixSpan, CellMerge, CellRangeSource, CellTextContent, CellValue,
    Column, ConditionalFormat, Row, Sheet, SheetPrintSettings, SheetScenario, SheetStyle,
    SheetTableSource, SparklineGroup, TableGroup, TableRange, TableStructure, TableVisibility,
    conditional_format::MAX_CONDITIONAL_FORMATS_PER_SHEET,
    dde::DdeSource,
    sparkline::MAX_SPARKLINE_GROUPS_PER_SHEET,
    structure::{
        MAX_EXPANDED_COLUMNS_PER_SHEET, MAX_EXPANDED_ROWS_PER_SHEET, MAX_TABLE_STRUCTURE_DEPTH,
    },
};
use crate::model::hyperlink::Link;
use litchi_core::{Error, Result};

const MAX_EXPANDED_CELLS_PER_ROW: usize = 1_048_576;
const MAX_EXPANDED_CELLS_PER_SHEET: usize = 4_194_304;
/// Interleaved runs of empty rows kept unmaterialised before the parser gives
/// up and expands them, so deferral cannot grow without bound.
const MAX_DEFERRED_BLANK_ROW_RUNS: usize = 4_096;
/// Longest run of cell-less rows still kept at the end of a table. Anything
/// longer is the full-height grid padding every ODF producer writes.
const MAX_TRAILING_EMPTY_ROWS: usize = 4_096;

struct StructureContext {
    display: Option<bool>,
    children: Vec<TableStructure>,
    header_start: Option<usize>,
}

impl StructureContext {
    fn root() -> Self {
        Self {
            display: None,
            children: Vec::new(),
            header_start: None,
        }
    }
}

struct StructureStack {
    contexts: Vec<StructureContext>,
}

impl StructureStack {
    fn new() -> Self {
        Self {
            contexts: vec![StructureContext::root()],
        }
    }

    fn begin_group(&mut self, display: bool) -> Result<()> {
        if self.contexts.len() > MAX_TABLE_STRUCTURE_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "table structure exceeds the {MAX_TABLE_STRUCTURE_DEPTH} level nesting safety limit"
            )));
        }
        if self
            .contexts
            .last()
            .is_some_and(|context| context.header_start.is_some())
        {
            return Err(Error::InvalidFormat(
                "table groups cannot be nested inside a header container".to_string(),
            ));
        }
        self.contexts.push(StructureContext {
            display: Some(display),
            children: Vec::new(),
            header_start: None,
        });
        Ok(())
    }

    fn end_group(&mut self) -> Result<()> {
        if self.contexts.len() <= 1 {
            return Err(Error::InvalidFormat(
                "table group end has no matching start".to_string(),
            ));
        }
        let context = self.contexts.pop().expect("non-root context was checked");
        if context.header_start.is_some() {
            return Err(Error::InvalidFormat(
                "table header container is not closed before its group".to_string(),
            ));
        }
        if context.children.is_empty() {
            return Err(Error::InvalidFormat(
                "table groups must contain at least one row or column".to_string(),
            ));
        }
        self.contexts
            .last_mut()
            .expect("root context is retained")
            .children
            .push(TableStructure::Group(TableGroup {
                display: context.display.expect("group contexts have display state"),
                children: context.children,
            }));
        Ok(())
    }

    fn begin_header(&mut self, position: usize) -> Result<()> {
        let context = self.contexts.last_mut().expect("root context is retained");
        if context.header_start.replace(position).is_some() {
            return Err(Error::InvalidFormat(
                "table header containers cannot be nested".to_string(),
            ));
        }
        Ok(())
    }

    fn end_header(&mut self, position: usize) -> Result<()> {
        let context = self.contexts.last_mut().expect("root context is retained");
        let start = context.header_start.take().ok_or_else(|| {
            Error::InvalidFormat("table header end has no matching start".to_string())
        })?;
        if position <= start {
            return Err(Error::InvalidFormat(
                "table header containers must not be empty".to_string(),
            ));
        }
        Ok(())
    }

    fn add_range(&mut self, start: usize, end: usize) -> Result<()> {
        let range = TableRange::new(start, end)?;
        let context = self.contexts.last_mut().expect("root context is retained");
        let entry = if context.header_start.is_some() {
            TableStructure::Header(range)
        } else {
            TableStructure::Range(range)
        };
        if let Some(previous) = context.children.last_mut() {
            match (previous, &entry) {
                (TableStructure::Range(previous), TableStructure::Range(next))
                | (TableStructure::Header(previous), TableStructure::Header(next))
                    if previous.end == next.start =>
                {
                    previous.end = next.end;
                    return Ok(());
                },
                _ => {},
            }
        }
        context.children.push(entry);
        Ok(())
    }

    fn finish(mut self) -> Result<Vec<TableStructure>> {
        if self.contexts.len() != 1 {
            return Err(Error::InvalidFormat(
                "table group is not closed before the table ends".to_string(),
            ));
        }
        let root = self.contexts.pop().expect("one root context was checked");
        if root.header_start.is_some() {
            return Err(Error::InvalidFormat(
                "table header container is not closed before the table ends".to_string(),
            ));
        }
        Ok(root.children)
    }
}

/// Builder for constructing Sheet during parsing
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

    fn with_formatting(
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

    fn set_scenario(&mut self, scenario: SheetScenario) -> Result<()> {
        if self.scenario.replace(scenario).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one scenario".to_string(),
            ));
        }
        Ok(())
    }

    fn add_conditional_format(&mut self, format: ConditionalFormat) -> Result<()> {
        if self.conditional_formats.len() >= MAX_CONDITIONAL_FORMATS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "sheet exceeds the {MAX_CONDITIONAL_FORMATS_PER_SHEET} conditional format safety limit"
            )));
        }
        self.conditional_formats.push(format);
        Ok(())
    }

    fn add_sparkline_group(&mut self, group: SparklineGroup) -> Result<()> {
        if self.sparkline_groups.len() >= MAX_SPARKLINE_GROUPS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "sheet exceeds the {MAX_SPARKLINE_GROUPS_PER_SHEET} sparkline group safety limit"
            )));
        }
        self.sparkline_groups.push(group);
        Ok(())
    }

    fn set_dde_source(&mut self, source: DdeSource) -> Result<()> {
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

    fn set_table_source(&mut self, source: SheetTableSource) -> Result<()> {
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

    fn set_title(&mut self, title: String) -> Result<()> {
        if self.title.replace(title).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one title".to_string(),
            ));
        }
        Ok(())
    }

    fn set_description(&mut self, description: String) -> Result<()> {
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
    fn begin_row_group(&mut self, display: bool) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.begin_group(display)
    }

    fn end_row_group(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.end_group()
    }

    fn begin_row_header(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.begin_header(self.rows.len())
    }

    fn end_row_header(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.end_header(self.rows.len())
    }

    fn add_repeated_row(&mut self, row: Row, repeated: usize) -> Result<()> {
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

    fn add_repeated_column(&mut self, column: Column, repeated: usize) -> Result<()> {
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

    pub fn build(mut self) -> Result<Sheet> {
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
            protection: super::super::SheetProtection::default(),
        })
    }
}

/// Builder for constructing Row during parsing
pub(crate) struct RowBuilder {
    cells: Vec<Cell>,
    repeated: usize,
    style_name: Option<String>,
    default_cell_style_name: Option<String>,
    visibility: TableVisibility,
    /// Number of attribute-free filler cells read but not yet materialised.
    deferred_blank_cells: usize,
    /// The filler cell to clone when the deferred run has to be materialised.
    deferred_blank_cell: Option<Cell>,
}

impl RowBuilder {
    #[cfg(test)]
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            repeated: 1,
            style_name: None,
            default_cell_style_name: None,
            visibility: TableVisibility::Visible,
            deferred_blank_cells: 0,
            deferred_blank_cell: None,
        }
    }

    pub fn add_cell(&mut self, mut cell: Cell) {
        cell.col = self.cells.len();
        self.cells.push(cell);
    }

    fn add_repeated_cells(
        &mut self,
        builder: &CellBuilder,
        text: &str,
        rich_text: Option<&CellTextContent>,
    ) -> Result<()> {
        // Producers pad every row out to the full sheet width with attribute-free
        // `<table:table-cell/>` fillers. Defer those instead of materialising them:
        // an interior run is still expanded when real content follows, but a
        // trailing run is dropped by `build`, which is what makes ordinary
        // spreadsheets fit inside the expansion safety limits at all.
        if builder.is_blank(text, rich_text) {
            self.deferred_blank_cells = self
                .deferred_blank_cells
                .checked_add(builder.repeated)
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "table cell repetition overflows address space".to_string(),
                    )
                })?;
            if self.deferred_blank_cell.is_none() {
                self.deferred_blank_cell = Some(builder.build(text, rich_text));
            }
            return Ok(());
        }
        self.flush_deferred_blank_cells()?;
        let expanded = self
            .cells
            .len()
            .checked_add(builder.repeated)
            .ok_or_else(|| {
                Error::InvalidFormat("table cell repetition overflows address space".to_string())
            })?;
        if expanded > MAX_EXPANDED_CELLS_PER_ROW {
            return Err(Error::InvalidFormat(format!(
                "expanded row exceeds the {MAX_EXPANDED_CELLS_PER_ROW} cell safety limit"
            )));
        }
        for _ in 0..builder.repeated {
            self.add_cell(builder.build(text, rich_text));
        }
        Ok(())
    }

    /// Materialise the deferred blank run because real content follows it, so
    /// the column index of that content stays correct.
    fn flush_deferred_blank_cells(&mut self) -> Result<()> {
        let deferred = std::mem::take(&mut self.deferred_blank_cells);
        let Some(template) = self.deferred_blank_cell.take() else {
            return Ok(());
        };
        if deferred == 0 {
            return Ok(());
        }
        let expanded = self.cells.len().checked_add(deferred).ok_or_else(|| {
            Error::InvalidFormat("table cell repetition overflows address space".to_string())
        })?;
        if expanded > MAX_EXPANDED_CELLS_PER_ROW {
            return Err(Error::InvalidFormat(format!(
                "expanded row exceeds the {MAX_EXPANDED_CELLS_PER_ROW} cell safety limit"
            )));
        }
        self.cells.reserve(deferred);
        for _ in 0..deferred {
            self.add_cell(template.clone());
        }
        Ok(())
    }

    pub fn build(mut self) -> Row {
        // Row index will be set by the parent SheetBuilder
        // For now, set to 0 and update cells
        for cell in &mut self.cells {
            cell.row = 0; // Will be updated by parent
        }

        Row {
            cells: self.cells,
            index: 0, // Will be set by parent
            style_name: self.style_name,
            default_cell_style_name: self.default_cell_style_name,
            visibility: self.visibility,
        }
    }
}

/// Builder for constructing Cell during parsing
pub(crate) struct CellBuilder {
    value_type: Option<String>,
    value_str: Option<String>,
    currency: Option<String>,
    formula: Option<String>,
    validation_name: Option<String>,
    style_name: Option<String>,
    matrix_span: Option<CellMatrixSpan>,
    protect: Option<bool>,
    protected: Option<bool>,
    repeated: usize,
    merge: CellMerge,
    annotation: Option<super::super::Annotation>,
    hyperlinks: Vec<Link>,
    range_source: Option<CellRangeSource>,
    detective: Option<CellDetective>,
}

impl CellBuilder {
    /// Whether this cell carries no user data whatsoever.
    ///
    /// A blank cell is exactly the attribute-free `<table:table-cell/>` filler
    /// producers emit to pad a row out to the full sheet width. Anything that a
    /// user could have authored — a value, formula, style, annotation,
    /// hyperlink, validation, protection flag, merge role, or text — makes the
    /// cell meaningful and therefore not blank.
    fn is_blank(&self, text_content: &str, rich_text: Option<&CellTextContent>) -> bool {
        self.value_type.is_none()
            && self.value_str.is_none()
            && self.currency.is_none()
            && self.formula.is_none()
            && self.validation_name.is_none()
            && self.style_name.is_none()
            && self.matrix_span.is_none()
            && self.protect.is_none()
            && self.protected.is_none()
            && self.annotation.is_none()
            && self.hyperlinks.is_empty()
            && self.range_source.is_none()
            && self.detective.is_none()
            && self.merge == CellMerge::None
            && text_content.is_empty()
            && rich_text.is_none()
    }

    pub fn build(&self, text_content: &str, rich_text: Option<&CellTextContent>) -> Cell {
        let value = self.parse_value(text_content);

        Cell {
            value,
            text: text_content.to_string(),
            // Clone necessary: formula may be reused for repeated cells
            formula: self.formula.clone(),
            annotation: self.annotation.clone(),
            hyperlinks: self.hyperlinks.clone(),
            rich_text: rich_text.cloned(),
            range_source: self.range_source.clone(),
            detective: self.detective.clone(),
            validation_name: self.validation_name.clone(),
            style_name: self.style_name.clone(),
            matrix_span: self.matrix_span,
            merge: self.merge,
            protect: self.protect,
            protected: self.protected,
            row: 0, // Will be set by parent
            col: 0, // Will be set by parent
        }
    }

    fn parse_value(&self, text_content: &str) -> CellValue {
        match self.value_type.as_deref() {
            Some("float") | Some("double") | Some("decimal") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        CellValue::Number(num)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("currency") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        let currency_code = self.currency.as_deref().unwrap_or("USD").to_string();
                        CellValue::Currency(num, currency_code)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("percentage") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        CellValue::Percentage(num)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("boolean") => {
                if let Some(ref val_str) = self.value_str {
                    match val_str.as_str() {
                        "true" => CellValue::Boolean(true),
                        "false" => CellValue::Boolean(false),
                        _ => CellValue::Text(text_content.to_string()),
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("date") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Date(val_str.to_string())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("time") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Time(val_str.to_string())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            _ => {
                if text_content.trim().is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
        }
    }
}
