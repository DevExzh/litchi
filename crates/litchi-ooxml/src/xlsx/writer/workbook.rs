//! Workbook data structure for XLSX.
use crate::pivot::{PivotDataField, PivotFieldRole, PivotTable, PivotValueFunction};
use crate::xlsx::Cell;
use litchi_core::sheet::CellValue;
use litchi_core::sheet::Result as SheetResult;
use litchi_core::xml::escape_xml;
use std::collections::HashMap;
use std::collections::HashSet;
use std::fmt::Write as FmtWrite;

use super::sheet::{MutableWorksheet, NamedRange, Visibility};
use super::strings::MutableSharedStrings;
use super::styles::StylesBuilder;
use crate::xlsx::ProtectionPasswordVerifier;
pub use crate::xlsx::workbook_protection::WorkbookProtectionMetadata as WorkbookProtection;
use crate::xlsx::workbook_protection::write_workbook_protection;
use litchi_xlsx::threaded_comments::People;

/// Type alias for cell position to style index mapping.
type CellStyleMap = HashMap<(u32, u32), usize>;

pub(crate) fn render_pivot_table_sheet_cells(
    pivot: &WritablePivotTable,
    worksheets: &mut [MutableWorksheet],
) -> SheetResult<()> {
    if pivot.row_fields.len() != 1 || pivot.column_fields.len() != 1 || pivot.data_fields.len() != 1
    {
        return Ok(());
    }

    let row_field_name = pivot.row_fields[0].as_str();
    let col_field_name = pivot.column_fields[0].as_str();
    let (data_field_name, data_func) = (&pivot.data_fields[0].0, pivot.data_fields[0].1);
    if matches!(data_func, PivotValueFunction::Custom) {
        return Ok(());
    }

    let row_field_idx = pivot
        .field_names
        .iter()
        .position(|n| n == row_field_name)
        .ok_or_else(|| {
            format!(
                "Pivot row field '{}' not found in field_names",
                row_field_name
            )
        })?;
    let col_field_idx = pivot
        .field_names
        .iter()
        .position(|n| n == col_field_name)
        .ok_or_else(|| {
            format!(
                "Pivot column field '{}' not found in field_names",
                col_field_name
            )
        })?;
    let data_field_idx = pivot
        .field_names
        .iter()
        .position(|n| n == data_field_name)
        .ok_or_else(|| {
            format!(
                "Pivot data field '{}' not found in field_names",
                data_field_name
            )
        })?;

    let source_ws_idx = worksheets
        .iter()
        .position(|ws| ws.name() == pivot.source_sheet)
        .ok_or_else(|| format!("Pivot source sheet '{}' not found", pivot.source_sheet))?;
    let dest_ws_idx = pivot.dest_sheet_index;
    if dest_ws_idx >= worksheets.len() {
        return Err(format!(
            "Pivot destination sheet index {} out of bounds (sheets={})",
            dest_ws_idx,
            worksheets.len()
        )
        .into());
    }

    // Copy out the source data we need first (avoid simultaneous mutable borrows).
    let ((start_row, start_col), (end_row, end_col)) = parse_a1_range(&pivot.source_ref)?;
    let header_row = start_row;
    let data_start_row = header_row + 1;

    let mut row_keys: Vec<String> = Vec::new();
    let mut col_keys: Vec<String> = Vec::new();
    let mut seen_rows: HashSet<String> = HashSet::new();
    let mut seen_cols: HashSet<String> = HashSet::new();

    let mut records: Vec<(String, String, Option<f64>)> = Vec::new();
    {
        let source_ws = &worksheets[source_ws_idx];
        if data_start_row <= end_row {
            for r in data_start_row..=end_row {
                let row_cell = source_ws
                    .cell_value(r, start_col + row_field_idx as u32)
                    .unwrap_or(CellValue::EMPTY);
                let col_cell = source_ws
                    .cell_value(r, start_col + col_field_idx as u32)
                    .unwrap_or(CellValue::EMPTY);
                let data_cell = source_ws
                    .cell_value(r, start_col + data_field_idx as u32)
                    .unwrap_or(CellValue::EMPTY);

                let row_key = match row_cell {
                    CellValue::String(s) => s.clone(),
                    CellValue::Int(i) => i.to_string(),
                    CellValue::Float(f) | CellValue::DateTime(f) => f.to_string(),
                    CellValue::Bool(b) => (if *b { "TRUE" } else { "FALSE" }).to_string(),
                    CellValue::Error(e) => e.clone(),
                    CellValue::Empty => "".to_string(),
                    CellValue::Formula { cached_value, .. } => cached_value
                        .as_deref()
                        .and_then(|cv| match cv {
                            CellValue::String(s) => Some(s.clone()),
                            CellValue::Int(i) => Some(i.to_string()),
                            CellValue::Float(f) | CellValue::DateTime(f) => Some(f.to_string()),
                            CellValue::Bool(b) => {
                                Some((if *b { "TRUE" } else { "FALSE" }).to_string())
                            },
                            CellValue::Error(e) => Some(e.clone()),
                            CellValue::Empty => Some("".to_string()),
                            CellValue::Formula { .. } => None,
                        })
                        .unwrap_or_default(),
                };

                let col_key = match col_cell {
                    CellValue::String(s) => s.clone(),
                    CellValue::Int(i) => i.to_string(),
                    CellValue::Float(f) | CellValue::DateTime(f) => f.to_string(),
                    CellValue::Bool(b) => (if *b { "TRUE" } else { "FALSE" }).to_string(),
                    CellValue::Error(e) => e.clone(),
                    CellValue::Empty => "".to_string(),
                    CellValue::Formula { cached_value, .. } => cached_value
                        .as_deref()
                        .and_then(|cv| match cv {
                            CellValue::String(s) => Some(s.clone()),
                            CellValue::Int(i) => Some(i.to_string()),
                            CellValue::Float(f) | CellValue::DateTime(f) => Some(f.to_string()),
                            CellValue::Bool(b) => {
                                Some((if *b { "TRUE" } else { "FALSE" }).to_string())
                            },
                            CellValue::Error(e) => Some(e.clone()),
                            CellValue::Empty => Some("".to_string()),
                            CellValue::Formula { .. } => None,
                        })
                        .unwrap_or_default(),
                };

                let data_num = match data_cell {
                    CellValue::Int(i) => Some(*i as f64),
                    CellValue::Float(f) | CellValue::DateTime(f) => Some(*f),
                    CellValue::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
                    CellValue::Empty => None,
                    CellValue::String(_) | CellValue::Error(_) => None,
                    CellValue::Formula { cached_value, .. } => {
                        cached_value.as_deref().and_then(|cv| match cv {
                            CellValue::Int(i) => Some(*i as f64),
                            CellValue::Float(f) | CellValue::DateTime(f) => Some(*f),
                            CellValue::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
                            _ => None,
                        })
                    },
                };

                if seen_rows.insert(row_key.clone()) {
                    row_keys.push(row_key.clone());
                }
                if seen_cols.insert(col_key.clone()) {
                    col_keys.push(col_key.clone());
                }

                records.push((row_key, col_key, data_num));
            }
        }
    }

    let rows = row_keys.len();
    let cols = col_keys.len();
    if rows == 0 || cols == 0 {
        return Ok(());
    }

    let mut sums = vec![vec![0.0f64; cols]; rows];
    let mut counts = vec![vec![0.0f64; cols]; rows];
    let mut mins = vec![vec![f64::INFINITY; cols]; rows];
    let mut maxs = vec![vec![f64::NEG_INFINITY; cols]; rows];
    let mut has = vec![vec![false; cols]; rows];
    let mut row_index: HashMap<String, usize> = HashMap::with_capacity(rows);
    let mut col_index: HashMap<String, usize> = HashMap::with_capacity(cols);
    for (i, k) in row_keys.iter().enumerate() {
        row_index.insert(k.clone(), i);
    }
    for (j, k) in col_keys.iter().enumerate() {
        col_index.insert(k.clone(), j);
    }

    for (rk, ck, v) in records {
        let Some(i) = row_index.get(&rk).copied() else {
            continue;
        };
        let Some(j) = col_index.get(&ck).copied() else {
            continue;
        };
        match data_func {
            PivotValueFunction::Sum => {
                if let Some(v) = v {
                    sums[i][j] += v;
                }
            },
            PivotValueFunction::Count => {
                if v.is_some() {
                    sums[i][j] += 1.0;
                }
            },
            PivotValueFunction::Average => {
                if let Some(v) = v {
                    sums[i][j] += v;
                    counts[i][j] += 1.0;
                }
            },
            PivotValueFunction::Min => {
                if let Some(v) = v {
                    has[i][j] = true;
                    mins[i][j] = mins[i][j].min(v);
                }
            },
            PivotValueFunction::Max => {
                if let Some(v) = v {
                    has[i][j] = true;
                    maxs[i][j] = maxs[i][j].max(v);
                }
            },
            PivotValueFunction::Custom => {},
        }
    }

    // Compute final values.
    let mut values = vec![vec![0.0f64; cols]; rows];
    for i in 0..rows {
        for j in 0..cols {
            values[i][j] = match data_func {
                PivotValueFunction::Average => {
                    if counts[i][j] == 0.0 {
                        0.0
                    } else {
                        sums[i][j] / counts[i][j]
                    }
                },
                PivotValueFunction::Min => {
                    if has[i][j] {
                        mins[i][j]
                    } else {
                        0.0
                    }
                },
                PivotValueFunction::Max => {
                    if has[i][j] {
                        maxs[i][j]
                    } else {
                        0.0
                    }
                },
                _ => sums[i][j],
            };
        }
    }

    // Placement
    let start_ref = strip_dollar(&pivot.location_ref);
    let start_ref = start_ref.split(':').next().unwrap_or(start_ref.as_str());
    let (start_col_1, start_row_1) = Cell::reference_to_coords(start_ref)?;

    let dest_ws = &mut worksheets[dest_ws_idx];

    let col_header_row = start_row_1 + 1;
    let data_start_row = start_row_1 + 2;
    let row_label_col = start_col_1;
    let data_start_col = start_col_1 + 1;
    let total_col = data_start_col + cols as u32;
    let grand_total_row = data_start_row + rows as u32;

    // headers
    dest_ws.set_cell_value(start_row_1, row_label_col, row_field_name);
    dest_ws.set_cell_value(start_row_1, row_label_col + 1, col_field_name);
    dest_ws.set_cell_value(start_row_1, data_start_col, data_field_name.as_str());

    // column keys + total header
    for (j, ck) in col_keys.iter().enumerate() {
        dest_ws.set_cell_value(col_header_row, data_start_col + j as u32, ck.as_str());
    }
    dest_ws.set_cell_value(col_header_row, total_col, "Grand Total");

    // row keys + values (+ row totals)
    for (i, rk) in row_keys.iter().enumerate() {
        let out_row = data_start_row + i as u32;
        dest_ws.set_cell_value(out_row, row_label_col, rk.as_str());
        let mut row_total = 0.0f64;
        for (j, v) in values[i].iter().enumerate() {
            row_total += *v;
            dest_ws.set_cell_value(out_row, data_start_col + j as u32, *v);
        }
        dest_ws.set_cell_value(out_row, total_col, row_total);
    }

    // column totals + grand total
    dest_ws.set_cell_value(grand_total_row, row_label_col, "Grand Total");
    let mut grand_total = 0.0f64;
    for (j, _) in col_keys.iter().enumerate() {
        let mut col_total = 0.0f64;
        for row in &values {
            col_total += row[j];
        }
        grand_total += col_total;
        dest_ws.set_cell_value(grand_total_row, data_start_col + j as u32, col_total);
    }
    dest_ws.set_cell_value(grand_total_row, total_col, grand_total);

    let _ = end_col;
    Ok(())
}

/// Escape worksheet names for use in definedName references.
///
/// This doubles single quotes, e.g. "Bob's Sheet" -> "Bob''s Sheet".
fn escape_sheet_name(name: &str) -> String {
    let mut escaped = String::with_capacity(name.len());
    for ch in name.chars() {
        if ch == '\'' {
            escaped.push_str("''");
        } else {
            escaped.push(ch);
        }
    }
    escaped
}

#[derive(Debug, Clone)]
pub struct WritablePivotTable {
    pub name: String,
    pub source_sheet: String,
    pub source_ref: String,
    pub dest_sheet_index: usize,
    pub location_ref: String,
    pub field_names: Vec<String>,
    pub row_fields: Vec<String>,
    pub column_fields: Vec<String>,
    pub filter_fields: Vec<String>,
    pub data_fields: Vec<(String, PivotValueFunction)>,
}

/// Mutable workbook for writing.
///
/// This is managed internally by the Workbook struct.
#[derive(Debug)]
pub struct MutableWorkbookData {
    /// Worksheets
    pub worksheets: Vec<MutableWorksheet>,
    /// Shared strings table
    pub shared_strings: MutableSharedStrings,
    /// Named ranges
    pub named_ranges: Vec<NamedRange>,
    /// Workbook protection
    pub protection: Option<WorkbookProtection>,
    /// Force formula recalculation on open
    pub force_formula_recalculation: bool,
    /// Calculation mode: "auto", "manual", or "autoNoTable"
    pub calculation_mode: String,
    /// Whether the workbook has been modified
    pub modified: bool,
    pub pivot_tables: Vec<WritablePivotTable>,
    /// Person list for threaded comments
    pub person_list: Option<People>,
    /// Chart sheets authored in this workbook, in insertion order
    pub chart_sheets: Vec<super::chart_sheet::MutableChartSheet>,
    /// Sheet slots in workbook order, interleaving worksheets and
    /// chartsheets by insertion position
    pub(crate) sheet_order: Vec<SheetSlot>,
}

/// One slot in the workbook's ordered sheet list.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SheetSlot {
    /// Index into `MutableWorkbookData::worksheets`
    Worksheet(usize),
    /// Index into `MutableWorkbookData::chart_sheets`
    ChartSheet(usize),
}

/// Maximum number of worksheets in one workbook, symmetric with the
/// chartsheet limit (`MAX_CHART_SHEETS`); Excel's practical limit is lower.
pub(crate) const MAX_WORKSHEETS: usize = super::chart_sheet::MAX_CHART_SHEETS;

impl MutableWorkbookData {
    /// Create a new workbook with one default worksheet.
    pub fn new() -> Self {
        let mut data = Self {
            worksheets: Vec::new(),
            shared_strings: MutableSharedStrings::new(),
            named_ranges: Vec::new(),
            protection: None,
            force_formula_recalculation: false,
            calculation_mode: "auto".to_string(),
            modified: false,
            pivot_tables: Vec::new(),
            person_list: None,
            chart_sheets: Vec::new(),
            sheet_order: Vec::new(),
        };

        // Add a default worksheet
        data.add_worksheet("Sheet1".to_string());

        data
    }

    /// Add a new worksheet.
    ///
    /// The name is not validated: invalid or duplicate names produce
    /// workbooks Excel rejects. Prefer [`Self::try_add_worksheet`] for
    /// validated creation.
    pub fn add_worksheet(&mut self, name: String) -> &mut MutableWorksheet {
        let sheet_id = (self.worksheets.len() + self.chart_sheets.len() + 1) as u32;
        let worksheet = MutableWorksheet::new(name, sheet_id);
        self.worksheets.push(worksheet);
        self.sheet_order
            .push(SheetSlot::Worksheet(self.worksheets.len() - 1));
        self.modified = true;
        self.worksheets.last_mut().unwrap()
    }

    /// Add a new worksheet after validating its name.
    ///
    /// The name must satisfy Excel's sheet-name rules (1-31 characters,
    /// none of `: \ / ? * [ ]`) and be unique case-insensitively across
    /// worksheets and chartsheets — the same rules [`Self::add_chart_sheet`]
    /// enforces.
    pub fn try_add_worksheet(&mut self, name: String) -> SheetResult<&mut MutableWorksheet> {
        self.validate_new_sheet_name(&name)?;
        if self.worksheets.len() >= MAX_WORKSHEETS {
            return Err("worksheet count limit exceeded".into());
        }
        Ok(self.add_worksheet(name))
    }

    /// Insert a new worksheet at the requested workbook-order position
    /// (`sheet` inside `sheets`, ECMA-376 §18.2.19 and §18.2.20).
    ///
    /// `index` counts worksheets and chartsheets together in workbook
    /// order, matching the order emitted on save; passing the current
    /// sheet count appends like [`Self::add_worksheet`]. The name is
    /// validated like in [`Self::try_add_worksheet`]. The new sheet
    /// receives a fresh `sheetId`; relationship IDs and part names are
    /// derived from the final sheet order when the workbook is saved.
    ///
    /// # Arguments
    /// * `index` - Zero-based workbook-order position
    /// * `name` - Worksheet name
    pub fn insert_worksheet(
        &mut self,
        index: usize,
        name: String,
    ) -> SheetResult<&mut MutableWorksheet> {
        if index > self.sheet_order.len() {
            return Err(format!(
                "worksheet insertion index {index} out of bounds (max: {})",
                self.sheet_order.len()
            )
            .into());
        }
        self.validate_new_sheet_name(&name)?;
        if self.worksheets.len() >= MAX_WORKSHEETS {
            return Err("worksheet count limit exceeded".into());
        }
        let sheet_id = (self.worksheets.len() + self.chart_sheets.len() + 1) as u32;
        let worksheet = MutableWorksheet::new(name, sheet_id);
        self.worksheets.push(worksheet);
        self.sheet_order
            .insert(index, SheetSlot::Worksheet(self.worksheets.len() - 1));
        self.modified = true;
        Ok(self.worksheets.last_mut().unwrap())
    }

    /// Whether a worksheet or chartsheet with `name` already exists
    /// (case-insensitive, matching Excel's sheet-name uniqueness).
    fn sheet_name_taken(&self, name: &str) -> bool {
        self.worksheets
            .iter()
            .map(|worksheet| worksheet.name())
            .chain(self.chart_sheets.iter().map(|sheet| sheet.name()))
            .any(|existing| existing.eq_ignore_ascii_case(name))
    }

    /// Validate a candidate sheet name: Excel's name rules plus
    /// case-insensitive uniqueness across worksheets and chartsheets.
    fn validate_new_sheet_name(&self, name: &str) -> SheetResult<()> {
        super::chart_sheet::validate_sheet_name(name)?;
        if self.sheet_name_taken(name) {
            return Err(format!("workbook already has a sheet named '{name}'").into());
        }
        Ok(())
    }

    /// Remove a worksheet by its worksheets-relative `index` (matching
    /// [`Self::worksheet_mut`]) and return it.
    ///
    /// The sheet disappears from the `sheets` sequence (ECMA-376 §18.2.20),
    /// so its part, workbook relationship, and content-type override are
    /// simply not emitted on the next save. Sheet-scoped defined names
    /// (`definedName@localSheetId` is the zero-based workbook-order sheet
    /// index, ECMA-376 §18.2.5) are remapped the way LibreOffice resolves
    /// sheet deletion: names scoped to the removed sheet are dropped and
    /// names scoped to later sheets shift one position up. Pivot-table
    /// destination indices shift likewise; removing a sheet that a pivot
    /// table targets is rejected so a report is never silently relocated.
    pub fn remove_worksheet(&mut self, index: usize) -> SheetResult<MutableWorksheet> {
        if index >= self.worksheets.len() {
            return Err(format!(
                "worksheet index {index} out of bounds (max: {})",
                self.worksheets.len().saturating_sub(1)
            )
            .into());
        }
        if self
            .pivot_tables
            .iter()
            .any(|pivot| pivot.dest_sheet_index == index)
        {
            return Err(
                format!("cannot remove worksheet {index}: a pivot table targets it").into(),
            );
        }
        let position = self
            .workbook_position(index)
            .ok_or_else(|| format!("worksheet {index} is missing from the workbook sheet order"))?;

        let removed = self.worksheets.remove(index);
        self.sheet_order.remove(position);
        for slot in &mut self.sheet_order {
            if let SheetSlot::Worksheet(ws_index) = slot
                && *ws_index > index
            {
                *ws_index -= 1;
            }
        }

        let position = position as u32;
        self.remap_sheet_scopes_after_removal(position);
        for pivot in &mut self.pivot_tables {
            if pivot.dest_sheet_index > index {
                pivot.dest_sheet_index -= 1;
            }
        }

        self.modified = true;
        Ok(removed)
    }

    /// Remove a chartsheet by its chartsheets-relative `index` and return it.
    ///
    /// Symmetric to [`Self::remove_worksheet`]: the chartsheet leaves the
    /// `sheets` sequence (ECMA-376 §18.2.20), so its part — and the drawing
    /// and hosted chart parts wired only from it (`chartsheet`,
    /// ECMA-376 §18.3.1.12) — are not emitted on the next save. Defined
    /// names scoped to its workbook position are dropped and later scopes
    /// shift up (`definedName@localSheetId`, ECMA-376 §18.2.5). Pivot
    /// tables never target chartsheets, so no pivot handling is needed.
    pub fn remove_chart_sheet(
        &mut self,
        index: usize,
    ) -> SheetResult<super::chart_sheet::MutableChartSheet> {
        if index >= self.chart_sheets.len() {
            return Err(format!(
                "chartsheet index {index} out of bounds (max: {})",
                self.chart_sheets.len().saturating_sub(1)
            )
            .into());
        }
        let position = self
            .sheet_slot_position(SheetSlot::ChartSheet(index))
            .ok_or_else(|| {
                format!("chartsheet {index} is missing from the workbook sheet order")
            })?;

        let removed = self.chart_sheets.remove(index);
        self.sheet_order.remove(position);
        for slot in &mut self.sheet_order {
            if let SheetSlot::ChartSheet(cs_index) = slot
                && *cs_index > index
            {
                *cs_index -= 1;
            }
        }

        self.remap_sheet_scopes_after_removal(position as u32);
        self.modified = true;
        Ok(removed)
    }

    /// Remap sheet-scoped defined names after the sheet at workbook-order
    /// `position` was removed (`definedName@localSheetId`, ECMA-376 §18.2.5):
    /// names scoped to the removed sheet are dropped and names scoped to
    /// later sheets shift one position up.
    fn remap_sheet_scopes_after_removal(&mut self, position: u32) {
        self.named_ranges
            .retain_mut(|range| match range.local_sheet_id {
                Some(local) if local == position => false,
                Some(local) if local > position => {
                    range.local_sheet_id = Some(local - 1);
                    true
                },
                _ => true,
            });
    }

    /// Workbook-order position of a sheet slot in the `sheets` sequence
    /// (ECMA-376 §18.2.20).
    fn sheet_slot_position(&self, slot: SheetSlot) -> Option<usize> {
        self.sheet_order
            .iter()
            .position(|candidate| *candidate == slot)
    }

    /// Workbook-order position of a worksheet in the `sheets` sequence
    /// (ECMA-376 §18.2.20), counting chartsheets as well.
    fn workbook_position(&self, worksheet_index: usize) -> Option<usize> {
        self.sheet_slot_position(SheetSlot::Worksheet(worksheet_index))
    }

    /// Add a new chartsheet hosting the given chart.
    ///
    /// The name must be a valid Excel sheet name and unique (case-insensitive)
    /// across worksheets and chartsheets. Charts with external-data parts,
    /// user-shapes parts, or additional relationships are rejected; pivot
    /// charts created with `WorksheetChart::into_pivot_chart` are supported
    /// and validated at save time like worksheet pivot charts.
    pub fn add_chart_sheet(
        &mut self,
        name: &str,
        chart: super::sheet::WorksheetChart,
    ) -> SheetResult<&mut super::chart_sheet::MutableChartSheet> {
        super::chart_sheet::validate_chart_sheet_chart(&chart)?;
        self.validate_new_sheet_name(name)?;
        if self.chart_sheets.len() >= super::chart_sheet::MAX_CHART_SHEETS {
            return Err("chartsheet count limit exceeded".into());
        }
        let sheet_id = (self.worksheets.len() + self.chart_sheets.len() + 1) as u32;
        self.chart_sheets
            .push(super::chart_sheet::MutableChartSheet::new(
                name.to_string(),
                sheet_id,
                chart,
            ));
        self.sheet_order
            .push(SheetSlot::ChartSheet(self.chart_sheets.len() - 1));
        self.modified = true;
        Ok(self.chart_sheets.last_mut().unwrap())
    }

    /// Sheet slots in workbook order (worksheets and chartsheets interleaved).
    pub(crate) fn sheet_order(&self) -> &[SheetSlot] {
        &self.sheet_order
    }

    /// Get a worksheet by index.
    pub fn worksheet_mut(&mut self, index: usize) -> SheetResult<&mut MutableWorksheet> {
        self.worksheets
            .get_mut(index)
            .ok_or_else(|| "Worksheet index out of bounds".into())
    }

    /// Get the number of worksheets.
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Define a named range.
    ///
    /// Named ranges allow you to refer to cells or ranges by meaningful names.
    pub fn define_name(&mut self, name: &str, reference: &str) {
        self.named_ranges.push(NamedRange {
            name: name.to_string(),
            reference: reference.to_string(),
            ..NamedRange::default()
        });
        self.modified = true;
    }

    /// Define a sheet-scoped named range.
    ///
    /// Sheet-scoped names are only visible within the specified worksheet.
    pub fn define_name_local(&mut self, name: &str, reference: &str, sheet_id: u32) {
        self.named_ranges.push(NamedRange {
            name: name.to_string(),
            reference: reference.to_string(),
            local_sheet_id: Some(sheet_id),
            ..NamedRange::default()
        });
        self.modified = true;
    }

    /// Define a named range with a comment.
    pub fn define_name_with_comment(&mut self, name: &str, reference: &str, comment: &str) {
        self.named_ranges.push(NamedRange {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: Some(comment.to_string()),
            ..NamedRange::default()
        });
        self.modified = true;
    }

    /// Remove a named range by name.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let initial_len = self.named_ranges.len();
        self.named_ranges.retain(|r| r.name != name);
        let removed = self.named_ranges.len() < initial_len;
        if removed {
            self.modified = true;
        }
        removed
    }

    /// Get all named ranges.
    pub fn named_ranges(&self) -> &[NamedRange] {
        &self.named_ranges
    }

    pub fn add_pivot_table(&mut self, pivot: PivotTable) -> SheetResult<()> {
        let source_sheet = pivot
            .source_sheet
            .clone()
            .ok_or_else(|| "PivotTable.source_sheet is required for writing".to_string())?;
        let source_ref = pivot
            .source_ref
            .clone()
            .ok_or_else(|| "PivotTable.source_ref is required for writing".to_string())?;
        if pivot.field_names.is_empty() {
            return Err("PivotTable.field_names is required for writing".into());
        }

        let dest_sheet_index = self
            .worksheets
            .iter()
            .position(|ws| ws.name() == pivot.sheet_name)
            .ok_or_else(|| format!("Pivot destination sheet '{}' not found", pivot.sheet_name))?;

        let mut row_fields: Vec<PivotFieldRole> = pivot.row_fields.clone();
        row_fields.sort_by_key(|r| r.position);
        let mut column_fields: Vec<PivotFieldRole> = pivot.column_fields.clone();
        column_fields.sort_by_key(|r| r.position);
        let mut filter_fields: Vec<PivotFieldRole> = pivot.filter_fields.clone();
        filter_fields.sort_by_key(|r| r.position);

        let row_fields: Vec<String> = row_fields.into_iter().map(|r| r.field_name).collect();
        let column_fields: Vec<String> = column_fields.into_iter().map(|r| r.field_name).collect();
        let filter_fields: Vec<String> = filter_fields.into_iter().map(|r| r.field_name).collect();

        let data_fields: Vec<(String, PivotValueFunction)> = pivot
            .data_fields
            .iter()
            .map(|df: &PivotDataField| (df.field_name.clone(), df.function))
            .collect();

        self.pivot_tables.push(WritablePivotTable {
            name: pivot.name,
            source_sheet,
            source_ref,
            dest_sheet_index,
            location_ref: pivot.location_ref,
            field_names: pivot.field_names,
            row_fields,
            column_fields,
            filter_fields,
            data_fields,
        });
        self.modified = true;
        Ok(())
    }

    /// Synchronize worksheet print settings with internal defined names.
    ///
    /// This maps worksheet-level page setup (print area, repeating rows/columns)
    /// to the Excel-reserved sheet-scoped defined names `_xlnm.Print_Area` and
    /// `_xlnm.Print_Titles` in `workbook.xml`.
    pub(crate) fn sync_print_settings_to_defined_names(&mut self) {
        // Preserve all existing named ranges except the internal print-related
        // ones for each sheet, which we rebuild from the current worksheet
        // settings.
        let mut new_ranges =
            Vec::with_capacity(self.named_ranges.len() + self.worksheets.len() * 2);

        for range in &self.named_ranges {
            let is_internal_print_name = (range.name == "_xlnm.Print_Area"
                || range.name == "_xlnm.Print_Titles")
                && range.local_sheet_id.is_some();

            if !is_internal_print_name {
                new_ranges.push(range.clone());
            }
        }

        // Rebuild per-sheet print area and print titles based on the
        // MutableWorksheet settings. localSheetId is the zero-based
        // workbook-order sheet index (ECMA-376 §18.2.5), which counts
        // chartsheets and shifts when sheets are inserted or removed.
        for (ws_index, ws) in self.worksheets.iter().enumerate() {
            let Some(position) = self.workbook_position(ws_index) else {
                continue;
            };
            let position = position as u32;
            let sheet_name = ws.name();
            let escaped_sheet_name = escape_sheet_name(sheet_name);

            // Print area -> _xlnm.Print_Area
            if let Some(range) = ws.get_print_area() {
                let reference = format!("'{}'!{}", escaped_sheet_name, range);
                new_ranges.push(NamedRange {
                    name: "_xlnm.Print_Area".to_string(),
                    reference,
                    local_sheet_id: Some(position),
                    ..NamedRange::default()
                });
            }

            // Repeating rows/columns -> _xlnm.Print_Titles
            let repeating_cols = ws.get_repeating_columns();
            let repeating_rows = ws.get_repeating_rows();

            if repeating_cols.is_some() || repeating_rows.is_some() {
                let mut parts = Vec::new();

                if let Some(cols) = repeating_cols {
                    parts.push(format!("'{}'!{}", escaped_sheet_name, cols));
                }
                if let Some(rows) = repeating_rows {
                    parts.push(format!("'{}'!{}", escaped_sheet_name, rows));
                }

                let reference = parts.join(",");
                new_ranges.push(NamedRange {
                    name: "_xlnm.Print_Titles".to_string(),
                    reference,
                    local_sheet_id: Some(position),
                    ..NamedRange::default()
                });
            }
        }

        self.named_ranges = new_ranges;
        self.modified = true;
    }

    /// Check if the workbook has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified || self.worksheets.iter().any(|ws| ws.is_modified())
    }

    /// Build styles from all worksheets and return a StylesBuilder and cell position -> style index mappings.
    ///
    /// Returns a tuple of (StylesBuilder, Vec of per-worksheet CellStyleMap).
    pub fn build_styles(&self) -> SheetResult<(StylesBuilder, Vec<CellStyleMap>)> {
        let mut builder = StylesBuilder::new();
        let mut worksheet_style_indices = Vec::new();

        // For each worksheet, collect cell formats and build style indices
        for ws in &self.worksheets {
            let mut style_map = CellStyleMap::new();

            // Iterate through all cells with formats
            for (pos, format) in ws.cell_formats() {
                // Add format to builder and get its style index
                let style_index = builder.add_cell_format(format);
                style_map.insert(*pos, style_index);
            }

            worksheet_style_indices.push(style_map);
        }

        Ok((builder, worksheet_style_indices))
    }

    /// Generate workbook.xml content with actual relationship IDs.
    ///
    /// # Arguments
    /// * `worksheet_rel_ids` - Vector of relationship IDs for worksheets (e.g., ["rId1", "rId2", ...])
    ///
    /// Retained for tests; the save pipeline uses
    /// [`Self::generate_workbook_xml_ordered`].
    #[cfg(test)]
    pub(crate) fn generate_workbook_xml_with_external_rels(
        &self,
        worksheet_rel_ids: &[String],
        pivot_cache_rel_ids: &[(u32, String)],
        external_reference_ids: &[String],
    ) -> SheetResult<String> {
        let sheet_entries: Vec<(String, u32, String)> = self
            .worksheets
            .iter()
            .enumerate()
            .map(|(index, ws)| {
                let rel_id = worksheet_rel_ids
                    .get(index)
                    .map(|s| s.as_str())
                    .unwrap_or("rId1"); // Fallback, shouldn't happen
                (ws.name().to_string(), ws.sheet_id(), rel_id.to_string())
            })
            .collect();
        self.generate_workbook_xml_ordered(
            &sheet_entries,
            pivot_cache_rel_ids,
            external_reference_ids,
        )
    }

    /// Generate workbook.xml with sheet entries given in workbook order as
    /// `(name, sheet_id, relationship_id)` triples, covering both worksheets
    /// and chartsheets.
    pub(crate) fn generate_workbook_xml_ordered(
        &self,
        sheet_entries: &[(String, u32, String)],
        pivot_cache_rel_ids: &[(u32, String)],
        external_reference_ids: &[String],
    ) -> SheetResult<String> {
        let mut xml = String::with_capacity(2048);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);

        xml.push_str(r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);

        // Add fileVersion (recommended by Excel for compatibility)
        xml.push_str(
            r#"<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="16925"/>"#,
        );

        // Add workbookPr (required by Excel)
        xml.push_str(r#"<workbookPr defaultThemeVersion="166925" hidePivotFieldList="0"/>"#);

        // Write workbook protection if configured (must come after workbookPr per OOXML spec)
        if let Some(ref protection) = self.protection {
            xml.push_str(
                &write_workbook_protection(protection).map_err(|error| error.to_string())?,
            );
        }

        // Add bookViews (required by Excel)
        xml.push_str("<bookViews>");
        let active_tab = self
            .worksheets
            .iter()
            .position(|ws| ws.is_active())
            .unwrap_or(0);
        write!(
            xml,
            r#"<workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="10000" activeTab="{}" uid="{{00000000-0000-0000-0000-000000000000}}"/>"#,
            active_tab
        )
        .map_err(|e| format!("XML write error: {}", e))?;
        xml.push_str("</bookViews>");

        xml.push_str("<sheets>");
        for (name, sheet_id, rel_id) in sheet_entries {
            let visibility = self
                .worksheets
                .iter()
                .find(|worksheet| worksheet.sheet_id() == *sheet_id)
                .map(MutableWorksheet::visibility)
                .unwrap_or_default();

            // Sheet elements are always self-closing (tab colors are in worksheet XML, not here)
            if visibility == Visibility::Visible {
                write!(
                    xml,
                    r#"<sheet name="{}" sheetId="{}" r:id="{}"/>"#,
                    escape_xml(name),
                    sheet_id,
                    rel_id
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            } else {
                write!(
                    xml,
                    r#"<sheet name="{}" sheetId="{}" state="{}" r:id="{}"/>"#,
                    escape_xml(name),
                    sheet_id,
                    visibility.as_str(),
                    rel_id
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            }
        }
        xml.push_str("</sheets>");

        if !external_reference_ids.is_empty() {
            xml.push_str("<externalReferences>");
            for relationship_id in external_reference_ids {
                write!(
                    xml,
                    r#"<externalReference r:id="{}"/>"#,
                    escape_xml(relationship_id)
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            }
            xml.push_str("</externalReferences>");
        }

        // Write defined names (named ranges)
        if !self.named_ranges.is_empty() {
            xml.push_str("<definedNames>");
            for named_range in &self.named_ranges {
                xml.push_str("<definedName name=\"");
                xml.push_str(&escape_xml(&named_range.name));
                xml.push('"');

                // Add localSheetId if it's a sheet-scoped name
                if let Some(sheet_id) = named_range.local_sheet_id {
                    write!(xml, " localSheetId=\"{}\"", sheet_id)
                        .map_err(|e| format!("XML write error: {}", e))?;
                }

                // Add comment if present
                if let Some(ref comment) = named_range.comment {
                    write!(xml, " comment=\"{}\"", escape_xml(comment))
                        .map_err(|e| format!("XML write error: {}", e))?;
                }

                write_optional_defined_name_attribute(
                    &mut xml,
                    "customMenu",
                    named_range.custom_menu.as_deref(),
                )?;
                write_optional_defined_name_attribute(
                    &mut xml,
                    "description",
                    named_range.description.as_deref(),
                )?;
                write_optional_defined_name_attribute(
                    &mut xml,
                    "help",
                    named_range.help.as_deref(),
                )?;
                write_optional_defined_name_attribute(
                    &mut xml,
                    "statusBar",
                    named_range.status_bar.as_deref(),
                )?;
                write_optional_defined_name_attribute(
                    &mut xml,
                    "shortcutKey",
                    named_range.shortcut_key.as_deref(),
                )?;
                write_defined_name_bool(&mut xml, "hidden", named_range.hidden)?;
                write_defined_name_bool(&mut xml, "function", named_range.function)?;
                write_defined_name_bool(&mut xml, "vbProcedure", named_range.vb_procedure)?;
                write_defined_name_bool(&mut xml, "xlm", named_range.xlm)?;
                if let Some(function_group_id) = named_range.function_group_id {
                    write!(xml, " functionGroupId=\"{}\"", function_group_id)
                        .map_err(|e| format!("XML write error: {}", e))?;
                }
                write_defined_name_bool(
                    &mut xml,
                    "publishToServer",
                    named_range.publish_to_server,
                )?;
                write_defined_name_bool(
                    &mut xml,
                    "workbookParameter",
                    named_range.workbook_parameter,
                )?;

                xml.push('>');
                xml.push_str(&escape_xml(&named_range.reference));
                xml.push_str("</definedName>");
            }
            xml.push_str("</definedNames>");
        }

        // Add calculation properties (recommended for Excel compatibility)
        xml.push_str(r#"<calcPr calcId="171027"/>"#);

        if !pivot_cache_rel_ids.is_empty() {
            xml.push_str("<pivotCaches>");
            for (cache_id, rel_id) in pivot_cache_rel_ids {
                write!(
                    xml,
                    r#"<pivotCache cacheId="{}" r:id="{}"/>"#,
                    cache_id, rel_id
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            }
            xml.push_str("</pivotCaches>");
        }

        xml.push_str("</workbook>");

        Ok(xml)
    }

    // ===== Workbook-level Features =====

    /// Hide a worksheet by index.
    pub fn hide_sheet(&mut self, index: usize) -> SheetResult<()> {
        if let Some(ws) = self.worksheets.get_mut(index) {
            ws.set_hidden(true);
            self.modified = true;
            Ok(())
        } else {
            Err("Worksheet index out of bounds".into())
        }
    }

    /// Unhide a worksheet by index.
    pub fn unhide_sheet(&mut self, index: usize) -> SheetResult<()> {
        if let Some(ws) = self.worksheets.get_mut(index) {
            ws.set_hidden(false);
            self.modified = true;
            Ok(())
        } else {
            Err("Worksheet index out of bounds".into())
        }
    }

    /// Check if a worksheet is hidden.
    pub fn is_sheet_hidden(&self, index: usize) -> Option<bool> {
        self.worksheets.get(index).map(|ws| ws.is_hidden())
    }

    /// Move a worksheet to a new position.
    pub fn move_sheet(&mut self, from_index: usize, to_index: usize) -> SheetResult<()> {
        if from_index >= self.worksheets.len() || to_index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        let worksheet = self.worksheets.remove(from_index);
        self.worksheets.insert(to_index, worksheet);
        self.modified = true;
        Ok(())
    }

    /// Set sheet visibility state.
    pub fn set_sheet_visibility(
        &mut self,
        index: usize,
        visibility: Visibility,
    ) -> SheetResult<()> {
        if let Some(ws) = self.worksheets.get_mut(index) {
            ws.set_visibility(visibility);
            self.modified = true;
            Ok(())
        } else {
            Err("Worksheet index out of bounds".into())
        }
    }

    /// Get sheet visibility state.
    pub fn get_sheet_visibility(&self, index: usize) -> Option<Visibility> {
        self.worksheets.get(index).map(|ws| ws.visibility())
    }

    /// Set the active worksheet index.
    pub fn set_active_sheet(&mut self, index: usize) {
        // Mark all worksheets as not active
        for ws in &mut self.worksheets {
            ws.set_active(false);
        }

        // Set the specified worksheet as active
        if let Some(ws) = self.worksheets.get_mut(index) {
            ws.set_active(true);
            self.modified = true;
        }
    }

    /// Force formula recalculation when the workbook is opened.
    pub fn set_force_formula_recalculation(&mut self, force: bool) {
        self.force_formula_recalculation = force;
        self.modified = true;
    }

    /// Get whether formula recalculation is forced.
    pub fn get_force_formula_recalculation(&self) -> bool {
        self.force_formula_recalculation
    }

    /// Set the calculation mode for the workbook.
    pub fn set_calculation_mode(&mut self, mode: &str) {
        self.calculation_mode = mode.to_string();
        self.modified = true;
    }

    /// Get the calculation mode for the workbook.
    pub fn get_calculation_mode(&self) -> Option<&str> {
        Some(&self.calculation_mode)
    }

    /// Protect the workbook with optional password.
    ///
    /// # Arguments
    /// * `password` - Optional password (will be hashed)
    /// * `lock_structure` - Prevent adding/deleting sheets
    /// * `lock_windows` - Prevent resizing/moving workbook window
    pub fn protect_workbook(
        &mut self,
        password: Option<&str>,
        lock_structure: bool,
        lock_windows: bool,
    ) {
        use super::sheet::MutableWorksheet;

        let mut protection = WorkbookProtection::new();
        protection.set_workbook_verifier(password.map(|value| {
            ProtectionPasswordVerifier::Legacy(MutableWorksheet::hash_password(value))
        }));
        protection.set_structure_locked(lock_structure);
        protection.set_windows_locked(lock_windows);
        self.protection = Some(protection);
        self.modified = true;
    }

    /// Replace the complete typed workbook-protection configuration.
    ///
    /// This accepts caller-supplied legacy or strong verifier metadata for both
    /// workbook and revision locks. It never derives or checks a password.
    pub fn set_workbook_protection(&mut self, protection: WorkbookProtection) {
        self.protection = Some(protection);
        self.modified = true;
    }

    /// Unprotect the workbook.
    pub fn unprotect_workbook(&mut self) {
        self.protection = None;
        self.modified = true;
    }

    /// Check if the workbook is protected.
    pub fn is_protected(&self) -> bool {
        self.protection.is_some()
    }

    /// Get the workbook protection configuration.
    pub fn get_protection(&self) -> Option<&WorkbookProtection> {
        self.protection.as_ref()
    }
}

fn write_optional_defined_name_attribute(
    xml: &mut String,
    name: &str,
    value: Option<&str>,
) -> SheetResult<()> {
    if let Some(value) = value {
        write!(xml, " {name}=\"{}\"", escape_xml(value))
            .map_err(|error| format!("XML write error: {error}"))?;
    }
    Ok(())
}

fn write_defined_name_bool(xml: &mut String, name: &str, value: bool) -> SheetResult<()> {
    if value {
        write!(xml, " {name}=\"1\"").map_err(|error| format!("XML write error: {error}"))?;
    }
    Ok(())
}

impl Default for MutableWorkbookData {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, Default)]
pub(crate) struct PivotCacheFieldStats {
    contains_blank: bool,
    contains_string: bool,
    contains_number: bool,
    contains_integer: bool,
    min_value: Option<f64>,
    max_value: Option<f64>,
    string_items: Vec<String>,
    string_index: HashMap<String, u32>,
    number_items: Vec<String>,
    number_index: HashMap<String, u32>,
    bool_items: Vec<bool>,
    bool_index: HashMap<bool, u32>,
    has_mixed_types: bool,
}

impl PivotCacheFieldStats {
    fn observe(&mut self, v: &CellValue) {
        match v {
            CellValue::Empty => {
                self.contains_blank = true;
            },
            CellValue::Bool(b) => {
                if self.contains_string || self.contains_number {
                    self.has_mixed_types = true;
                }
                self.index_bool(*b);
            },
            CellValue::Int(i) => {
                if self.contains_string {
                    self.has_mixed_types = true;
                }
                self.contains_number = true;
                self.contains_integer = true;
                self.index_number(&i.to_string());
                let f = *i as f64;
                self.min_value = Some(self.min_value.map_or(f, |m| m.min(f)));
                self.max_value = Some(self.max_value.map_or(f, |m| m.max(f)));
            },
            CellValue::Float(f) | CellValue::DateTime(f) => {
                if self.contains_string {
                    self.has_mixed_types = true;
                }
                if !self.contains_number {
                    // If we only see whole-number floats, Excel considers it integer.
                    self.contains_integer = true;
                }
                self.contains_number = true;
                self.index_number(&f.to_string());
                if f.fract() != 0.0 {
                    self.contains_integer = false;
                }
                self.min_value = Some(self.min_value.map_or(*f, |m| m.min(*f)));
                self.max_value = Some(self.max_value.map_or(*f, |m| m.max(*f)));
            },
            CellValue::String(s) => {
                if self.contains_number {
                    self.has_mixed_types = true;
                }
                self.contains_string = true;
                self.index_string(s);
            },
            CellValue::Error(e) => {
                if self.contains_number {
                    self.has_mixed_types = true;
                }
                self.contains_string = true;
                self.index_string(e);
            },
            CellValue::Formula { cached_value, .. } => {
                if let Some(v) = cached_value.as_deref() {
                    self.observe(v);
                } else {
                    self.contains_blank = true;
                }
            },
        }
    }

    fn index_string(&mut self, s: &str) -> u32 {
        if let Some(idx) = self.string_index.get(s) {
            return *idx;
        }
        let idx = self.string_items.len() as u32;
        self.string_items.push(s.to_string());
        self.string_index.insert(s.to_string(), idx);
        idx
    }

    fn index_bool(&mut self, b: bool) -> u32 {
        if let Some(idx) = self.bool_index.get(&b) {
            return *idx;
        }
        let idx = self.bool_items.len() as u32;
        self.bool_items.push(b);
        self.bool_index.insert(b, idx);
        idx
    }

    fn index_number(&mut self, s: &str) -> u32 {
        if let Some(idx) = self.number_index.get(s) {
            return *idx;
        }
        let idx = self.number_items.len() as u32;
        self.number_items.push(s.to_string());
        self.number_index.insert(s.to_string(), idx);
        idx
    }
}

fn strip_dollar(s: &str) -> String {
    s.chars().filter(|&c| c != '$').collect()
}

fn parse_a1_range(range_ref: &str) -> SheetResult<((u32, u32), (u32, u32))> {
    let range_ref = strip_dollar(range_ref);
    let (start_ref, end_ref) = match range_ref.split_once(':') {
        Some((a, b)) => (a, b),
        None => (range_ref.as_str(), range_ref.as_str()),
    };

    let (start_col_1, start_row_1) = Cell::reference_to_coords(start_ref)?;
    let (end_col_1, end_row_1) = Cell::reference_to_coords(end_ref)?;

    let mut start_row = start_row_1.saturating_sub(1);
    let mut start_col = start_col_1.saturating_sub(1);
    let mut end_row = end_row_1.saturating_sub(1);
    let mut end_col = end_col_1.saturating_sub(1);

    if start_row > end_row {
        std::mem::swap(&mut start_row, &mut end_row);
    }
    if start_col > end_col {
        std::mem::swap(&mut start_col, &mut end_col);
    }

    Ok(((start_row, start_col), (end_row, end_col)))
}

fn write_pivot_record_value_indexed(
    xml: &mut String,
    v: &CellValue,
    stats: &mut PivotCacheFieldStats,
) -> SheetResult<()> {
    match v {
        CellValue::Empty => {
            stats.observe(v);
            if stats.contains_number {
                xml.push_str("<m/>");
                return Ok(());
            }

            // Reference the missing item in sharedItems if present; otherwise fall back to <m/>.
            // We always place <m/> at the end when contains_blank is true.
            let idx = if stats.contains_string {
                stats.string_items.len() as u32
            } else if !stats.bool_items.is_empty() {
                stats.bool_items.len() as u32
            } else {
                // no sharedItems list
                xml.push_str("<m/>");
                return Ok(());
            };
            write!(xml, r#"<x v="{}"/>"#, idx).map_err(|e| format!("XML write error: {}", e))?;
        },
        CellValue::Bool(b) => {
            let idx = stats.index_bool(*b);
            stats.observe(v);
            write!(xml, r#"<x v="{}"/>"#, idx).map_err(|e| format!("XML write error: {}", e))?;
        },
        CellValue::Int(i) => {
            stats.observe(v);
            write!(xml, r#"<n v="{}"/>"#, i).map_err(|e| format!("XML write error: {}", e))?;
        },
        CellValue::Float(f) | CellValue::DateTime(f) => {
            stats.observe(v);
            write!(xml, r#"<n v="{}"/>"#, f).map_err(|e| format!("XML write error: {}", e))?;
        },
        CellValue::String(s) => {
            let idx = stats.index_string(s);
            stats.observe(v);
            write!(xml, r#"<x v="{}"/>"#, idx).map_err(|e| format!("XML write error: {}", e))?;
        },
        CellValue::Error(e) => {
            let idx = stats.index_string(e);
            stats.observe(v);
            write!(xml, r#"<x v="{}"/>"#, idx).map_err(|e| format!("XML write error: {}", e))?;
        },
        CellValue::Formula { cached_value, .. } => {
            if let Some(v) = cached_value.as_deref() {
                write_pivot_record_value_indexed(xml, v, stats)?;
            } else {
                stats.observe(&CellValue::Empty);
                xml.push_str("<m/>");
            }
        },
    }
    Ok(())
}

pub(crate) fn generate_pivot_cache_definition_xml(
    pivot: &WritablePivotTable,
    records_rel_id: Option<&str>,
    record_count: u32,
    field_stats: &[PivotCacheFieldStats],
) -> SheetResult<String> {
    let mut xml = String::with_capacity(768);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision""#);
    if let Some(rel_id) = records_rel_id {
        write!(xml, r#" r:id="{}""#, rel_id).map_err(|e| format!("XML write error: {}", e))?;
    }
    xml.push_str(r#" refreshedBy="litchi" refreshedDate="0" createdVersion="4" refreshedVersion="4" minRefreshableVersion="3""#);
    write!(xml, r#" recordCount="{}""#, record_count)
        .map_err(|e| format!("XML write error: {}", e))?;
    xml.push_str(r#" xr:uid="{00000000-0000-0000-0000-000000000000}""#);
    xml.push('>');
    xml.push_str(r#"<cacheSource type="worksheet">"#);
    write!(
        xml,
        r#"<worksheetSource ref="{}" sheet="{}"/>"#,
        escape_xml(&pivot.source_ref),
        escape_xml(&pivot.source_sheet)
    )
    .map_err(|e| format!("XML write error: {}", e))?;
    xml.push_str("</cacheSource>");
    write!(xml, r#"<cacheFields count="{}">"#, pivot.field_names.len())
        .map_err(|e| format!("XML write error: {}", e))?;

    for (idx, name) in pivot.field_names.iter().enumerate() {
        let stats = field_stats.get(idx).cloned().unwrap_or_default();
        write!(
            xml,
            r#"<cacheField name="{}" numFmtId="0">"#,
            escape_xml(name)
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        if stats.contains_number {
            xml.push_str(r#"<sharedItems"#);
            if stats.has_mixed_types {
                xml.push_str(r#" containsSemiMixedTypes="1""#);
            } else {
                xml.push_str(r#" containsSemiMixedTypes="0""#);
            }
            xml.push_str(r#" containsString="0" containsNumber="1""#);
            if stats.contains_integer {
                xml.push_str(r#" containsInteger="1""#);
            } else {
                xml.push_str(r#" containsInteger="0""#);
            }
            if let Some(min_v) = stats.min_value {
                write!(xml, r#" minValue="{}""#, min_v)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(max_v) = stats.max_value {
                write!(xml, r#" maxValue="{}""#, max_v)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            xml.push_str("/>");
        } else if stats.contains_string {
            let shared_count = stats.string_items.len() + if stats.contains_blank { 1 } else { 0 };
            write!(xml, r#"<sharedItems count="{}">"#, shared_count)
                .map_err(|e| format!("XML write error: {}", e))?;
            for s in &stats.string_items {
                write!(xml, r#"<s v="{}"/>"#, escape_xml(s))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if stats.contains_blank {
                xml.push_str("<m/>");
            }
            xml.push_str("</sharedItems>");
        } else if !stats.bool_items.is_empty() {
            xml.push_str("<sharedItems");
            let shared_count = stats.bool_items.len() + if stats.contains_blank { 1 } else { 0 };
            write!(xml, r#" count="{}""#, shared_count)
                .map_err(|e| format!("XML write error: {}", e))?;
            xml.push('>');
            for b in &stats.bool_items {
                write!(xml, r#"<b v="{}"/>"#, if *b { 1 } else { 0 })
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if stats.contains_blank {
                xml.push_str("<m/>");
            }
            xml.push_str("</sharedItems>");
        } else {
            xml.push_str(r#"<sharedItems count="0"/>"#);
        }

        xml.push_str("</cacheField>");
    }

    xml.push_str("</cacheFields>");

    // Common x14 extension block present in many Excel-generated files.
    xml.push_str(r#"<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}"><x14:pivotCacheDefinition/></ext></extLst>"#);
    xml.push_str("</pivotCacheDefinition>");
    Ok(xml)
}

pub(crate) fn generate_pivot_cache_records_xml(
    pivot: &WritablePivotTable,
    worksheets: &[MutableWorksheet],
) -> SheetResult<(String, u32, Vec<PivotCacheFieldStats>)> {
    let source_ws = worksheets
        .iter()
        .find(|ws| ws.name() == pivot.source_sheet)
        .ok_or_else(|| format!("Pivot source sheet '{}' not found", pivot.source_sheet))?;

    let ((start_row, start_col), (end_row, end_col)) = parse_a1_range(&pivot.source_ref)?;
    let col_count = (end_col - start_col + 1) as usize;
    if col_count != pivot.field_names.len() {
        return Err(format!(
            "Pivot field_names len {} does not match source_ref column count {} ({})",
            pivot.field_names.len(),
            col_count,
            pivot.source_ref
        )
        .into());
    }

    let mut field_stats = vec![PivotCacheFieldStats::default(); col_count];

    // First row is assumed to be headers. Records are the data rows.
    let data_start_row = start_row + 1;

    let mut xml = String::with_capacity(1024);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision""#);

    let mut record_count: u32 = 0;
    let mut records_body = String::with_capacity(1024);

    if data_start_row <= end_row {
        for row in data_start_row..=end_row {
            records_body.push_str("<r>");
            for (i, col) in (start_col..=end_col).enumerate() {
                let v = source_ws.cell_value(row, col).unwrap_or(CellValue::EMPTY);
                write_pivot_record_value_indexed(&mut records_body, v, &mut field_stats[i])?;
            }
            records_body.push_str("</r>");
            record_count += 1;
        }
    }

    write!(xml, r#" count="{}">"#, record_count).map_err(|e| format!("XML write error: {}", e))?;
    xml.push_str(&records_body);
    xml.push_str("</pivotCacheRecords>");

    Ok((xml, record_count, field_stats))
}

pub(crate) fn generate_pivot_table_definition_xml(
    pivot: &WritablePivotTable,
    cache_id: u32,
    field_stats: &[PivotCacheFieldStats],
) -> SheetResult<String> {
    let mut xml = String::with_capacity(1024);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{00000000-0000-0000-0000-000000000000}""#);
    write!(
        xml,
        r#" name="{}" cacheId="{}""#,
        escape_xml(&pivot.name),
        cache_id
    )
    .map_err(|e| format!("XML write error: {}", e))?;
    xml.push_str(r#" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="4" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="4" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">"#);

    let field_index = build_field_index(&pivot.field_names);
    let row_indexes = resolve_field_indexes(&field_index, &pivot.row_fields);
    let col_indexes = resolve_field_indexes(&field_index, &pivot.column_fields);
    let filter_indexes = resolve_field_indexes(&field_index, &pivot.filter_fields);

    let shared_count_for = |idx: u32| -> u32 {
        field_stats
            .get(idx as usize)
            .map(|s| {
                if s.contains_string {
                    s.string_items.len() as u32 + if s.contains_blank { 1 } else { 0 }
                } else if s.contains_number {
                    s.number_items.len() as u32 + if s.contains_blank { 1 } else { 0 }
                } else if !s.bool_items.is_empty() {
                    s.bool_items.len() as u32 + if s.contains_blank { 1 } else { 0 }
                } else {
                    0
                }
            })
            .unwrap_or(0)
    };

    // Excel typically expects a location range, not a single cell.
    // For the common simple case (1 row field, 1 col field, 1 data field):
    // width = 2 (row labels + grand total) + colSharedCount
    // height = 2 (header + grand total) + rowSharedCount
    let start_ref = strip_dollar(&pivot.location_ref);
    let start_ref = start_ref.split(':').next().unwrap_or(start_ref.as_str());
    let (start_col_1, start_row_1) = Cell::reference_to_coords(start_ref)?;

    let row_shared = row_indexes
        .first()
        .copied()
        .map(shared_count_for)
        .unwrap_or(0);
    let col_shared = col_indexes
        .first()
        .copied()
        .map(shared_count_for)
        .unwrap_or(0);
    let width = 2 + col_shared;
    let height = 2 + row_shared;

    let end_col_1 = start_col_1 + width.saturating_sub(1);
    let end_row_1 = start_row_1 + height.saturating_sub(1);
    let location_ref = format!(
        "{}{}:{}{}",
        Cell::column_to_letters(start_col_1),
        start_row_1,
        Cell::column_to_letters(end_col_1),
        end_row_1
    );

    write!(
        xml,
        r#"<location ref="{}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>"#,
        escape_xml(&location_ref)
    )
    .map_err(|e| format!("XML write error: {}", e))?;

    let mut data_fields = Vec::new();
    for (name, func) in &pivot.data_fields {
        if let Some(idx) = field_index.get(name.as_str()) {
            data_fields.push((*idx, *func));
        }
    }

    // pivotFields must appear after location and should reflect axes.
    let field_count = pivot.field_names.len();
    let mut axis_by_field: Vec<Option<&'static str>> = vec![None; field_count];
    for idx in &row_indexes {
        if let Some(slot) = axis_by_field.get_mut(*idx as usize) {
            *slot = Some("axisRow");
        }
    }
    for idx in &col_indexes {
        if let Some(slot) = axis_by_field.get_mut(*idx as usize) {
            *slot = Some("axisCol");
        }
    }
    for idx in &filter_indexes {
        if let Some(slot) = axis_by_field.get_mut(*idx as usize) {
            *slot = Some("axisPage");
        }
    }

    let mut is_data_field: Vec<bool> = vec![false; field_count];
    for (idx, _) in &data_fields {
        if let Some(slot) = is_data_field.get_mut(*idx as usize) {
            *slot = true;
        }
    }

    write!(xml, r#"<pivotFields count="{}">"#, field_count)
        .map_err(|e| format!("XML write error: {}", e))?;
    for i in 0..field_count {
        xml.push_str("<pivotField");
        if let Some(axis) = axis_by_field[i] {
            write!(xml, r#" axis="{}""#, axis).map_err(|e| format!("XML write error: {}", e))?;
        }
        if is_data_field[i] {
            xml.push_str(r#" dataField="1""#);
        }

        xml.push_str(r#" showAll="0""#);

        let needs_items = axis_by_field[i].is_some();
        if needs_items {
            let shared_count = shared_count_for(i as u32);
            xml.push('>');
            write!(xml, r#"<items count="{}">"#, shared_count + 1)
                .map_err(|e| format!("XML write error: {}", e))?;
            for idx in 0..shared_count {
                write!(xml, r#"<item x="{}"/>"#, idx)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            xml.push_str(r#"<item t="default"/>"#);
            xml.push_str("</items>");
            xml.push_str("</pivotField>");
        } else {
            xml.push_str(r#"/>"#);
        }
    }
    xml.push_str("</pivotFields>");

    if !row_indexes.is_empty() {
        write!(xml, r#"<rowFields count="{}">"#, row_indexes.len())
            .map_err(|e| format!("XML write error: {}", e))?;
        for idx in &row_indexes {
            write!(xml, r#"<field x="{}"/>"#, idx)
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("</rowFields>");
    }

    // Minimal rowItems (helps Excel accept the part).
    if row_indexes.len() == 1 {
        let shared = shared_count_for(row_indexes[0]);
        write!(xml, r#"<rowItems count="{}">"#, shared + 1)
            .map_err(|e| format!("XML write error: {}", e))?;
        for idx in 0..shared {
            if idx == 0 {
                xml.push_str(r#"<i><x/></i>"#);
            } else {
                write!(xml, r#"<i><x v="{}"/></i>"#, idx)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
        }
        xml.push_str(r#"<i t="grand"><x/></i>"#);
        xml.push_str("</rowItems>");
    }

    if !col_indexes.is_empty() {
        write!(xml, r#"<colFields count="{}">"#, col_indexes.len())
            .map_err(|e| format!("XML write error: {}", e))?;
        for idx in &col_indexes {
            write!(xml, r#"<field x="{}"/>"#, idx)
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("</colFields>");
    }

    // Minimal colItems (helps Excel accept the part).
    if col_indexes.len() == 1 {
        let shared = shared_count_for(col_indexes[0]);
        write!(xml, r#"<colItems count="{}">"#, shared + 1)
            .map_err(|e| format!("XML write error: {}", e))?;
        for idx in 0..shared {
            if idx == 0 {
                xml.push_str(r#"<i><x/></i>"#);
            } else {
                write!(xml, r#"<i><x v="{}"/></i>"#, idx)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
        }
        xml.push_str(r#"<i t="grand"><x/></i>"#);
        xml.push_str("</colItems>");
    }

    if !filter_indexes.is_empty() {
        write!(xml, r#"<pageFields count="{}">"#, filter_indexes.len())
            .map_err(|e| format!("XML write error: {}", e))?;
        for idx in &filter_indexes {
            write!(xml, r#"<pageField fld="{}"/>"#, idx)
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("</pageFields>");
    }

    if !data_fields.is_empty() {
        write!(xml, r#"<dataFields count="{}">"#, data_fields.len())
            .map_err(|e| format!("XML write error: {}", e))?;
        for (idx, func) in &data_fields {
            let subtotal = subtotal_from_function(*func);
            write!(
                xml,
                r#"<dataField name="{}" fld="{}" baseField="0" baseItem="0"/>"#,
                escape_xml(&pivot.field_names[*idx as usize]),
                idx
            )
            .map_err(|e| format!("XML write error: {}", e))?;
            let _ = subtotal;
        }
        xml.push_str("</dataFields>");
    }

    // Style info is commonly present and helps Excel accept the part.
    xml.push_str(
        r#"<pivotTableStyleInfo name="PivotStyleMedium4" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>"#,
    );

    // Excel pivotTableDefinition extensions
    xml.push_str(r#"<extLst>"#);
    xml.push_str(r#"<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}"><x14:pivotTableDefinition hideValuesRow="1"/></ext>"#);
    xml.push_str(r#"<ext uri="{B0B4B2F1-4D3B-4B4A-8E00-000000000000}"><pivotTableDefinition16 xmlns="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout"/></ext>"#);
    xml.push_str(r#"</extLst>"#);

    xml.push_str("</pivotTableDefinition>");
    Ok(xml)
}

fn build_field_index(field_names: &[String]) -> HashMap<String, u32> {
    let mut map = HashMap::with_capacity(field_names.len());
    for (i, name) in field_names.iter().enumerate() {
        map.insert(name.clone(), i as u32);
    }
    map
}

fn resolve_field_indexes(field_index: &HashMap<String, u32>, selected: &[String]) -> Vec<u32> {
    let mut result = Vec::new();
    for name in selected {
        if let Some(idx) = field_index.get(name) {
            result.push(*idx);
        }
    }
    result
}

fn subtotal_from_function(func: PivotValueFunction) -> &'static str {
    match func {
        PivotValueFunction::Sum => "sum",
        PivotValueFunction::Count => "count",
        PivotValueFunction::Average => "average",
        PivotValueFunction::Min => "min",
        PivotValueFunction::Max => "max",
        PivotValueFunction::Custom => "sum",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escape_sheet_name() {
        assert_eq!(escape_sheet_name("Sheet1"), "Sheet1");
        assert_eq!(escape_sheet_name("Bob's Sheet"), "Bob''s Sheet");
        assert_eq!(escape_sheet_name("Data'Value"), "Data''Value");
    }

    #[test]
    fn test_mutable_workbook_data_new() {
        let wb = MutableWorkbookData::new();
        assert_eq!(wb.worksheet_count(), 1);
        assert_eq!(wb.worksheets[0].name(), "Sheet1");
        // Note: is_modified() also checks worksheets, which may be marked modified on creation
    }

    #[test]
    fn test_mutable_workbook_add_worksheet() {
        let mut wb = MutableWorkbookData::new();
        let ws = wb.add_worksheet("Data".to_string());
        assert_eq!(ws.name(), "Data");
        assert_eq!(wb.worksheet_count(), 2);
        assert!(wb.is_modified());
    }

    #[test]
    fn test_mutable_workbook_worksheet_mut() {
        let mut wb = MutableWorkbookData::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_name("Updated".to_string());
        assert_eq!(wb.worksheets[0].name(), "Updated");
    }

    #[test]
    fn test_mutable_workbook_worksheet_mut_out_of_bounds() {
        let mut wb = MutableWorkbookData::new();
        assert!(wb.worksheet_mut(99).is_err());
    }

    #[test]
    fn test_define_name() {
        let mut wb = MutableWorkbookData::new();
        wb.define_name("SalesRange", "Sheet1!$A$1:$D$10");
        assert_eq!(wb.named_ranges.len(), 1);
        assert_eq!(wb.named_ranges[0].name, "SalesRange");
        assert_eq!(wb.named_ranges[0].reference, "Sheet1!$A$1:$D$10");
        assert!(wb.named_ranges[0].local_sheet_id.is_none());
    }

    #[test]
    fn test_define_name_local() {
        let mut wb = MutableWorkbookData::new();
        wb.define_name_local("LocalRange", "A1:B10", 0);
        assert_eq!(wb.named_ranges[0].local_sheet_id, Some(0));
    }

    #[test]
    fn test_define_name_with_comment() {
        let mut wb = MutableWorkbookData::new();
        wb.define_name_with_comment("NamedRange", "A1:C5", "This is a comment");
        assert_eq!(
            wb.named_ranges[0].comment,
            Some("This is a comment".to_string())
        );
    }

    #[test]
    fn test_remove_name() {
        let mut wb = MutableWorkbookData::new();
        wb.define_name("Range1", "A1:A10");
        wb.define_name("Range2", "B1:B10");
        assert_eq!(wb.named_ranges.len(), 2);

        let removed = wb.remove_name("Range1");
        assert!(removed);
        assert_eq!(wb.named_ranges.len(), 1);
        assert_eq!(wb.named_ranges[0].name, "Range2");

        let not_removed = wb.remove_name("NonExistent");
        assert!(!not_removed);
    }

    #[test]
    fn test_named_ranges_accessor() {
        let mut wb = MutableWorkbookData::new();
        wb.define_name("Test", "A1");
        let ranges = wb.named_ranges();
        assert_eq!(ranges.len(), 1);
    }

    #[test]
    fn test_hide_unhide_sheet() {
        let mut wb = MutableWorkbookData::new();
        wb.hide_sheet(0).unwrap();
        assert!(wb.worksheets[0].is_hidden());

        wb.unhide_sheet(0).unwrap();
        assert!(!wb.worksheets[0].is_hidden());
    }

    #[test]
    fn typed_visibility_is_persisted_in_workbook_xml() {
        let mut wb = MutableWorkbookData::new();
        assert_eq!(wb.get_sheet_visibility(0), Some(Visibility::Visible));

        wb.set_sheet_visibility(0, Visibility::VeryHidden).unwrap();
        assert_eq!(wb.get_sheet_visibility(0), Some(Visibility::VeryHidden));
        assert_eq!(wb.is_sheet_hidden(0), Some(true));

        let xml = wb
            .generate_workbook_xml_with_external_rels(&["rId1".to_string()], &[], &[])
            .unwrap();
        assert!(
            xml.contains(r#"<sheet name="Sheet1" sheetId="1" state="veryHidden" r:id="rId1"/>"#)
        );

        wb.set_sheet_visibility(0, Visibility::Visible).unwrap();
        let xml = wb
            .generate_workbook_xml_with_external_rels(&["rId1".to_string()], &[], &[])
            .unwrap();
        assert!(xml.contains(r#"<sheet name="Sheet1" sheetId="1" r:id="rId1"/>"#));
        assert!(!xml.contains(" state="));
    }

    #[test]
    fn test_hide_sheet_out_of_bounds() {
        let mut wb = MutableWorkbookData::new();
        assert!(wb.hide_sheet(99).is_err());
    }

    #[test]
    fn test_is_sheet_hidden() {
        let mut wb = MutableWorkbookData::new();
        assert_eq!(wb.is_sheet_hidden(0), Some(false));

        wb.hide_sheet(0).unwrap();
        assert_eq!(wb.is_sheet_hidden(0), Some(true));

        assert!(wb.is_sheet_hidden(99).is_none());
    }

    #[test]
    fn test_set_calculation_mode() {
        let mut wb = MutableWorkbookData::new();
        wb.set_calculation_mode("manual");
        assert_eq!(wb.calculation_mode, "manual");

        wb.set_calculation_mode("autoNoTable");
        assert_eq!(wb.calculation_mode, "autoNoTable");
    }

    #[test]
    fn test_protection_direct() {
        let mut wb = MutableWorkbookData::new();
        assert!(!wb.is_protected());

        let mut protection = WorkbookProtection::new();
        protection.set_structure_locked(true);
        wb.protection = Some(protection);

        assert!(wb.is_protected());
        assert!(wb.protection.as_ref().unwrap().structure_locked());
        assert!(!wb.protection.as_ref().unwrap().windows_locked());
    }

    #[test]
    fn test_remove_protection_direct() {
        let mut wb = MutableWorkbookData::new();
        let mut protection = WorkbookProtection::new();
        protection.set_structure_locked(true);
        wb.protection = Some(protection);
        assert!(wb.is_protected());

        wb.protection = None;
        assert!(!wb.is_protected());
    }

    #[test]
    fn test_force_formula_recalculation_direct() {
        let mut wb = MutableWorkbookData::new();
        assert!(!wb.force_formula_recalculation);

        wb.force_formula_recalculation = true;
        assert!(wb.force_formula_recalculation);
    }

    #[test]
    fn test_generate_workbook_xml_basic() {
        let wb = MutableWorkbookData::new();
        let rel_ids = vec!["rId1".to_string()];
        let xml = wb
            .generate_workbook_xml_with_external_rels(&rel_ids, &[], &[])
            .unwrap();

        assert!(xml.contains(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#));
        assert!(xml.contains(
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#
        ));
        assert!(xml.contains(r#"<sheets>"#));
        assert!(xml.contains(r#"<sheet name="Sheet1" sheetId="1" r:id="rId1"/>"#));
        assert!(xml.contains(r#"</workbook>"#));
    }

    #[test]
    fn test_generate_workbook_xml_with_protection() {
        let mut wb = MutableWorkbookData::new();
        let mut protection = WorkbookProtection::new();
        protection.set_workbook_verifier(Some(ProtectionPasswordVerifier::Legacy(0xABCD)));
        protection.set_revisions_verifier(Some(ProtectionPasswordVerifier::Strong(
            crate::xlsx::StrongProtectionPasswordVerifier::new(
                "SHA-512",
                vec![1, 2],
                vec![3, 4],
                100_000,
            )
            .unwrap(),
        )));
        protection.set_structure_locked(true);
        protection.set_windows_locked(true);
        protection.set_revision_locked(true);
        wb.set_workbook_protection(protection);

        let xml = wb
            .generate_workbook_xml_with_external_rels(&["rId1".to_string()], &[], &[])
            .unwrap();

        assert!(xml.contains(
            r#"<workbookProtection workbookPassword="ABCD" revisionsAlgorithmName="SHA-512" revisionsHashValue="AQI=" revisionsSaltValue="AwQ=" revisionsSpinCount="100000" lockStructure="1" lockWindows="1" lockRevision="1"/>"#
        ));
    }

    #[test]
    fn test_generate_workbook_xml_with_defined_names() {
        let mut wb = MutableWorkbookData::new();
        wb.define_name("Range1", "Sheet1!$A$1:$B$10");
        // local_sheet_id is the zero-based workbook sheet index used by OOXML.
        wb.define_name_local("LocalRange", "A1:C5", 0);

        let xml = wb
            .generate_workbook_xml_with_external_rels(&["rId1".to_string()], &[], &[])
            .unwrap();

        assert!(xml.contains(r#"<definedNames>"#));
        assert!(xml.contains(r#"<definedName name="Range1">"#));
        assert!(xml.contains(r#"<definedName name="LocalRange" localSheetId="0">"#));
        assert!(xml.contains(r#"</definedNames>"#));
    }

    #[test]
    fn test_build_field_index() {
        let fields = vec![
            "FieldA".to_string(),
            "FieldB".to_string(),
            "FieldC".to_string(),
        ];
        let index = build_field_index(&fields);

        assert_eq!(index.get("FieldA"), Some(&0u32));
        assert_eq!(index.get("FieldB"), Some(&1u32));
        assert_eq!(index.get("FieldC"), Some(&2u32));
    }

    #[test]
    fn test_resolve_field_indexes() {
        let mut index = HashMap::new();
        index.insert("A".to_string(), 0u32);
        index.insert("B".to_string(), 1u32);
        index.insert("C".to_string(), 2u32);

        let selected = vec!["A".to_string(), "C".to_string()];
        let resolved = resolve_field_indexes(&index, &selected);

        assert_eq!(resolved, vec![0, 2]);
    }

    #[test]
    fn test_subtotal_from_function() {
        assert_eq!(subtotal_from_function(PivotValueFunction::Sum), "sum");
        assert_eq!(subtotal_from_function(PivotValueFunction::Count), "count");
        assert_eq!(
            subtotal_from_function(PivotValueFunction::Average),
            "average"
        );
        assert_eq!(subtotal_from_function(PivotValueFunction::Min), "min");
        assert_eq!(subtotal_from_function(PivotValueFunction::Max), "max");
        assert_eq!(subtotal_from_function(PivotValueFunction::Custom), "sum");
    }

    #[test]
    fn test_parse_a1_range_valid() {
        let result = parse_a1_range("A1:D10").unwrap();
        // Returns 0-based indices (row, col) format
        assert_eq!(result, ((0, 0), (9, 3)));
    }

    #[test]
    fn test_parse_a1_range_invalid() {
        assert!(parse_a1_range("invalid").is_err());
        assert!(parse_a1_range("").is_err());
    }

    #[test]
    fn remove_worksheet_shifts_slots_names_and_pivots() {
        let mut wb = MutableWorkbookData::new();
        wb.add_worksheet("Sheet2".to_string());
        wb.add_worksheet("Sheet3".to_string());
        wb.define_name_local("Scoped1", "Sheet1!$A$1", 0);
        wb.define_name_local("Scoped2", "Sheet2!$A$1", 1);
        wb.define_name_local("Scoped3", "Sheet3!$A$1", 2);
        wb.define_name("Global", "Sheet1!$A$1");
        wb.pivot_tables.push(WritablePivotTable {
            name: "Pivot1".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_ref: "A1:B2".to_string(),
            dest_sheet_index: 2,
            location_ref: "A3".to_string(),
            field_names: vec!["F".to_string()],
            row_fields: vec!["F".to_string()],
            column_fields: Vec::new(),
            filter_fields: Vec::new(),
            data_fields: Vec::new(),
        });

        // Removing the pivot target is rejected; nothing changes.
        assert!(wb.remove_worksheet(2).is_err());
        assert_eq!(wb.worksheets.len(), 3);
        // Out-of-range removal is rejected.
        assert!(wb.remove_worksheet(9).is_err());

        let removed = wb.remove_worksheet(1).unwrap();
        assert_eq!(removed.name(), "Sheet2");
        assert_eq!(wb.worksheets.len(), 2);
        assert_eq!(
            wb.sheet_order,
            vec![SheetSlot::Worksheet(0), SheetSlot::Worksheet(1)]
        );

        let scopes: Vec<(&str, Option<u32>)> = wb
            .named_ranges
            .iter()
            .map(|range| (range.name.as_str(), range.local_sheet_id))
            .collect();
        assert!(scopes.contains(&("Scoped1", Some(0))));
        assert!(!scopes.iter().any(|(name, _)| *name == "Scoped2"));
        assert!(scopes.contains(&("Scoped3", Some(1))));
        assert!(scopes.contains(&("Global", None)));

        assert_eq!(wb.pivot_tables[0].dest_sheet_index, 1);
    }

    fn bar_chart() -> crate::xlsx::WorksheetChart {
        crate::xlsx::WorksheetChart::bar_chart(
            "Sales",
            "Sheet1!$A$2:$A$3",
            "Sheet1!$B$2:$B$3",
            crate::xlsx::ChartAnchor::new(0, 0, 10, 15),
        )
        .unwrap()
    }

    #[test]
    fn remove_chart_sheet_shifts_slots_and_scopes() {
        let mut wb = MutableWorkbookData::new();
        wb.add_chart_sheet("Chart1", bar_chart()).unwrap();
        wb.add_worksheet("Sheet2".to_string());
        wb.add_chart_sheet("Chart2", bar_chart()).unwrap();
        wb.define_name_local("Scoped1", "Sheet1!$A$1", 0);
        wb.define_name_local("ScopedChart1", "Sheet1!$A$1", 1);
        wb.define_name_local("ScopedSheet2", "Sheet2!$A$1", 2);
        wb.define_name_local("ScopedChart2", "Sheet1!$A$1", 3);

        // Out-of-range removal is rejected; nothing changes.
        assert!(wb.remove_chart_sheet(2).is_err());
        assert_eq!(wb.chart_sheets.len(), 2);

        let removed = wb.remove_chart_sheet(0).unwrap();
        assert_eq!(removed.name(), "Chart1");
        assert_eq!(wb.chart_sheets.len(), 1);
        assert_eq!(
            wb.sheet_order,
            vec![
                SheetSlot::Worksheet(0),
                SheetSlot::Worksheet(1),
                SheetSlot::ChartSheet(0)
            ]
        );

        let scopes: Vec<(&str, Option<u32>)> = wb
            .named_ranges
            .iter()
            .map(|range| (range.name.as_str(), range.local_sheet_id))
            .collect();
        assert!(scopes.contains(&("Scoped1", Some(0))));
        assert!(!scopes.iter().any(|(name, _)| *name == "ScopedChart1"));
        assert!(scopes.contains(&("ScopedSheet2", Some(1))));
        assert!(scopes.contains(&("ScopedChart2", Some(2))));
    }

    #[test]
    fn worksheet_creation_validates_names_uniformly() {
        let mut wb = MutableWorkbookData::new();
        assert!(wb.try_add_worksheet("a/b".to_string()).is_err());
        assert!(wb.try_add_worksheet(String::new()).is_err());
        assert!(wb.try_add_worksheet("x".repeat(32)).is_err());
        assert!(wb.insert_worksheet(0, "a?b".to_string()).is_err());
        // Duplicate of the default Sheet1, in any letter case.
        assert!(wb.try_add_worksheet("sheet1".to_string()).is_err());
        assert!(wb.insert_worksheet(1, "SHEET1".to_string()).is_err());
        assert_eq!(wb.worksheets.len(), 1, "failed creations must not mutate");

        wb.try_add_worksheet("Summary".to_string()).unwrap();
        wb.insert_worksheet(0, "First".to_string()).unwrap();
        assert_eq!(wb.worksheets.len(), 3);

        // The unchecked legacy append path is unchanged.
        wb.add_worksheet("Anything Goes".to_string());
        assert_eq!(wb.worksheets.len(), 4);
    }
}
