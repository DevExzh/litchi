/// Writer module for creating and modifying Excel worksheets.
use crate::sheet::{CellValue, Result as SheetResult};
use std::collections::HashMap;
use std::fmt::Write as FmtWrite;

// Import shared formatting types
pub use super::super::format::{
    CellBorder, CellBorderLineStyle, CellBorderSide, CellFill, CellFillPatternType, CellFont,
    CellFormat, Chart, ChartType, DataValidation, DataValidationOperator, DataValidationType,
};
// Import from other writer modules
use super::strings::MutableSharedStrings;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Freeze panes configuration.
///
/// Freezes rows and columns in place while scrolling.
#[derive(Debug, Clone)]
pub struct FreezePanes {
    /// Number of columns to freeze from the left
    pub freeze_cols: u32,
    /// Number of rows to freeze from the top
    pub freeze_rows: u32,
}

/// Named range definition.
///
/// Associates a name with a cell or range of cells for easier formula references.
#[derive(Debug, Clone)]
pub struct NamedRange {
    /// Name of the range (e.g., "TaxRate", "SalesData")
    pub name: String,
    /// Reference formula (e.g., "Sheet1!$A$1:$B$10", "Sheet1!$C$5")
    pub reference: String,
    /// Optional comment/description
    pub comment: Option<String>,
    /// Whether this is a workbook-scoped or sheet-scoped name
    /// If None, it's workbook-scoped; if Some(sheet_index), it's sheet-scoped
    pub local_sheet_id: Option<u32>,
}

/// A mutable worksheet for writing and modification.
///
/// Provides methods to set cell values, formulas, and formatting.
#[derive(Debug)]
pub struct MutableWorksheet {
    /// Worksheet name
    name: String,
    /// Sheet ID
    sheet_id: u32,
    /// Cell data (row, col) -> value
    cells: HashMap<(u32, u32), CellValue>,
    /// Cell formatting
    cell_formats: HashMap<(u32, u32), CellFormat>,
    /// Merged cell ranges (start_row, start_col, end_row, end_col)
    merged_cells: Vec<(u32, u32, u32, u32)>,
    /// Charts in this worksheet
    charts: Vec<Chart>,
    /// Data validation rules
    validations: Vec<DataValidation>,
    /// Column widths (col -> width in characters)
    column_widths: HashMap<u32, f64>,
    /// Hidden columns
    hidden_columns: std::collections::HashSet<u32>,
    /// Row heights (row -> height in points)
    row_heights: HashMap<u32, f64>,
    /// Hidden rows
    hidden_rows: std::collections::HashSet<u32>,
    /// Freeze panes configuration
    freeze_panes: Option<FreezePanes>,
    /// Whether the worksheet has been modified
    modified: bool,
}

impl MutableWorksheet {
    /// Create a new empty worksheet.
    pub fn new(name: String, sheet_id: u32) -> Self {
        Self {
            name,
            sheet_id,
            cells: HashMap::new(),
            cell_formats: HashMap::new(),
            merged_cells: Vec::new(),
            charts: Vec::new(),
            validations: Vec::new(),
            column_widths: HashMap::new(),
            hidden_columns: std::collections::HashSet::new(),
            row_heights: HashMap::new(),
            hidden_rows: std::collections::HashSet::new(),
            freeze_panes: None,
            modified: false,
        }
    }

    /// Get the worksheet name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Set the worksheet name.
    pub fn set_name(&mut self, name: String) {
        self.name = name;
        self.modified = true;
    }

    /// Get the sheet ID.
    pub fn sheet_id(&self) -> u32 {
        self.sheet_id
    }

    /// Set a cell value.
    pub fn set_cell_value<V: Into<CellValue>>(&mut self, row: u32, col: u32, value: V) {
        self.cells.insert((row, col), value.into());
        self.modified = true;
    }

    /// Set a cell formula.
    pub fn set_cell_formula(&mut self, row: u32, col: u32, formula: &str) {
        self.cells.insert(
            (row, col),
            CellValue::Formula {
                formula: formula.to_string(),
                cached_value: None,
            },
        );
        self.modified = true;
    }

    /// Set a cell formula with a cached result value.
    pub fn set_cell_formula_with_cache<V: Into<CellValue>>(
        &mut self,
        row: u32,
        col: u32,
        formula: &str,
        cached_value: V,
    ) {
        self.cells.insert(
            (row, col),
            CellValue::Formula {
                formula: formula.to_string(),
                cached_value: Some(Box::new(cached_value.into())),
            },
        );
        self.modified = true;
    }

    /// Set cell formatting.
    pub fn set_cell_format(&mut self, row: u32, col: u32, format: CellFormat) {
        self.cell_formats.insert((row, col), format);
        self.modified = true;
    }

    /// Merge cells in a rectangular range.
    pub fn merge_cells(&mut self, start_row: u32, start_col: u32, end_row: u32, end_col: u32) {
        self.merged_cells
            .push((start_row, start_col, end_row, end_col));
        self.modified = true;
    }

    /// Add a chart to the worksheet.
    pub fn add_chart(
        &mut self,
        chart_type: ChartType,
        title: &str,
        data_range: &str,
        position: (u32, u32, u32, u32),
        show_legend: bool,
    ) {
        self.charts.push(Chart {
            chart_type,
            title: Some(title.to_string()),
            data_range: data_range.to_string(),
            position,
            show_legend,
        });
        self.modified = true;
    }

    /// Add data validation to a cell range.
    #[allow(clippy::too_many_arguments)]
    pub fn add_data_validation(
        &mut self,
        range: &str,
        validation_type: DataValidationType,
        show_input_message: bool,
        input_title: Option<&str>,
        input_message: Option<&str>,
        show_error_alert: bool,
        error_title: Option<&str>,
        error_message: Option<&str>,
    ) {
        self.validations.push(DataValidation {
            range: range.to_string(),
            validation_type,
            show_input_message,
            input_title: input_title.map(|s| s.to_string()),
            input_message: input_message.map(|s| s.to_string()),
            show_error_alert,
            error_title: error_title.map(|s| s.to_string()),
            error_message: error_message.map(|s| s.to_string()),
        });
        self.modified = true;
    }

    /// Get a cell value.
    pub fn cell_value(&self, row: u32, col: u32) -> Option<&CellValue> {
        self.cells.get(&(row, col))
    }

    /// Clear a cell.
    pub fn clear_cell(&mut self, row: u32, col: u32) {
        self.cells.remove(&(row, col));
        self.modified = true;
    }

    /// Clear all cells in the worksheet.
    pub fn clear_all(&mut self) {
        self.cells.clear();
        self.modified = true;
    }

    /// Get the number of non-empty cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Set column width in characters (Excel default is 8.43).
    pub fn set_column_width(&mut self, col: u32, width: f64) {
        self.column_widths.insert(col, width);
        self.modified = true;
    }

    /// Hide a column.
    pub fn hide_column(&mut self, col: u32) {
        self.hidden_columns.insert(col);
        self.modified = true;
    }

    /// Show a previously hidden column.
    pub fn show_column(&mut self, col: u32) {
        self.hidden_columns.remove(&col);
        self.modified = true;
    }

    /// Set row height in points (Excel default is 15).
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.row_heights.insert(row, height);
        self.modified = true;
    }

    /// Hide a row.
    pub fn hide_row(&mut self, row: u32) {
        self.hidden_rows.insert(row);
        self.modified = true;
    }

    /// Show a previously hidden row.
    pub fn show_row(&mut self, row: u32) {
        self.hidden_rows.remove(&row);
        self.modified = true;
    }

    /// Freeze panes at the specified position.
    pub fn freeze_panes(&mut self, freeze_rows: u32, freeze_cols: u32) {
        if freeze_rows > 0 || freeze_cols > 0 {
            self.freeze_panes = Some(FreezePanes {
                freeze_rows,
                freeze_cols,
            });
            self.modified = true;
        }
    }

    /// Remove freeze panes.
    pub fn unfreeze_panes(&mut self) {
        self.freeze_panes = None;
        self.modified = true;
    }

    /// Check if the worksheet has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified
    }

    /// Get the used range dimensions (min_row, min_col, max_row, max_col).
    pub fn used_range(&self) -> Option<(u32, u32, u32, u32)> {
        if self.cells.is_empty() {
            return None;
        }

        let mut min_row = u32::MAX;
        let mut max_row = 0;
        let mut min_col = u32::MAX;
        let mut max_col = 0;

        for &(row, col) in self.cells.keys() {
            min_row = min_row.min(row);
            max_row = max_row.max(row);
            min_col = min_col.min(col);
            max_col = max_col.max(col);
        }

        Some((min_row, min_col, max_row, max_col))
    }

    /// Serialize the worksheet to XML.
    ///
    /// # Arguments
    /// * `shared_strings` - Mutable shared strings table
    /// * `style_indices` - Optional map of cell positions to style indices
    pub fn to_xml(
        &self,
        shared_strings: &mut MutableSharedStrings,
        style_indices: &HashMap<(u32, u32), usize>,
    ) -> SheetResult<String> {
        let mut xml = String::with_capacity(4096);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);

        // Write sheet dimensions
        if let Some((min_row, min_col, max_row, max_col)) = self.used_range() {
            let min_ref = format!("{}{}", Self::column_to_letters(min_col), min_row);
            let max_ref = format!("{}{}", Self::column_to_letters(max_col), max_row);
            write!(
                xml,
                r#"<dimension ref="{}:{}"/>"#,
                escape_xml(&min_ref),
                escape_xml(&max_ref)
            )
            .map_err(|e| format!("XML write error: {}", e))?;
        }

        // Write sheet views (including freeze panes if set)
        xml.push_str("<sheetViews><sheetView workbookViewId=\"0\"");

        // Add freeze panes if configured
        if let Some(ref freeze) = self.freeze_panes {
            xml.push('>');

            let y_split = freeze.freeze_rows;
            let x_split = freeze.freeze_cols;

            let active_pane = match (x_split > 0, y_split > 0) {
                (true, true) => "bottomRight",
                (true, false) => "topRight",
                (false, true) => "bottomLeft",
                (false, false) => "",
            };

            let top_left_cell = format!("{}{}", Self::column_to_letters(x_split + 1), y_split + 1);

            write!(
                xml,
                r#"<pane xSplit="{}" ySplit="{}" topLeftCell="{}" activePane="{}" state="frozen"/>"#,
                x_split, y_split, top_left_cell, active_pane
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            if !active_pane.is_empty() {
                write!(
                    xml,
                    r#"<selection pane="{}" activeCell="{}" sqref="{}"/>"#,
                    active_pane, top_left_cell, top_left_cell
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("</sheetView>");
        } else {
            xml.push_str("/>");
        }

        xml.push_str("</sheetViews>");
        xml.push_str("<sheetFormatPr defaultRowHeight=\"15\"/>\n");

        // Write sheet data
        xml.push_str("<sheetData>");
        self.write_sheet_data(&mut xml, shared_strings, style_indices)?;
        xml.push_str("</sheetData>");

        // Write merged cells
        if !self.merged_cells.is_empty() {
            write!(xml, r#"<mergeCells count="{}">"#, self.merged_cells.len())
                .map_err(|e| format!("XML write error: {}", e))?;

            for (start_row, start_col, end_row, end_col) in &self.merged_cells {
                let start_ref = format!("{}{}", Self::column_to_letters(*start_col), start_row);
                let end_ref = format!("{}{}", Self::column_to_letters(*end_col), end_row);
                write!(xml, r#"<mergeCell ref="{}:{}"/>"#, start_ref, end_ref)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("</mergeCells>");
        }

        xml.push_str("</worksheet>");

        Ok(xml)
    }

    /// Get cell formats for all cells (used by workbook to build styles).
    pub fn cell_formats(&self) -> &HashMap<(u32, u32), CellFormat> {
        &self.cell_formats
    }

    /// Write sheet data (rows and cells).
    fn write_sheet_data(
        &self,
        xml: &mut String,
        shared_strings: &mut MutableSharedStrings,
        style_indices: &HashMap<(u32, u32), usize>,
    ) -> SheetResult<()> {
        if self.cells.is_empty() {
            return Ok(());
        }

        // Group cells by row
        let mut rows: HashMap<u32, Vec<(u32, &CellValue)>> = HashMap::new();
        for (&(row, col), value) in &self.cells {
            rows.entry(row).or_default().push((col, value));
        }

        // Sort rows
        let mut row_nums: Vec<u32> = rows.keys().copied().collect();
        row_nums.sort_unstable();

        for row_num in row_nums {
            let mut cells = rows[&row_num].clone();
            cells.sort_unstable_by_key(|(col, _)| *col);

            write!(xml, r#"<row r="{}">"#, row_num)
                .map_err(|e| format!("XML write error: {}", e))?;

            for (col_num, value) in cells {
                let cell_ref = format!("{}{}", Self::column_to_letters(col_num), row_num);
                // Get the style index for this cell (if any)
                let style_index = style_indices.get(&(row_num, col_num)).copied();
                self.write_cell(xml, &cell_ref, value, shared_strings, style_index)?;
            }

            xml.push_str("</row>");
        }

        Ok(())
    }

    /// Write a single cell to XML.
    fn write_cell(
        &self,
        xml: &mut String,
        cell_ref: &str,
        value: &CellValue,
        shared_strings: &mut MutableSharedStrings,
        style_index: Option<usize>,
    ) -> SheetResult<()> {
        // Helper to add style attribute if present
        let style_attr = if let Some(idx) = style_index {
            format!(r#" s="{}""#, idx)
        } else {
            String::new()
        };

        match value {
            CellValue::Empty => {},
            CellValue::String(s) => {
                let string_index = shared_strings.add_string(s);
                write!(
                    xml,
                    r#"<c r="{}"{} t="s"><v>{}</v></c>"#,
                    cell_ref, style_attr, string_index
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Int(i) => {
                write!(
                    xml,
                    r#"<c r="{}"{}>  <v>{}</v></c>"#,
                    cell_ref, style_attr, i
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Float(f) => {
                write!(
                    xml,
                    r#"<c r="{}"{}>  <v>{}</v></c>"#,
                    cell_ref, style_attr, f
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Bool(b) => {
                write!(
                    xml,
                    r#"<c r="{}"{} t="b"><v>{}</v></c>"#,
                    cell_ref,
                    style_attr,
                    if *b { "1" } else { "0" }
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::DateTime(d) => {
                write!(
                    xml,
                    r#"<c r="{}"{}>  <v>{}</v></c>"#,
                    cell_ref, style_attr, d
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Error(e) => {
                write!(
                    xml,
                    r#"<c r="{}"{} t="e"><v>{}</v></c>"#,
                    cell_ref,
                    style_attr,
                    escape_xml(e)
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Formula {
                formula,
                cached_value,
            } => {
                xml.push_str(&format!(r#"<c r="{}"{}">"#, cell_ref, style_attr));
                write!(xml, "<f>{}</f>", escape_xml(formula))
                    .map_err(|e| format!("XML write error: {}", e))?;

                if let Some(cached) = cached_value {
                    match &**cached {
                        CellValue::String(s) => {
                            let string_index = shared_strings.add_string(s);
                            write!(xml, "<v>{}</v>", string_index)
                                .map_err(|e| format!("XML write error: {}", e))?;
                        },
                        CellValue::Int(i) => {
                            write!(xml, "<v>{}</v>", i)
                                .map_err(|e| format!("XML write error: {}", e))?;
                        },
                        CellValue::Float(f) => {
                            write!(xml, "<v>{}</v>", f)
                                .map_err(|e| format!("XML write error: {}", e))?;
                        },
                        CellValue::Bool(b) => {
                            write!(xml, "<v>{}</v>", if *b { "1" } else { "0" })
                                .map_err(|e| format!("XML write error: {}", e))?;
                        },
                        _ => {},
                    }
                }
                xml.push_str("</c>");
            },
        }

        Ok(())
    }

    /// Convert column number to Excel column letters (e.g., 1 -> "A", 26 -> "Z", 27 -> "AA").
    pub(crate) fn column_to_letters(col: u32) -> String {
        let mut letters = String::new();
        let mut col = col;

        while col > 0 {
            col -= 1;
            let letter = ((col % 26) as u8 + b'A') as char;
            letters.insert(0, letter);
            col /= 26;
        }

        letters
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_worksheet() {
        let ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        assert_eq!(ws.name(), "Sheet1");
        assert_eq!(ws.sheet_id(), 1);
        assert_eq!(ws.cell_count(), 0);
    }

    #[test]
    fn test_set_cell_value() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_cell_value(1, 1, "Hello");
        ws.set_cell_value(1, 2, 42);
        ws.set_cell_value(2, 1, 3.15);

        assert_eq!(ws.cell_count(), 3);
        assert!(matches!(ws.cell_value(1, 1), Some(CellValue::String(_))));
    }

    #[test]
    fn test_column_to_letters() {
        assert_eq!(MutableWorksheet::column_to_letters(1), "A");
        assert_eq!(MutableWorksheet::column_to_letters(26), "Z");
        assert_eq!(MutableWorksheet::column_to_letters(27), "AA");
        assert_eq!(MutableWorksheet::column_to_letters(702), "ZZ");
    }
}
