//! Mutable spreadsheet structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODS spreadsheets that allows
//! for in-place modification of sheets, rows, and cells.

use crate::common::{Metadata, Result, xml::escape_xml};
use crate::odf::core::{OdfStructure, PackageWriter};
use crate::odf::ods::{Cell, CellValue, Row, Sheet, Spreadsheet};
use std::path::Path;

/// A mutable ODS spreadsheet that supports in-place modifications.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::{Spreadsheet, MutableSpreadsheet};
///
/// # fn main() -> litchi::Result<()> {
/// let spreadsheet = Spreadsheet::open("input.ods")?;
/// let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)?;
///
/// // Modify the spreadsheet
/// mutable.add_sheet("NewSheet")?;
/// mutable.save("output.ods")?;
/// # Ok(())
/// # }
/// ```
pub struct MutableSpreadsheet {
    /// Mutable sheets
    sheets: Vec<Sheet>,
    /// Document metadata
    metadata: Metadata,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML
    styles_xml: Option<String>,
}

impl MutableSpreadsheet {
    fn sheet_max_cols(sheet: &Sheet) -> usize {
        sheet.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
    }

    fn has_formulas(&self) -> bool {
        self.sheets
            .iter()
            .flat_map(|s| s.rows.iter())
            .flat_map(|r| r.cells.iter())
            .any(|c| c.formula.is_some())
    }

    fn push_table_columns(out: &mut String, max_cols: usize) {
        if max_cols <= 1 {
            out.push_str("<table:table-column/>");
        } else {
            out.push_str(&format!(
                r#"<table:table-column table:number-columns-repeated="{}"/>"#,
                max_cols
            ));
        }
    }

    fn push_cell(out: &mut String, cell: &Cell) {
        let formula_attr = cell
            .formula
            .as_deref()
            .map(|f| format!(" table:formula=\"{}\"", escape_xml(f)))
            .unwrap_or_default();

        match &cell.value {
            CellValue::Text(_) => {
                out.push_str(&format!(
                    r#"<table:table-cell{} office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
                    formula_attr,
                    escape_xml(&cell.text)
                ));
            },
            CellValue::Number(f) => {
                out.push_str(&format!(
                    r#"<table:table-cell{} office:value-type="float" office:value="{}"><text:p>{}</text:p></table:table-cell>"#,
                    formula_attr,
                    f,
                    escape_xml(&cell.text)
                ));
            },
            CellValue::Currency(f, currency) => {
                out.push_str(&format!(
                    r#"<table:table-cell{} office:value-type="currency" office:value="{}" office:currency="{}"><text:p>{}</text:p></table:table-cell>"#,
                    formula_attr,
                    f,
                    escape_xml(currency),
                    escape_xml(&cell.text)
                ));
            },
            CellValue::Percentage(f) => {
                out.push_str(&format!(
                    r#"<table:table-cell{} office:value-type="percentage" office:value="{}"><text:p>{}</text:p></table:table-cell>"#,
                    formula_attr,
                    f,
                    escape_xml(&cell.text)
                ));
            },
            CellValue::Date(d) => {
                out.push_str(&format!(
                    r#"<table:table-cell{} office:value-type="date" office:date-value="{}"><text:p>{}</text:p></table:table-cell>"#,
                    formula_attr,
                    escape_xml(d),
                    escape_xml(&cell.text)
                ));
            },
            CellValue::Time(t) => {
                out.push_str(&format!(
                    r#"<table:table-cell{} office:value-type="time" office:time-value="{}"><text:p>{}</text:p></table:table-cell>"#,
                    formula_attr,
                    escape_xml(t),
                    escape_xml(&cell.text)
                ));
            },
            CellValue::Boolean(b) => {
                out.push_str(&format!(
                    r#"<table:table-cell{} office:value-type="boolean" office:boolean-value="{}"><text:p>{}</text:p></table:table-cell>"#,
                    formula_attr,
                    b,
                    escape_xml(&cell.text)
                ));
            },
            CellValue::Empty => {
                if cell.formula.is_some() {
                    out.push_str(&format!(
                        r#"<table:table-cell{} office:value-type="float" office:value="0"><text:p>0</text:p></table:table-cell>"#,
                        formula_attr
                    ));
                } else {
                    out.push_str("<table:table-cell/>");
                }
            },
        }
    }

    /// Create a mutable spreadsheet from an existing Spreadsheet.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::{Spreadsheet, MutableSpreadsheet};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)?;
    /// mutable.add_sheet("NewSheet")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_spreadsheet(mut spreadsheet: Spreadsheet) -> Result<Self> {
        let sheets = spreadsheet.sheets()?;
        let metadata = spreadsheet.metadata()?;
        let mimetype = "application/vnd.oasis.opendocument.spreadsheet".to_string();

        // Extract styles XML from the spreadsheet's package (requires accessing internal package)
        // For now, we'll use None and rely on default styles
        // TODO: Add method to Spreadsheet to expose get_file for extracting styles.xml

        Ok(Self {
            sheets,
            metadata,
            mimetype,
            styles_xml: None,
        })
    }

    /// Create a new empty mutable spreadsheet.
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.spreadsheet".to_string(),
            styles_xml: None,
        }
    }

    /// Get all sheets.
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// Get mutable reference to sheets.
    pub fn sheets_mut(&mut self) -> &mut Vec<Sheet> {
        &mut self.sheets
    }

    /// Get metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Get mutable reference to metadata.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }

    /// Add a new sheet.
    pub fn add_sheet(&mut self, name: &str) -> Result<()> {
        let sheet = Sheet {
            name: name.to_string(),
            rows: Vec::new(),
        };
        self.sheets.push(sheet);
        Ok(())
    }

    /// Remove a sheet at index.
    pub fn remove_sheet(&mut self, index: usize) -> Result<Sheet> {
        if index < self.sheets.len() {
            Ok(self.sheets.remove(index))
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                index
            )))
        }
    }

    /// Add a row to a sheet.
    pub fn add_row(&mut self, sheet_index: usize, cells: Vec<Cell>) -> Result<()> {
        if sheet_index < self.sheets.len() {
            let row_index = self.sheets[sheet_index].rows.len();
            let row = Row {
                cells,
                index: row_index,
            };
            self.sheets[sheet_index].rows.push(row);
            Ok(())
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Remove a row from a sheet.
    pub fn remove_row(&mut self, sheet_index: usize, row_index: usize) -> Result<Row> {
        if sheet_index < self.sheets.len() {
            let sheet = &mut self.sheets[sheet_index];
            if row_index < sheet.rows.len() {
                Ok(sheet.rows.remove(row_index))
            } else {
                Err(crate::common::Error::InvalidFormat(format!(
                    "Row index {} out of bounds",
                    row_index
                )))
            }
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Set a cell value.
    ///
    /// # Arguments
    ///
    /// * `sheet_index` - Index of the sheet
    /// * `row` - Row index
    /// * `col` - Column index  
    /// * `value` - New cell value
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::{MutableSpreadsheet, CellValue};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut spreadsheet = MutableSpreadsheet::new();
    /// spreadsheet.add_sheet("Sheet1")?;
    /// spreadsheet.set_cell(0, 0, 0, CellValue::Text("Hello".to_string()))?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_cell(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        value: CellValue,
    ) -> Result<()> {
        if sheet_index < self.sheets.len() {
            let sheet = &mut self.sheets[sheet_index];

            // Ensure row exists
            while sheet.rows.len() <= row {
                let row_index = sheet.rows.len();
                sheet.rows.push(Row {
                    cells: Vec::new(),
                    index: row_index,
                });
            }

            let row_data = &mut sheet.rows[row];

            // Ensure cell exists
            while row_data.cells.len() <= col {
                let col_index = row_data.cells.len();
                row_data.cells.push(Cell {
                    value: CellValue::Empty,
                    text: String::new(),
                    formula: None,
                    row,
                    col: col_index,
                });
            }

            // Set the cell value
            row_data.cells[col].value = value.clone();
            row_data.cells[col].text = match value {
                CellValue::Empty => String::new(),
                CellValue::Text(ref s) => s.clone(),
                CellValue::Number(n) => n.to_string(),
                CellValue::Boolean(b) => b.to_string(),
                CellValue::Date(ref d) => d.clone(),
                CellValue::Currency(n, ref currency) => format!("{} {}", n, currency),
                CellValue::Percentage(n) => format!("{}%", n * 100.0),
                CellValue::Time(ref t) => t.clone(),
            };

            Ok(())
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Clear a cell value.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableSpreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut spreadsheet = MutableSpreadsheet::new();
    /// spreadsheet.add_sheet("Sheet1")?;
    /// spreadsheet.clear_cell(0, 0, 0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_cell(&mut self, sheet_index: usize, row: usize, col: usize) -> Result<()> {
        if sheet_index < self.sheets.len() {
            let sheet = &mut self.sheets[sheet_index];
            if row < sheet.rows.len() {
                let row_data = &mut sheet.rows[row];
                if col < row_data.cells.len() {
                    row_data.cells[col].value = CellValue::Empty;
                    row_data.cells[col].text = String::new();
                    Ok(())
                } else {
                    Err(crate::common::Error::InvalidFormat(format!(
                        "Column index {} out of bounds",
                        col
                    )))
                }
            } else {
                Err(crate::common::Error::InvalidFormat(format!(
                    "Row index {} out of bounds",
                    row
                )))
            }
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Clear all content from a sheet.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableSpreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut spreadsheet = MutableSpreadsheet::new();
    /// spreadsheet.add_sheet("Sheet1")?;
    /// spreadsheet.clear_sheet(0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_sheet(&mut self, sheet_index: usize) -> Result<()> {
        if sheet_index < self.sheets.len() {
            self.sheets[sheet_index].rows.clear();
            Ok(())
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Generate content.xml from current state.
    fn generate_content_xml(&self) -> String {
        let mut body = String::new();

        for sheet in &self.sheets {
            let escaped_name = escape_xml(&sheet.name);
            body.push_str(&format!(r#"<table:table table:name="{}">"#, escaped_name));

            Self::push_table_columns(&mut body, Self::sheet_max_cols(sheet));

            for row in &sheet.rows {
                body.push_str("<table:table-row>");
                for cell in &row.cells {
                    Self::push_cell(&mut body, cell);
                }
                body.push_str("</table:table-row>");
            }

            body.push_str("</table:table>");
        }

        let of_ns = if self.has_formulas() {
            " xmlns:of=\"urn:oasis:names:tc:opendocument:xmlns:of:1.2\""
        } else {
            ""
        };

        let mut out = String::with_capacity(body.len() + 256);
        out.push_str(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#,
        );
        out.push_str(of_ns);
        out.push_str(
            r#" office:version="1.3"><office:font-face-decls/><office:automatic-styles/><office:body><office:spreadsheet>"#,
        );
        out.push_str(&body);
        out.push_str(r#"</office:spreadsheet></office:body></office:document-content>"#);
        out
    }

    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><dc:date>{}</dc:date></office:meta></office:document-meta>"#,
            now
        )
    }

    /// Save the modified spreadsheet.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert to bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        writer.set_mimetype(&self.mimetype)?;

        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        let default_styles = OdfStructure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        writer.finish_to_bytes()
    }
}

impl Default for MutableSpreadsheet {
    fn default() -> Self {
        Self::new()
    }
}
