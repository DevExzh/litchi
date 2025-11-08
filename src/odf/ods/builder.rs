//! OpenDocument Spreadsheet builder.
//!
//! This module provides a builder pattern for creating new ODS spreadsheets from scratch.

use crate::common::{Metadata, Result};
use crate::odf::core::{OdfStructure, PackageWriter};
use crate::odf::ods::{Cell, CellValue, Row, Sheet};
use std::path::Path;

/// Builder for creating new ODS spreadsheets.
///
/// This builder allows you to create ODS spreadsheets programmatically by adding
/// sheets, rows, and cells, then saving them to a file or bytes.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::SpreadsheetBuilder;
///
/// # fn main() -> litchi::Result<()> {
/// let mut builder = SpreadsheetBuilder::new();
/// builder.add_sheet("Sheet1")?;
/// builder.add_row_with_values(&["Name", "Age", "City"])?;
/// builder.add_row_with_values(&["Alice", "30", "New York"])?;
/// builder.save("spreadsheet.ods")?;
/// # Ok(())
/// # }
/// ```
pub struct SpreadsheetBuilder {
    sheets: Vec<Sheet>,
    metadata: Metadata,
}

impl Default for SpreadsheetBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl SpreadsheetBuilder {
    /// Create a new spreadsheet builder
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::SpreadsheetBuilder;
    ///
    /// let builder = SpreadsheetBuilder::new();
    /// ```
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            metadata: Metadata::default(),
        }
    }

    /// Set document metadata
    ///
    /// # Arguments
    ///
    /// * `metadata` - Document metadata (title, author, etc.)
    pub fn set_metadata(&mut self, metadata: Metadata) {
        self.metadata = metadata;
    }

    /// Add a new sheet to the spreadsheet
    ///
    /// # Arguments
    ///
    /// * `name` - Name of the sheet
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_sheet(&mut self, name: &str) -> Result<&mut Self> {
        let sheet = Sheet {
            name: name.to_string(),
            rows: Vec::new(),
        };
        self.sheets.push(sheet);
        Ok(self)
    }

    /// Add a row to the current sheet with string values
    ///
    /// # Arguments
    ///
    /// * `values` - String values for the cells in the row
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.add_row_with_values(&["A", "B", "C"])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_row_with_values(&mut self, values: &[&str]) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        let row_index = if let Some(sheet) = self.sheets.last() {
            sheet.rows.len()
        } else {
            0
        };

        let cells: Vec<Cell> = values
            .iter()
            .enumerate()
            .map(|(col, &value)| Cell {
                text: value.to_string(),
                value: CellValue::Text(value.to_string()),
                formula: None,
                row: row_index,
                col,
            })
            .collect();

        let row = Row {
            cells,
            index: row_index,
        };

        if let Some(sheet) = self.sheets.last_mut() {
            sheet.rows.push(row);
        }

        Ok(self)
    }

    /// Add a row with numbers
    ///
    /// # Arguments
    ///
    /// * `values` - Numeric values for the cells in the row
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.add_row_with_numbers(&[1.0, 2.5, 3.14])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_row_with_numbers(&mut self, values: &[f64]) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        let row_index = if let Some(sheet) = self.sheets.last() {
            sheet.rows.len()
        } else {
            0
        };

        let cells: Vec<Cell> = values
            .iter()
            .enumerate()
            .map(|(col, &value)| Cell {
                text: value.to_string(),
                value: CellValue::Number(value),
                formula: None,
                row: row_index,
                col,
            })
            .collect();

        let row = Row {
            cells,
            index: row_index,
        };

        if let Some(sheet) = self.sheets.last_mut() {
            sheet.rows.push(row);
        }

        Ok(self)
    }

    /// Add a row with mixed values (numbers, text, booleans)
    ///
    /// # Arguments
    ///
    /// * `values` - Cell values for the row
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::{SpreadsheetBuilder, CellValue};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.add_row_with_cell_values(&[
    ///     CellValue::Text("Product".to_string()),
    ///     CellValue::Number(99.99),
    ///     CellValue::Boolean(true),
    /// ])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_row_with_cell_values(&mut self, values: &[CellValue]) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        let row_index = if let Some(sheet) = self.sheets.last() {
            sheet.rows.len()
        } else {
            0
        };

        let cells: Vec<Cell> = values
            .iter()
            .enumerate()
            .map(|(col, value)| {
                let text = match value {
                    CellValue::Number(n) => n.to_string(),
                    CellValue::Text(t) => t.clone(),
                    CellValue::Boolean(b) => b.to_string(),
                    CellValue::Date(d) => d.clone(),
                    CellValue::Currency(n, code) => format!("{} {}", n, code),
                    CellValue::Percentage(n) => format!("{}%", n),
                    CellValue::Time(t) => t.clone(),
                    CellValue::Empty => String::new(),
                };
                Cell {
                    text,
                    value: value.clone(),
                    formula: None,
                    row: row_index,
                    col,
                }
            })
            .collect();

        let row = Row {
            cells,
            index: row_index,
        };

        if let Some(sheet) = self.sheets.last_mut() {
            sheet.rows.push(row);
        }

        Ok(self)
    }

    /// Set a cell value at a specific position in the current sheet
    ///
    /// # Arguments
    ///
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Cell value
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::{SpreadsheetBuilder, CellValue};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.set_cell(0, 0, CellValue::Number(42.0))?;
    /// builder.set_cell(0, 1, CellValue::Text("Hello".to_string()))?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellValue) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        if let Some(sheet) = self.sheets.last_mut() {
            // Ensure we have enough rows
            while sheet.rows.len() <= row {
                sheet.rows.push(Row {
                    cells: Vec::new(),
                    index: sheet.rows.len(),
                });
            }

            let row_data = &mut sheet.rows[row];

            // Ensure we have enough cells in the row
            while row_data.cells.len() <= col {
                row_data.cells.push(Cell {
                    text: String::new(),
                    value: CellValue::Empty,
                    formula: None,
                    row,
                    col: row_data.cells.len(),
                });
            }

            // Set the cell value
            let text = match &value {
                CellValue::Number(n) => n.to_string(),
                CellValue::Text(t) => t.clone(),
                CellValue::Boolean(b) => b.to_string(),
                CellValue::Date(d) => d.clone(),
                CellValue::Currency(n, code) => format!("{} {}", n, code),
                CellValue::Percentage(n) => format!("{}%", n),
                CellValue::Time(t) => t.clone(),
                CellValue::Empty => String::new(),
            };

            row_data.cells[col] = Cell {
                text,
                value,
                formula: None,
                row,
                col,
            };
        }

        Ok(self)
    }

    /// Set a cell formula at a specific position in the current sheet
    ///
    /// # Arguments
    ///
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `formula` - The formula (e.g., "=SUM(A1:A10)")
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.set_cell_formula(0, 0, "=SUM(A2:A10)")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_cell_formula(&mut self, row: usize, col: usize, formula: &str) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        if let Some(sheet) = self.sheets.last_mut() {
            // Ensure we have enough rows
            while sheet.rows.len() <= row {
                sheet.rows.push(Row {
                    cells: Vec::new(),
                    index: sheet.rows.len(),
                });
            }

            let row_data = &mut sheet.rows[row];

            // Ensure we have enough cells in the row
            while row_data.cells.len() <= col {
                row_data.cells.push(Cell {
                    text: String::new(),
                    value: CellValue::Empty,
                    formula: None,
                    row,
                    col: row_data.cells.len(),
                });
            }

            // Set the formula
            row_data.cells[col].formula = Some(formula.to_string());
        }

        Ok(self)
    }

    /// Select a specific sheet by index for subsequent operations
    ///
    /// # Arguments
    ///
    /// * `index` - Sheet index (0-based)
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.add_sheet("Sheet2")?;
    /// builder.select_sheet(0)?; // Go back to Sheet1
    /// builder.add_row_with_values(&["Data for Sheet1"])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn select_sheet(&mut self, index: usize) -> Result<&mut Self> {
        if index >= self.sheets.len() {
            return Err(crate::Error::Other(format!(
                "Sheet index {} out of bounds (have {} sheets)",
                index,
                self.sheets.len()
            )));
        }

        // Move the selected sheet to the end (current working sheet)
        let sheet = self.sheets.remove(index);
        self.sheets.push(sheet);

        Ok(self)
    }

    /// Add a row with typed cell values
    ///
    /// # Arguments
    ///
    /// * `cells` - Cell values for the row
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::{SpreadsheetBuilder, Cell, CellValue};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    ///
    /// let cells = vec![
    ///     Cell {
    ///         text: "100".to_string(),
    ///         value: CellValue::Number(100.0),
    ///         formula: None,
    ///         row: 0,
    ///         col: 0,
    ///     },
    /// ];
    /// builder.add_row(cells)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_row(&mut self, cells: Vec<Cell>) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        let row_index = if let Some(sheet) = self.sheets.last() {
            sheet.rows.len()
        } else {
            0
        };

        let row = Row {
            cells,
            index: row_index,
        };

        if let Some(sheet) = self.sheets.last_mut() {
            sheet.rows.push(row);
        }

        Ok(self)
    }

    /// Add a Sheet element directly
    ///
    /// # Arguments
    ///
    /// * `sheet` - A complete `Sheet` element to add
    pub fn add_sheet_element(&mut self, sheet: Sheet) -> Result<&mut Self> {
        self.sheets.push(sheet);
        Ok(self)
    }

    /// Generate the content.xml body for spreadsheet
    fn generate_content_body(&self) -> String {
        let mut body = String::new();

        for sheet in &self.sheets {
            body.push_str(&format!(
                r#"      <table:table table:name="{}" table:style-name="ta1">
"#,
                Self::escape_xml(&sheet.name)
            ));

            // Add column definitions
            body.push_str(
                r#"        <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
"#,
            );

            // Add rows
            for row in &sheet.rows {
                body.push_str("        <table:table-row>\n");

                // Add cells
                for cell in &row.cells {
                    match &cell.value {
                        CellValue::Text(_) => {
                            body.push_str(&format!(
                                r#"          <table:table-cell office:value-type="string">
            <text:p>{}</text:p>
          </table:table-cell>
"#,
                                Self::escape_xml(&cell.text)
                            ));
                        },
                        CellValue::Number(f) => {
                            body.push_str(&format!(
                                r#"          <table:table-cell office:value-type="float" office:value="{}">
            <text:p>{}</text:p>
          </table:table-cell>
"#,
                                f,
                                Self::escape_xml(&cell.text)
                            ));
                        },
                        CellValue::Currency(f, currency) => {
                            body.push_str(&format!(
                                r#"          <table:table-cell office:value-type="currency" office:value="{}" office:currency="{}">
            <text:p>{}</text:p>
          </table:table-cell>
"#,
                                f,
                                Self::escape_xml(currency),
                                Self::escape_xml(&cell.text)
                            ));
                        },
                        CellValue::Percentage(f) => {
                            body.push_str(&format!(
                                r#"          <table:table-cell office:value-type="percentage" office:value="{}">
            <text:p>{}</text:p>
          </table:table-cell>
"#,
                                f,
                                Self::escape_xml(&cell.text)
                            ));
                        },
                        CellValue::Date(d) => {
                            body.push_str(&format!(
                                r#"          <table:table-cell office:value-type="date" office:date-value="{}">
            <text:p>{}</text:p>
          </table:table-cell>
"#,
                                d,
                                Self::escape_xml(&cell.text)
                            ));
                        },
                        CellValue::Time(t) => {
                            body.push_str(&format!(
                                r#"          <table:table-cell office:value-type="time" office:time-value="{}">
            <text:p>{}</text:p>
          </table:table-cell>
"#,
                                t,
                                Self::escape_xml(&cell.text)
                            ));
                        },
                        CellValue::Boolean(b) => {
                            body.push_str(&format!(
                                r#"          <table:table-cell office:value-type="boolean" office:boolean-value="{}">
            <text:p>{}</text:p>
          </table:table-cell>
"#,
                                b,
                                Self::escape_xml(&cell.text)
                            ));
                        },
                        CellValue::Empty => {
                            // Empty cell
                            body.push_str("          <table:table-cell/>\n");
                        },
                    }
                }

                body.push_str("        </table:table-row>\n");
            }

            body.push_str("      </table:table>\n");
        }

        body
    }

    /// Generate the complete content.xml for spreadsheet
    fn generate_content_xml(&self) -> String {
        let body = self.generate_content_body();

        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                          xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                          xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                          xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
                          xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                          xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
                          xmlns:xlink="http://www.w3.org/1999/xlink"
                          xmlns:dc="http://purl.org/dc/elements/1.1/"
                          xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                          xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
                          xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
                          xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
                          xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
                          xmlns:math="http://www.w3.org/1998/Math/MathML"
                          xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
                          xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
                          xmlns:ooo="http://openoffice.org/2004/office"
                          xmlns:ooow="http://openoffice.org/2004/writer"
                          xmlns:oooc="http://openoffice.org/2004/calc"
                          xmlns:dom="http://www.w3.org/2001/xml-events"
                          office:version="1.3">
  <office:scripts/>
  <office:font-face-decls/>
  <office:automatic-styles/>
  <office:body>
    <office:spreadsheet>
{}    </office:spreadsheet>
  </office:body>
</office:document-content>
"#,
            body
        )
    }

    /// Generate meta.xml with metadata
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();

        let mut meta = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                       xmlns:xlink="http://www.w3.org/1999/xlink"
                       xmlns:dc="http://purl.org/dc/elements/1.1/"
                       xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                       office:version="1.3">
  <office:meta>
    <meta:generator>Litchi/0.0.1</meta:generator>
    <meta:creation-date>{}</meta:creation-date>
    <dc:date>{}</dc:date>
"#,
            now, now
        );

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            meta.push_str(&format!(
                "    <dc:title>{}</dc:title>\n",
                Self::escape_xml(title)
            ));
        }

        if let Some(ref author) = self.metadata.author {
            meta.push_str(&format!(
                "    <dc:creator>{}</dc:creator>\n",
                Self::escape_xml(author)
            ));
        }

        meta.push_str("  </office:meta>\n");
        meta.push_str("</office:document-meta>\n");

        meta
    }

    /// Escape XML special characters
    fn escape_xml(text: &str) -> String {
        text.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
    }

    /// Build the spreadsheet and return as bytes
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// let bytes = builder.build()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn build(self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype("application/vnd.oasis.opendocument.spreadsheet")?;

        // Add content.xml
        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml
        let styles_xml = OdfStructure::default_styles_xml();
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Finish and return bytes
        writer.finish_to_bytes()
    }

    /// Build and save the spreadsheet to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODS file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.save("output.ods")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let bytes = self.build()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}
