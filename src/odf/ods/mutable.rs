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
            body.push_str(&xml_minifier::minified_xml_format!(
                r#"<table:table table:name="{}">"#,
                escaped_name
            ));

            body.push_str(xml_minifier::minified_xml_str!(
                r#"<table:table-column table:style-name="co1"/>"#
            ));

            for row in &sheet.rows {
                body.push_str("<table:table-row>");

                for cell in &row.cells {
                    match &cell.value {
                        CellValue::Text(_) => {
                            let escaped_text = escape_xml(&cell.text);
                            body.push_str(&xml_minifier::minified_xml_format!(
                                r#"<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
                                escaped_text
                            ));
                        },
                        CellValue::Number(f) => {
                            let escaped_text = escape_xml(&cell.text);
                            body.push_str(&xml_minifier::minified_xml_format!(
                                r#"<table:table-cell office:value-type="float" office:value="{}"><text:p>{}</text:p></table:table-cell>"#,
                                f,
                                escaped_text
                            ));
                        },
                        CellValue::Empty => {
                            body.push_str(xml_minifier::minified_xml_str!(
                                r#"<table:table-cell/>"#
                            ));
                        },
                        _ => {
                            // Handle other types similarly
                            let escaped_text = escape_xml(&cell.text);
                            body.push_str(&xml_minifier::minified_xml_format!(
                                r#"<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
                                escaped_text
                            ));
                        },
                    }
                }

                body.push_str("</table:table-row>");
            }

            body.push_str("</table:table>");
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:spreadsheet>{}</office:spreadsheet></office:body></office:document-content>"#,
            body
        )
    }

    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        xml_minifier::minified_xml_format!(
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
