//! Mutable spreadsheet structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODS spreadsheets that allows
//! for in-place modification of sheets, rows, and cells.

use crate::core::{OdfStructure, PackageWriter};
use crate::ods::{
    Cell, CellAnnotation, CellValue, NamedDefinition, NamedDefinitionScope, NamedExpression,
    NamedRange, Row, Sheet, Spreadsheet,
    named_expression::{ensure_unique, write_named_definitions},
};
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::path::Path;

/// A mutable ODS spreadsheet that supports in-place modifications.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::{Spreadsheet, MutableSpreadsheet};
///
/// # fn main() -> litchi_core::Result<()> {
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
    /// Global and sheet-local named ranges and expressions.
    named_definitions: Vec<NamedDefinition>,
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
            || self
                .named_definitions
                .iter()
                .any(|definition| matches!(definition, NamedDefinition::Expression(_)))
    }

    fn has_annotations(&self) -> bool {
        self.sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .any(Cell::has_annotation)
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
        super::cell::write_cell_xml(out, cell);
    }

    /// Create a mutable spreadsheet from an existing Spreadsheet.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::{Spreadsheet, MutableSpreadsheet};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)?;
    /// mutable.add_sheet("NewSheet")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_spreadsheet(mut spreadsheet: Spreadsheet) -> Result<Self> {
        let sheets = spreadsheet.sheets()?;
        let metadata = spreadsheet.metadata()?;
        let named_definitions = spreadsheet.named_definitions().to_vec();
        let mimetype = "application/vnd.oasis.opendocument.spreadsheet".to_string();

        // Extract styles XML from the spreadsheet's package (requires accessing internal package)
        // For now, we'll use None and rely on default styles
        // TODO: Add method to Spreadsheet to expose get_file for extracting styles.xml

        Ok(Self {
            sheets,
            metadata,
            mimetype,
            styles_xml: None,
            named_definitions,
        })
    }

    /// Create a new empty mutable spreadsheet.
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.spreadsheet".to_string(),
            styles_xml: None,
            named_definitions: Vec::new(),
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

    /// Return all global and sheet-local named definitions.
    pub fn named_definitions(&self) -> &[NamedDefinition] {
        &self.named_definitions
    }

    /// Add a named range.
    pub fn add_named_range(&mut self, range: NamedRange) -> Result<()> {
        self.add_named_definition(range.into())
    }

    /// Add a named OpenFormula expression.
    pub fn add_named_expression(&mut self, expression: NamedExpression) -> Result<()> {
        self.add_named_definition(expression.into())
    }

    /// Add either kind of named definition.
    pub fn add_named_definition(&mut self, definition: NamedDefinition) -> Result<()> {
        definition.validate()?;
        self.validate_scope(definition.scope())?;
        ensure_unique(&self.named_definitions, &definition)?;
        self.named_definitions.push(definition);
        Ok(())
    }

    /// Remove a named definition and return it if it exists.
    pub fn remove_named_definition(
        &mut self,
        name: &str,
        scope: &NamedDefinitionScope,
    ) -> Option<NamedDefinition> {
        let index = self
            .named_definitions
            .iter()
            .position(|definition| definition.name() == name && definition.scope() == scope)?;
        Some(self.named_definitions.remove(index))
    }

    fn validate_scope(&self, scope: &NamedDefinitionScope) -> Result<()> {
        if let NamedDefinitionScope::Sheet(sheet_name) = scope
            && !self.sheets.iter().any(|sheet| sheet.name == *sheet_name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "named definition refers to missing sheet '{sheet_name}'"
            )));
        }
        Ok(())
    }

    fn validate_named_definitions(&self) -> Result<()> {
        for (index, definition) in self.named_definitions.iter().enumerate() {
            definition.validate()?;
            self.validate_scope(definition.scope())?;
            if self.named_definitions[..index].iter().any(|existing| {
                existing.name() == definition.name() && existing.scope() == definition.scope()
            }) {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "duplicate named definition '{}' in {:?}",
                    definition.name(),
                    definition.scope()
                )));
            }
        }
        Ok(())
    }

    fn validate_annotations(&self) -> Result<()> {
        for annotation in self
            .sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .filter_map(Cell::annotation)
        {
            annotation.validate()?;
        }
        Ok(())
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
            let sheet = self.sheets.remove(index);
            self.named_definitions
                .retain(|definition| definition.scope().sheet_name() != Some(sheet.name.as_str()));
            Ok(sheet)
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
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
            Err(litchi_core::Error::InvalidFormat(format!(
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
                Err(litchi_core::Error::InvalidFormat(format!(
                    "Row index {} out of bounds",
                    row_index
                )))
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::{MutableSpreadsheet, CellValue};
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
                    annotation: None,
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
            Err(litchi_core::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Attach or replace an annotation on a cell.
    pub fn set_cell_annotation(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        annotation: CellAnnotation,
    ) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.rows.len() <= row {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
            });
        }
        let row_data = &mut sheet.rows[row];
        while row_data.cells.len() <= col {
            row_data.cells.push(Cell {
                value: CellValue::Empty,
                text: String::new(),
                formula: None,
                annotation: None,
                row,
                col: row_data.cells.len(),
            });
        }
        row_data.cells[col].annotation = Some(annotation);
        Ok(())
    }

    /// Remove and return an annotation from a cell.
    pub fn remove_cell_annotation(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
    ) -> Result<Option<CellAnnotation>> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        Ok(sheet
            .rows
            .get_mut(row)
            .and_then(|row| row.cells.get_mut(col))
            .and_then(Cell::take_annotation))
    }

    /// Clear a cell value.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableSpreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
                    Err(litchi_core::Error::InvalidFormat(format!(
                        "Column index {} out of bounds",
                        col
                    )))
                }
            } else {
                Err(litchi_core::Error::InvalidFormat(format!(
                    "Row index {} out of bounds",
                    row
                )))
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutableSpreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
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

            write_named_definitions(
                &mut body,
                self.named_definitions.iter().filter(|definition| {
                    definition.scope().sheet_name() == Some(sheet.name.as_str())
                }),
            );

            body.push_str("</table:table>");
        }

        write_named_definitions(
            &mut body,
            self.named_definitions
                .iter()
                .filter(|definition| definition.scope() == &NamedDefinitionScope::Global),
        );

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
        if self.has_annotations() {
            out.push_str(r#" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""#);
        }
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
        self.validate_named_definitions()?;
        self.validate_annotations()?;
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ods::{NamedRangeUsage, SpreadsheetBuilder};

    #[test]
    fn mutable_spreadsheet_preserves_and_edits_named_definitions() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder
            .add_named_range(
                NamedRange::new(
                    "LocalPrintArea",
                    "$Sheet1.$A$1:.$C$9",
                    NamedDefinitionScope::sheet("Sheet1"),
                )
                .unwrap()
                .with_usage(NamedRangeUsage::PrintRange),
            )
            .unwrap();
        let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();

        let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        assert_eq!(mutable.named_definitions().len(), 1);
        mutable
            .add_named_expression(
                NamedExpression::new("GlobalValue", "of:=42", NamedDefinitionScope::Global)
                    .unwrap(),
            )
            .unwrap();

        let output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert!(
            output
                .named_range("LocalPrintArea", &NamedDefinitionScope::sheet("Sheet1"))
                .is_some()
        );
        assert!(
            output
                .named_expression("GlobalValue", &NamedDefinitionScope::Global)
                .is_some()
        );
    }

    #[test]
    fn removing_sheet_removes_its_local_definitions_only() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Sheet1").unwrap();
        mutable
            .add_named_range(
                NamedRange::new(
                    "Local",
                    "$Sheet1.$A$1",
                    NamedDefinitionScope::sheet("Sheet1"),
                )
                .unwrap(),
            )
            .unwrap();
        mutable
            .add_named_expression(
                NamedExpression::new("Global", "of:=1", NamedDefinitionScope::Global).unwrap(),
            )
            .unwrap();

        mutable.remove_sheet(0).unwrap();
        assert_eq!(mutable.named_definitions().len(), 1);
        assert_eq!(mutable.named_definitions()[0].name(), "Global");
    }

    #[test]
    fn mutable_spreadsheet_adds_edits_and_removes_annotations() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Sheet1").unwrap();
        let mut annotation = CellAnnotation::new("review this");
        annotation.set_creator(Some("Reviewer"));
        mutable.set_cell_annotation(0, 3, 4, annotation).unwrap();

        mutable.sheets_mut()[0].rows[3].cells[4]
            .annotation_mut()
            .unwrap()
            .push_paragraph("follow-up");
        let mut round_trip = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let sheets = round_trip.sheets().unwrap();
        let annotation = sheets[0].rows[3].cells[4].annotation().unwrap();
        assert_eq!(annotation.creator().as_deref(), Some("Reviewer"));
        assert_eq!(annotation.text(), "review this\nfollow-up");

        assert!(mutable.remove_cell_annotation(0, 3, 4).unwrap().is_some());
        assert!(mutable.remove_cell_annotation(0, 3, 4).unwrap().is_none());
    }
}
