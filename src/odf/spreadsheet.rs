//! OpenDocument Spreadsheet (.ods) support.
//!
//! This module provides a unified API for working with OpenDocument spreadsheets,
//! equivalent to Microsoft Excel spreadsheets.

use crate::common::{Error, Result, Metadata};
use crate::odf::core::{Content, Meta, Package, Styles, Manifest};
use std::io::Cursor;
use std::path::Path;

/// Cell data types supported by ODF spreadsheets
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell
    Empty,
    /// Text string
    Text(String),
    /// Numeric value
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Date/time value
    Date(String), // Using string for now, could be parsed to chrono types later
    /// Currency value
    Currency(f64, String), // (value, currency_code)
    /// Percentage value
    Percentage(f64),
    /// Time duration
    Time(String),
}

/// A cell in an ODS sheet
#[derive(Clone, Debug)]
pub struct Cell {
    /// The cell value
    pub value: CellValue,
    /// The raw text content of the cell
    pub text: String,
    /// The formula in the cell (if any)
    pub formula: Option<String>,
    /// The row index (0-based)
    pub row: usize,
    /// The column index (0-based)
    pub col: usize,
}

/// An OpenDocument spreadsheet (.ods)
pub struct Spreadsheet {
    _package: Package<Cursor<Vec<u8>>>,
    _content: Content,
    _styles: Option<Styles>,
    _meta: Option<Meta>,
    _manifest: Manifest,
}

impl Spreadsheet {
    /// Open an ODS spreadsheet from a file path
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref();

        // Read the entire file into memory
        let bytes = std::fs::read(path)?;
        Self::from_bytes(bytes)
    }

    /// Create a Spreadsheet from a byte buffer
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let mut package = Package::from_reader(cursor)?;

        // Verify this is a spreadsheet
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.spreadsheet") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODS file: MIME type is {}", mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        let manifest = package.manifest().clone();

        Ok(Self {
            _package: package,
            _content: content,
            _styles: styles,
            _meta: meta,
            _manifest: manifest,
        })
    }

    /// Get the number of sheets in the spreadsheet
    pub fn sheet_count(&mut self) -> Result<usize> {
        let sheets = self.sheets()?;
        Ok(sheets.len())
    }

    /// Get all sheets in the spreadsheet
    pub fn sheets(&mut self) -> Result<Vec<Sheet>> {
        // For now, use a simpler approach to avoid crashes
        // TODO: Implement proper table parsing with the new elements
        let content_bytes = self._package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(content.xml_content());
        let mut buf = Vec::new();
        let mut sheets = Vec::new();

        // Parser state
        let mut current_sheet: Option<SheetBuilder> = None;
        let mut current_row: Option<RowBuilder> = None;
        let mut current_cell: Option<CellBuilder> = None;
        let mut in_text_element = false;
        let mut text_content = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"table:table" => {
                            let name = Self::extract_table_name(e)?;
                            current_sheet = Some(SheetBuilder::new(name));
                        }
                        b"table:table-row" => {
                            if current_sheet.is_some() {
                                current_row = Some(RowBuilder::new());
                            }
                        }
                        b"table:table-cell" => {
                            if current_row.is_some() {
                                let cell_builder = Self::parse_cell_attributes(e)?;
                                current_cell = Some(cell_builder);
                                text_content.clear();
                            }
                        }
                        b"text:p" | b"text:span" => {
                            if current_cell.is_some() {
                                in_text_element = true;
                                if e.name().as_ref() == b"text:p" {
                                    text_content.clear();
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(ref t)) => {
                    if in_text_element && current_cell.is_some() {
                        let text = String::from_utf8(t.to_vec()).unwrap_or_default();
                        text_content.push_str(&text);
                    }
                }
                Ok(Event::End(ref e)) => {
                    match e.name().as_ref() {
                        b"text:p" | b"text:span" => {
                            if in_text_element {
                                in_text_element = false;
                            }
                        }
                        b"table:table-cell" => {
                            if let Some(cell_builder) = current_cell.take() {
                                let repeated = cell_builder.repeated;
                                let cell = cell_builder.build(text_content.clone());
                                if let Some(ref mut row_builder) = current_row {
                                    // Handle repeated cells
                                    for _ in 0..repeated {
                                        row_builder.add_cell(cell.clone());
                                    }
                                }
                            }
                        }
                        b"table:table-row" => {
                            if let Some(row_builder) = current_row.take() {
                                let row = row_builder.build();
                                if let Some(ref mut sheet_builder) = current_sheet {
                                    sheet_builder.add_row(row);
                                }
                            }
                        }
                        b"table:table" => {
                            if let Some(sheet_builder) = current_sheet.take() {
                                let sheet = sheet_builder.build();
                                sheets.push(sheet);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(crate::common::Error::InvalidFormat(format!("XML parsing error: {}", e))),
                _ => {}
            }
            buf.clear();
        }

        Ok(sheets)
    }

    /// Extract table name from table:table element
    fn extract_table_name(e: &quick_xml::events::BytesStart) -> Result<String> {
        for attr_result in e.attributes() {
            let attr = attr_result.map_err(|_| crate::common::Error::InvalidFormat("Invalid attribute".to_string()))?;
            if attr.key.as_ref() == b"table:name" {
                return String::from_utf8(attr.value.to_vec())
                    .map_err(|_| crate::common::Error::InvalidFormat("Invalid UTF-8 in table name".to_string()));
            }
        }
        Ok("Sheet1".to_string()) // Default name
    }

    /// Parse cell attributes and create a CellBuilder
    fn parse_cell_attributes(e: &quick_xml::events::BytesStart) -> Result<CellBuilder> {
        let mut value_type = None;
        let mut value_str = None;
        let mut currency = None;
        let mut formula = None;
        let mut repeated = 1;

        for attr_result in e.attributes() {
            let attr = attr_result.map_err(|_| crate::common::Error::InvalidFormat("Invalid attribute".to_string()))?;
            match attr.key.as_ref() {
                b"office:value-type" => {
                    value_type = Some(String::from_utf8(attr.value.to_vec())
                        .map_err(|_| crate::common::Error::InvalidFormat("Invalid UTF-8".to_string()))?);
                }
                b"office:value" => {
                    value_str = Some(String::from_utf8(attr.value.to_vec())
                        .map_err(|_| crate::common::Error::InvalidFormat("Invalid UTF-8".to_string()))?);
                }
                b"office:currency" => {
                    currency = Some(String::from_utf8(attr.value.to_vec())
                        .map_err(|_| crate::common::Error::InvalidFormat("Invalid UTF-8".to_string()))?);
                }
                b"table:formula" => {
                    formula = Some(String::from_utf8(attr.value.to_vec())
                        .map_err(|_| crate::common::Error::InvalidFormat("Invalid UTF-8".to_string()))?);
                }
                b"table:number-columns-repeated" => {
                    if let Ok(rep) = String::from_utf8(attr.value.to_vec())
                        .map_err(|_| crate::common::Error::InvalidFormat("Invalid UTF-8".to_string()))?
                        .parse::<usize>() {
                        repeated = rep;
                    }
                }
                _ => {}
            }
        }

        Ok(CellBuilder {
            value_type,
            value_str,
            currency,
            formula,
            repeated,
        })
    }

    /// Get a sheet by name
    pub fn sheet_by_name(&mut self, name: &str) -> Result<Option<Sheet>> {
        let sheets = self.sheets()?;
        Ok(sheets.into_iter().find(|sheet| sheet.name == name))
    }

    /// Get a sheet by index
    pub fn sheet_by_index(&mut self, index: usize) -> Result<Option<Sheet>> {
        let sheets = self.sheets()?;
        Ok(sheets.into_iter().nth(index))
    }

    /// Extract all text content from the spreadsheet
    pub fn text(&mut self) -> Result<String> {
        let sheets = self.sheets()?;
        let mut all_text = Vec::new();

        for sheet in sheets {
            for row in sheet.rows {
                for cell in row.cells {
                    if !cell.text.trim().is_empty() {
                        all_text.push(cell.text.trim().to_string());
                    }
                }
            }
        }

        Ok(all_text.join("\n"))
    }

    /// Export spreadsheet data as CSV
    pub fn to_csv(&mut self) -> Result<String> {
        let sheets = self.sheets()?;
        let mut csv_output = String::new();

        for (sheet_index, sheet) in sheets.iter().enumerate() {
            if sheet_index > 0 {
                csv_output.push_str("\n\n"); // Separate sheets with double newline
            }

            for (row_index, row) in sheet.rows.iter().enumerate() {
                if row_index > 0 {
                    csv_output.push('\n');
                }

                for (col_index, cell) in row.cells.iter().enumerate() {
                    if col_index > 0 {
                        csv_output.push(',');
                    }

                    // Escape CSV special characters and wrap in quotes if needed
                    let cell_text = &cell.text;
                    if cell_text.contains(',') || cell_text.contains('"') || cell_text.contains('\n') {
                        let escaped = cell_text.replace('"', "\"\"");
                        csv_output.push('"');
                        csv_output.push_str(&escaped);
                        csv_output.push('"');
                    } else {
                        csv_output.push_str(cell_text);
                    }
                }
            }
        }

        Ok(csv_output)
    }

    /// Get document metadata
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self._meta {
            Ok(meta.extract_metadata())
        } else {
            Ok(Metadata::default())
        }
    }
}

/// A sheet (worksheet) in an ODS spreadsheet
#[derive(Clone)]
pub struct Sheet {
    pub name: String,
    pub rows: Vec<Row>,
}

impl Sheet {
    /// Get the name of the sheet
    pub fn name(&self) -> Result<String> {
        Ok(self.name.clone())
    }

    /// Get all rows in the sheet
    pub fn rows(&self) -> Result<Vec<Row>> {
        Ok(self.rows.clone())
    }

    /// Get the number of rows in the sheet
    pub fn row_count(&self) -> Result<usize> {
        Ok(self.rows.len())
    }

    /// Get the number of columns in the sheet
    pub fn column_count(&self) -> Result<usize> {
        let max_cols = self.rows.iter()
            .map(|row| row.cells.len())
            .max()
            .unwrap_or(0);
        Ok(max_cols)
    }
}

/// A row in an ODS sheet
#[derive(Clone)]
pub struct Row {
    pub cells: Vec<Cell>,
    pub index: usize,
}

impl Row {
    /// Get all cells in the row
    pub fn cells(&self) -> Result<Vec<Cell>> {
        Ok(self.cells.clone())
    }

    /// Get a cell by column index
    pub fn cell(&self, col: usize) -> Result<Option<Cell>> {
        if col < self.cells.len() {
            Ok(Some(self.cells[col].clone()))
        } else {
            Ok(None)
        }
    }

    /// Get the row index
    pub fn index(&self) -> usize {
        self.index
    }
}

/// Builder for constructing Sheet during parsing
struct SheetBuilder {
    name: String,
    rows: Vec<Row>,
}

impl SheetBuilder {
    fn new(name: String) -> Self {
        Self {
            name,
            rows: Vec::new(),
        }
    }

    fn add_row(&mut self, mut row: Row) {
        let row_index = self.rows.len();
        row.index = row_index;
        // Update row index for all cells in this row
        for cell in &mut row.cells {
            cell.row = row_index;
        }
        self.rows.push(row);
    }

    fn build(self) -> Sheet {
        Sheet {
            name: self.name,
            rows: self.rows,
        }
    }
}

/// Builder for constructing Row during parsing
struct RowBuilder {
    cells: Vec<Cell>,
}

impl RowBuilder {
    fn new() -> Self {
        Self {
            cells: Vec::new(),
        }
    }

    fn add_cell(&mut self, mut cell: Cell) {
        cell.col = self.cells.len();
        self.cells.push(cell);
    }

    fn build(mut self) -> Row {
        // Row index will be set by the parent SheetBuilder
        // For now, set to 0 and update cells
        for cell in &mut self.cells {
            cell.row = 0; // Will be updated by parent
        }

        Row {
            cells: self.cells,
            index: 0, // Will be set by parent
        }
    }
}

/// Builder for constructing Cell during parsing
struct CellBuilder {
    value_type: Option<String>,
    value_str: Option<String>,
    currency: Option<String>,
    formula: Option<String>,
    repeated: usize,
}

impl CellBuilder {
    fn build(self, text_content: String) -> Cell {
        let value = self.parse_value(&text_content);

        Cell {
            value,
            text: text_content,
            formula: self.formula,
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
            }
            Some("currency") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        let currency_code = self.currency.clone().unwrap_or_else(|| "USD".to_string());
                        CellValue::Currency(num, currency_code)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            }
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
            }
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
            }
            Some("date") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Date(val_str.clone())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            }
            Some("time") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Time(val_str.clone())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            }
            _ => {
                if text_content.trim().is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::Text(text_content.to_string())
                }
            }
        }
    }
}

impl Cell {
    /// Get the text content of the cell
    pub fn text(&self) -> Result<String> {
        Ok(self.text.clone())
    }

    /// Get the cell value
    pub fn value(&self) -> Result<CellValue> {
        Ok(self.value.clone())
    }

    /// Get the numeric value of the cell (if applicable)
    pub fn numeric_value(&self) -> Result<Option<f64>> {
        match &self.value {
            CellValue::Number(n) => Ok(Some(*n)),
            CellValue::Currency(n, _) => Ok(Some(*n)),
            CellValue::Percentage(p) => Ok(Some(*p)),
            _ => Ok(None),
        }
    }

    /// Get the formula in the cell
    pub fn formula(&self) -> Result<Option<String>> {
        Ok(self.formula.clone())
    }

    /// Get the cell coordinates (row, column)
    pub fn coordinates(&self) -> (usize, usize) {
        (self.row, self.col)
    }

    /// Check if the cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self.value, CellValue::Empty)
    }
}
