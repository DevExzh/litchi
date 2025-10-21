//! ODS-specific parsing utilities.

use super::{Cell, CellValue, Row, Sheet};
use crate::common::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

/// Parser for ODS-specific structures.
///
/// This provides parsing logic specific to spreadsheets,
/// including sheet, row, and cell parsing with proper type detection.
pub(crate) struct OdsParser;

impl OdsParser {
    /// Parse all sheets from ODS content.xml
    pub fn parse_sheets(xml_content: &str) -> Result<Vec<Sheet>> {
        let mut reader = Reader::from_str(xml_content);
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
                Ok(Event::Start(ref e)) => match e.name().as_ref() {
                    b"table:table" => {
                        let name = Self::extract_table_name(e)?;
                        current_sheet = Some(SheetBuilder::new(name));
                    },
                    b"table:table-row" => {
                        if current_sheet.is_some() {
                            current_row = Some(RowBuilder::new());
                        }
                    },
                    b"table:table-cell" => {
                        if current_row.is_some() {
                            let cell_builder = Self::parse_cell_attributes(e)?;
                            current_cell = Some(cell_builder);
                            text_content.clear();
                        }
                    },
                    b"text:p" | b"text:span" => {
                        if current_cell.is_some() {
                            in_text_element = true;
                            if e.name().as_ref() == b"text:p" {
                                text_content.clear();
                            }
                        }
                    },
                    _ => {},
                },
                Ok(Event::Text(ref t)) => {
                    if in_text_element && current_cell.is_some() {
                        let text = String::from_utf8(t.to_vec()).unwrap_or_default();
                        text_content.push_str(&text);
                    }
                },
                Ok(Event::End(ref e)) => {
                    match e.name().as_ref() {
                        b"text:p" | b"text:span" => {
                            if in_text_element {
                                in_text_element = false;
                            }
                        },
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
                        },
                        b"table:table-row" => {
                            if let Some(row_builder) = current_row.take() {
                                let row = row_builder.build();
                                if let Some(ref mut sheet_builder) = current_sheet {
                                    sheet_builder.add_row(row);
                                }
                            }
                        },
                        b"table:table" => {
                            if let Some(sheet_builder) = current_sheet.take() {
                                let sheet = sheet_builder.build();
                                sheets.push(sheet);
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(Error::InvalidFormat(format!("XML parsing error: {}", e)));
                },
                _ => {},
            }
            buf.clear();
        }

        Ok(sheets)
    }

    /// Extract table name from table:table element
    fn extract_table_name(e: &quick_xml::events::BytesStart) -> Result<String> {
        for attr_result in e.attributes() {
            let attr =
                attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute".to_string()))?;
            if attr.key.as_ref() == b"table:name" {
                return String::from_utf8(attr.value.to_vec())
                    .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in table name".to_string()));
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
            let attr =
                attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute".to_string()))?;
            match attr.key.as_ref() {
                b"office:value-type" => {
                    value_type = Some(
                        String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
                    );
                },
                b"office:value" => {
                    value_str = Some(
                        String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
                    );
                },
                b"office:currency" => {
                    currency = Some(
                        String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
                    );
                },
                b"table:formula" => {
                    formula = Some(
                        String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
                    );
                },
                b"table:number-columns-repeated" => {
                    if let Ok(rep) = String::from_utf8(attr.value.to_vec())
                        .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?
                        .parse::<usize>()
                    {
                        repeated = rep;
                    }
                },
                _ => {},
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
}

/// Builder for constructing Sheet during parsing
pub(crate) struct SheetBuilder {
    name: String,
    rows: Vec<Row>,
}

impl SheetBuilder {
    pub fn new(name: String) -> Self {
        Self {
            name,
            rows: Vec::new(),
        }
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

    pub fn build(self) -> Sheet {
        Sheet {
            name: self.name,
            rows: self.rows,
        }
    }
}

/// Builder for constructing Row during parsing
pub(crate) struct RowBuilder {
    cells: Vec<Cell>,
}

impl RowBuilder {
    pub fn new() -> Self {
        Self { cells: Vec::new() }
    }

    pub fn add_cell(&mut self, mut cell: Cell) {
        cell.col = self.cells.len();
        self.cells.push(cell);
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
        }
    }
}

/// Builder for constructing Cell during parsing
pub(crate) struct CellBuilder {
    value_type: Option<String>,
    value_str: Option<String>,
    currency: Option<String>,
    formula: Option<String>,
    repeated: usize,
}

impl CellBuilder {
    pub fn build(self, text_content: String) -> Cell {
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
            },
            Some("currency") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        let currency_code =
                            self.currency.clone().unwrap_or_else(|| "USD".to_string());
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
                    CellValue::Date(val_str.clone())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("time") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Time(val_str.clone())
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
