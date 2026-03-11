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
                                if let Some(ref mut row_builder) = current_row {
                                    // Handle repeated cells - build cell once for first occurrence
                                    if repeated > 0 {
                                        let cell = cell_builder.build(&text_content);
                                        row_builder.add_cell(cell);
                                        // Clone only for additional repetitions
                                        for _ in 1..repeated {
                                            row_builder.add_cell(cell_builder.build(&text_content));
                                        }
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
    pub fn build(&self, text_content: &str) -> Cell {
        let value = self.parse_value(text_content);

        Cell {
            value,
            text: text_content.to_string(),
            // Clone necessary: formula may be reused for repeated cells
            formula: self.formula.clone(),
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

#[cfg(test)]
mod tests {
    use super::*;

    const TEST_SHEETS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Hello</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="float" office:value="42">
                        <text:p>42</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_MULTIPLE_SHEETS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>First Sheet</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
            <table:table table:name="Sheet2">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Second Sheet</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_CELL_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="TypesTest">
                <table:table-row>
                    <table:table-cell office:value-type="string"><text:p>Text</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="3.14"><text:p>3.14</text:p></table:table-cell>
                    <table:table-cell office:value-type="currency" office:value="100" office:currency="EUR"><text:p>€100</text:p></table:table-cell>
                    <table:table-cell office:value-type="percentage" office:value="0.5"><text:p>50%</text:p></table:table-cell>
                    <table:table-cell office:value-type="boolean" office:value="true"><text:p>TRUE</text:p></table:table-cell>
                    <table:table-cell office:value-type="date" office:value="2024-03-15"><text:p>2024-03-15</text:p></table:table-cell>
                    <table:table-cell office:value-type="time" office:value="PT12H30M00S"><text:p>12:30:00</text:p></table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_FORMULA_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="FormulaTest">
                <table:table-row>
                    <table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="20"><text:p>20</text:p></table:table-cell>
                    <table:table-cell table:formula="=SUM([.A1]:[.B1])" office:value-type="float" office:value="30">
                        <text:p>30</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_REPEATED_CELLS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="RepeatedTest">
                <table:table-row>
                    <table:table-cell table:number-columns-repeated="3" office:value-type="string">
                        <text:p>Repeated</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>Single</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_EMPTY_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="EmptySheet">
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_SPAN_TEXT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="SpanTest">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Normal text <text:span>spanned text</text:span> more text</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    #[test]
    fn test_parse_sheets_basic() {
        let sheets = OdsParser::parse_sheets(TEST_SHEETS_XML).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[0].rows.len(), 1);
    }

    #[test]
    fn test_parse_multiple_sheets() {
        let sheets = OdsParser::parse_sheets(TEST_MULTIPLE_SHEETS_XML).unwrap();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[1].name, "Sheet2");
    }

    #[test]
    fn test_parse_cell_types() {
        let sheets = OdsParser::parse_sheets(TEST_CELL_TYPES_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 7);

        // Text cell
        match &row.cells[0].value {
            CellValue::Text(t) => assert_eq!(t, "Text"),
            _ => panic!("Expected Text"),
        }

        // Float/Number cell
        match &row.cells[1].value {
            CellValue::Number(n) => assert!((n - 3.14).abs() < f64::EPSILON),
            _ => panic!("Expected Number"),
        }

        // Currency cell
        match &row.cells[2].value {
            CellValue::Currency(amount, currency) => {
                assert!((amount - 100.0).abs() < f64::EPSILON);
                assert_eq!(currency, "EUR");
            },
            _ => panic!("Expected Currency"),
        }

        // Percentage cell
        match &row.cells[3].value {
            CellValue::Percentage(p) => assert!((p - 0.5).abs() < f64::EPSILON),
            _ => panic!("Expected Percentage"),
        }

        // Boolean cell
        match &row.cells[4].value {
            CellValue::Boolean(b) => assert!(*b),
            _ => panic!("Expected Boolean"),
        }

        // Date cell
        match &row.cells[5].value {
            CellValue::Date(d) => assert_eq!(d, "2024-03-15"),
            _ => panic!("Expected Date"),
        }

        // Time cell
        match &row.cells[6].value {
            CellValue::Time(t) => assert_eq!(t, "PT12H30M00S"),
            _ => panic!("Expected Time"),
        }
    }

    #[test]
    fn test_parse_formula() {
        let sheets = OdsParser::parse_sheets(TEST_FORMULA_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 3);

        // Cell with formula
        assert_eq!(row.cells[2].formula, Some("=SUM([.A1]:[.B1])".to_string()));
        match &row.cells[2].value {
            CellValue::Number(n) => assert!((n - 30.0).abs() < f64::EPSILON),
            _ => panic!("Expected Number for formula result"),
        }
    }

    #[test]
    fn test_parse_repeated_cells() {
        let sheets = OdsParser::parse_sheets(TEST_REPEATED_CELLS_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        // 3 repeated cells + 1 single = 4 cells
        assert_eq!(row.cells.len(), 4);

        for i in 0..3 {
            match &row.cells[i].value {
                CellValue::Text(t) => assert_eq!(t, "Repeated"),
                _ => panic!("Expected Text for repeated cell {i}"),
            }
        }

        match &row.cells[3].value {
            CellValue::Text(t) => assert_eq!(t, "Single"),
            _ => panic!("Expected Text for single cell"),
        }
    }

    #[test]
    fn test_parse_empty_sheet() {
        let sheets = OdsParser::parse_sheets(TEST_EMPTY_SHEET_XML).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "EmptySheet");
        assert_eq!(sheets[0].rows.len(), 0);
    }

    #[test]
    fn test_parse_span_text() {
        let sheets = OdsParser::parse_sheets(TEST_SPAN_TEXT_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 1);

        // Text should include content from both text:p and text:span
        assert!(row.cells[0].text.contains("Normal text"));
        assert!(row.cells[0].text.contains("spanned text"));
    }

    #[test]
    fn test_extract_table_name_default() {
        // XML without table:name attribute
        let xml = r#"<?xml version="1.0"?>
<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <table:table>
    </table:table>
</office:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1"); // Default name
    }

    #[test]
    fn test_sheet_builder() {
        let mut builder = SheetBuilder::new("TestSheet".to_string());

        let row1 = Row {
            cells: vec![],
            index: 0,
        };
        builder.add_row(row1);

        let row2 = Row {
            cells: vec![Cell {
                value: CellValue::Text("A1".to_string()),
                text: "A1".to_string(),
                formula: None,
                row: 0,
                col: 0,
            }],
            index: 0,
        };
        builder.add_row(row2);

        let sheet = builder.build();
        assert_eq!(sheet.name, "TestSheet");
        assert_eq!(sheet.rows.len(), 2);
        assert_eq!(sheet.rows[0].index, 0);
        assert_eq!(sheet.rows[1].index, 1);
    }

    #[test]
    fn test_row_builder() {
        let mut builder = RowBuilder::new();

        let cell1 = Cell {
            value: CellValue::Text("A".to_string()),
            text: "A".to_string(),
            formula: None,
            row: 0,
            col: 0,
        };
        builder.add_cell(cell1);

        let cell2 = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: None,
            row: 0,
            col: 0,
        };
        builder.add_cell(cell2);

        let row = builder.build();
        assert_eq!(row.cells.len(), 2);
        assert_eq!(row.cells[0].col, 0);
        assert_eq!(row.cells[1].col, 1);
    }

    #[test]
    fn test_cell_builder_float_types() {
        // Test "float" value type
        let builder = CellBuilder {
            value_type: Some("float".to_string()),
            value_str: Some("123.45".to_string()),
            currency: None,
            formula: None,
            repeated: 1,
        };
        let cell = builder.build("123.45");
        match cell.value {
            CellValue::Number(n) => assert!((n - 123.45).abs() < f64::EPSILON),
            _ => panic!("Expected Number for float"),
        }

        // Test "double" value type
        let builder = CellBuilder {
            value_type: Some("double".to_string()),
            value_str: Some("99.99".to_string()),
            currency: None,
            formula: None,
            repeated: 1,
        };
        let cell = builder.build("99.99");
        match cell.value {
            CellValue::Number(n) => assert!((n - 99.99).abs() < f64::EPSILON),
            _ => panic!("Expected Number for double"),
        }

        // Test "decimal" value type
        let builder = CellBuilder {
            value_type: Some("decimal".to_string()),
            value_str: Some("0.001".to_string()),
            currency: None,
            formula: None,
            repeated: 1,
        };
        let cell = builder.build("0.001");
        match cell.value {
            CellValue::Number(n) => assert!((n - 0.001).abs() < f64::EPSILON),
            _ => panic!("Expected Number for decimal"),
        }
    }

    #[test]
    fn test_cell_builder_invalid_number_fallback() {
        let builder = CellBuilder {
            value_type: Some("float".to_string()),
            value_str: Some("not-a-number".to_string()),
            currency: None,
            formula: None,
            repeated: 1,
        };
        let cell = builder.build("some text");
        match cell.value {
            CellValue::Text(t) => assert_eq!(t, "some text"),
            _ => panic!("Expected Text fallback for invalid number"),
        }
    }

    #[test]
    fn test_cell_builder_boolean_variations() {
        // Test "false" boolean
        let builder = CellBuilder {
            value_type: Some("boolean".to_string()),
            value_str: Some("false".to_string()),
            currency: None,
            formula: None,
            repeated: 1,
        };
        let cell = builder.build("FALSE");
        match cell.value {
            CellValue::Boolean(b) => assert!(!b),
            _ => panic!("Expected Boolean false"),
        }

        // Test invalid boolean value (falls back to text)
        let builder = CellBuilder {
            value_type: Some("boolean".to_string()),
            value_str: Some("maybe".to_string()),
            currency: None,
            formula: None,
            repeated: 1,
        };
        let cell = builder.build("maybe");
        match cell.value {
            CellValue::Text(t) => assert_eq!(t, "maybe"),
            _ => panic!("Expected Text for invalid boolean"),
        }
    }

    #[test]
    fn test_cell_builder_empty_text() {
        let builder = CellBuilder {
            value_type: None,
            value_str: None,
            currency: None,
            formula: None,
            repeated: 1,
        };
        let cell = builder.build("   ");
        match cell.value {
            CellValue::Empty => {},
            _ => panic!("Expected Empty for whitespace-only text"),
        }
    }

    #[test]
    fn test_cell_builder_currency_default() {
        let builder = CellBuilder {
            value_type: Some("currency".to_string()),
            value_str: Some("50".to_string()),
            currency: None, // No currency specified
            formula: None,
            repeated: 1,
        };
        let cell = builder.build("$50");
        match cell.value {
            CellValue::Currency(amount, currency) => {
                assert!((amount - 50.0).abs() < f64::EPSILON);
                assert_eq!(currency, "USD"); // Default
            },
            _ => panic!("Expected Currency with default USD"),
        }
    }

    #[test]
    fn test_parse_invalid_xml() {
        let invalid_xml = "<invalid>unclosed tag";
        let result = OdsParser::parse_sheets(invalid_xml);
        // The parser may return Ok with empty sheets or Err depending on implementation
        // Either behavior is acceptable - we just verify it doesn't panic
        match result {
            Ok(sheets) => {
                // If parsing succeeds, we should get 0 sheets
                assert_eq!(sheets.len(), 0);
            },
            Err(_) => {
                // Error is also acceptable
            },
        }
    }
}
