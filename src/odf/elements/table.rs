//! Table-related ODF elements.
//!
//! This module provides classes for table elements like tables, rows, cells,
//! and other table-related content.

use super::element::{Element, ElementBase};
use crate::common::{Error, Result};
use crate::odf::spreadsheet::CellValue;

/// A table element
#[derive(Debug, Clone)]
pub struct Table {
    element: Element,
}

impl Default for Table {
    fn default() -> Self {
        Self::new()
    }
}

impl Table {
    /// Create a new table
    pub fn new() -> Self {
        Self {
            element: Element::new("table:table"),
        }
    }

    /// Create table from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "table:table" {
            return Err(Error::InvalidFormat("Element is not a table".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the table name
    pub fn name(&self) -> Option<&str> {
        self.element.get_attribute("table:name")
    }

    /// Set the table name
    pub fn set_name(&mut self, name: &str) {
        self.element.set_attribute("table:name", name);
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("table:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("table:style-name", name);
    }

    /// Get all rows in the table
    pub fn rows(&self) -> Result<Vec<TableRow>> {
        let mut rows = Vec::new();
        for child in self.element.children() {
            if child.tag_name() == "table:table-row"
                && let Ok(row) = TableRow::from_element(unsafe { &*(child as *const _ as *const Element) }.clone()) {
                    rows.push(row);
                }
        }
        Ok(rows)
    }

    /// Get the number of rows
    pub fn row_count(&self) -> Result<usize> {
        Ok(self.rows()?.len())
    }

    /// Get a row by index
    pub fn row(&self, index: usize) -> Result<Option<TableRow>> {
        let rows = self.rows()?;
        Ok(rows.into_iter().nth(index))
    }

    /// Add a row to the table
    pub fn add_row(&mut self, row: TableRow) {
        self.element.add_child(Box::new(row.element));
    }

    /// Get the number of columns (based on the widest row)
    pub fn column_count(&self) -> Result<usize> {
        let rows = self.rows()?;
        let max_cols = rows.iter()
            .map(|row| row.cells().map(|cells| cells.len()).unwrap_or(0))
            .max()
            .unwrap_or(0);
        Ok(max_cols)
    }
}

impl From<Table> for Element {
    fn from(table: Table) -> Element {
        table.element
    }
}

/// A table row element
#[derive(Debug, Clone)]
pub struct TableRow {
    element: Element,
}

impl Default for TableRow {
    fn default() -> Self {
        Self::new()
    }
}

impl TableRow {
    /// Create a new table row
    pub fn new() -> Self {
        Self {
            element: Element::new("table:table-row"),
        }
    }

    /// Create table row from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "table:table-row" {
            return Err(Error::InvalidFormat("Element is not a table row".to_string()));
        }
        Ok(Self { element })
    }

    /// Get all cells in the row
    pub fn cells(&self) -> Result<Vec<TableCell>> {
        let mut cells = Vec::new();
        for child in self.element.children() {
            if child.tag_name() == "table:table-cell"
                && let Ok(cell) = TableCell::from_element(unsafe { &*(child as *const _ as *const Element) }.clone()) {
                    cells.push(cell);
                }
        }
        Ok(cells)
    }

    /// Get a cell by column index
    pub fn cell(&self, index: usize) -> Result<Option<TableCell>> {
        let cells = self.cells()?;
        Ok(cells.into_iter().nth(index))
    }

    /// Add a cell to the row
    pub fn add_cell(&mut self, cell: TableCell) {
        self.element.add_child(Box::new(cell.element));
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("table:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("table:style-name", name);
    }
}

impl From<TableRow> for Element {
    fn from(row: TableRow) -> Element {
        row.element
    }
}

/// A table cell element
#[derive(Debug, Clone)]
pub struct TableCell {
    element: Element,
}

impl Default for TableCell {
    fn default() -> Self {
        Self::new()
    }
}

impl TableCell {
    /// Create a new table cell
    pub fn new() -> Self {
        Self {
            element: Element::new("table:table-cell"),
        }
    }

    /// Create table cell from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "table:table-cell" {
            return Err(Error::InvalidFormat("Element is not a table cell".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the text content of the cell
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive().trim().to_string())
    }

    /// Set the text content of the cell
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the cell value (parsed from attributes and content)
    pub fn value(&self) -> Result<CellValue> {
        // Check for value type
        let value_type = self.element.get_attribute("office:value-type");

        match value_type {
            Some("float") | Some("double") | Some("decimal") => {
                if let Some(val_str) = self.element.get_attribute("office:value")
                    && let Ok(num) = val_str.parse::<f64>() {
                        return Ok(CellValue::Number(num));
                    }
            }
            Some("currency") => {
                if let Some(val_str) = self.element.get_attribute("office:value")
                    && let Ok(num) = val_str.parse::<f64>() {
                        let currency = self.element.get_attribute("office:currency")
                            .unwrap_or("USD");
                        return Ok(CellValue::Currency(num, currency.to_string()));
                    }
            }
            Some("percentage") => {
                if let Some(val_str) = self.element.get_attribute("office:value")
                    && let Ok(num) = val_str.parse::<f64>() {
                        return Ok(CellValue::Percentage(num));
                    }
            }
            Some("boolean") => {
                if let Some(val_str) = self.element.get_attribute("office:value") {
                    match val_str {
                        "true" => return Ok(CellValue::Boolean(true)),
                        "false" => return Ok(CellValue::Boolean(false)),
                        _ => {}
                    }
                }
            }
            Some("date") => {
                if let Some(val_str) = self.element.get_attribute("office:value") {
                    return Ok(CellValue::Date(val_str.to_string()));
                }
            }
            Some("time") => {
                if let Some(val_str) = self.element.get_attribute("office:value") {
                    return Ok(CellValue::Time(val_str.to_string()));
                }
            }
            Some("string") | _ => {
                let text = self.text()?;
                if text.trim().is_empty() {
                    return Ok(CellValue::Empty);
                } else {
                    return Ok(CellValue::Text(text));
                }
            }
        }

        // Fallback to text parsing
        let text = self.text()?;
        if text.trim().is_empty() {
            Ok(CellValue::Empty)
        } else {
            Ok(CellValue::Text(text))
        }
    }

    /// Get the formula in the cell
    pub fn formula(&self) -> Option<&str> {
        self.element.get_attribute("table:formula")
    }

    /// Set the formula in the cell
    pub fn set_formula(&mut self, formula: &str) {
        self.element.set_attribute("table:formula", formula);
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("table:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("table:style-name", name);
    }

    /// Get the number of columns this cell spans
    pub fn colspan(&self) -> usize {
        self.element.get_int_attribute("table:number-columns-spanned")
            .map(|n| n as usize)
            .unwrap_or(1)
    }

    /// Set the number of columns this cell spans
    pub fn set_colspan(&mut self, span: usize) {
        self.element.set_attribute("table:number-columns-spanned", &span.to_string());
    }

    /// Get the number of rows this cell spans
    pub fn rowspan(&self) -> usize {
        self.element.get_int_attribute("table:number-rows-spanned")
            .map(|n| n as usize)
            .unwrap_or(1)
    }

    /// Set the number of rows this cell spans
    pub fn set_rowspan(&mut self, span: usize) {
        self.element.set_attribute("table:number-rows-spanned", &span.to_string());
    }

    /// Check if the cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self.value(), Ok(CellValue::Empty))
    }
}

impl From<TableCell> for Element {
    fn from(cell: TableCell) -> Element {
        cell.element
    }
}

/// A table column element
#[derive(Debug, Clone)]
pub struct TableColumn {
    element: Element,
}

impl Default for TableColumn {
    fn default() -> Self {
        Self::new()
    }
}

impl TableColumn {
    /// Create a new table column
    pub fn new() -> Self {
        Self {
            element: Element::new("table:table-column"),
        }
    }

    /// Create table column from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "table:table-column" {
            return Err(Error::InvalidFormat("Element is not a table column".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("table:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("table:style-name", name);
    }

    /// Get the default cell style name
    pub fn default_cell_style_name(&self) -> Option<&str> {
        self.element.get_attribute("table:default-cell-style-name")
    }

    /// Set the default cell style name
    pub fn set_default_cell_style_name(&mut self, name: &str) {
        self.element.set_attribute("table:default-cell-style-name", name);
    }

    /// Get the number of columns this column definition represents
    pub fn repeated(&self) -> usize {
        self.element.get_int_attribute("table:number-columns-repeated")
            .map(|n| n as usize)
            .unwrap_or(1)
    }

    /// Set the number of columns this column definition represents
    pub fn set_repeated(&mut self, count: usize) {
        self.element.set_attribute("table:number-columns-repeated", &count.to_string());
    }
}

impl From<TableColumn> for Element {
    fn from(col: TableColumn) -> Element {
        col.element
    }
}

/// Collection of table elements for easy parsing
pub struct TableElements;

impl TableElements {
    /// Parse all tables from document content (content.xml)
    pub fn parse_tables_from_content(xml_content: &str) -> Result<Vec<Table>> {
        Self::parse_tables(xml_content)
    }

    /// Parse all tables from XML content
    pub fn parse_tables(xml_content: &str) -> Result<Vec<Table>> {
        let mut reader = quick_xml::Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut tables = Vec::new();
        let mut stack: Vec<Element> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let tag_name = String::from_utf8(e.name().as_ref().to_vec())
                        .unwrap_or_default();

                    if tag_name == "table:table" {
                        let mut element = Element::new(&tag_name);

                        // Parse attributes
                        for attr_result in e.attributes() {
                            if let Ok(attr) = attr_result
                                && let (Ok(key), Ok(value)) = (
                                    String::from_utf8(attr.key.as_ref().to_vec()),
                                    String::from_utf8(attr.value.to_vec())
                                ) {
                                    element.set_attribute(&key, &value);
                                }
                        }

                        stack.push(element);
                    } else if !stack.is_empty() {
                        // Handle nested elements within table
                        let mut element = Element::new(&tag_name);

                        // Parse attributes
                        for attr_result in e.attributes() {
                            if let Ok(attr) = attr_result
                                && let (Ok(key), Ok(value)) = (
                                    String::from_utf8(attr.key.as_ref().to_vec()),
                                    String::from_utf8(attr.value.to_vec())
                                ) {
                                    element.set_attribute(&key, &value);
                                }
                        }

                        stack.push(element);
                    }
                }
                Ok(quick_xml::events::Event::Text(ref t)) => {
                    if let Some(current) = stack.last_mut()
                        && let Ok(text) = String::from_utf8(t.to_vec()) {
                            let current_text = current.text().to_string();
                            current.set_text(&format!("{}{}", current_text, text));
                        }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let tag_name = String::from_utf8(e.name().as_ref().to_vec())
                        .unwrap_or_default();

                    if tag_name == "table:table" {
                        if let Some(table_element) = stack.pop()
                            && let Ok(table) = Table::from_element(table_element) {
                                tables.push(table);
                            }
                    } else if !stack.is_empty() {
                        let element = stack.pop().unwrap();
                        if let Some(parent) = stack.last_mut() {
                            parent.add_child(Box::new(element));
                        }
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(tables)
    }

    /// Parse table from XML content with proper handling of repeated cells
    pub fn parse_table_with_expansion(xml_content: &str, table_name: Option<&str>) -> Result<Option<Table>> {
        let tables = Self::parse_tables(xml_content)?;

        for table in tables {
            if table_name.is_none() || table.name() == table_name {
                // Expand repeated cells
                let mut expanded_table = Table::new();
                if let Some(name) = table.name() {
                    expanded_table.set_name(name);
                }
                if let Some(style) = table.style_name() {
                    expanded_table.set_style_name(style);
                }

                for row in table.rows()? {
                    let mut expanded_row = TableRow::new();
                    if let Some(style) = row.style_name() {
                        expanded_row.set_style_name(style);
                    }

                    for cell in row.cells()? {
                        let repeated = cell.element.get_int_attribute("table:number-columns-repeated")
                            .map(|n| n as usize)
                            .unwrap_or(1);

                        for _ in 0..repeated {
                            let mut new_cell = TableCell::new();
                            new_cell.set_text(cell.text()?.as_str());

                            // Copy other attributes
                            if let Some(formula) = cell.formula() {
                                new_cell.set_formula(formula);
                            }
                            if let Some(style) = cell.style_name() {
                                new_cell.set_style_name(style);
                            }
                            if cell.colspan() > 1 {
                                new_cell.set_colspan(cell.colspan());
                            }
                            if cell.rowspan() > 1 {
                                new_cell.set_rowspan(cell.rowspan());
                            }

                            // Copy value attributes
                            for (key, value) in cell.element.attributes() {
                                if key.starts_with("office:") {
                                    new_cell.element.set_attribute(key, value);
                                }
                            }

                            expanded_row.add_cell(new_cell);
                        }
                    }

                    expanded_table.add_row(expanded_row);
                }

                return Ok(Some(expanded_table));
            }
        }

        Ok(None)
    }
}
