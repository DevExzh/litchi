//! Table-related ODF elements.
//!
//! This module provides classes for table elements like tables, rows, cells,
//! and other table-related content.

use super::element::{Element, ElementBase};
use crate::common::{Error, Result};
use crate::odf::ods::CellValue;

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
        for child in self.element.children.iter() {
            if child.tag_name() == "table:table-row"
                && let Ok(row) = TableRow::from_element(child.clone())
            {
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

    /// Get a row by index (alias for unified API)
    pub fn row_at(&self, index: usize) -> Result<Option<TableRow>> {
        self.row(index)
    }

    /// Add a row to the table
    pub fn add_row(&mut self, row: TableRow) {
        self.element.add_child(row.element);
    }

    /// Add a column definition to the table
    ///
    /// This must be called before adding rows. Columns are defined at the beginning
    /// of the table in ODF format.
    ///
    /// # Arguments
    ///
    /// * `column` - Table column definition
    pub fn add_column(&mut self, column: TableColumn) {
        // Columns must be inserted before rows
        // Find the first row index
        let first_row_idx = self
            .element
            .children
            .iter()
            .position(|child| child.tag_name() == "table:table-row");

        let col_element: Element = column.into();
        if let Some(idx) = first_row_idx {
            self.element.children.insert(idx, col_element);
        } else {
            self.element.add_child(col_element);
        }
    }

    /// Set the number of columns and create default column definitions
    ///
    /// # Arguments
    ///
    /// * `count` - Number of columns
    pub fn set_column_count(&mut self, count: usize) {
        for _ in 0..count {
            self.add_column(TableColumn::new());
        }
    }

    /// Get the number of columns (based on the widest row)
    pub fn column_count(&self) -> Result<usize> {
        let rows = self.rows()?;
        let max_cols = rows
            .iter()
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
            return Err(Error::InvalidFormat(
                "Element is not a table row".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get all cells in the row
    pub fn cells(&self) -> Result<Vec<TableCell>> {
        let mut cells = Vec::new();
        for child in self.element.children.iter() {
            if child.tag_name() == "table:table-cell"
                && let Ok(cell) = TableCell::from_element(child.clone())
            {
                cells.push(cell);
            }
        }
        Ok(cells)
    }

    /// Get the number of cells in the row
    pub fn cell_count(&self) -> Result<usize> {
        Ok(self.cells()?.len())
    }

    /// Get a cell by column index
    pub fn cell(&self, index: usize) -> Result<Option<TableCell>> {
        let cells = self.cells()?;
        Ok(cells.into_iter().nth(index))
    }

    /// Get a cell by column index (alias for unified API)
    pub fn cell_at(&self, index: usize) -> Result<Option<TableCell>> {
        self.cell(index)
    }

    /// Add a cell to the row
    pub fn add_cell(&mut self, cell: TableCell) {
        self.element.add_child(cell.element);
    }

    /// Get the style name (for row height)
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("table:style-name")
    }

    /// Set the style name (for row height)
    ///
    /// To set a specific height, you need to create a row style in the
    /// document's automatic styles section.
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("table:style-name", name);
    }

    /// Get the number of times this row is repeated.
    ///
    /// In ODF, rows can have a `table:number-rows-repeated` attribute to indicate
    /// that the row should be repeated multiple times. This method returns that count.
    ///
    /// # Returns
    ///
    /// The number of times this row appears (defaults to 1 if not specified).
    pub fn repeat_count(&self) -> usize {
        self.element
            .get_int_attribute("table:number-rows-repeated")
            .map(|n| n as usize)
            .unwrap_or(1)
    }

    /// Set the number of times this row should be repeated.
    pub fn set_repeat_count(&mut self, count: usize) {
        if count > 1 {
            self.element
                .set_attribute("table:number-rows-repeated", &count.to_string());
        } else {
            self.element.remove_attribute("table:number-rows-repeated");
        }
    }

    /// Get access to the underlying element for advanced operations.
    ///
    /// This is used internally by expansion utilities and other advanced features.
    #[allow(dead_code)] // Used by table expansion utilities
    pub(crate) fn element(&self) -> &Element {
        &self.element
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
            return Err(Error::InvalidFormat(
                "Element is not a table cell".to_string(),
            ));
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
                    && let Ok(num) = val_str.parse::<f64>()
                {
                    return Ok(CellValue::Number(num));
                }
            },
            Some("currency") => {
                if let Some(val_str) = self.element.get_attribute("office:value")
                    && let Ok(num) = val_str.parse::<f64>()
                {
                    let currency = self
                        .element
                        .get_attribute("office:currency")
                        .unwrap_or("USD");
                    return Ok(CellValue::Currency(num, currency.to_string()));
                }
            },
            Some("percentage") => {
                if let Some(val_str) = self.element.get_attribute("office:value")
                    && let Ok(num) = val_str.parse::<f64>()
                {
                    return Ok(CellValue::Percentage(num));
                }
            },
            Some("boolean") => {
                if let Some(val_str) = self.element.get_attribute("office:value") {
                    match val_str {
                        "true" => return Ok(CellValue::Boolean(true)),
                        "false" => return Ok(CellValue::Boolean(false)),
                        _ => {},
                    }
                }
            },
            Some("date") => {
                if let Some(val_str) = self.element.get_attribute("office:value") {
                    return Ok(CellValue::Date(val_str.to_string()));
                }
            },
            Some("time") => {
                if let Some(val_str) = self.element.get_attribute("office:value") {
                    return Ok(CellValue::Time(val_str.to_string()));
                }
            },
            _ => {
                let text = self.text()?;
                if text.trim().is_empty() {
                    return Ok(CellValue::Empty);
                } else {
                    return Ok(CellValue::Text(text));
                }
            },
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
        self.element
            .get_int_attribute("table:number-columns-spanned")
            .map(|n| n as usize)
            .unwrap_or(1)
    }

    /// Set the number of columns this cell spans
    pub fn set_colspan(&mut self, span: usize) {
        self.element
            .set_attribute("table:number-columns-spanned", &span.to_string());
    }

    /// Get the number of rows this cell spans
    pub fn rowspan(&self) -> usize {
        self.element
            .get_int_attribute("table:number-rows-spanned")
            .map(|n| n as usize)
            .unwrap_or(1)
    }

    /// Set the number of rows this cell spans
    pub fn set_rowspan(&mut self, span: usize) {
        self.element
            .set_attribute("table:number-rows-spanned", &span.to_string());
    }

    /// Check if the cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self.value(), Ok(CellValue::Empty))
    }

    /// Get the number of times this cell is repeated.
    ///
    /// In ODF, cells can have a `table:number-columns-repeated` attribute to indicate
    /// that the cell should be repeated multiple times horizontally. This method returns that count.
    ///
    /// # Returns
    ///
    /// The number of times this cell appears (defaults to 1 if not specified).
    pub fn repeat_count(&self) -> usize {
        self.element
            .get_int_attribute("table:number-columns-repeated")
            .map(|n| n as usize)
            .unwrap_or(1)
    }

    /// Set the number of times this cell should be repeated.
    pub fn set_repeat_count(&mut self, count: usize) {
        if count > 1 {
            self.element
                .set_attribute("table:number-columns-repeated", &count.to_string());
        } else {
            self.element
                .remove_attribute("table:number-columns-repeated");
        }
    }

    /// Get access to the underlying element for advanced operations.
    ///
    /// This is used internally by expansion utilities and other advanced features.
    #[allow(dead_code)] // Used by table expansion utilities
    pub(crate) fn element(&self) -> &Element {
        &self.element
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
            return Err(Error::InvalidFormat(
                "Element is not a table column".to_string(),
            ));
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
        self.element
            .set_attribute("table:default-cell-style-name", name);
    }

    /// Get the number of columns this column definition represents
    pub fn repeated(&self) -> usize {
        self.element
            .get_int_attribute("table:number-columns-repeated")
            .map(|n| n as usize)
            .unwrap_or(1)
    }

    /// Set the number of columns this column definition represents
    pub fn set_repeated(&mut self, count: usize) {
        self.element
            .set_attribute("table:number-columns-repeated", &count.to_string());
    }
}

impl From<TableColumn> for Element {
    fn from(col: TableColumn) -> Element {
        col.element
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ========== Table Tests ==========
    #[test]
    fn test_table_new() {
        let table = Table::new();
        assert!(table.name().is_none());
        assert!(table.style_name().is_none());
        assert_eq!(table.row_count().unwrap(), 0);
        assert_eq!(table.column_count().unwrap(), 0);
    }

    #[test]
    fn test_table_name() {
        let mut table = Table::new();
        table.set_name("Sheet1");
        assert_eq!(table.name(), Some("Sheet1"));
    }

    #[test]
    fn test_table_style_name() {
        let mut table = Table::new();
        table.set_style_name("TableStyle");
        assert_eq!(table.style_name(), Some("TableStyle"));
    }

    #[test]
    fn test_table_from_element() {
        let element = Element::new("table:table");
        let table = Table::from_element(element).unwrap();
        assert!(table.name().is_none());
    }

    #[test]
    fn test_table_from_element_wrong_tag() {
        let element = Element::new("table:row");
        assert!(Table::from_element(element).is_err());
    }

    #[test]
    fn test_table_add_row() {
        let mut table = Table::new();
        let row = TableRow::new();
        table.add_row(row);
        assert_eq!(table.row_count().unwrap(), 1);
    }

    #[test]
    fn test_table_row_access() {
        let mut table = Table::new();
        let mut row = TableRow::new();
        let cell = TableCell::new();
        row.add_cell(cell);
        table.add_row(row);

        assert!(table.row(0).unwrap().is_some());
        assert!(table.row_at(0).unwrap().is_some());
        assert!(table.row(1).unwrap().is_none());
    }

    #[test]
    fn test_table_add_column() {
        let mut table = Table::new();
        let col = TableColumn::new();
        table.add_column(col);
        // Column doesn't affect row-based column count
        assert_eq!(table.column_count().unwrap(), 0);
    }

    #[test]
    fn test_table_set_column_count() {
        let mut table = Table::new();
        table.set_column_count(3);
        // Column count is based on widest row, so still 0
        assert_eq!(table.column_count().unwrap(), 0);
    }

    #[test]
    fn test_table_rows() {
        let mut table = Table::new();
        let row1 = TableRow::new();
        let row2 = TableRow::new();
        table.add_row(row1);
        table.add_row(row2);

        let rows = table.rows().unwrap();
        assert_eq!(rows.len(), 2);
    }

    #[test]
    fn test_table_column_count_with_data() {
        let mut table = Table::new();
        let mut row = TableRow::new();
        row.add_cell(TableCell::new());
        row.add_cell(TableCell::new());
        row.add_cell(TableCell::new());
        table.add_row(row);

        assert_eq!(table.column_count().unwrap(), 3);
    }

    // ========== TableRow Tests ==========
    #[test]
    fn test_table_row_new() {
        let row = TableRow::new();
        assert_eq!(row.cell_count().unwrap(), 0);
        assert!(row.style_name().is_none());
        assert_eq!(row.repeat_count(), 1);
    }

    #[test]
    fn test_table_row_from_element() {
        let element = Element::new("table:table-row");
        let row = TableRow::from_element(element).unwrap();
        assert_eq!(row.cell_count().unwrap(), 0);
    }

    #[test]
    fn test_table_row_from_element_wrong_tag() {
        let element = Element::new("table:table-cell");
        assert!(TableRow::from_element(element).is_err());
    }

    #[test]
    fn test_table_row_add_cell() {
        let mut row = TableRow::new();
        let cell = TableCell::new();
        row.add_cell(cell);
        assert_eq!(row.cell_count().unwrap(), 1);
    }

    #[test]
    fn test_table_row_cells() {
        let mut row = TableRow::new();
        row.add_cell(TableCell::new());
        row.add_cell(TableCell::new());

        let cells = row.cells().unwrap();
        assert_eq!(cells.len(), 2);
    }

    #[test]
    fn test_table_row_cell_access() {
        let mut row = TableRow::new();
        let mut cell = TableCell::new();
        cell.set_text("Test");
        row.add_cell(cell);

        assert!(row.cell(0).unwrap().is_some());
        assert!(row.cell_at(0).unwrap().is_some());
        assert!(row.cell(1).unwrap().is_none());
    }

    #[test]
    fn test_table_row_style_name() {
        let mut row = TableRow::new();
        row.set_style_name("RowStyle");
        assert_eq!(row.style_name(), Some("RowStyle"));
    }

    #[test]
    fn test_table_row_repeat_count() {
        let mut row = TableRow::new();
        assert_eq!(row.repeat_count(), 1);

        row.set_repeat_count(5);
        assert_eq!(row.repeat_count(), 5);

        row.set_repeat_count(1);
        assert_eq!(row.repeat_count(), 1);
    }

    // ========== TableCell Tests ==========
    #[test]
    fn test_table_cell_new() {
        let cell = TableCell::new();
        assert_eq!(cell.text().unwrap(), "");
        assert!(cell.formula().is_none());
        assert!(cell.style_name().is_none());
        assert_eq!(cell.colspan(), 1);
        assert_eq!(cell.rowspan(), 1);
        assert_eq!(cell.repeat_count(), 1);
    }

    #[test]
    fn test_table_cell_from_element() {
        let element = Element::new("table:table-cell");
        let cell = TableCell::from_element(element).unwrap();
        assert_eq!(cell.text().unwrap(), "");
    }

    #[test]
    fn test_table_cell_from_element_wrong_tag() {
        let element = Element::new("table:table-row");
        assert!(TableCell::from_element(element).is_err());
    }

    #[test]
    fn test_table_cell_set_text() {
        let mut cell = TableCell::new();
        cell.set_text("Hello World");
        assert_eq!(cell.text().unwrap(), "Hello World");
    }

    #[test]
    fn test_table_cell_formula() {
        let mut cell = TableCell::new();
        assert!(cell.formula().is_none());

        cell.set_formula("=SUM(A1:B2)");
        assert_eq!(cell.formula(), Some("=SUM(A1:B2)"));
    }

    #[test]
    fn test_table_cell_style_name() {
        let mut cell = TableCell::new();
        cell.set_style_name("CellStyle");
        assert_eq!(cell.style_name(), Some("CellStyle"));
    }

    #[test]
    fn test_table_cell_colspan() {
        let mut cell = TableCell::new();
        assert_eq!(cell.colspan(), 1);

        cell.set_colspan(3);
        assert_eq!(cell.colspan(), 3);
    }

    #[test]
    fn test_table_cell_rowspan() {
        let mut cell = TableCell::new();
        assert_eq!(cell.rowspan(), 1);

        cell.set_rowspan(2);
        assert_eq!(cell.rowspan(), 2);
    }

    #[test]
    fn test_table_cell_repeat_count() {
        let mut cell = TableCell::new();
        assert_eq!(cell.repeat_count(), 1);

        cell.set_repeat_count(4);
        assert_eq!(cell.repeat_count(), 4);

        cell.set_repeat_count(1);
        assert_eq!(cell.repeat_count(), 1);
    }

    #[test]
    fn test_table_cell_is_empty() {
        let cell = TableCell::new();
        assert!(cell.is_empty());

        let mut cell = TableCell::new();
        cell.set_text("Content");
        assert!(!cell.is_empty());
    }

    #[test]
    fn test_table_cell_value_empty() {
        let cell = TableCell::new();
        assert!(matches!(cell.value().unwrap(), CellValue::Empty));
    }

    #[test]
    fn test_table_cell_value_text() {
        let mut cell = TableCell::new();
        cell.set_text("Hello");
        assert!(matches!(cell.value().unwrap(), CellValue::Text(_)));
    }

    #[test]
    fn test_table_cell_value_number() {
        let mut element = Element::new("table:table-cell");
        element.set_attribute("office:value-type", "float");
        element.set_attribute("office:value", "42.5");
        let cell = TableCell::from_element(element).unwrap();

        match cell.value().unwrap() {
            CellValue::Number(n) => assert!((n - 42.5).abs() < f64::EPSILON),
            _ => panic!("Expected Number"),
        }
    }

    #[test]
    fn test_table_cell_value_boolean() {
        let mut element = Element::new("table:table-cell");
        element.set_attribute("office:value-type", "boolean");
        element.set_attribute("office:value", "true");
        let cell = TableCell::from_element(element).unwrap();

        match cell.value().unwrap() {
            CellValue::Boolean(b) => assert!(b),
            _ => panic!("Expected Boolean"),
        }
    }

    #[test]
    fn test_table_cell_value_date() {
        let mut element = Element::new("table:table-cell");
        element.set_attribute("office:value-type", "date");
        element.set_attribute("office:value", "2024-03-15");
        let cell = TableCell::from_element(element).unwrap();

        match cell.value().unwrap() {
            CellValue::Date(d) => assert_eq!(d, "2024-03-15"),
            _ => panic!("Expected Date"),
        }
    }

    // ========== TableColumn Tests ==========
    #[test]
    fn test_table_column_new() {
        let col = TableColumn::new();
        assert!(col.style_name().is_none());
        assert!(col.default_cell_style_name().is_none());
        assert_eq!(col.repeated(), 1);
    }

    #[test]
    fn test_table_column_from_element() {
        let element = Element::new("table:table-column");
        let col = TableColumn::from_element(element).unwrap();
        assert_eq!(col.repeated(), 1);
    }

    #[test]
    fn test_table_column_from_element_wrong_tag() {
        let element = Element::new("table:table-cell");
        assert!(TableColumn::from_element(element).is_err());
    }

    #[test]
    fn test_table_column_style_name() {
        let mut col = TableColumn::new();
        col.set_style_name("ColumnStyle");
        assert_eq!(col.style_name(), Some("ColumnStyle"));
    }

    #[test]
    fn test_table_column_default_cell_style() {
        let mut col = TableColumn::new();
        col.set_default_cell_style_name("DefaultCell");
        assert_eq!(col.default_cell_style_name(), Some("DefaultCell"));
    }

    #[test]
    fn test_table_column_repeated() {
        let mut col = TableColumn::new();
        assert_eq!(col.repeated(), 1);

        col.set_repeated(5);
        assert_eq!(col.repeated(), 5);
    }

    // ========== TableElements Tests ==========
    #[test]
    fn test_table_elements_parse_tables_empty() {
        let xml = r#"<office:document xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"></office:document>"#;
        let tables = TableElements::parse_tables(xml).unwrap();
        assert!(tables.is_empty());
    }

    #[test]
    fn test_table_elements_parse_tables_single() {
        let xml = r#"<office:document xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:table table:name="Table1">
                <table:table-row>
                    <table:table-cell>Cell 1</table:table-cell>
                </table:table-row>
            </table:table>
        </office:document>"#;

        let tables = TableElements::parse_tables(xml).unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name(), Some("Table1"));
    }

    #[test]
    fn test_table_elements_parse_tables_multiple() {
        let xml = r#"<office:document xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:table table:name="Table1"></table:table>
            <table:table table:name="Table2"></table:table>
        </office:document>"#;

        let tables = TableElements::parse_tables(xml).unwrap();
        assert_eq!(tables.len(), 2);
    }

    #[test]
    fn test_table_elements_parse_from_content() {
        let xml = r#"<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:table table:name="Sheet1"></table:table>
        </office:document-content>"#;

        let tables = TableElements::parse_tables_from_content(xml).unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name(), Some("Sheet1"));
    }

    #[test]
    fn test_table_elements_parse_table_with_attributes() {
        let xml = r#"<office:document xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:table table:name="TestTable" table:style-name="TableStyle">
            </table:table>
        </office:document>"#;

        let tables = TableElements::parse_tables(xml).unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name(), Some("TestTable"));
        assert_eq!(tables[0].style_name(), Some("TableStyle"));
    }

    #[test]
    fn test_table_roundtrip() {
        // Test converting Table to Element and back
        let mut table = Table::new();
        table.set_name("TestTable");
        table.set_style_name("TestStyle");

        let element: Element = table.into();
        let table2 = Table::from_element(element).unwrap();

        assert_eq!(table2.name(), Some("TestTable"));
        assert_eq!(table2.style_name(), Some("TestStyle"));
    }

    #[test]
    fn test_table_row_roundtrip() {
        let mut row = TableRow::new();
        row.set_style_name("RowStyle");
        row.set_repeat_count(3);

        let element: Element = row.into();
        let row2 = TableRow::from_element(element).unwrap();

        assert_eq!(row2.style_name(), Some("RowStyle"));
        assert_eq!(row2.repeat_count(), 3);
    }

    #[test]
    fn test_table_cell_roundtrip() {
        let mut cell = TableCell::new();
        cell.set_text("Test");
        cell.set_formula("=A1+B1");
        cell.set_style_name("CellStyle");
        cell.set_colspan(2);
        cell.set_rowspan(3);

        let element: Element = cell.into();
        let cell2 = TableCell::from_element(element).unwrap();

        assert_eq!(cell2.text().unwrap(), "Test");
        assert_eq!(cell2.formula(), Some("=A1+B1"));
        assert_eq!(cell2.style_name(), Some("CellStyle"));
        assert_eq!(cell2.colspan(), 2);
        assert_eq!(cell2.rowspan(), 3);
    }

    #[test]
    fn test_table_column_roundtrip() {
        let mut col = TableColumn::new();
        col.set_style_name("ColStyle");
        col.set_default_cell_style_name("DefaultCell");
        col.set_repeated(5);

        let element: Element = col.into();
        let col2 = TableColumn::from_element(element).unwrap();

        assert_eq!(col2.style_name(), Some("ColStyle"));
        assert_eq!(col2.default_cell_style_name(), Some("DefaultCell"));
        assert_eq!(col2.repeated(), 5);
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
                    let tag_name =
                        String::from_utf8(e.name().as_ref().to_vec()).unwrap_or_default();

                    if tag_name == "table:table" {
                        let mut element = Element::new(&tag_name);

                        // Parse attributes
                        for attr_result in e.attributes() {
                            if let Ok(attr) = attr_result
                                && let (Ok(key), Ok(value)) = (
                                    String::from_utf8(attr.key.as_ref().to_vec()),
                                    String::from_utf8(attr.value.to_vec()),
                                )
                            {
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
                                    String::from_utf8(attr.value.to_vec()),
                                )
                            {
                                element.set_attribute(&key, &value);
                            }
                        }

                        stack.push(element);
                    }
                },
                Ok(quick_xml::events::Event::Text(ref t)) => {
                    if let Some(current) = stack.last_mut()
                        && let Ok(text) = String::from_utf8(t.to_vec())
                    {
                        let current_text = current.text().to_string();
                        current.set_text(&format!("{}{}", current_text, text));
                    }
                },
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let tag_name =
                        String::from_utf8(e.name().as_ref().to_vec()).unwrap_or_default();

                    if tag_name == "table:table" {
                        if let Some(table_element) = stack.pop()
                            && let Ok(table) = Table::from_element(table_element)
                        {
                            tables.push(table);
                        }
                    } else if !stack.is_empty() {
                        let element = stack.pop().unwrap();
                        if let Some(parent) = stack.last_mut() {
                            parent.add_child(element);
                        }
                    }
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(tables)
    }

    /// Parse table from XML content with proper handling of repeated cells
    #[allow(dead_code)]
    pub fn parse_table_with_expansion(
        xml_content: &str,
        table_name: Option<&str>,
    ) -> Result<Option<Table>> {
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
                        let repeated = cell
                            .element
                            .get_int_attribute("table:number-columns-repeated")
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
