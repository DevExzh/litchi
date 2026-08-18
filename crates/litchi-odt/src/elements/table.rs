//! Table-related ODF elements.
//!
//! This module provides classes for table elements like tables, rows, cells,
//! and other table-related content.

use super::element::{Element, ElementBase, try_owned_string};
use crate::CellValue;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::BytesStart;
use std::fmt::Write as _;

fn try_push<T>(items: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    items
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })?;
    items.push(value);
    Ok(())
}

fn try_usize_string(value: usize, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(usize::BITS as usize / 3 + 1)
        .map_err(|source| Error::Allocation { resource, source })?;
    write!(&mut output, "{value}")
        .map_err(|_error| Error::InvalidFormat("failed to format ODT integer".to_string()))?;
    Ok(output)
}

fn append_text_control(
    reader: &quick_xml::Reader<&[u8]>,
    source: &BytesStart<'_>,
    element: &mut Element,
) -> Result<()> {
    match source.name().as_ref() {
        b"text:s" => {
            let mut count = 1usize;
            for attribute in source.attributes() {
                let attribute = attribute.map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODT table text:s attribute: {error}"))
                })?;
                if attribute.key.as_ref() == b"text:c" {
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODT table text:s count: {error}"))
                        })?;
                    count = value.parse().map_err(|_error| {
                        Error::InvalidFormat(
                            "ODT table text:s count must be a non-negative integer".to_string(),
                        )
                    })?;
                }
            }
            if count > 1_000_000 {
                return Err(Error::InvalidFormat(
                    "ODT table text:s count exceeds 1000000".to_string(),
                ));
            }
            element.try_append_spaces(count, "ODT table text control")
        },
        b"text:tab" => element.try_append_text("\t", "ODT table text control"),
        b"text:line-break" => element.try_append_text("\n", "ODT table text control"),
        _ => Ok(()),
    }
}

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
    pub(crate) fn try_new() -> Result<Self> {
        Ok(Self {
            element: Element::try_new("table:table")?,
        })
    }

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

    pub(crate) fn try_set_name(&mut self, name: &str) -> Result<()> {
        self.element
            .try_set_attribute("table:name", name, "ODT expanded table name")
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("table:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("table:style-name", name);
    }

    pub(crate) fn try_set_style_name(&mut self, name: &str) -> Result<()> {
        self.element
            .try_set_attribute("table:style-name", name, "ODT expanded table style name")
    }

    /// Get all rows in the table
    pub fn rows(&self) -> Result<Vec<TableRow>> {
        let mut rows = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "table:table-row" {
                let row = TableRow::from_element(child.try_clone()?)?;
                try_push(&mut rows, row, "ODT table row projection")?;
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

    pub(crate) fn try_add_row(&mut self, row: TableRow) -> Result<()> {
        self.element
            .try_add_child(row.element, "ODT expanded table rows")
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
        let mut max_cols = 0usize;
        for row in self.rows()? {
            max_cols = max_cols.max(row.cells()?.len());
        }
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
    pub(crate) fn try_new() -> Result<Self> {
        Ok(Self {
            element: Element::try_new("table:table-row")?,
        })
    }

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
        for child in &self.element.children {
            if child.tag_name() == "table:table-cell" {
                let cell = TableCell::from_element(child.try_clone()?)?;
                try_push(&mut cells, cell, "ODT table cell projection")?;
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

    pub(crate) fn try_add_cell(&mut self, cell: TableCell) -> Result<()> {
        self.element
            .try_add_child(cell.element, "ODT expanded table cells")
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

    pub(crate) fn try_set_style_name(&mut self, name: &str) -> Result<()> {
        self.element.try_set_attribute(
            "table:style-name",
            name,
            "ODT expanded table-row style name",
        )
    }

    pub(crate) fn try_repeat_count(&self) -> Result<usize> {
        self.element
            .get_attribute("table:number-rows-repeated")
            .map(|value| {
                value.parse::<usize>().map_err(|_error| {
                    Error::InvalidFormat(
                        "table:number-rows-repeated must be a non-negative integer".to_string(),
                    )
                })
            })
            .transpose()
            .map(|value| value.unwrap_or(1))
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
            .and_then(|value| usize::try_from(value).ok())
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
    pub(crate) fn try_new() -> Result<Self> {
        Ok(Self {
            element: Element::try_new("table:table-cell")?,
        })
    }

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
        let text = self.element.try_get_text_recursive()?;
        let trimmed = text.trim();
        if trimmed.len() == text.len() {
            return Ok(text);
        }
        let mut output = String::new();
        output
            .try_reserve(trimmed.len())
            .map_err(|source| Error::Allocation {
                resource: "ODT table-cell text projection",
                source,
            })?;
        output.push_str(trimmed);
        Ok(output)
    }

    /// Set the text content of the cell
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    pub(crate) fn try_set_text(&mut self, text: &str) -> Result<()> {
        self.element
            .try_set_text(text, "ODT expanded table-cell text")
    }

    /// Get the cell value (parsed from attributes and content)
    pub fn value(&self) -> Result<CellValue> {
        // Check for value type
        let value_type = self.element.get_attribute("office:value-type");

        match value_type {
            Some("float" | "double" | "decimal") => {
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
                    return Ok(CellValue::Currency(
                        num,
                        try_owned_string(currency, "ODT table-cell currency")?,
                    ));
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
                    return Ok(CellValue::Date(try_owned_string(
                        val_str,
                        "ODT table-cell date",
                    )?));
                }
            },
            Some("time") => {
                if let Some(val_str) = self.element.get_attribute("office:value") {
                    return Ok(CellValue::Time(try_owned_string(
                        val_str,
                        "ODT table-cell time",
                    )?));
                }
            },
            _ => {
                let text = self.text()?;
                if text.trim().is_empty() {
                    return Ok(CellValue::Empty);
                }
                return Ok(CellValue::Text(text));
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

    pub(crate) fn try_set_formula(&mut self, formula: &str) -> Result<()> {
        self.element
            .try_set_attribute("table:formula", formula, "ODT expanded table-cell formula")
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("table:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("table:style-name", name);
    }

    pub(crate) fn try_set_style_name(&mut self, name: &str) -> Result<()> {
        self.element.try_set_attribute(
            "table:style-name",
            name,
            "ODT expanded table-cell style name",
        )
    }

    /// Get the number of columns this cell spans
    pub fn colspan(&self) -> usize {
        self.element
            .get_int_attribute("table:number-columns-spanned")
            .and_then(|value| usize::try_from(value).ok())
            .unwrap_or(1)
    }

    /// Set the number of columns this cell spans
    pub fn set_colspan(&mut self, span: usize) {
        self.element
            .set_attribute("table:number-columns-spanned", &span.to_string());
    }

    pub(crate) fn try_set_colspan(&mut self, span: usize) -> Result<()> {
        let value = try_usize_string(span, "ODT expanded table-cell column span")?;
        self.element.try_set_attribute(
            "table:number-columns-spanned",
            &value,
            "ODT expanded table-cell column span",
        )
    }

    /// Get the number of rows this cell spans
    pub fn rowspan(&self) -> usize {
        self.element
            .get_int_attribute("table:number-rows-spanned")
            .and_then(|value| usize::try_from(value).ok())
            .unwrap_or(1)
    }

    /// Set the number of rows this cell spans
    pub fn set_rowspan(&mut self, span: usize) {
        self.element
            .set_attribute("table:number-rows-spanned", &span.to_string());
    }

    pub(crate) fn try_set_rowspan(&mut self, span: usize) -> Result<()> {
        let value = try_usize_string(span, "ODT expanded table-cell row span")?;
        self.element.try_set_attribute(
            "table:number-rows-spanned",
            &value,
            "ODT expanded table-cell row span",
        )
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
            .and_then(|value| usize::try_from(value).ok())
            .unwrap_or(1)
    }

    pub(crate) fn try_repeat_count(&self) -> Result<usize> {
        self.element
            .get_attribute("table:number-columns-repeated")
            .map(|value| {
                value.parse::<usize>().map_err(|_error| {
                    Error::InvalidFormat(
                        "table:number-columns-repeated must be a non-negative integer".to_string(),
                    )
                })
            })
            .transpose()
            .map(|value| value.unwrap_or(1))
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
            .and_then(|value| usize::try_from(value).ok())
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
        let mut tables = Vec::new();
        let mut stack: Vec<Element> = Vec::new();

        loop {
            match reader.read_event() {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let raw_name = e.name();
                    let tag_name = std::str::from_utf8(raw_name.as_ref()).map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid UTF-8 in ODT table element name: {error}"
                        ))
                    })?;

                    if tag_name == "table:table" {
                        let mut element = Element::try_new(tag_name)?;

                        // Parse attributes
                        for attr_result in e.attributes() {
                            let attr = attr_result.map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid ODT table attribute: {error}"
                                ))
                            })?;
                            let key = std::str::from_utf8(attr.key.as_ref()).map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid UTF-8 in ODT table attribute name: {error}"
                                ))
                            })?;
                            let value =
                                std::str::from_utf8(attr.value.as_ref()).map_err(|error| {
                                    Error::InvalidFormat(format!(
                                        "invalid UTF-8 in ODT table attribute value: {error}"
                                    ))
                                })?;
                            element.try_set_attribute(key, value, "ODT table attribute")?;
                        }

                        append_text_control(&reader, e, &mut element)?;
                        try_push(&mut stack, element, "ODT table parser stack")?;
                    } else if !stack.is_empty() {
                        // Handle nested elements within table
                        let mut element = Element::try_new(tag_name)?;

                        // Parse attributes
                        for attr_result in e.attributes() {
                            let attr = attr_result.map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid ODT table attribute: {error}"
                                ))
                            })?;
                            let key = std::str::from_utf8(attr.key.as_ref()).map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid UTF-8 in ODT table attribute name: {error}"
                                ))
                            })?;
                            let value =
                                std::str::from_utf8(attr.value.as_ref()).map_err(|error| {
                                    Error::InvalidFormat(format!(
                                        "invalid UTF-8 in ODT table attribute value: {error}"
                                    ))
                                })?;
                            element.try_set_attribute(key, value, "ODT table attribute")?;
                        }

                        append_text_control(&reader, e, &mut element)?;
                        try_push(&mut stack, element, "ODT table parser stack")?;
                    }
                },
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let raw_name = e.name();
                    let tag_name = std::str::from_utf8(raw_name.as_ref()).map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid UTF-8 in empty ODT table element name: {error}"
                        ))
                    })?;
                    if matches!(tag_name, "text:s" | "text:tab" | "text:line-break")
                        && let Some(parent) = stack.last_mut()
                    {
                        append_text_control(&reader, e, parent)?;
                        continue;
                    }
                    if tag_name == "table:table" || !stack.is_empty() {
                        let mut element = Element::try_new(tag_name)?;
                        for attr_result in e.attributes() {
                            let attr = attr_result.map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid empty ODT table attribute: {error}"
                                ))
                            })?;
                            let key = std::str::from_utf8(attr.key.as_ref()).map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid UTF-8 in empty ODT table attribute name: {error}"
                                ))
                            })?;
                            let value =
                                std::str::from_utf8(attr.value.as_ref()).map_err(|error| {
                                    Error::InvalidFormat(format!(
                                        "invalid UTF-8 in empty ODT table attribute value: {error}"
                                    ))
                                })?;
                            element.try_set_attribute(key, value, "ODT table attribute")?;
                        }
                        append_text_control(&reader, e, &mut element)?;
                        if let Some(parent) = stack.last_mut() {
                            parent.try_add_child(element, "ODT empty table child projection")?;
                        } else {
                            let table = Table::from_element(element)?;
                            try_push(&mut tables, table, "ODT table projection")?;
                        }
                    }
                },
                Ok(quick_xml::events::Event::Text(ref t)) => {
                    if let Some(current) = stack.last_mut() {
                        let text = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODT table text: {error}"))
                        })?;
                        current.try_append_text(&text, "ODT table text projection")?;
                    }
                },
                Ok(quick_xml::events::Event::CData(ref value)) => {
                    if let Some(current) = stack.last_mut() {
                        let text = value
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| {
                                Error::InvalidFormat(format!("invalid ODT table CDATA: {error}"))
                            })?;
                        current.try_append_text(&text, "ODT table CDATA projection")?;
                    }
                },
                Ok(quick_xml::events::Event::GeneralRef(ref reference)) => {
                    if let Some(current) = stack.last_mut() {
                        let text = super::parser::decode_reference(reference)?;
                        current.try_append_text(&text, "ODT table entity projection")?;
                    }
                },
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let raw_name = e.name();
                    let tag_name = std::str::from_utf8(raw_name.as_ref()).map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid UTF-8 in ODT table end name: {error}"
                        ))
                    })?;

                    if tag_name == "table:table" {
                        let table_element = stack.pop().ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODT table end tag has no matching start tag".to_string(),
                            )
                        })?;
                        let table = Table::from_element(table_element)?;
                        try_push(&mut tables, table, "ODT table projection")?;
                    } else if let Some(element) = stack.pop() {
                        if let Some(parent) = stack.last_mut() {
                            parent.try_add_child(element, "ODT table child projection")?;
                        }
                    }
                },
                Ok(quick_xml::events::Event::Eof) => {
                    if !stack.is_empty() {
                        return Err(Error::InvalidFormat(
                            "ODT table XML ended with open table elements".to_string(),
                        ));
                    }
                    break;
                },
                Err(error) => {
                    return Err(Error::InvalidFormat(format!(
                        "invalid ODT table XML: {error}"
                    )));
                },
                _ => {},
            }
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
                            .and_then(|value| usize::try_from(value).ok())
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
                                    new_cell.element.try_set_attribute(
                                        key,
                                        value,
                                        "ODT expanded table-cell attribute",
                                    )?;
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
    fn test_table_elements_preserve_cdata_entities_and_empty_text_controls() {
        let xml = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><table:table><table:table-row><table:table-cell><text:p><![CDATA[A]]>&amp;<text:s text:c="2"/><text:tab/><text:line-break/>B</text:p></table:table-cell></table:table-row></table:table></office:document>"#;

        let tables = TableElements::parse_tables(xml).unwrap();
        let rows = tables[0].rows().unwrap();
        let cells = rows[0].cells().unwrap();
        assert_eq!(cells[0].text().unwrap(), "A&  \t\nB");
    }

    #[test]
    fn test_table_elements_propagate_malformed_xml() {
        let xml = r#"<table:table xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><table:table-row></table:table>"#;
        assert!(TableElements::parse_tables(xml).is_err());
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
