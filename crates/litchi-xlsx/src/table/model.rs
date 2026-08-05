//! Typed SpreadsheetML table models and semantic validation.

use std::collections::HashSet;

use crate::error::{Result, invalid};
use crate::sort::SortState;

use super::{
    MAX_COLUMNS, MAX_EXCEL_COLUMN, MAX_EXCEL_ROW, MAX_SORT_CONDITIONS, MAX_TEXT_BYTES, limit,
};

/// Table style information for visual formatting.
#[derive(Debug, Clone)]
pub struct TableStyleInfo {
    /// Style name (e.g., "TableStyleMedium2")
    pub name: Option<String>,
    /// Show first column with special formatting
    pub show_first_column: Option<bool>,
    /// Show last column with special formatting
    pub show_last_column: Option<bool>,
    /// Show alternating row stripes
    pub show_row_stripes: Option<bool>,
    /// Show alternating column stripes
    pub show_column_stripes: Option<bool>,
}

impl TableStyleInfo {
    pub fn new() -> Self {
        Self {
            name: None,
            show_first_column: None,
            show_last_column: None,
            show_row_stripes: None,
            show_column_stripes: None,
        }
    }

    /// Parse table style info from XML tag.
    pub fn parse(tag: &str) -> Option<Self> {
        let name = Self::extract_attribute(tag, "name");
        let show_first_column = Self::extract_attribute(tag, "showFirstColumn")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        let show_last_column = Self::extract_attribute(tag, "showLastColumn")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        let show_row_stripes = Self::extract_attribute(tag, "showRowStripes")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        let show_column_stripes = Self::extract_attribute(tag, "showColumnStripes")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));

        Some(Self {
            name,
            show_first_column,
            show_last_column,
            show_row_stripes,
            show_column_stripes,
        })
    }

    fn extract_attribute(tag: &str, attr: &str) -> Option<String> {
        let search_str = format!("{}=\"", attr);
        let start = tag.find(&search_str)? + search_str.len();
        let end = tag[start..].find('"')? + start;
        Some(tag[start..end].to_string())
    }
}

impl Default for TableStyleInfo {
    fn default() -> Self {
        Self::new()
    }
}

/// Totals row function types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TotalsRowFunction {
    Sum,
    Min,
    Max,
    Average,
    Count,
    CountNums,
    StdDev,
    Var,
    Custom,
}

impl TotalsRowFunction {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Sum => "sum",
            Self::Min => "min",
            Self::Max => "max",
            Self::Average => "average",
            Self::Count => "count",
            Self::CountNums => "countNums",
            Self::StdDev => "stdDev",
            Self::Var => "var",
            Self::Custom => "custom",
        }
    }

    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "sum" => Some(Self::Sum),
            "min" => Some(Self::Min),
            "max" => Some(Self::Max),
            "average" => Some(Self::Average),
            "count" => Some(Self::Count),
            "countNums" => Some(Self::CountNums),
            "stdDev" => Some(Self::StdDev),
            "var" => Some(Self::Var),
            "custom" => Some(Self::Custom),
            _ => None,
        }
    }
}

/// A formula for a table column (calculated or totals row).
#[derive(Debug, Clone)]
pub struct TableFormula {
    /// Whether this is an array formula
    pub array: Option<bool>,
    /// Formula text
    pub text: String,
}

/// A single column in a table.
#[derive(Debug, Clone)]
pub struct TableColumn {
    /// Column ID (1-based)
    pub id: u32,
    /// Unique name (optional)
    pub unique_name: Option<String>,
    /// Display name
    pub name: String,
    /// Totals row function
    pub totals_row_function: Option<TotalsRowFunction>,
    /// Totals row label (for custom totals)
    pub totals_row_label: Option<String>,
    /// Calculated column formula
    pub calculated_column_formula: Option<TableFormula>,
    /// Totals row formula
    pub totals_row_formula: Option<TableFormula>,
}

impl TableColumn {
    /// Create a new table column.
    pub fn new(id: u32, name: impl Into<String>) -> Self {
        Self {
            id,
            unique_name: None,
            name: name.into(),
            totals_row_function: None,
            totals_row_label: None,
            calculated_column_formula: None,
            totals_row_formula: None,
        }
    }
}

/// Table type enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableType {
    Worksheet,
    Xml,
    QueryTable,
}

impl TableType {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Worksheet => "worksheet",
            Self::Xml => "xml",
            Self::QueryTable => "queryTable",
        }
    }

    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "worksheet" => Some(Self::Worksheet),
            "xml" => Some(Self::Xml),
            "queryTable" => Some(Self::QueryTable),
            _ => None,
        }
    }
}

/// An Excel table (structured data range).
///
/// Tables provide structured references in formulas and enhanced formatting.
#[derive(Debug, Clone)]
pub struct Table {
    /// Table ID (unique within workbook)
    pub id: u32,
    /// Internal name (used in formulas)
    pub name: String,
    /// Display name (shown in Excel UI)
    pub display_name: String,
    /// Comment/description
    pub comment: Option<String>,
    /// Cell range (e.g., "A1:D10")
    pub ref_range: String,
    /// Table type
    pub table_type: Option<TableType>,
    /// Number of header rows (usually 1)
    pub header_row_count: Option<u32>,
    /// Number of totals rows
    pub totals_row_count: Option<u32>,
    /// Whether totals row is shown
    pub totals_row_shown: Option<bool>,
    /// Published to server
    pub published: Option<bool>,
    /// Table columns
    pub columns: Vec<TableColumn>,
    /// Auto-filter configuration
    pub auto_filter_range: Option<String>,
    /// Sort state
    pub sort_state: Option<SortState>,
    /// Table style information
    pub style_info: Option<TableStyleInfo>,
}

impl Table {
    /// Create a new table with the given ID, name, and range.
    pub fn new(id: u32, name: impl Into<String>, ref_range: impl Into<String>) -> Self {
        let name_str = name.into();
        Self {
            id,
            name: name_str.clone(),
            display_name: name_str,
            comment: None,
            ref_range: ref_range.into(),
            table_type: None,
            header_row_count: Some(1),
            totals_row_count: None,
            totals_row_shown: None,
            published: None,
            columns: Vec::new(),
            auto_filter_range: None,
            sort_state: None,
            style_info: None,
        }
    }

    /// Initialize columns from range (creates default Column1, Column2, etc.).
    pub fn initialize_columns(&mut self) {
        if !self.columns.is_empty() {
            return;
        }

        // Parse range to determine column count
        if let Some((min_col, _min_row, max_col, _max_row)) = parse_range(&self.ref_range) {
            let Some(col_count) = max_col
                .checked_sub(min_col)
                .and_then(|width| width.checked_add(1))
                .filter(|count| *count as usize <= MAX_COLUMNS)
            else {
                return;
            };
            for i in 0..col_count {
                let Some(col_id) = min_col.checked_add(i) else {
                    return;
                };
                self.columns
                    .push(TableColumn::new(col_id, format!("Column{}", col_id)));
            }
        }

        // Set auto-filter if we have headers
        if self.header_row_count.unwrap_or(0) > 0 && self.auto_filter_range.is_none() {
            self.auto_filter_range = Some(self.ref_range.clone());
        }
    }

    /// Get column names.
    pub fn column_names(&self) -> Vec<&str> {
        self.columns.iter().map(|c| c.name.as_str()).collect()
    }
}

/// Validate a table before it is authored as SpreadsheetML.
pub fn validate_table(table: &Table) -> Result<()> {
    if table.id == 0 {
        return Err(invalid("table ID must be positive"));
    }
    bounded_nonempty(&table.name, "table name")?;
    bounded_nonempty(&table.display_name, "table display name")?;
    bounded(&table.comment, "table comment")?;
    validate_table_range(&table.ref_range, "table reference")?;
    if table.columns.is_empty() {
        return Err(invalid("table must contain at least one column"));
    }
    if table.columns.len() > MAX_COLUMNS {
        return Err(limit("table columns"));
    }
    let width = table_range_width(&table.ref_range)?;
    if usize::try_from(width) != Ok(table.columns.len()) {
        return Err(invalid(format!(
            "table range spans {width} columns, but {} columns were provided",
            table.columns.len()
        )));
    }
    if let Some(auto_filter_range) = table.auto_filter_range.as_deref() {
        ensure_range_contains(
            &table.ref_range,
            auto_filter_range,
            "table autoFilter reference",
        )?;
    }
    if let Some(sort_state) = table.sort_state.as_ref() {
        let auto_filter_range = table
            .auto_filter_range
            .as_deref()
            .ok_or_else(|| invalid("table sort state requires an auto-filter"))?;
        ensure_range_contains(
            auto_filter_range,
            &sort_state.ref_range,
            "table sortState reference",
        )?;
        if sort_state.conditions.len() > MAX_SORT_CONDITIONS {
            return Err(limit("table sort conditions"));
        }
        bounded_text(&sort_state.ref_range, "table sortState reference")?;
        for condition in &sort_state.conditions {
            ensure_range_contains(
                &sort_state.ref_range,
                &condition.ref_range,
                "table sort condition reference",
            )?;
            condition
                .validate()
                .map_err(|error| invalid(error.to_string()))?;
            if let Some(custom_list) = condition.custom_list.as_deref() {
                bounded_text(custom_list, "table custom sort list")?;
            }
        }
    }
    let mut ids = HashSet::with_capacity(table.columns.len());
    let mut names = HashSet::with_capacity(table.columns.len());
    for column in &table.columns {
        if column.id == 0 || !ids.insert(column.id) {
            return Err(invalid(format!(
                "invalid or duplicate table column ID {}",
                column.id
            )));
        }
        bounded_nonempty(&column.name, "table column name")?;
        if !names.insert(column.name.to_ascii_lowercase()) {
            return Err(invalid(format!(
                "empty or duplicate table column name '{}'",
                column.name
            )));
        }
        bounded(&column.unique_name, "table column unique name")?;
        bounded(&column.totals_row_label, "table totals row label")?;
        validate_formula(
            column.calculated_column_formula.as_ref(),
            "calculated column",
        )?;
        validate_formula(column.totals_row_formula.as_ref(), "totals row")?;
    }
    if let Some(style) = table.style_info.as_ref() {
        bounded(&style.name, "table style name")?;
    }
    Ok(())
}

fn validate_formula(formula: Option<&TableFormula>, description: &str) -> Result<()> {
    if let Some(formula) = formula {
        bounded_text(&formula.text, description)?;
    }
    Ok(())
}

fn bounded(value: &Option<String>, description: &str) -> Result<()> {
    if let Some(value) = value.as_deref() {
        bounded_text(value, description)?;
    }
    Ok(())
}

fn bounded_text(value: &str, description: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(limit(description));
    }
    Ok(())
}

fn bounded_nonempty(value: &str, description: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{description} cannot be empty")));
    }
    bounded_text(value, description)
}

pub(super) fn validate_table_range(range: &str, description: &str) -> Result<()> {
    let mut references = range.split(':');
    let first = references
        .next()
        .ok_or_else(|| invalid(format!("empty {description}")))?;
    let (first_column, first_row) = checked_cell_ref(first, description)?;
    if let Some(second) = references.next() {
        let (second_column, second_row) = checked_cell_ref(second, description)?;
        if second_column < first_column || second_row < first_row {
            return Err(invalid(format!(
                "{description} range '{range}' is descending"
            )));
        }
    }
    if references.next().is_some() {
        return Err(invalid(format!("invalid {description} range '{range}'")));
    }
    Ok(())
}

pub(super) fn ensure_range_contains(
    container: &str,
    nested: &str,
    description: &str,
) -> Result<()> {
    let (container_start, container_end) = table_range_bounds(container)?;
    let (nested_start, nested_end) = table_range_bounds(nested)?;
    if nested_start.0 < container_start.0
        || nested_start.1 < container_start.1
        || nested_end.0 > container_end.0
        || nested_end.1 > container_end.1
    {
        return Err(invalid(format!(
            "{description} '{nested}' is outside containing range '{container}'"
        )));
    }
    Ok(())
}

fn table_range_bounds(range: &str) -> Result<((u32, u32), (u32, u32))> {
    let mut references = range.split(':');
    let first = references
        .next()
        .ok_or_else(|| invalid("empty table range"))?;
    let second = references.next().unwrap_or(first);
    if references.next().is_some() {
        return Err(invalid(format!("invalid table range '{range}'")));
    }
    let start = checked_cell_ref(first, "table range")?;
    let end = checked_cell_ref(second, "table range")?;
    if end.0 < start.0 || end.1 < start.1 {
        return Err(invalid(format!("table range '{range}' is descending")));
    }
    Ok((start, end))
}

pub(super) fn table_range_width(range: &str) -> Result<u32> {
    let ((first_column, _), (second_column, _)) = table_range_bounds(range)?;
    second_column
        .checked_sub(first_column)
        .and_then(|width| width.checked_add(1))
        .ok_or_else(|| invalid(format!("table range '{range}' has descending columns")))
}

fn checked_cell_ref(reference: &str, description: &str) -> Result<(u32, u32)> {
    parse_cell_ref(&reference.replace('$', "")).ok_or_else(|| {
        invalid(format!(
            "invalid {description} cell reference '{reference}'"
        ))
    })
}

/// Parse a cell range like "A1:D10" into (min_col, min_row, max_col, max_row).
/// Returns 1-based indices.
pub(super) fn parse_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    let normalized = range.replace('$', "");
    let parts: Vec<&str> = normalized.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (min_col, min_row) = parse_cell_ref(parts[0])?;
    let (max_col, max_row) = parse_cell_ref(parts[1])?;

    Some((min_col, min_row, max_col, max_row))
}

/// Parse a cell reference like "A1" into (col, row) with 1-based indices.
pub(super) fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let bytes = cell_ref.as_bytes();
    let column_end = bytes.iter().position(u8::is_ascii_digit)?;
    if column_end == 0 || column_end == bytes.len() {
        return None;
    }
    let mut column = 0u32;
    for byte in &bytes[..column_end] {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        let digit = u32::from(byte.to_ascii_uppercase() - b'A' + 1);
        column = column.checked_mul(26)?.checked_add(digit)?;
    }
    if column == 0 || column > MAX_EXCEL_COLUMN {
        return None;
    }
    let row_bytes = &bytes[column_end..];
    if !row_bytes.iter().all(u8::is_ascii_digit) {
        return None;
    }
    let row = std::str::from_utf8(row_bytes).ok()?.parse::<u32>().ok()?;
    if row == 0 || row > MAX_EXCEL_ROW {
        return None;
    }
    Some((column, row))
}
