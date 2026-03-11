//! XLSX table structures and parsing.
//!
//! Tables in Excel provide structured references and enhanced formatting for data ranges.

use crate::ooxml::xlsx::sort::SortState;

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
            let col_count = max_col - min_col + 1;
            for i in 0..col_count {
                let col_id = min_col + i;
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

/// Parse a cell range like "A1:D10" into (min_col, min_row, max_col, max_row).
/// Returns 1-based indices.
fn parse_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (min_col, min_row) = parse_cell_ref(parts[0])?;
    let (max_col, max_row) = parse_cell_ref(parts[1])?;

    Some((min_col, min_row, max_col, max_row))
}

/// Parse a cell reference like "A1" into (col, row) with 1-based indices.
fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row_str = String::new();

    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }

    let row = row_str.parse::<u32>().ok()?;
    Some((col, row))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_style_info_new() {
        let style = TableStyleInfo::new();
        assert!(style.name.is_none());
        assert!(style.show_first_column.is_none());
        assert!(style.show_last_column.is_none());
        assert!(style.show_row_stripes.is_none());
        assert!(style.show_column_stripes.is_none());
    }

    #[test]
    fn test_table_style_info_default() {
        let style: TableStyleInfo = Default::default();
        assert!(style.name.is_none());
    }

    #[test]
    fn test_table_style_info_parse() {
        let tag = r#"name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0""#;
        let style = TableStyleInfo::parse(tag).unwrap();
        assert_eq!(style.name, Some("TableStyleMedium2".to_string()));
        assert_eq!(style.show_first_column, Some(true));
        assert_eq!(style.show_last_column, Some(false));
        assert_eq!(style.show_row_stripes, Some(true));
        assert_eq!(style.show_column_stripes, Some(false));
    }

    #[test]
    fn test_table_style_info_parse_partial() {
        let tag = r#"name="TableStyleLight1" showRowStripes="true""#;
        let style = TableStyleInfo::parse(tag).unwrap();
        assert_eq!(style.name, Some("TableStyleLight1".to_string()));
        assert_eq!(style.show_row_stripes, Some(true));
        assert!(style.show_first_column.is_none());
    }

    #[test]
    fn test_totals_row_function_as_str() {
        assert_eq!(TotalsRowFunction::Sum.as_str(), "sum");
        assert_eq!(TotalsRowFunction::Min.as_str(), "min");
        assert_eq!(TotalsRowFunction::Max.as_str(), "max");
        assert_eq!(TotalsRowFunction::Average.as_str(), "average");
        assert_eq!(TotalsRowFunction::Count.as_str(), "count");
        assert_eq!(TotalsRowFunction::CountNums.as_str(), "countNums");
        assert_eq!(TotalsRowFunction::StdDev.as_str(), "stdDev");
        assert_eq!(TotalsRowFunction::Var.as_str(), "var");
        assert_eq!(TotalsRowFunction::Custom.as_str(), "custom");
    }

    #[test]
    fn test_totals_row_function_parse() {
        assert_eq!(
            TotalsRowFunction::parse("sum"),
            Some(TotalsRowFunction::Sum)
        );
        assert_eq!(
            TotalsRowFunction::parse("min"),
            Some(TotalsRowFunction::Min)
        );
        assert_eq!(
            TotalsRowFunction::parse("max"),
            Some(TotalsRowFunction::Max)
        );
        assert_eq!(
            TotalsRowFunction::parse("average"),
            Some(TotalsRowFunction::Average)
        );
        assert_eq!(
            TotalsRowFunction::parse("count"),
            Some(TotalsRowFunction::Count)
        );
        assert_eq!(
            TotalsRowFunction::parse("countNums"),
            Some(TotalsRowFunction::CountNums)
        );
        assert_eq!(
            TotalsRowFunction::parse("stdDev"),
            Some(TotalsRowFunction::StdDev)
        );
        assert_eq!(
            TotalsRowFunction::parse("var"),
            Some(TotalsRowFunction::Var)
        );
        assert_eq!(
            TotalsRowFunction::parse("custom"),
            Some(TotalsRowFunction::Custom)
        );
        assert_eq!(TotalsRowFunction::parse("invalid"), None);
    }

    #[test]
    fn test_table_column_new() {
        let col = TableColumn::new(1u32, "Sales");
        assert_eq!(col.id, 1);
        assert_eq!(col.name, "Sales");
        assert!(col.unique_name.is_none());
        assert!(col.totals_row_function.is_none());
        assert!(col.totals_row_label.is_none());
        assert!(col.calculated_column_formula.is_none());
        assert!(col.totals_row_formula.is_none());
    }

    #[test]
    fn test_table_type_as_str() {
        assert_eq!(TableType::Worksheet.as_str(), "worksheet");
        assert_eq!(TableType::Xml.as_str(), "xml");
        assert_eq!(TableType::QueryTable.as_str(), "queryTable");
    }

    #[test]
    fn test_table_type_parse() {
        assert_eq!(TableType::parse("worksheet"), Some(TableType::Worksheet));
        assert_eq!(TableType::parse("xml"), Some(TableType::Xml));
        assert_eq!(TableType::parse("queryTable"), Some(TableType::QueryTable));
        assert_eq!(TableType::parse("invalid"), None);
    }

    #[test]
    fn test_table_new() {
        let table = Table::new(1u32, "Table1", "A1:D10");
        assert_eq!(table.id, 1);
        assert_eq!(table.name, "Table1");
        assert_eq!(table.display_name, "Table1");
        assert_eq!(table.ref_range, "A1:D10");
        assert_eq!(table.header_row_count, Some(1));
        assert!(table.columns.is_empty());
        assert!(table.comment.is_none());
        assert!(table.table_type.is_none());
    }

    #[test]
    fn test_table_initialize_columns() {
        let mut table = Table::new(1u32, "Table1", "A1:D10");
        table.initialize_columns();
        assert_eq!(table.columns.len(), 4);
        assert_eq!(table.columns[0].name, "Column1");
        assert_eq!(table.columns[3].name, "Column4");
        assert!(table.auto_filter_range.is_some());
    }

    #[test]
    fn test_table_column_names() {
        let mut table = Table::new(1u32, "Table1", "A1:C5");
        table.initialize_columns();
        let names = table.column_names();
        assert_eq!(names, vec!["Column1", "Column2", "Column3"]);
    }

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((1, 1)));
        assert_eq!(parse_cell_ref("B2"), Some((2, 2)));
        assert_eq!(parse_cell_ref("Z10"), Some((26, 10)));
        assert_eq!(parse_cell_ref("AA1"), Some((27, 1)));
        assert_eq!(parse_cell_ref("AB100"), Some((28, 100)));
    }

    #[test]
    fn test_parse_range() {
        assert_eq!(parse_range("A1:D10"), Some((1, 1, 4, 10)));
        assert_eq!(parse_range("B2:C5"), Some((2, 2, 3, 5)));
        assert_eq!(parse_range("A1"), None); // Missing colon
        assert_eq!(parse_range(""), None);
    }
}
