//! Structured reference parsing for Excel-style table formulas.
//!
//! This module parses Excel's structured reference syntax used in tables:
//! - `TableName` - entire table including headers
//! - `TableName[#All]` - entire table (explicit)
//! - `TableName[#Data]` - data rows only (excludes headers/totals)
//! - `TableName[#Headers]` - header row only
//! - `TableName[#Totals]` - totals row only
//! - `TableName[@]` - current row (intersection with implicit row context)
//! - `TableName[@ColumnName]` - specific column in current row
//! - `TableName[ColumnName]` - entire column (data rows)
//! - `TableName[[Column1]:[Column2]]` - multiple columns range
//! - `TableName[#Headers],[ColumnName]` - header cell for column
//! - `TableName[#Totals],[ColumnName]` - totals cell for column

/// Parsed structured reference components.
#[derive(Debug, Clone, PartialEq)]
pub enum StructuredReference {
    /// Entire table reference (includes headers, data, optionally totals).
    WholeTable { table_name: String },
    /// Table data rows only (excludes headers and totals).
    DataOnly { table_name: String },
    /// Table headers row only.
    Headers { table_name: String },
    /// Table totals row only.
    Totals { table_name: String },
    /// All rows (headers + data + totals).
    All { table_name: String },
    /// Current row intersection ([@]).
    ThisRow { table_name: String },
    /// Specific column (data rows only, no header).
    Column {
        table_name: String,
        column_name: String,
    },
    /// Specific column in current row ([@ColumnName]).
    ColumnThisRow {
        table_name: String,
        column_name: String,
    },
    /// Multiple columns range ([[Col1]:[Col2]]).
    ColumnRange {
        table_name: String,
        start_column: String,
        end_column: String,
    },
    /// Header cell for a specific column.
    HeaderColumn {
        table_name: String,
        column_name: String,
    },
    /// Totals cell for a specific column.
    TotalsColumn {
        table_name: String,
        column_name: String,
    },
}

/// Parse a structured reference string.
///
/// Returns `Some(StructuredReference)` if the input matches structured reference syntax,
/// or `None` if it's not a structured reference.
pub fn parse_structured_reference(input: &str) -> Option<StructuredReference> {
    let s = input.trim();
    if s.is_empty() {
        return None;
    }

    // Check if contains '[' - basic indicator of structured reference
    if !s.contains('[') {
        // Could be a simple table name reference
        if is_valid_table_name(s) {
            return Some(StructuredReference::WholeTable {
                table_name: s.to_string(),
            });
        }
        return None;
    }

    // Split on first '['
    let bracket_pos = s.find('[')?;
    let table_name = s[..bracket_pos].trim();

    if table_name.is_empty() || !is_valid_table_name(table_name) {
        return None;
    }

    let rest = &s[bracket_pos..];
    if !rest.ends_with(']') {
        return None;
    }

    // Remove outer brackets
    let inner = &rest[1..rest.len() - 1];

    parse_table_specifier(table_name, inner)
}

fn parse_table_specifier(table_name: &str, spec: &str) -> Option<StructuredReference> {
    let spec = spec.trim();

    if spec.is_empty() {
        return Some(StructuredReference::WholeTable {
            table_name: table_name.to_string(),
        });
    }

    // Check for comma-separated compound specifiers like [#Headers],[ColumnName]
    if spec.contains(',') {
        return parse_compound_specifier(table_name, spec);
    }

    // Check for column range [[Col1]:[Col2]]
    if spec.starts_with('[') && spec.contains("]:") {
        return parse_column_range(table_name, spec);
    }

    // Single specifier
    match spec {
        "#All" | "#ALL" => Some(StructuredReference::All {
            table_name: table_name.to_string(),
        }),
        "#Data" | "#DATA" => Some(StructuredReference::DataOnly {
            table_name: table_name.to_string(),
        }),
        "#Headers" | "#HEADERS" => Some(StructuredReference::Headers {
            table_name: table_name.to_string(),
        }),
        "#Totals" | "#TOTALS" => Some(StructuredReference::Totals {
            table_name: table_name.to_string(),
        }),
        "@" => Some(StructuredReference::ThisRow {
            table_name: table_name.to_string(),
        }),
        s if s.starts_with('@') => {
            // [@ColumnName]
            let col_name = s[1..].trim();
            if col_name.is_empty() {
                return None;
            }
            Some(StructuredReference::ColumnThisRow {
                table_name: table_name.to_string(),
                column_name: col_name.to_string(),
            })
        },
        _ => {
            // Regular column name
            if is_valid_column_name(spec) {
                Some(StructuredReference::Column {
                    table_name: table_name.to_string(),
                    column_name: spec.to_string(),
                })
            } else {
                None
            }
        },
    }
}

fn parse_compound_specifier(table_name: &str, spec: &str) -> Option<StructuredReference> {
    // spec is like: #Headers],[Amount or #Totals],[Amount
    // Find the pattern ],[
    let split_pattern = "],[";
    let split_pos = spec.find(split_pattern)?;

    let first = spec[..split_pos]
        .trim()
        .trim_matches(|c| c == '[' || c == ']');
    let second = spec[split_pos + 3..]
        .trim()
        .trim_matches(|c| c == '[' || c == ']');

    // [#Headers],[ColumnName]
    if matches!(first, "#Headers" | "#HEADERS") && is_valid_column_name(second) {
        return Some(StructuredReference::HeaderColumn {
            table_name: table_name.to_string(),
            column_name: second.to_string(),
        });
    }

    // [#Totals],[ColumnName]
    if matches!(first, "#Totals" | "#TOTALS") && is_valid_column_name(second) {
        return Some(StructuredReference::TotalsColumn {
            table_name: table_name.to_string(),
            column_name: second.to_string(),
        });
    }

    None
}

fn parse_column_range(table_name: &str, spec: &str) -> Option<StructuredReference> {
    // [[Column1]:[Column2]]
    if !spec.starts_with('[') || !spec.ends_with(']') {
        return None;
    }

    let inner = &spec[1..spec.len() - 1];

    // Find the split point by looking for ]:[
    let split_pattern = "]:[";
    let split_pos = inner.find(split_pattern)?;

    let start_part = &inner[..split_pos];
    let end_part = &inner[split_pos + 3..]; // Skip ]:[

    let start_col = start_part.trim_matches(|c| c == '[' || c == ']').trim();
    let end_col = end_part.trim_matches(|c| c == '[' || c == ']').trim();

    if is_valid_column_name(start_col) && is_valid_column_name(end_col) {
        Some(StructuredReference::ColumnRange {
            table_name: table_name.to_string(),
            start_column: start_col.to_string(),
            end_column: end_col.to_string(),
        })
    } else {
        None
    }
}

fn is_valid_table_name(name: &str) -> bool {
    if name.is_empty() {
        return false;
    }

    // Table names can contain letters, numbers, underscores, and periods
    // but cannot start with a number
    let first = name.chars().next().unwrap();
    if first.is_ascii_digit() {
        return false;
    }

    name.chars()
        .all(|c| c.is_alphanumeric() || c == '_' || c == '.')
}

fn is_valid_column_name(name: &str) -> bool {
    !name.is_empty() && !name.starts_with('#') && !name.starts_with('@')
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_whole_table() {
        assert_eq!(
            parse_structured_reference("SalesTable"),
            Some(StructuredReference::WholeTable {
                table_name: "SalesTable".to_string(),
            })
        );
    }

    #[test]
    fn test_all_specifier() {
        assert_eq!(
            parse_structured_reference("SalesTable[#All]"),
            Some(StructuredReference::All {
                table_name: "SalesTable".to_string(),
            })
        );
    }

    #[test]
    fn test_data_only() {
        assert_eq!(
            parse_structured_reference("SalesTable[#Data]"),
            Some(StructuredReference::DataOnly {
                table_name: "SalesTable".to_string(),
            })
        );
    }

    #[test]
    fn test_headers() {
        assert_eq!(
            parse_structured_reference("SalesTable[#Headers]"),
            Some(StructuredReference::Headers {
                table_name: "SalesTable".to_string(),
            })
        );
    }

    #[test]
    fn test_totals() {
        assert_eq!(
            parse_structured_reference("SalesTable[#Totals]"),
            Some(StructuredReference::Totals {
                table_name: "SalesTable".to_string(),
            })
        );
    }

    #[test]
    fn test_this_row() {
        assert_eq!(
            parse_structured_reference("SalesTable[@]"),
            Some(StructuredReference::ThisRow {
                table_name: "SalesTable".to_string(),
            })
        );
    }

    #[test]
    fn test_column() {
        assert_eq!(
            parse_structured_reference("SalesTable[Amount]"),
            Some(StructuredReference::Column {
                table_name: "SalesTable".to_string(),
                column_name: "Amount".to_string(),
            })
        );
    }

    #[test]
    fn test_column_this_row() {
        assert_eq!(
            parse_structured_reference("SalesTable[@Amount]"),
            Some(StructuredReference::ColumnThisRow {
                table_name: "SalesTable".to_string(),
                column_name: "Amount".to_string(),
            })
        );
    }

    #[test]
    fn test_column_range() {
        assert_eq!(
            parse_structured_reference("SalesTable[[Q1]:[Q4]]"),
            Some(StructuredReference::ColumnRange {
                table_name: "SalesTable".to_string(),
                start_column: "Q1".to_string(),
                end_column: "Q4".to_string(),
            })
        );
    }

    #[test]
    fn test_header_column() {
        assert_eq!(
            parse_structured_reference("SalesTable[#Headers],[Amount]"),
            Some(StructuredReference::HeaderColumn {
                table_name: "SalesTable".to_string(),
                column_name: "Amount".to_string(),
            })
        );
    }

    #[test]
    fn test_totals_column() {
        assert_eq!(
            parse_structured_reference("SalesTable[#Totals],[Amount]"),
            Some(StructuredReference::TotalsColumn {
                table_name: "SalesTable".to_string(),
                column_name: "Amount".to_string(),
            })
        );
    }

    #[test]
    fn test_invalid_table_name() {
        assert_eq!(parse_structured_reference("123Invalid"), None);
    }

    #[test]
    fn test_case_insensitive_specifiers() {
        assert_eq!(
            parse_structured_reference("Table1[#DATA]"),
            Some(StructuredReference::DataOnly {
                table_name: "Table1".to_string(),
            })
        );
    }
}
