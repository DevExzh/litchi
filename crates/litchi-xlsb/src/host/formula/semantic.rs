//! Formula semantic rendering independent of package relationship lookup.

use super::validation::validate_table_name;
use super::{TableNamedColumns, TableRowType};

pub(super) fn format_pivot_identifier(name: &str) -> String {
    if !name.eq_ignore_ascii_case("All")
        && !name.eq_ignore_ascii_case("Blank")
        && validate_table_name(name).is_ok()
    {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn escape_structured_column(name: &str) -> String {
    let mut escaped = String::with_capacity(name.len());
    for ch in name.chars() {
        if matches!(ch, '#' | '[' | ']' | '\'' | '@') {
            escaped.push('\'');
        }
        escaped.push(ch);
    }
    escaped
}

pub(super) fn format_structured_reference(
    table: &str,
    row_type: TableRowType,
    columns: &TableNamedColumns,
    square_bracket_space: bool,
    comma_space: bool,
) -> String {
    let mut items = Vec::new();
    match row_type {
        TableRowType::Data => {},
        TableRowType::All => items.push("[#All]".to_string()),
        TableRowType::Headers => items.push("[#Headers]".to_string()),
        TableRowType::DataAlternate => items.push("[#Data]".to_string()),
        TableRowType::DataAndHeaders => {
            items.push("[#Headers]".to_string());
            items.push("[#Data]".to_string());
        },
        TableRowType::Totals => items.push("[#Totals]".to_string()),
        TableRowType::DataAndTotals => {
            items.push("[#Data]".to_string());
            items.push("[#Totals]".to_string());
        },
        TableRowType::Current => items.push("[#This Row]".to_string()),
    }
    let has_range = matches!(columns, TableNamedColumns::Range { .. });
    match columns {
        TableNamedColumns::All => {},
        TableNamedColumns::One(name) => {
            items.push(format!("[{}]", escape_structured_column(name)));
        },
        TableNamedColumns::Range { first, last } => {
            items.push(format!(
                "[{}]:[{}]",
                escape_structured_column(first),
                escape_structured_column(last)
            ));
        },
    }
    if items.is_empty() {
        return table.to_string();
    }
    let separator = if comma_space { ", " } else { "," };
    let body = items.join(separator);
    let specifiers = if items.len() == 1 && !has_range {
        if square_bracket_space {
            format!("[ {} ]", &body[1..body.len() - 1])
        } else {
            body
        }
    } else if square_bracket_space {
        format!("[ {body} ]")
    } else {
        format!("[{body}]")
    };
    format!("{table}{specifiers}")
}
