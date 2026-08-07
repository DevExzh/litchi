//! XLSB-owned pivot-source resolution for authored charts.

use crate::package::error::{Error, Result};

/// Validate and canonicalize an authored pivot-table source name.
pub(crate) fn resolve_authored_pivot_source_name(
    pivot_source_name: &str,
    host_sheet_name: &str,
    pivot_tables: &[(String, String)],
) -> Result<String> {
    let (sheet_prefix, table_name) = split_pivot_source_name(pivot_source_name);
    if table_name.is_empty() {
        return Err(Error::InvalidFormat(
            "authored pivot chart has an empty pivot-table name".to_string(),
        ));
    }
    let folded = table_name.to_lowercase();
    let candidates = || {
        pivot_tables
            .iter()
            .filter(|(name, _)| name.to_lowercase() == folded)
    };
    let found = match sheet_prefix {
        Some(prefix) => {
            let wanted = prefix.to_lowercase();
            candidates()
                .find(|(_, sheet)| sheet.to_lowercase() == wanted)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "authored pivot chart references pivot table '{table_name}' on sheet '{prefix}', which does not host it"
                    ))
                })?
        },
        None => {
            let host = host_sheet_name.to_lowercase();
            if let Some(found) = candidates().find(|(_, sheet)| sheet.to_lowercase() == host) {
                found
            } else {
                let mut matches = candidates();
                match (matches.next(), matches.next()) {
                    (Some(found), None) => found,
                    (None, _) => {
                        return Err(Error::InvalidFormat(format!(
                            "authored pivot chart references missing pivot table '{table_name}'"
                        )));
                    },
                    (Some(_), Some(_)) => {
                        return Err(Error::InvalidFormat(format!(
                            "authored pivot chart pivot-table name '{table_name}' is ambiguous"
                        )));
                    },
                }
            }
        },
    };
    Ok(format!("{}!{}", quote_sheet_name(&found.1), found.0))
}

#[cfg(test)]
mod tests {
    use super::resolve_authored_pivot_source_name;

    #[test]
    fn unqualified_name_falls_back_to_unique_table_on_another_sheet() {
        let tables = [("Sales".to_string(), "Other".to_string())];
        let resolved = resolve_authored_pivot_source_name("Sales", "Host", &tables);

        assert!(matches!(resolved.as_deref(), Ok("Other!Sales")));
    }
}

fn split_pivot_source_name(name: &str) -> (Option<String>, String) {
    let Some(bang) = name.rfind('!') else {
        return (None, name.to_string());
    };
    let mut sheet = &name[..bang];
    if let Some(rest) = sheet.strip_prefix('[')
        && let Some(end) = rest.find(']')
    {
        sheet = &rest[end + 1..];
    }
    (
        Some(unquote_sheet_name(sheet)),
        name[bang + 1..].to_string(),
    )
}

fn unquote_sheet_name(sheet: &str) -> String {
    if sheet.len() >= 2 && sheet.starts_with('\'') && sheet.ends_with('\'') {
        sheet[1..sheet.len() - 1].replace("''", "'")
    } else {
        sheet.to_string()
    }
}

fn quote_sheet_name(sheet: &str) -> String {
    let needs_quotes = sheet.is_empty()
        || sheet
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_digit())
        || !sheet
            .chars()
            .all(|character| character == '_' || character == '.' || character.is_alphanumeric());
    if needs_quotes {
        format!("'{}'", sheet.replace('\'', "''"))
    } else {
        sheet.to_string()
    }
}
