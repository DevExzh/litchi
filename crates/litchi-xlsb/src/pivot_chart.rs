//! Small pivot-chart authoring helpers used by the XLSB chart facade.

use crate::package::error::{Error, Result};

/// Default `c:fmtId` written for authored pivot charts.
pub(crate) const DEFAULT_PIVOT_CHART_FORMAT_ID: u32 = 0;

const PIVOT_OPTIONS_EXTENSION_URI: &str = "{781A3756-C4B2-4CAC-9D66-4F8C8630D5DC}";
const C14_CHART_NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";

/// Return the default all-visible pivot-options extension list.
pub(crate) fn default_pivot_options_extension_xml() -> Vec<u8> {
    format!(
        r#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ext uri="{PIVOT_OPTIONS_EXTENSION_URI}" xmlns:c14="{C14_CHART_NAMESPACE}"><c14:pivotOptions><c14:dropZoneVisible val="1"/><c14:dropZoneCategories val="1"/><c14:dropZoneData val="1"/><c14:dropZoneSeries val="1"/><c14:dropZoneAxis val="1"/><c14:dropZoneValues val="1"/></c14:pivotOptions></c:ext></c:extLst>"#
    )
    .into_bytes()
}

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
    let mut candidates = pivot_tables
        .iter()
        .filter(|(name, _)| name.to_lowercase() == folded);
    let found = match sheet_prefix {
        Some(prefix) => {
            let wanted = prefix.to_lowercase();
            candidates
                .find(|(_, sheet)| sheet.to_lowercase() == wanted)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "authored pivot chart references pivot table '{table_name}' on sheet '{prefix}', which does not host it"
                    ))
                })?
        },
        None => {
            let host = host_sheet_name.to_lowercase();
            if let Some(found) = candidates.find(|(_, sheet)| sheet.to_lowercase() == host) {
                found
            } else {
                match (candidates.next(), candidates.next()) {
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
