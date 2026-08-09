//! OPC relationship graph loading and authored-name resolution for pivot charts.

use super::{
    CHARTSHEET_CONTENT_TYPE, CHARTSHEET_REL, Chart, MAX_DRAWING_PART_BYTES,
    MAX_DRAWINGS_PER_WORKSHEET, MAX_PIVOT_CHARTS_PER_WORKSHEET, MAX_WORKBOOK_BYTES,
    STRICT_CHARTSHEET_REL, SheetCharts, SheetKind, Source, invalid, limit,
};
use crate::drawing::parse;
use crate::error::Result;
use crate::pivot::{PivotTable, read_pivot_tables};
use crate::raw::parse_catalog;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Part};

use super::codec::{parse_binding, xml_error};

/// Load every pivot chart reachable through the workbook relationship graph.
pub fn load(package: &OpcPackage) -> Result<Vec<SheetCharts>> {
    let workbook_part = package.main_document_part()?;
    let sheets = parse_workbook_sheets(workbook_part.blob())?;
    let tables = read_pivot_tables(package).map_err(|error| invalid(error.to_string()))?;
    let mut output = Vec::new();
    for sheet in &sheets {
        let Some((part_name, sheet_kind, charts)) =
            load_for_sheet(package, workbook_part, sheet, &tables)?
        else {
            continue;
        };
        if !charts.is_empty() {
            output.push(SheetCharts {
                sheet_name: sheet.name.clone(),
                sheet_part_name: part_name.to_string(),
                sheet_kind,
                charts,
            });
        }
    }
    Ok(output)
}

/// Load pivot charts hosted by one named worksheet or chartsheet.
pub fn load_sheet(package: &OpcPackage, sheet_name: &str) -> Result<Vec<Chart>> {
    let workbook_part = package.main_document_part()?;
    let sheets = parse_workbook_sheets(workbook_part.blob())?;
    let tables = read_pivot_tables(package).map_err(|error| invalid(error.to_string()))?;
    let sheet = sheets
        .iter()
        .find(|sheet| sheet.name == sheet_name)
        .ok_or_else(|| invalid(format!("worksheet '{sheet_name}' not found")))?;
    let Some((_, _, charts)) = load_for_sheet(package, workbook_part, sheet, &tables)? else {
        return Ok(Vec::new());
    };
    Ok(charts)
}

fn parse_workbook_sheets(xml: &[u8]) -> Result<Vec<crate::raw::Sheet>> {
    if xml.len() > MAX_WORKBOOK_BYTES {
        return Err(limit("workbook XML bytes"));
    }
    Ok(parse_catalog(xml)
        .map_err(|error| invalid(error.to_string()))?
        .sheets)
}

fn load_for_sheet(
    package: &OpcPackage,
    workbook_part: &dyn Part,
    sheet: &crate::raw::Sheet,
    tables: &[PivotTable],
) -> Result<Option<(PackURI, SheetKind, Vec<Chart>)>> {
    let relationship = workbook_part
        .rels()
        .get(&sheet.relationship_id)
        .ok_or_else(|| {
            invalid(format!(
                "worksheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
    let (sheet_kind, expected_content_type) = match relationship.reltype() {
        rt::WORKSHEET | rt::STRICT_WORKSHEET => (SheetKind::Worksheet, ct::SML_WORKSHEET),
        CHARTSHEET_REL | STRICT_CHARTSHEET_REL => (SheetKind::Chartsheet, CHARTSHEET_CONTENT_TYPE),
        // Dialog sheets, macro sheets, and other kinds cannot anchor charts
        // and are skipped gracefully.
        _ => return Ok(None),
    };
    if relationship.is_external() {
        return Err(invalid(format!(
            "worksheet '{}' relationship cannot be external",
            sheet.name
        )));
    }
    let sheet_uri = relationship.target_partname()?;
    let sheet_part = package.get_part(&sheet_uri)?;
    if sheet_part.content_type() != expected_content_type {
        return Err(invalid(format!(
            "sheet part has content type '{}', expected '{}'",
            sheet_part.content_type(),
            expected_content_type
        )));
    }
    let mut charts = Vec::new();
    let drawings: Vec<_> = sheet_part
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING))
        .collect();
    if drawings.len() > MAX_DRAWINGS_PER_WORKSHEET {
        return Err(limit("drawings per worksheet"));
    }
    for drawing_relationship in drawings {
        if drawing_relationship.is_external() {
            return Err(invalid("worksheet drawing relationship cannot be external"));
        }
        let drawing_uri = drawing_relationship.target_partname()?;
        let drawing_part = package.get_part(&drawing_uri)?;
        if drawing_part.content_type() != ct::OFC_DRAWING {
            return Err(invalid(format!(
                "drawing part has content type '{}', expected '{}'",
                drawing_part.content_type(),
                ct::OFC_DRAWING
            )));
        }
        if drawing_part.blob().len() > MAX_DRAWING_PART_BYTES {
            return Err(limit("drawing part bytes"));
        }
        let drawing_xml = std::str::from_utf8(drawing_part.blob()).map_err(xml_error)?;
        let drawing = parse(drawing_xml)?
            .ok_or_else(|| invalid(format!("drawing part '{drawing_uri}' has no wsDr root")))?;
        for chart in drawing.charts() {
            let chart_relationship =
                drawing_part
                    .rels()
                    .get(&chart.relationship_id)
                    .ok_or_else(|| {
                        invalid(format!(
                            "drawing chart references missing relationship '{}'",
                            chart.relationship_id
                        ))
                    })?;
            if !matches!(chart_relationship.reltype(), rt::CHART | rt::STRICT_CHART) {
                return Err(invalid(format!(
                    "drawing chart relationship '{}' has invalid type '{}'",
                    chart.relationship_id,
                    chart_relationship.reltype()
                )));
            }
            if chart_relationship.is_external() {
                return Err(invalid("drawing chart relationship cannot be external"));
            }
            let chart_uri = chart_relationship.target_partname()?;
            let chart_part = package.get_part(&chart_uri)?;
            if chart_part.content_type() != ct::DML_CHART {
                return Err(invalid(format!(
                    "chart part has content type '{}', expected '{}'",
                    chart_part.content_type(),
                    ct::DML_CHART
                )));
            }
            // Ordinary charts have no pivot source and are excluded.
            let Some(binding) = parse_binding(chart_part.blob())? else {
                continue;
            };
            if charts.len() >= MAX_PIVOT_CHARTS_PER_WORKSHEET {
                return Err(limit("pivot charts per worksheet"));
            }
            let pivot_table = resolve_table(&chart_uri, &binding.source, sheet, tables)?;
            charts.push(Chart {
                relationship_id: chart.relationship_id.clone(),
                part_name: chart_uri.to_string(),
                source: binding.source,
                series: binding.series,
                pivot_table: pivot_table.clone(),
            });
        }
    }
    Ok(Some((sheet_uri, sheet_kind, charts)))
}

/// Resolve a `c:pivotSource` name to the typed pivot-table model.
///
/// Names written by Excel are sheet-qualified (`[Book1.xlsx]Sheet1!Pivot1`);
/// an unqualified name resolves against the chart's own worksheet first and
/// then against the whole workbook, with ambiguity reported as an error.
fn resolve_table<'a>(
    chart_uri: &PackURI,
    source: &Source,
    sheet: &crate::raw::Sheet,
    tables: &'a [PivotTable],
) -> Result<&'a PivotTable> {
    let (sheet_prefix, table_name) = split_source_name(&source.name);
    if table_name.is_empty() {
        return Err(invalid(format!(
            "pivot chart '{chart_uri}' has an empty pivot-table name"
        )));
    }
    let folded = table_name.to_lowercase();
    let candidates = || {
        tables
            .iter()
            .filter(|table| table.name.to_lowercase() == folded)
    };
    if let Some(prefix) = sheet_prefix {
        let wanted = prefix.to_lowercase();
        candidates()
            .find(|table| table.sheet_name.to_lowercase() == wanted)
            .ok_or_else(|| {
                invalid(format!(
                    "pivot chart '{chart_uri}' references pivot table '{table_name}' on sheet '{prefix}', which does not host it"
                ))
            })
    } else {
        let host = sheet.name.to_lowercase();
        if let Some(table) = candidates().find(|table| table.sheet_name.to_lowercase() == host) {
            return Ok(table);
        }
        let mut matches = candidates();
        match (matches.next(), matches.next()) {
            (Some(table), None) => Ok(table),
            (None, _) => Err(invalid(format!(
                "pivot chart '{chart_uri}' references missing pivot table '{table_name}'"
            ))),
            (Some(_), Some(_)) => Err(invalid(format!(
                "pivot chart '{chart_uri}' pivot-table name '{table_name}' is ambiguous"
            ))),
        }
    }
}

/// Split a `c:pivotSource` name into its optional sheet qualifier and the
/// pivot-table name, stripping any `[workbook]` prefix and sheet quoting.
pub(crate) fn split_source_name(name: &str) -> (Option<String>, String) {
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

/// Quote a sheet name for use in a qualified pivot-source reference,
/// doubling embedded single quotes (the inverse of `unquote_sheet_name`).
#[cfg(test)]
fn quote_sheet_name(sheet: &str) -> String {
    let needs_quotes = sheet.is_empty()
        || sheet.chars().next().is_some_and(|c| c.is_ascii_digit())
        || !sheet
            .chars()
            .all(|c| c == '_' || c == '.' || c.is_alphanumeric());
    if needs_quotes {
        format!("'{}'", sheet.replace('\'', "''"))
    } else {
        sheet.to_string()
    }
}

/// Validate an authored pivot-source name against the workbook's authored
/// pivot tables (`(table name, hosting sheet name)` pairs) using the same
/// rules as [`resolve_table`], and return its canonical
/// sheet-qualified form so saved packages are valid by construction.
#[cfg(test)]
pub(crate) fn resolve_source_name(
    source_name: &str,
    host_sheet_name: &str,
    pivot_tables: &[(String, String)],
) -> Result<String> {
    let (sheet_prefix, table_name) = split_source_name(source_name);
    if table_name.is_empty() {
        return Err(invalid(
            "authored pivot chart has an empty pivot-table name",
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
                    invalid(format!(
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
                        return Err(invalid(format!(
                            "authored pivot chart references missing pivot table '{table_name}'"
                        )));
                    },
                    (Some(_), Some(_)) => {
                        return Err(invalid(format!(
                            "authored pivot chart pivot-table name '{table_name}' is ambiguous"
                        )));
                    },
                }
            }
        },
    };
    Ok(format!("{}!{}", quote_sheet_name(&found.1), found.0))
}
