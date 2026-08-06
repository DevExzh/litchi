//! Worksheet graph and package-boundary validation.

use super::model::{Cell, CellValue, Merge, Row, Sheet};
use litchi_core::{Error, Result};
use std::collections::HashSet;

/// Maximum XML/text payload retained by one worksheet field.
pub(crate) const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
/// Maximum number of logical rows accepted by the bounded worksheet facade.
pub(crate) const MAX_LOGICAL_ROWS: usize = 1_048_576;
/// Maximum number of logical columns accepted by the bounded worksheet facade.
pub(crate) const MAX_LOGICAL_COLUMNS: usize = 1_048_576;
/// Maximum number of logical cells accepted in one worksheet.
pub(crate) const MAX_LOGICAL_CELLS: usize = 4_194_304;
/// Maximum number of physical runs accepted in one worksheet.
pub(crate) const MAX_PHYSICAL_RUNS: usize = 262_144;
/// Maximum content XML size accepted by the worksheet codec.
pub(crate) const MAX_CONTENT_XML_BYTES: usize = 256 * 1024 * 1024;

pub(crate) fn validate_sheets(sheets: &[Sheet]) -> Result<()> {
    if sheets.len() > MAX_PHYSICAL_RUNS {
        return Err(Error::InvalidFormat(format!(
            "ODS sheet count exceeds the {MAX_PHYSICAL_RUNS} safety limit"
        )));
    }
    let mut names = HashSet::with_capacity(sheets.len());
    for sheet in sheets {
        if !names.insert(sheet.name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "ODS sheet name '{}' is duplicated",
                sheet.name
            )));
        }
        validate_sheet(sheet)?;
    }
    Ok(())
}

pub(crate) fn validate_sheet(sheet: &Sheet) -> Result<()> {
    validate_text(&sheet.name, "sheet name")?;
    if sheet.name.is_empty() {
        return Err(Error::InvalidFormat(
            "ODS sheet names must be non-empty".to_string(),
        ));
    }
    if let Some(style_name) = &sheet.style_name {
        validate_non_empty_text(style_name, "sheet style name")?;
    }
    if sheet.rows.len() > MAX_PHYSICAL_RUNS {
        return Err(Error::InvalidFormat(format!(
            "sheet '{}' exceeds the {MAX_PHYSICAL_RUNS} physical row-run safety limit",
            sheet.name
        )));
    }

    let mut logical_rows = 0usize;
    let mut logical_cells = 0usize;
    for row in &sheet.rows {
        validate_row(row)?;
        logical_rows = logical_rows.checked_add(row.repeat()).ok_or_else(|| {
            Error::InvalidFormat("ODS logical row count overflows address space".to_string())
        })?;
        if logical_rows > MAX_LOGICAL_ROWS {
            return Err(Error::InvalidFormat(format!(
                "sheet '{}' exceeds the {MAX_LOGICAL_ROWS} logical-row safety limit",
                sheet.name
            )));
        }
        let columns = row.logical_cell_count();
        if columns > MAX_LOGICAL_COLUMNS {
            return Err(Error::InvalidFormat(format!(
                "sheet '{}' exceeds the {MAX_LOGICAL_COLUMNS} logical-column safety limit",
                sheet.name
            )));
        }
        let row_cells = columns.checked_mul(row.repeat()).ok_or_else(|| {
            Error::InvalidFormat("ODS logical cell count overflows address space".to_string())
        })?;
        logical_cells = logical_cells.checked_add(row_cells).ok_or_else(|| {
            Error::InvalidFormat("ODS logical cell count overflows address space".to_string())
        })?;
        if logical_cells > MAX_LOGICAL_CELLS {
            return Err(Error::InvalidFormat(format!(
                "sheet '{}' exceeds the {MAX_LOGICAL_CELLS} logical-cell safety limit",
                sheet.name
            )));
        }
    }
    Ok(())
}

fn validate_row(row: &Row) -> Result<()> {
    validate_optional_name(row.style_name.as_deref(), "row style name")?;
    validate_optional_name(
        row.default_cell_style_name.as_deref(),
        "row default cell style name",
    )?;
    if row.cells.len() > MAX_PHYSICAL_RUNS {
        return Err(Error::InvalidFormat(format!(
            "row exceeds the {MAX_PHYSICAL_RUNS} physical cell-run safety limit"
        )));
    }
    validate_cell_runs(&row.cells)
}

pub(crate) fn validate_cell_runs(cells: &[Cell]) -> Result<()> {
    let mut columns = 0usize;
    for cell in cells {
        validate_cell(cell)?;
        columns = columns.checked_add(cell.repeat()).ok_or_else(|| {
            Error::InvalidFormat("ODS logical column count overflows address space".to_string())
        })?;
        if columns > MAX_LOGICAL_COLUMNS {
            return Err(Error::InvalidFormat(format!(
                "row exceeds the {MAX_LOGICAL_COLUMNS} logical-column safety limit"
            )));
        }
    }
    Ok(())
}

pub(crate) fn validate_cell(cell: &Cell) -> Result<()> {
    cell.value.validate()?;
    validate_text(&cell.text, "cell text")?;
    validate_optional_name(cell.style_name.as_deref(), "cell style name")?;
    if let Some(formula) = &cell.formula {
        validate_non_empty_text(formula, "cell formula")?;
    }
    match cell.merge {
        Merge::None => {},
        Merge::Span { rows, columns } => {
            if rows.get() == 1 && columns.get() == 1 {
                return Err(Error::InvalidFormat(
                    "ODS merge spans must cover more than one grid position".to_string(),
                ));
            }
        },
        Merge::Covered => {
            if !matches!(cell.value, CellValue::Empty)
                || !cell.text.is_empty()
                || cell.formula.is_some()
            {
                return Err(Error::InvalidFormat(
                    "ODS covered cells cannot carry a value, text, or formula".to_string(),
                ));
            }
        },
    }
    if matches!(cell.value, CellValue::Text(_)) && cell.text.is_empty() {
        // Empty strings are valid ODF strings; this branch documents that the
        // typed value is intentionally not collapsed into CellValue::Empty.
    }
    Ok(())
}

pub(crate) fn validate_content_xml_size(xml: &str) -> Result<()> {
    if xml.len() > MAX_CONTENT_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODS content.xml exceeds {MAX_CONTENT_XML_BYTES} bytes"
        )));
    }
    Ok(())
}

fn validate_optional_name(value: Option<&str>, label: &str) -> Result<()> {
    if let Some(value) = value {
        validate_non_empty_text(value, label)?;
    }
    Ok(())
}

fn validate_non_empty_text(value: &str, label: &str) -> Result<()> {
    validate_text(value, label)?;
    if value.is_empty() {
        return Err(Error::InvalidFormat(format!("{label} must be non-empty")));
    }
    Ok(())
}

pub(crate) fn validate_text(value: &str, label: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{label} exceeds the {MAX_TEXT_BYTES} byte safety limit"
        )));
    }
    if value.chars().any(|character| {
        matches!(character, '\u{0000}'..='\u{0008}' | '\u{000B}'..='\u{000C}' | '\u{000E}'..='\u{001F}')
    }) {
        return Err(Error::InvalidFormat(format!(
            "{label} contains an XML-forbidden control character"
        )));
    }
    Ok(())
}
