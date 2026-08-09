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
        let semantic_columns = semantic_cell_footprint(row)?;
        let row_cells = semantic_columns.checked_mul(row.repeat()).ok_or_else(|| {
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

/// Count stored payload and grid coverage without charging sparse empty runs.
///
/// ODF producers commonly terminate a sheet with one repeated row containing
/// unstyled empty cell runs across the remaining grid height. Those runs are
/// address-space padding, not materialized cells. Direct/default cell styles,
/// values, displayed text, formulas, and merge coverage remain semantic and
/// therefore consume the logical-cell budget.
fn semantic_cell_footprint(row: &Row) -> Result<usize> {
    let default_cell_style = row.default_cell_style_name.is_some();
    row.cells.iter().try_fold(0usize, |total, cell| {
        let per_cell = match cell.merge {
            Merge::Span { rows, columns } => {
                rows.get().checked_mul(columns.get()).ok_or_else(|| {
                    Error::InvalidFormat(
                        "ODS merged-cell coverage overflows address space".to_string(),
                    )
                })?
            },
            Merge::Covered => 1,
            Merge::None
                if default_cell_style
                    || cell.style_name.is_some()
                    || !matches!(cell.value, CellValue::Empty)
                    || !cell.text.is_empty()
                    || cell.formula.is_some() =>
            {
                1
            },
            Merge::None => 0,
        };
        let run_footprint = per_cell.checked_mul(cell.repeat()).ok_or_else(|| {
            Error::InvalidFormat("ODS logical cell footprint overflows address space".to_string())
        })?;
        total.checked_add(run_footprint).ok_or_else(|| {
            Error::InvalidFormat("ODS logical cell footprint overflows address space".to_string())
        })
    })
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
    if let Some(name) = value {
        validate_non_empty_text(name, label)?;
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

#[cfg(test)]
mod sparse_accounting_tests {
    use super::{MAX_LOGICAL_CELLS, MAX_LOGICAL_COLUMNS, MAX_LOGICAL_ROWS, validate_sheet};
    use crate::worksheet::{Cell, CellValue, Merge, Row, Sheet};
    use std::num::NonZeroUsize;

    fn row(cell: Cell, repeat: usize) -> Row {
        Row {
            cells: vec![cell],
            style_name: None,
            default_cell_style_name: None,
            repeat: NonZeroUsize::new(repeat).expect("test fixture or operation should succeed"),
        }
    }

    fn sheet(rows: Vec<Row>) -> Sheet {
        Sheet {
            name: "Sheet1".to_string(),
            rows,
            style_name: None,
        }
    }

    #[test]
    fn meaningful_cells_accept_exact_limit_and_reject_one_more() {
        let cell = Cell::repeated(CellValue::Text("x".to_string()), "x", MAX_LOGICAL_COLUMNS)
            .expect("test fixture or operation should succeed");
        let exact_rows = MAX_LOGICAL_CELLS / MAX_LOGICAL_COLUMNS;
        let exact = row(cell, exact_rows);
        assert!(validate_sheet(&sheet(vec![exact.clone()])).is_ok());

        let one_more = row(Cell::new(CellValue::Text("x".to_string()), "x"), 1);
        assert!(validate_sheet(&sheet(vec![exact, one_more])).is_err());
    }

    #[test]
    fn trailing_empty_repeated_grid_padding_is_sparse() {
        let mut padding = row(
            Cell::repeated(CellValue::Empty, "", 7)
                .expect("test fixture or operation should succeed"),
            MAX_LOGICAL_ROWS,
        );
        padding.style_name = Some("row-style".to_string());
        assert!(validate_sheet(&sheet(vec![padding])).is_ok());
    }

    #[test]
    fn cell_formatting_and_coverage_still_consume_the_budget() {
        let mut styled = row(
            Cell::repeated(CellValue::Empty, "", MAX_LOGICAL_COLUMNS)
                .expect("test fixture or operation should succeed"),
            5,
        );
        styled.default_cell_style_name = Some("Default".to_string());
        assert!(validate_sheet(&sheet(vec![styled])).is_err());

        let mut covered = Cell::repeated(CellValue::Empty, "", MAX_LOGICAL_COLUMNS)
            .expect("test fixture or operation should succeed");
        covered.merge = Merge::Covered;
        assert!(validate_sheet(&sheet(vec![row(covered, 5)])).is_err());

        let mut span = Cell::empty();
        span.merge = Merge::Span {
            rows: NonZeroUsize::new(MAX_LOGICAL_ROWS)
                .expect("test fixture or operation should succeed"),
            columns: NonZeroUsize::new(5).expect("test fixture or operation should succeed"),
        };
        assert!(validate_sheet(&sheet(vec![row(span, 1)])).is_err());
    }
}
