//! XLS find/replace and template merge (record 0295) over the shared text
//! surface (`crate::find::surface`).
//!
//! As XLSX and XLSB: the workbook searches every worksheet; a sheet, cell,
//! range, row or column its own cells; each stored text value is one group;
//! formula cells and other values are not searched. Every changed cell is
//! one native value assignment. The BIFF8 editor admits a bounded number of
//! staged changes per edit, so changes are committed in native edits of
//! [`CHUNK`] cells, each on the previous one's result, inside one unified
//! mutation (as CSV import does, record 0266).
//!
//! A cell showing a formatted shared string loses its formatting runs when
//! its text changes; the native editor refuses removing the last reference
//! to such a string.

use std::collections::{BTreeMap, HashSet};
use std::mem::size_of;

use litchi_core::sheet::CellValue;
use litchi_core::{FileFormat, Position};
use litchi_xls::cell_values::{Reference, Selector, Value as NativeValue};

use super::super::node::{Item, Node};
use super::super::owner::SheetHandle;
use super::super::{Nav, Xls};
use super::{WriteCommit, finish_commit, native_edit_error, native_value_error};
use crate::error::{Code, Error, Result};
use crate::find::Spec;
use crate::find::surface::{self, Change, Run};
use crate::merge::{Options, Planner, Record, Summary};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// The cells one native BIFF8 edit stages before it commits (as import).
const CHUNK: usize = 1_024;

/// The BIFF8 grid: 65,536 rows by 256 columns (inclusive last indexes).
const LAST_ROW: u32 = (1 << 16) - 1;
const LAST_COLUMN: u32 = (1 << 8) - 1;

/// Inclusive `(first_row, first_column, last_row, last_column)`.
type Area = (u32, u32, u32, u32);

const ALL: Area = (0, 0, LAST_ROW, LAST_COLUMN);

/// Find and replace text in the cells below `target`: the chained commits
/// (empty for no change) and the match count.
pub(crate) fn find_replace(
    file: &Xls,
    target: &TargetArg<'_>,
    spec: &Spec<'_>,
    limits: &SemanticLimits,
) -> Result<(Vec<WriteCommit>, usize)> {
    file.check_source()?;
    let result = (|| {
        let replacement = surface::text_only(
            spec,
            FileFormat::Xls,
            "a spreadsheet find replaces cell text",
        )?;
        file.admit_edit_for_write()?;
        let mut scopes: Vec<(usize, Area)> = Vec::new();
        surface::each_target(Nav::new(file), target, |node: Node<'_>| {
            let sheet = node.sheet_position();
            let area = match node.item() {
                Item::Workbook => {
                    scopes.extend((0..sheet_count(file)?).map(|sheet| (sheet, ALL)));
                    return Ok(());
                },
                Item::Sheet(_) => ALL,
                Item::Cell(cell) => (cell.at.0, cell.at.1, cell.at.0, cell.at.1),
                Item::Range(range) => range.rect,
                Item::Row(row) => (row.index, 0, row.index, LAST_COLUMN),
                Item::Column(column) => (0, column.index, LAST_ROW, column.index),
                Item::NamedRange(_) | Item::Feature(_) => {
                    return Err(Error::new(
                        Code::UnsupportedCapability,
                        "find searches the workbook, sheets, cells, ranges, rows and columns",
                    )
                    .with_property(crate::find::FIND));
                },
            };
            scopes.push((sheet, area));
            Ok(())
        })?;
        let cells = read(file, &scopes)?;
        let runs = runs(&cells);
        let (changes, matches) = surface::find(&runs, spec, replacement, FileFormat::Xls, limits)?;
        Ok((commit(file, &cells, changes, limits)?, matches))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Xls));
    file.check_source()?;
    result
}

/// Fill `record` into every text cell of every worksheet: the chained
/// commits (empty when nothing matched).
pub(crate) fn merge(
    file: &Xls,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Vec<WriteCommit>, Summary)> {
    file.check_source()?;
    let result = (|| {
        let scopes: Vec<(usize, Area)> =
            (0..sheet_count(file)?).map(|sheet| (sheet, ALL)).collect();
        let cells = read(file, &scopes)?;
        let runs = runs(&cells);
        let mut planner = Planner::new(record, options, limits);
        let changes = surface::merge(&runs, &mut planner, true, FileFormat::Xls, limits)?;
        let summary = planner.finish();
        if changes.is_empty() {
            return Ok((Vec::new(), summary));
        }
        Ok((commit(file, &cells, changes, limits)?, summary))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Xls));
    file.check_source()?;
    result
}

fn sheet_count(file: &Xls) -> Result<usize> {
    let mut count = 0;
    while SheetHandle::at(file, count)?.is_some() {
        count += 1;
    }
    Ok(count)
}

/// One stored text cell.
struct TextCell {
    sheet: usize,
    at: (u32, u32),
    text: String,
}

/// The stored, non-formula text cells of `scopes`, in sheet and cell
/// order, each once.
fn read(file: &Xls, scopes: &[(usize, Area)]) -> Result<Vec<TextCell>> {
    let mut by_sheet: BTreeMap<usize, Vec<Area>> = BTreeMap::new();
    for (sheet, area) in scopes {
        by_sheet.entry(*sheet).or_default().push(*area);
    }
    let mut cells = Vec::new();
    for (sheet, areas) in by_sheet {
        let Some(handle) = SheetHandle::at(file, sheet)? else {
            continue;
        };
        if !handle.is_grid() {
            continue;
        }
        let grid = handle.grid()?;
        let Some(extent) = grid.extents() else {
            continue;
        };
        let mut seen = HashSet::new();
        let mut found = Vec::new();
        for area in areas {
            let area = (
                area.0.max(extent.0),
                area.1.max(extent.1),
                area.2.min(extent.2),
                area.3.min(extent.3),
            );
            if area.0 > area.2 || area.1 > area.3 {
                continue;
            }
            for (at, cell) in grid.cells_in(area) {
                if cell.formula {
                    continue;
                }
                let CellValue::String(text) = &cell.value else {
                    continue;
                };
                if seen.insert(at) {
                    found.push(TextCell {
                        sheet,
                        at,
                        text: text.clone(),
                    });
                }
            }
        }
        found.sort_by_key(|cell| cell.at);
        cells.extend(found);
    }
    Ok(cells)
}

fn runs(cells: &[TextCell]) -> Vec<Run<'_>> {
    cells
        .iter()
        .enumerate()
        .map(|(group, cell)| Run::new(cell.text.as_str(), group))
        .collect()
}

/// Commit the changes in native edits of [`CHUNK`] cells, chained.
fn commit(
    file: &Xls,
    cells: &[TextCell],
    changes: Vec<Change>,
    limits: &SemanticLimits,
) -> Result<Vec<WriteCommit>> {
    let mut assignments = Vec::with_capacity(changes.len());
    for change in changes {
        let cell = &cells[change.run];
        crate::import::check_cell_length(&change.text, FileFormat::Xls)?;
        if change.text.encode_utf16().count() > 65_535 {
            return Err(
                Error::new(Code::InvalidValue, "the XLS cell text is too long")
                    .with_property("value"),
            );
        }
        assignments.push((
            cell.sheet,
            Reference::new(cell.at.0, cell.at.1).map_err(native_value_error)?,
            change.text,
        ));
    }
    let mut commits: Vec<WriteCommit> = Vec::new();
    for chunk in assignments.chunks(CHUNK) {
        let base = commits
            .last()
            .map_or_else(|| file.clone(), |commit| commit.after.clone());
        let used = commits.iter().fold(0usize, |total, commit| {
            total.saturating_add(commit.snapshot_bytes)
        });
        let remaining = limits.max_snapshot_bytes().saturating_sub(used);
        let chunk_limits = limits.with_snapshot_budget(remaining);
        let additional = chunk.iter().fold(0usize, |total, (_, _, text)| {
            total
                .saturating_add(text.len())
                .saturating_add(size_of::<NativeValue>())
        });
        base.preflight_snapshot_bytes(remaining, additional)?;
        let mut edit = base
            .editable_package_for_write_with_snapshot_limit(remaining)?
            .edit()
            .map_err(native_edit_error)?;
        for (sheet, reference, text) in chunk {
            edit.set_value(
                Selector::Position(Position::new(*sheet)),
                *reference,
                NativeValue::Text(text.clone()),
            )
            .map_err(native_edit_error)?;
        }
        let native = edit.commit().map_err(native_edit_error)?;
        commits.push(finish_commit(&base, native, chunk.len(), &chunk_limits)?);
    }
    Ok(commits)
}
