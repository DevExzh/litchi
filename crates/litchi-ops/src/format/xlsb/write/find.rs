//! XLSB find/replace and template merge (record 0295) over the shared text
//! surface (`crate::find::surface`).
//!
//! As XLSX (record 0273): the workbook searches every worksheet; a sheet,
//! cell, range, row or column its own cells. Each stored text value is one
//! group. Formula cells and other values are not searched. A rich-text cell
//! (an inline `BrtCellRString`, or a cell whose text is a formatted shared
//! string) is refused when it would change: a plain value would drop its
//! runs. Every changed cell is one native scalar assignment; one native
//! cell-values edit per changed worksheet, chained, in one unified commit.

use std::collections::{BTreeMap, HashSet};

use litchi_core::FileFormat;
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xlsb::sheet::{Rect, Selector as NativeSelector};

use super::super::node::{Grid, Item, Node};
use super::super::owner::SheetHandle;
use super::super::{Nav, Xlsb};
use super::{
    Commit, NativeCellOptions, NativeScalar, Reference, caller_value_error,
    cell_patch_container_bytes, cell_snapshot_budget, cell_transfer_limits, native_error,
    package_from_editor,
};
use crate::error::{Code, Error, Result};
use crate::find::Spec;
use crate::find::surface::{self, Change, Run};
use crate::merge::{Options, Planner, Record, Summary};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// The XLSB grid: 1,048,576 rows by 16,384 columns.
const MAX_ROWS: u32 = 1 << 20;
const MAX_COLUMNS: u32 = 1 << 14;

/// Find and replace text in the cells below `target`.
pub(crate) fn find_replace(
    file: &Xlsb,
    target: &TargetArg<'_>,
    spec: &Spec<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, usize)> {
    file.check_source()?;
    let result = (|| {
        // OfficeCLI: "Excel only supports 'find' with 'replace'".
        let replacement = surface::text_only(
            spec,
            FileFormat::Xlsb,
            "a spreadsheet find replaces cell text",
        )?;
        let mut scopes: Vec<(usize, Rect)> = Vec::new();
        surface::each_target(Nav::new(file), target, |node: Node<'_>| {
            let sheet = node.sheet_position();
            let area = match node.item() {
                Item::Workbook => {
                    for sheet in 0..file.catalog()?.len() {
                        scopes.push((sheet, Rect::ALL));
                    }
                    return Ok(());
                },
                Item::Sheet(_) => Rect::ALL,
                Item::Cell(cell) => rect(cell.at.0, cell.at.1, cell.at.0 + 1, cell.at.1 + 1)?,
                Item::Range(range) => range.rect,
                Item::Row(row) => rect(row.index, 0, row.index + 1, MAX_COLUMNS)?,
                Item::Column(column) => rect(0, column.index, MAX_ROWS, column.index + 1)?,
                _ => {
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
        let editor = file.editor_for_write()?;
        let rich = rich_texts(super::super::projection(&editor)?);
        let runs = runs(&cells, &rich);
        let (changes, matches) = surface::find(&runs, spec, replacement, FileFormat::Xlsb, limits)?;
        let commit = commit(
            file,
            package_from_editor(editor),
            &cells,
            changes,
            matches,
            limits,
        )?;
        Ok((commit, matches))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Xlsb));
    let freshness = file.check_source();
    freshness.and(result)
}

/// Fill `record` into every text cell of every worksheet. `None` means
/// nothing matched.
pub(crate) fn merge(
    file: &Xlsb,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    file.check_source()?;
    let result = (|| {
        let scopes: Vec<(usize, Rect)> = (0..file.catalog()?.len())
            .map(|sheet| (sheet, Rect::ALL))
            .collect();
        let cells = read(file, &scopes)?;
        let editor = file.editor_for_write()?;
        let rich = rich_texts(super::super::projection(&editor)?);
        let runs = runs(&cells, &rich);
        let mut planner = Planner::new(record, options, limits);
        let changes = surface::merge(&runs, &mut planner, true, FileFormat::Xlsb, limits)?;
        let summary = planner.finish();
        if changes.is_empty() {
            return Ok((None, summary));
        }
        let commit = commit(
            file,
            package_from_editor(editor),
            &cells,
            changes,
            summary.replacements().saturating_add(summary.fields()),
            limits,
        )?;
        Ok((Some(commit), summary))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Xlsb));
    let freshness = file.check_source();
    freshness.and(result)
}

/// One stored text cell.
struct TextCell {
    sheet: usize,
    at: (u32, u32),
    text: String,
    rich: bool,
}

/// The stored text cells of `scopes`, in sheet and cell order, each once.
fn read(file: &Xlsb, scopes: &[(usize, Rect)]) -> Result<Vec<TextCell>> {
    let mut by_sheet: BTreeMap<usize, Vec<Rect>> = BTreeMap::new();
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
        let grid = Grid::from_handle(&handle)?;
        let Some(extent) = grid.extents()? else {
            continue;
        };
        let mut seen = HashSet::new();
        let mut found = Vec::new();
        for area in areas {
            let Some(area) = intersect(area, extent) else {
                continue;
            };
            for (at, cell) in grid.cells(area)? {
                let CellValue::String(text) = cell.value() else {
                    continue;
                };
                if seen.insert(at) {
                    found.push(TextCell {
                        sheet,
                        at,
                        text: text.clone(),
                        rich: cell.rich_string().is_some_and(|rich| !rich.runs.is_empty()),
                    });
                }
            }
        }
        found.sort_by_key(|cell| cell.at);
        cells.extend(found);
    }
    Ok(cells)
}

/// The texts of the formatted shared strings: a cell showing one of them
/// may hold its runs.
fn rich_texts(workbook: &litchi_xlsb::Workbook) -> HashSet<&str> {
    workbook
        .shared_strings()
        .iter()
        .filter(|item| !item.runs.is_empty())
        .map(|item| item.text.as_str())
        .collect()
}

fn runs<'c>(cells: &'c [TextCell], rich: &HashSet<&str>) -> Vec<Run<'c>> {
    cells
        .iter()
        .enumerate()
        .map(|(group, cell)| {
            Run::new(cell.text.as_str(), group)
                .editable(!cell.rich && !rich.contains(cell.text.as_str()))
        })
        .collect()
}

/// One native cell-values edit per changed worksheet, chained.
fn commit(
    file: &Xlsb,
    package: litchi_xlsb::Package,
    cells: &[TextCell],
    changes: Vec<Change>,
    authored_operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let mut by_sheet: BTreeMap<usize, Vec<(Reference, String)>> = BTreeMap::new();
    for change in changes {
        let cell = &cells[change.run];
        crate::import::check_cell_length(&change.text, FileFormat::Xlsb)?;
        by_sheet.entry(cell.sheet).or_default().push((
            Reference::new(cell.at.0, cell.at.1).map_err(native_error)?,
            change.text,
        ));
    }
    let mut remaining_snapshot = cell_snapshot_budget(limits.max_snapshot_bytes())?;
    let canonical_limits = cell_transfer_limits(remaining_snapshot, limits)?;
    let canonical_sheet = match by_sheet.keys().next().copied() {
        Some(sheet) => Some(sheet),
        None => first_worksheet(file)?,
    };
    let mut package = package;
    let mut patches = Vec::new();
    let mut retained_snapshot = cell_patch_container_bytes();
    for (sheet, assignments) in by_sheet {
        let native_limits = cell_transfer_limits(remaining_snapshot, limits)?;
        let mut edit = package
            .edit_cell_values_with_options(
                NativeSelector::from(sheet),
                NativeCellOptions::new().transfer_limits(native_limits),
            )
            .map_err(native_error)?
            .resolve()
            .map_err(native_error)?
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSB worksheet position is out of range",
                )
            })?;
        for (reference, text) in assignments {
            edit.set_scalar(reference, NativeScalar::Text(text))
                .map_err(|error| caller_value_error(error, "value"))?;
        }
        let (next, patch) = edit.commit().map_err(native_error)?.into_parts();
        let retained = patch.retained_bytes();
        if retained > remaining_snapshot {
            let consumed = limits
                .max_snapshot_bytes()
                .saturating_sub(remaining_snapshot);
            return Err(super::snapshot_limit(
                consumed.saturating_add(retained),
                limits.max_snapshot_bytes(),
            ));
        }
        remaining_snapshot -= retained;
        retained_snapshot = retained_snapshot
            .checked_add(retained)
            .ok_or_else(|| super::snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
        package = next;
        patches.push(patch);
    }
    if patches.is_empty() {
        // Nothing changed: an exact no-op commit of the first worksheet.
        let canonical_sheet = canonical_sheet.ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the workbook has no worksheet to search",
            )
        })?;
        let edit = package
            .edit_cell_values_with_options(
                NativeSelector::from(canonical_sheet),
                NativeCellOptions::new().transfer_limits(canonical_limits),
            )
            .map_err(native_error)?
            .resolve()
            .map_err(native_error)?;
        let Some(edit) = edit else {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the workbook has no worksheet to search",
            ));
        };
        let (next, patch) = edit.commit().map_err(native_error)?.into_parts();
        let retained = patch.retained_bytes();
        if retained > remaining_snapshot {
            let consumed = limits
                .max_snapshot_bytes()
                .saturating_sub(remaining_snapshot);
            return Err(super::snapshot_limit(
                consumed.saturating_add(retained),
                limits.max_snapshot_bytes(),
            ));
        }
        retained_snapshot = retained_snapshot
            .checked_add(retained)
            .ok_or_else(|| super::snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
        package = next;
        patches.push(patch);
    }
    let canonical_sheet = canonical_sheet
        .ok_or_else(|| Error::new(Code::InternalError, "XLSB find lost its worksheet owner"))?;
    let commit = Commit::from_cell_parts(
        file,
        package,
        patches,
        (canonical_sheet, canonical_limits),
        retained_snapshot,
        limits.max_snapshot_bytes(),
        authored_operations,
    )?;
    Ok(commit)
}

/// Return the first complete-catalog position that is an actual worksheet.
/// Native XLSB selectors are worksheet-aware, so a no-match proof must never
/// silently fall back to selector zero or to a worksheet ordinal.
fn first_worksheet(file: &Xlsb) -> Result<Option<usize>> {
    let catalog = file.catalog()?;
    for position in 0..catalog.len() {
        let Some(handle) = SheetHandle::at(file, position)? else {
            continue;
        };
        if handle.is_grid() {
            return Ok(Some(position));
        }
    }
    Ok(None)
}

fn rect(row: u32, column: u32, end_row: u32, end_column: u32) -> Result<Rect> {
    Rect::at(row, column, end_row, end_column)
        .map_err(|_| Error::new(Code::OutOfRange, "the XLSB find area is outside the grid"))
}

fn intersect(left: Rect, right: Rect) -> Option<Rect> {
    let row = left.start().row().get().max(right.start().row().get());
    let column = left
        .start()
        .column()
        .get()
        .max(right.start().column().get());
    let end_row = left.end().0.min(right.end().0);
    let end_column = left.end().1.min(right.end().1);
    if row >= end_row || column >= end_column {
        return None;
    }
    Rect::at(row, column, end_row, end_column).ok()
}
