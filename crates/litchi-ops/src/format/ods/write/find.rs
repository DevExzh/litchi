//! ODS find/replace and template merge (record 0295) over the shared text
//! surface (`crate::find::surface`).
//!
//! As the other workbooks: the workbook searches every sheet; a sheet,
//! cell, range, row or column its own cells; each logical text cell is one
//! group (a repeated cell run is one cell per logical coordinate); formula
//! cells and other values are not searched. A changed cell keeps its style,
//! merge role and range source; its text is written as one plain paragraph,
//! as `set value` writes it. A cell holding hyperlinks is refused when it
//! would change (the link ranges index the old text). Every changed cell is
//! staged through the worksheet family of one leaf edit and published by
//! one commit.

use std::collections::BTreeMap;

use litchi_core::FileFormat;
use litchi_ods::{Cell, CellValue, Sheet};

use super::super::node::{Item, Node, Rect};
use super::super::{Nav, Ods};
use super::{Commit, finish, limit, no_op, stage_cells};
use crate::error::{Code, Error, Result};
use crate::find::Spec;
use crate::find::surface::{self, Change, Run};
use crate::merge::{Options, Planner, Record, Summary};
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};

/// The ODF spreadsheet grid as LibreOffice bounds it (inclusive).
const LAST_ROW: u32 = (1 << 20) - 1;
const LAST_COLUMN: u32 = (1 << 14) - 1;
const ALL: Rect = (0, 0, LAST_ROW, LAST_COLUMN);

/// Find and replace text in the cells below `target`.
pub(crate) fn find_replace(
    file: &Ods,
    target: &TargetArg<'_>,
    spec: &Spec<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, usize)> {
    file.check_source()?;
    let result = (|| {
        let replacement = surface::text_only(
            spec,
            FileFormat::Ods,
            "a spreadsheet find replaces cell text",
        )?;
        file.admit_edit()?;
        let mut scopes: Vec<(u32, Rect)> = Vec::new();
        let sheets = file.worksheets()?;
        surface::each_target(Nav::new(file), target, |node: Node<'_>| {
            let sheet = node.sheet_position();
            let area =
                match node.item() {
                    Item::Workbook => {
                        scopes.extend((0..sheets.len()).filter_map(|sheet| {
                            u32::try_from(sheet).ok().map(|sheet| (sheet, ALL))
                        }));
                        return Ok(());
                    },
                    Item::Sheet(_) => ALL,
                    Item::Cell(cell) => (cell.at.0, cell.at.1, cell.at.0, cell.at.1),
                    Item::Range(range) => range.rect,
                    Item::Merge(merge) => merge.rect,
                    Item::Row(row) => (row.index, 0, row.index, LAST_COLUMN),
                    Item::Column(column) => (0, column.index, LAST_ROW, column.index),
                    Item::NamedRange(_) | Item::Comment(_) | Item::Feature(_) => {
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
        let cells = read(sheets, &scopes, limits)?;
        let runs = runs(&cells);
        let (changes, matches) = surface::find(&runs, spec, replacement, FileFormat::Ods, limits)?;
        Ok((commit(file, &cells, changes, limits)?, matches))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Ods));
    file.check_source()?;
    result
}

/// Fill `record` into every text cell of every sheet. `None` means
/// nothing matched.
pub(crate) fn merge(
    file: &Ods,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    file.check_source()?;
    let result = (|| {
        file.admit_edit()?;
        let sheets = file.worksheets()?;
        let scopes: Vec<(u32, Rect)> = (0..sheets.len())
            .filter_map(|sheet| u32::try_from(sheet).ok().map(|sheet| (sheet, ALL)))
            .collect();
        let cells = read(sheets, &scopes, limits)?;
        let runs = runs(&cells);
        let mut planner = Planner::new(record, options, limits);
        let changes = surface::merge(&runs, &mut planner, true, FileFormat::Ods, limits)?;
        let summary = planner.finish();
        if changes.is_empty() {
            return Ok((None, summary));
        }
        Ok((Some(commit(file, &cells, changes, limits)?), summary))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Ods));
    file.check_source()?;
    result
}

/// One logical text cell.
struct TextCell<'s> {
    sheet: u32,
    at: (u32, u32),
    cell: &'s Cell,
}

/// The logical, non-formula text cells of `scopes`, in sheet and cell
/// order, each once.
fn read<'s>(
    sheets: &'s [Sheet],
    scopes: &[(u32, Rect)],
    limits: &SemanticLimits,
) -> Result<Vec<TextCell<'s>>> {
    let mut by_sheet: BTreeMap<u32, Vec<Rect>> = BTreeMap::new();
    for (sheet, area) in scopes {
        by_sheet.entry(*sheet).or_default().push(*area);
    }
    let mut cells = Vec::new();
    for (sheet, areas) in by_sheet {
        let Some(data) = sheets.get(sheet as usize) else {
            continue;
        };
        let mut row = 0usize;
        for line in &data.rows {
            let rows = row..row.saturating_add(line.repeat());
            row = rows.end;
            let mut column = 0usize;
            for cell in line.cells() {
                let columns = column..column.saturating_add(cell.repeat());
                column = columns.end;
                if cell.formula.is_some() || !matches!(cell.value, CellValue::Text(_)) {
                    continue;
                }
                for at_row in rows.clone() {
                    for at_column in columns.clone() {
                        let (Ok(at_row), Ok(at_column)) =
                            (u32::try_from(at_row), u32::try_from(at_column))
                        else {
                            continue;
                        };
                        if !areas.iter().any(|area| {
                            (area.0..=area.2).contains(&at_row)
                                && (area.1..=area.3).contains(&at_column)
                        }) {
                            continue;
                        }
                        if cells.len() >= limits.max_expanded_targets() {
                            return Err(limit(
                                LimitKind::ExpandedTargets,
                                cells.len().saturating_add(1),
                                limits.max_expanded_targets(),
                            ));
                        }
                        cells.push(TextCell {
                            sheet,
                            at: (at_row, at_column),
                            cell,
                        });
                    }
                }
            }
        }
    }
    Ok(cells)
}

fn runs<'c>(cells: &'c [TextCell<'_>]) -> Vec<Run<'c>> {
    cells
        .iter()
        .enumerate()
        .map(|(group, cell)| {
            Run::new(cell.cell.text.as_str(), group).editable(cell.cell.hyperlinks.is_empty())
        })
        .collect()
}

/// Stage every changed cell through one leaf edit and commit once.
fn commit(
    file: &Ods,
    cells: &[TextCell<'_>],
    changes: Vec<Change>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    if changes.is_empty() {
        return no_op(file, limits);
    }
    let mut staged: BTreeMap<u32, BTreeMap<(u32, u32), Cell>> = BTreeMap::new();
    for change in changes {
        let cell = &cells[change.run];
        crate::import::check_cell_length(&change.text, FileFormat::Ods)?;
        let mut replacement = cell.cell.clone();
        replacement.value = CellValue::Text(change.text.clone());
        replacement.text = change.text;
        replacement.repeat = std::num::NonZeroUsize::MIN;
        staged
            .entry(cell.sheet)
            .or_default()
            .insert(cell.at, replacement);
    }
    let mut edit = file.leaf_edit()?;
    stage_cells(&mut edit, staged)?;
    finish(file, edit, 1, limits)
}
