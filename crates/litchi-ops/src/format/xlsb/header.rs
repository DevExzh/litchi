//! Header-row predicate binding for XLSB workbooks (record 0010 §5.4).
//!
//! A binding first considers structured tables, then the same conservative
//! detected-table heuristic used by the XLSX adapter. Diagnostics contain only
//! structural counts and column positions; authored sheet, table, header and
//! predicate text never enters an error message or suggestion.

use std::collections::HashSet;

use litchi_core::selector::{Position, column_number};
use litchi_core::sheet::{Cell as CoreCell, CellValue};
use litchi_xlsb::sheet::Rect;

use super::children::coordinate;
use super::nav::Nav;
use super::node::{Grid, Item, Node, Sheet, display_text};
use super::owner::SheetHandle;
use crate::error::{Code, Error, Result};
use crate::path::report::distance;
use crate::path::vocabulary::Case;
use crate::query::inspect::{HeaderBinding, HeaderKey};
use crate::value::parse_numeric_text;

const LIST_ALL: usize = 20;
const NEAREST: usize = 5;

struct Header {
    sheet: usize,
    names: Vec<String>,
    columns: (u32, u32),
    /// 1-based first and last data rows, empty when first exceeds last.
    rows: (u32, u32),
    rect: Rect,
}

fn sheets(nav: Nav<'_>, scope: &Node<'_>) -> Result<Vec<(usize, SheetHandle)>> {
    let file = nav.file();
    let mut sheets = Vec::new();
    match scope.item() {
        Item::Workbook => {
            for position in 0..super::nav::tab_count(file)? {
                if let Some(handle) = SheetHandle::at(file, position)? {
                    if handle.is_grid() {
                        sheets.push((position, handle));
                    }
                }
            }
        },
        Item::Sheet(sheet) if sheet.handle.is_grid() => {
            sheets.push((sheet.position, sheet.handle.clone()));
        },
        _ => {},
    }
    Ok(sheets)
}

fn tables(sheets: &[(usize, SheetHandle)]) -> Result<Vec<Header>> {
    let mut headers = Vec::new();
    for (position, handle) in sheets {
        for table in handle.tables()? {
            let range = table.range;
            let rect = Rect::at(
                range.first_row,
                range.first_column,
                range.last_row.checked_add(1).ok_or_else(|| {
                    Error::new(
                        Code::OutOfRange,
                        "table range row exceeds the worksheet grid",
                    )
                })?,
                range.last_column.checked_add(1).ok_or_else(|| {
                    Error::new(
                        Code::OutOfRange,
                        "table range column exceeds the worksheet grid",
                    )
                })?,
            )
            .map_err(coordinate)?;
            let first_data = range.first_row.saturating_add(table.header_row_count);
            let last_data = range
                .last_row
                .checked_sub(table.totals_row_count)
                .unwrap_or(range.first_row.saturating_sub(1));
            let rows = if first_data <= last_data {
                (first_data.saturating_add(1), last_data.saturating_add(1))
            } else {
                (last_data.saturating_add(2), last_data.saturating_add(1))
            };
            let names = table
                .columns
                .iter()
                .map(|column| {
                    column
                        .name
                        .as_deref()
                        .or(column.caption.as_deref())
                        .unwrap_or("")
                        .to_owned()
                })
                .collect();
            headers.push(Header {
                sheet: *position,
                names,
                columns: (range.first_column, range.last_column),
                rows,
                rect,
            });
        }
    }
    Ok(headers)
}

fn detected(sheets: &[(usize, SheetHandle)], tables: &[Header]) -> Result<Vec<Header>> {
    let mut headers = Vec::new();
    for (position, handle) in sheets {
        let sheet = Sheet::new(handle.position(), handle.clone());
        let grid = sheet.grid()?;
        let real: Vec<Rect> = tables
            .iter()
            .filter(|table| table.sheet == *position)
            .map(|table| table.rect)
            .collect();
        for block in detect(grid, &real)? {
            headers.push(Header {
                sheet: *position,
                names: block.names,
                columns: (block.first_column, block.last_column),
                rows: (
                    block.anchor_row.saturating_add(2),
                    block.last_row.saturating_add(1),
                ),
                rect: block.rect,
            });
        }
    }
    Ok(headers)
}

struct Block {
    first_column: u32,
    last_column: u32,
    anchor_row: u32,
    last_row: u32,
    names: Vec<String>,
    rect: Rect,
}

fn occupied_at(grid: &Grid, row: Option<u32>, column: Option<u32>) -> Result<bool> {
    let Some((row, column)) = row.zip(column) else {
        return Ok(false);
    };
    Ok(grid
        .record((row, column))?
        .is_some_and(|cell| !display_text(cell).is_empty()))
}

fn contains(rect: Rect, (row, column): (u32, u32)) -> bool {
    let start = rect.start();
    let (end_row, end_column) = rect.end();
    (start.row().get()..end_row).contains(&row)
        && (start.column().get()..end_column).contains(&column)
}

fn detect(grid: &Grid, real: &[Rect]) -> Result<Vec<Block>> {
    let Some(extent) = grid.extents()? else {
        return Ok(Vec::new());
    };
    let (end_row, end_column) = extent.end();
    let mut claimed = Vec::new();
    let mut blocks = Vec::new();
    for ((anchor_row, anchor_column), cell) in grid.cells(extent)? {
        let text = display_text(cell);
        if text.is_empty() {
            continue;
        }
        if claimed
            .iter()
            .copied()
            .any(|rect| contains(rect, (anchor_row, anchor_column)))
            || occupied_at(grid, anchor_row.checked_sub(1), Some(anchor_column))?
            || occupied_at(grid, Some(anchor_row), anchor_column.checked_sub(1))?
        {
            continue;
        }
        let is_text = !matches!(cell.value(), CellValue::Formula { .. })
            && parse_numeric_text(&text).is_none();
        if !is_text {
            continue;
        }
        let mut last_column = anchor_column;
        while let Some(next_column) = last_column.checked_add(1) {
            if next_column >= end_column || !occupied_at(grid, Some(anchor_row), Some(next_column))?
            {
                break;
            }
            last_column = next_column;
        }
        if last_column == anchor_column {
            continue;
        }
        let mut last_row = anchor_row;
        while let Some(next_row) = last_row.checked_add(1) {
            if next_row >= end_row {
                break;
            }
            let mut row_has_content = false;
            for column in anchor_column..=last_column {
                if occupied_at(grid, Some(next_row), Some(column))? {
                    row_has_content = true;
                    break;
                }
            }
            if !row_has_content {
                break;
            }
            last_row = next_row;
        }
        if last_row == anchor_row {
            continue;
        }
        let rect = Rect::at(
            anchor_row,
            anchor_column,
            last_row.saturating_add(1),
            last_column.saturating_add(1),
        )
        .map_err(coordinate)?;
        if real.iter().any(|table| overlaps(*table, rect)) {
            continue;
        }
        claimed.push(rect);
        let mut names = Vec::new();
        for column in anchor_column..=last_column {
            let at = (anchor_row, column);
            names.push(
                grid.record(at)?
                    .map(|cell| display_text(cell).into_owned())
                    .unwrap_or_default(),
            );
        }
        blocks.push(Block {
            first_column: anchor_column,
            last_column,
            anchor_row,
            last_row,
            names,
            rect,
        });
    }
    Ok(blocks)
}

fn overlaps(left: Rect, right: Rect) -> bool {
    let left_start = left.start();
    let right_start = right.start();
    left_start.column().get() < right.end().1
        && right_start.column().get() < left.end().1
        && left_start.row().get() < right.end().0
        && right_start.row().get() < left.end().0
}

fn resolve(header: &Header, keys: &[HeaderKey<'_>]) -> Result<Option<Vec<u32>>> {
    let mut columns = Vec::with_capacity(keys.len());
    for key in keys {
        let mut matching = header
            .names
            .iter()
            .enumerate()
            .filter(|(_, name)| Case::Caseless.matches(name, key.name));
        if let Some((index, _)) = matching.next() {
            if let Some((other, _)) = matching.next() {
                let matches = matching.count().saturating_add(2);
                return Err(Error::new(
                    Code::InvalidSelector,
                    format!("a header key matches {matches} columns in one table"),
                )
                .with_suggestion(format!(
                    "address it by column letter instead, for example '{}' or '{}'",
                    letters(header.columns.0.saturating_add(index as u32)),
                    letters(header.columns.0.saturating_add(other as u32))
                )));
            }
            columns.push(header.columns.0.saturating_add(index as u32));
            continue;
        }
        let letters =
            key.name.len() <= 3 && key.name.bytes().all(|byte| byte.is_ascii_alphabetic());
        match letters.then(|| column_number(key.name)).flatten() {
            Some(number) if (header.columns.0..=header.columns.1).contains(&(number - 1)) => {
                columns.push(number - 1);
            },
            _ => return Ok(None),
        }
    }
    Ok(Some(columns))
}

fn candidates<'h>(
    headers: &'h [Header],
    keys: &[HeaderKey<'_>],
) -> Result<Vec<(&'h Header, Vec<u32>)>> {
    let mut candidates = Vec::new();
    for header in headers {
        if let Some(columns) = resolve(header, keys)? {
            candidates.push((header, columns));
        }
    }
    Ok(candidates)
}

pub(super) fn bind(
    nav: Nav<'_>,
    scope: &Node<'_>,
    keys: &[HeaderKey<'_>],
) -> Result<HeaderBinding> {
    let sheets = sheets(nav, scope)?;
    let tables = tables(&sheets)?;
    let mut found = candidates(&tables, keys)?;
    let detected = if found.is_empty() {
        detected(&sheets, &tables)?
    } else {
        Vec::new()
    };
    if found.is_empty() {
        found = candidates(&detected, keys)?;
    }
    match found.as_slice() {
        [] => Err(no_column(scope, keys, &tables, &detected)),
        [(header, columns)] => Ok(HeaderBinding {
            sheet: Position::new(header.sheet),
            rows: header.rows,
            columns: columns
                .iter()
                .map(|column| column.saturating_add(1))
                .collect(),
        }),
        several => {
            let worksheet_count = several
                .iter()
                .map(|(header, _)| header.sheet)
                .collect::<HashSet<_>>()
                .len();
            let suggestion = if worksheet_count == 1 {
                "scope the query to one worksheet or address one table by its cell range"
            } else {
                "scope the query to one worksheet or address one table by its cell range"
            };
            Err(Error::new(
                Code::InvalidSelector,
                format!(
                    "row predicate is ambiguous: {} matching tables across {} worksheets",
                    several.len(),
                    worksheet_count,
                ),
            )
            .with_suggestion(suggestion))
        },
    }
}

pub(super) fn names_column(nav: Nav<'_>, scope: &Node<'_>, key: &str) -> Result<bool> {
    let sheets = sheets(nav, scope)?;
    let tables = tables(&sheets)?;
    let named = |headers: &[Header]| {
        headers.iter().any(|header| {
            header
                .names
                .iter()
                .any(|name| Case::Caseless.matches(name.trim(), key))
        })
    };
    if named(&tables) {
        return Ok(true);
    }
    Ok(named(&detected(&sheets, &tables)?))
}

fn no_column(
    scope: &Node<'_>,
    keys: &[HeaderKey<'_>],
    tables: &[Header],
    detected: &[Header],
) -> Error {
    let place = match scope.item() {
        Item::Sheet(_) => "the selected worksheet",
        _ => "the workbook",
    };
    let headers: Vec<&Header> = tables
        .iter()
        .chain(detected)
        .filter(|header| header.names.iter().any(|name| !name.trim().is_empty()))
        .collect();
    let named_columns = headers
        .iter()
        .map(|header| {
            header
                .names
                .iter()
                .filter(|name| !name.trim().is_empty())
                .count()
        })
        .sum::<usize>();
    let message = format!(
        "row predicate matched no table in {place} containing all {} requested column(s) (examined {} table(s) and {} detected table(s), with {} named column(s))",
        keys.len(),
        tables.len(),
        detected.len(),
        named_columns,
    );
    if headers.is_empty() {
        return Error::new(Code::NotFound, message).with_suggestion(
            "create a table over the data, or address the cells directly with a bounded cell range",
        );
    }
    let wide = headers.iter().any(|header| {
        header
            .names
            .iter()
            .filter(|name| !name.trim().is_empty())
            .count()
            > LIST_ALL
    });
    let positions = if wide {
        nearest_positions(&headers, keys)
    } else {
        available_positions(&headers)
    };
    let positions = positions.into_iter().map(letters).collect::<Vec<_>>();
    let suggestion = if wide {
        format!("nearby column positions: {}", positions.join(", "))
    } else {
        format!("available column positions: {}", positions.join(", "))
    };
    Error::new(Code::NotFound, message).with_suggestion(suggestion)
}

fn available_positions(headers: &[&Header]) -> Vec<u32> {
    let mut positions: Vec<u32> = headers
        .iter()
        .flat_map(|header| {
            header
                .names
                .iter()
                .enumerate()
                .filter(|(_, name)| !name.trim().is_empty())
                .map(|(index, _)| header.columns.0.saturating_add(index as u32))
        })
        .collect();
    positions.sort_unstable();
    positions.dedup();
    positions
}

fn nearest_positions(headers: &[&Header], keys: &[HeaderKey<'_>]) -> Vec<u32> {
    let wanted: Vec<Vec<char>> = keys
        .iter()
        .map(|key| key.name.to_lowercase().chars().collect())
        .collect();
    let mut scored: Vec<(usize, u32)> = headers
        .iter()
        .flat_map(|header| {
            header
                .names
                .iter()
                .enumerate()
                .filter(|(_, name)| !name.trim().is_empty())
                .map(|(index, name)| {
                    let lower: Vec<char> = name.to_lowercase().chars().collect();
                    let score = wanted
                        .iter()
                        .map(|key| distance(key, &lower))
                        .min()
                        .unwrap_or(usize::MAX);
                    (score, header.columns.0.saturating_add(index as u32))
                })
        })
        .collect();
    scored.sort_by_key(|(score, _)| *score);
    let mut positions = Vec::with_capacity(NEAREST);
    for (_, position) in scored {
        if !positions.contains(&position) {
            positions.push(position);
            if positions.len() == NEAREST {
                break;
            }
        }
    }
    positions.sort_unstable();
    positions
}

fn letters(column: u32) -> String {
    super::node::Label::column(column).as_str().to_owned()
}
