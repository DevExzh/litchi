//! Header-row predicates on workbooks (record 0010 §5.4): `row[Salary>5000
//! and Region=EMEA]` names table columns by their headers.
//!
//! The binding follows OfficeCLI (`ExcelHandler.Query.RowWhere.cs`
//! `QueryRowsByColumnPredicate`, `ExcelHandler.Query.cs` `DetectTables`):
//!
//! 1. Every worksheet table (ListObject) in the scope whose header owns
//!    every key is a candidate: a key matches a column name without regard
//!    to case, or is column letters inside the table. Its data rows exclude
//!    the header and totals rows.
//! 2. Only when no table qualifies are *detected* tables tried: a block whose
//!    top-left cell has nothing above or to its left, whose first row is at
//!    least two contiguous non-empty cells starting with text, and which has
//!    at least one data row, ending at the first row that is blank across the
//!    header span; blocks overlapping a table are skipped.
//! 3. No candidate is `not_found`, reporting structural table counts and
//!    column positions (all positions up to 20, else the five nearest); more
//!    than one is `invalid_selector`.
//!
//! The binding runs once per query: table parts are read once, and detection
//! makes one pass over each worksheet's stored records.

use std::collections::{HashMap, HashSet};

use litchi_core::selector::{Position, column_number};
use litchi_xlsx::{Address, Rect};

use super::children::coordinate;
use super::nav::Nav;
use super::node::{Grid, Item, Node, SheetHandle, display_text};
use crate::error::{Code, Error, Result};
use crate::path::report::distance;
use crate::path::vocabulary::Case;
use crate::query::inspect::{HeaderBinding, HeaderKey};
use crate::value::parse_numeric_text;

/// OfficeCLI lists every column of a table this narrow.
const LIST_ALL: usize = 20;
/// Otherwise it suggests this many nearest columns.
const NEAREST: usize = 5;

/// One table a header-row predicate may bind to.
struct Header {
    sheet: usize,
    names: Vec<String>,
    /// Zero-based first and last columns.
    columns: (u32, u32),
    /// 1-based first and last data rows (empty when first > last).
    rows: (u32, u32),
    /// The whole block, for overlap with tables.
    rect: Rect,
}

/// The worksheets a binding searches: every worksheet for the workbook
/// root, the scope itself for a worksheet scope.
fn sheets(nav: Nav<'_>, scope: &Node<'_>) -> Result<Vec<(usize, SheetHandle)>> {
    let file = nav.file();
    let mut sheets = Vec::new();
    match scope.item() {
        Item::Workbook => {
            for position in 0..file.0.owner.len() {
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

/// The worksheet tables (ListObjects) of `sheets`.
fn tables(sheets: &[(usize, SheetHandle)]) -> Result<Vec<Header>> {
    let mut headers = Vec::new();
    for (position, handle) in sheets {
        for part in handle.tables()? {
            let table = part.table();
            // Malformed bands are typed read failures, not empty candidates.
            let rect = table.range()?;
            let (first_column, last_column) =
                (rect.start().column().get(), rect.end().1.saturating_sub(1));
            let rows = match table.data_range()? {
                Some(data) => (data.start().row().get() + 1, data.end().0),
                None => (rect.end().0 + 1, rect.end().0),
            };
            headers.push(Header {
                sheet: *position,
                names: table
                    .columns
                    .iter()
                    .map(|column| column.name.clone())
                    .collect(),
                columns: (first_column, last_column),
                rows,
                rect,
            });
        }
    }
    Ok(headers)
}

/// The detected tables of `sheets` that overlap none of `tables`.
fn detected(sheets: &[(usize, SheetHandle)], tables: &[Header]) -> Result<Vec<Header>> {
    let mut headers = Vec::new();
    for (position, handle) in sheets {
        let grid = handle.grid()?;
        let real: Vec<Rect> = tables
            .iter()
            .filter(|table| table.sheet == *position)
            .map(|table| table.rect)
            .collect();
        for block in detect(&grid, &real)? {
            headers.push(Header {
                sheet: *position,
                names: block.names,
                columns: (block.first_column, block.last_column),
                rows: (block.anchor_row + 2, block.last_row + 1),
                rect: block.rect,
            });
        }
    }
    Ok(headers)
}

/// One detected table: zero-based coordinates.
struct Block {
    first_column: u32,
    last_column: u32,
    anchor_row: u32,
    last_row: u32,
    names: Vec<String>,
    rect: Rect,
}

/// OfficeCLI's `DetectTables`: strict, high-precision detection of
/// table-shaped blocks of non-empty cells.
fn detect(grid: &Grid, real: &[Rect]) -> Result<Vec<Block>> {
    let Some(extent) = grid.extents()?.stored() else {
        return Ok(Vec::new());
    };
    // One pass: occupied coordinates `(column, row)` and whether each holds
    // text (not a formula, not a number).
    let mut occupied: HashMap<(u32, u32), bool> = HashMap::new();
    for (at, cell) in grid.cells(extent)? {
        let text = display_text(cell);
        if text.is_empty() {
            continue;
        }
        let is_text =
            !matches!(cell, litchi_xlsx::Cell::Formula(_)) && parse_numeric_text(text).is_none();
        occupied.insert((at.column().get(), at.row().get()), is_text);
    }
    let is_occupied = |column: Option<u32>, row: Option<u32>| match (column, row) {
        (Some(column), Some(row)) => occupied.contains_key(&(column, row)),
        _ => false,
    };
    let mut anchors: Vec<(u32, u32)> = occupied
        .keys()
        .copied()
        .filter(|&(column, row)| {
            !is_occupied(Some(column), row.checked_sub(1))
                && !is_occupied(column.checked_sub(1), Some(row))
        })
        .collect();
    anchors.sort_unstable_by_key(|&(column, row)| (row, column));

    let mut claimed: HashSet<(u32, u32)> = HashSet::new();
    let mut blocks = Vec::new();
    for (anchor_column, anchor_row) in anchors {
        if claimed.contains(&(anchor_column, anchor_row))
            || occupied.get(&(anchor_column, anchor_row)) != Some(&true)
        {
            continue;
        }
        let mut column = anchor_column;
        while occupied.contains_key(&(column, anchor_row)) {
            column += 1;
        }
        let last_column = column - 1;
        if last_column < anchor_column + 1 {
            continue;
        }
        let mut last_row = anchor_row;
        loop {
            let row = last_row + 1;
            if !(anchor_column..=last_column).any(|column| occupied.contains_key(&(column, row))) {
                break;
            }
            last_row = row;
        }
        if last_row == anchor_row {
            continue;
        }
        let start = Address::at(anchor_row, anchor_column).map_err(coordinate)?;
        let rect = Rect::new(start, last_row + 1, last_column + 1).map_err(coordinate)?;
        if real.iter().any(|table| overlaps(*table, rect)) {
            continue;
        }
        for row in anchor_row..=last_row {
            for column in anchor_column..=last_column {
                claimed.insert((column, row));
            }
        }
        let mut names = Vec::new();
        for column in anchor_column..=last_column {
            let at = Address::at(anchor_row, column).map_err(coordinate)?;
            names.push(grid.record(at)?.map_or("", display_text).to_owned());
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
    let (left_start, right_start) = (left.start(), right.start());
    left_start.column().get() < right.end().1
        && right_start.column().get() < left.end().1
        && left_start.row().get() < right.end().0
        && right_start.row().get() < left.end().0
}

/// The absolute zero-based column of every key in `header`, or `None` when
/// the header does not own them all. A header name wins over column letters.
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
                let matches = matching.count() + 2;
                return Err(Error::new(
                    Code::InvalidSelector,
                    format!("a header key matches {matches} columns in one table"),
                )
                .with_suggestion(format!(
                    "address it by column letter instead, for example '{}' or '{}'",
                    letters(header.columns.0 + index as u32),
                    letters(header.columns.0 + other as u32)
                )));
            }
            columns.push(header.columns.0 + index as u32);
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

/// The headers that own every key, with each key's column.
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

/// Binds a header-row predicate's keys to the one table in `scope` that owns
/// them all.
pub(super) fn bind(
    nav: Nav<'_>,
    scope: &Node<'_>,
    keys: &[HeaderKey<'_>],
) -> Result<HeaderBinding> {
    let sheets = sheets(nav, scope)?;
    let tables = tables(&sheets)?;
    let mut found = candidates(&tables, keys)?;
    // Detected tables are a heuristic: they are tried only when no table
    // (ListObject) owns the columns, so a real table never loses to a guess.
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
            columns: columns.iter().map(|column| column + 1).collect(),
        }),
        several => {
            let worksheet_count = several
                .iter()
                .map(|(header, _)| header.sheet)
                .collect::<HashSet<_>>()
                .len();
            let suggestion = if worksheet_count == 1 {
                "the matches are within one worksheet; scope the query to one worksheet or address one table by its cell range"
            } else {
                "scope the query to one worksheet or address one table by its cell range"
            };
            Err(Error::new(
                Code::InvalidSelector,
                format!(
                    "row[col op val] is ambiguous: the columns match in {} tables across {} worksheets",
                    several.len(),
                    worksheet_count,
                ),
            )
            .with_suggestion(suggestion))
        },
    }
}

/// Whether a table or detected table in `scope` has a column named `key`,
/// which makes the bare row-property key `row[key …]` ambiguous.
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

/// The `not_found` error of an unbound predicate.
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
        "row[col op val] matched no table in {place} containing all {} requested column(s) (examined {} table(s) and {} detected table(s), with {} named column(s))",
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

/// The distinct zero-based positions of non-empty header cells.
fn available_positions(headers: &[&Header]) -> Vec<u32> {
    let mut positions: Vec<u32> = headers
        .iter()
        .flat_map(|header| {
            header
                .names
                .iter()
                .enumerate()
                .filter(|(_, name)| !name.trim().is_empty())
                .map(|(index, _)| header.columns.0 + index as u32)
        })
        .collect();
    positions.sort_unstable();
    positions.dedup();
    positions
}

/// The distinct zero-based positions nearest to any requested key by
/// optimal-string-alignment distance.
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
                    (score, header.columns.0 + index as u32)
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

/// Column letters for a zero-based column.
fn letters(column: u32) -> String {
    super::node::Label::column(column).as_str().to_owned()
}
