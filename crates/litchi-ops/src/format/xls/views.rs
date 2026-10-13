//! OfficeCLI-shaped textual and JSON views for legacy XLS workbooks.

use std::fmt;
use std::ops::ControlFlow;

use litchi_core::sheet::CellValue;
use serde::Serialize;
use serde::ser::{SerializeMap, Serializer};

use super::Xls;
use super::node::{cell_type, display_value};
use super::owner::{Grid, SheetHandle};
use crate::error::{Code, Error, Result};
use crate::format::workbook::Label;
use crate::path::vocabulary::Case;
use crate::view::mode::{Mode, Options};

struct Worksheet<'f> {
    position: usize,
    name: String,
    handle: SheetHandle<'f>,
}

impl Worksheet<'_> {
    fn grid(&self) -> Result<Grid> {
        self.handle.grid()
    }
}

pub(crate) fn view_to<W: fmt::Write>(
    file: &Xls,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    file.finish(|| {
        let mut out = TrimEnd::new(out);
        match mode {
            Mode::Text => text(file, options, &mut out)?,
            Mode::Annotated => annotated(file, options, &mut out)?,
            Mode::Outline => outline(file, &mut out)?,
            Mode::Stats => stats(file)?.write(&mut out).map_err(sink)?,
            _ => return Err(mode.unsupported()),
        }
        Ok(())
    })
}

pub(crate) fn view_json(file: &Xls, mode: Mode, options: &Options) -> Result<Json> {
    file.finish(|| {
        Ok(match mode {
            Mode::Text => Json::Text(text_json(file, options)?),
            Mode::Annotated => {
                let mut content = String::new();
                {
                    let mut out = TrimEnd::new(&mut content);
                    annotated(file, options, &mut out)?;
                }
                Json::Annotated(Annotated {
                    view: "annotated",
                    content,
                })
            },
            Mode::Outline => Json::Outline(outline_json(file)?),
            Mode::Stats => Json::Stats(stats(file)?.json()),
            _ => return Err(mode.unsupported()),
        })
    })
}

fn sink(_: fmt::Error) -> Error {
    Error::new(Code::IoError, "the XLS view output sink refused to write")
        .with_format(litchi_core::FileFormat::Xls)
}

fn worksheets(file: &Xls) -> Result<Vec<Worksheet<'_>>> {
    let count = super::nav::tab_count(file)?;
    let mut sheets = Vec::new();
    sheets.try_reserve_exact(count).map_err(|_| {
        Error::new(
            Code::InternalError,
            "could not reserve the XLS view catalog",
        )
        .with_format(litchi_core::FileFormat::Xls)
    })?;
    for position in 0..count {
        if let Some(handle) = SheetHandle::at(file, position)? {
            if !handle.is_grid() {
                continue;
            }
            sheets.push(Worksheet {
                position,
                name: handle.name().to_owned(),
                handle,
            });
        }
    }
    Ok(sheets)
}

fn same_sheet(name: &str, wanted: &str) -> bool {
    Case::Caseless.matches(name, wanted)
}

fn row_indices(grid: &Grid) -> Vec<u32> {
    grid.row_indices()
}

fn visit_rows(
    grid: &Grid,
    rows: &[u32],
    mut visit: impl FnMut(u32, u32, &Grid) -> Result<ControlFlow<()>>,
) -> Result<()> {
    for (ordinal, row) in rows.iter().copied().enumerate() {
        let ordinal = u32::try_from(ordinal).unwrap_or(u32::MAX);
        if matches!(visit(row, ordinal, grid)?, ControlFlow::Break(())) {
            break;
        }
    }
    Ok(())
}

fn text(file: &Xls, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let sheets = worksheets(file)?;
    let clip = options.range.as_ref();
    let mut emitted = 0_u32;
    for (sheet_index, sheet) in sheets.iter().enumerate() {
        if clip.is_some_and(|range| !same_sheet(&sheet.name, &range.sheet)) {
            continue;
        }
        writeln!(out, "=== Sheet: {} ===", sheet.name).map_err(sink)?;
        let grid = sheet.grid()?;
        let rows = row_indices(&grid);
        let total = rows.len();
        let mut stop = false;
        visit_rows(&grid, &rows, |row, ordinal, grid| {
            if options
                .start
                .is_some_and(|start| ordinal.saturating_add(1) < start)
            {
                return Ok(ControlFlow::Continue(()));
            }
            if options
                .end
                .is_some_and(|end| ordinal.saturating_add(1) > end)
            {
                return Ok(ControlFlow::Break(()));
            }
            if clip.is_some_and(|range| row < range.first.1 || row > range.last.1) {
                return Ok(ControlFlow::Continue(()));
            }
            if options.max_lines.is_some_and(|max| emitted >= max) {
                truncation(out, emitted, total).map_err(sink)?;
                stop = true;
                return Ok(ControlFlow::Break(()));
            }
            write!(out, "[/sheet[{}]/row[{}]] ", sheet.position + 1, row + 1).map_err(sink)?;
            let mut first = true;
            for ((cell_row, column), cell) in grid.cells_in((row, 0, row, u32::MAX)) {
                if matches!(&cell.value, CellValue::Empty)
                    || !options.shows_column(column)
                    || clip.is_some_and(|range| column < range.first.0 || column > range.last.0)
                {
                    continue;
                }
                if !first {
                    out.write_str("\t").map_err(sink)?;
                }
                first = false;
                write!(
                    out,
                    "{}={}",
                    Label::cell(column, cell_row).as_str(),
                    display_value(&cell.value)
                )
                .map_err(sink)?;
            }
            out.write_char('\n').map_err(sink)?;
            emitted = emitted.saturating_add(1);
            Ok(ControlFlow::Continue(()))
        })?;
        if stop {
            return Ok(());
        }
        if sheet_index + 1 < sheets.len() {
            out.write_char('\n').map_err(sink)?;
        }
    }
    Ok(())
}

fn annotated(file: &Xls, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let mut emitted = 0_u32;
    for sheet in worksheets(file)? {
        writeln!(out, "=== Sheet: {} ===", sheet.name).map_err(sink)?;
        let grid = sheet.grid()?;
        let rows = row_indices(&grid);
        let mut stop = false;
        visit_rows(&grid, &rows, |row, ordinal, grid| {
            if options
                .start
                .is_some_and(|start| ordinal.saturating_add(1) < start)
            {
                return Ok(ControlFlow::Continue(()));
            }
            if options
                .end
                .is_some_and(|end| ordinal.saturating_add(1) > end)
            {
                return Ok(ControlFlow::Break(()));
            }
            if options.max_lines.is_some_and(|max| emitted >= max) {
                truncation(out, emitted, rows.len()).map_err(sink)?;
                stop = true;
                return Ok(ControlFlow::Break(()));
            }
            for ((cell_row, column), cell) in grid.cells_in((row, 0, row, u32::MAX)) {
                if !options.shows_column(column) {
                    continue;
                }
                let display = display_value(&cell.value);
                write!(
                    out,
                    "  {}: [{display}] ← ",
                    Label::cell(column, cell_row).as_str()
                )
                .map_err(sink)?;
                if cell.formula {
                    out.write_str("formula (source unavailable)")
                } else {
                    out.write_str(cell_type(&cell.value))
                }
                .map_err(sink)?;
                if is_error(&cell.value) {
                    out.write_str(" ⚠ error").map_err(sink)?;
                }
                if display.is_empty() {
                    out.write_str(" ⚠ empty").map_err(sink)?;
                }
                out.write_char('\n').map_err(sink)?;
            }
            emitted = emitted.saturating_add(1);
            Ok(ControlFlow::Continue(()))
        })?;
        if stop {
            return Ok(());
        }
    }
    Ok(())
}

fn truncation(out: &mut impl fmt::Write, emitted: u32, total: usize) -> fmt::Result {
    writeln!(
        out,
        "... (showed {emitted} rows, {total} total in sheet, use --start/--end to view more)"
    )
}

#[derive(Debug, Clone, Serialize)]
struct SheetOutline {
    name: String,
    rows: usize,
    cols: u32,
    #[serde(rename = "formulaCells")]
    formula_cells: usize,
    #[serde(rename = "errorCells")]
    error_cells: usize,
}

fn sheet_outline(sheet: &Worksheet<'_>) -> Result<SheetOutline> {
    let grid = sheet.grid()?;
    let rows = row_indices(&grid);
    let mut formula_cells = 0;
    let mut error_cells = 0;
    for ((_, _), cell) in grid.cells_in((0, 0, u32::MAX, u32::MAX)) {
        formula_cells += usize::from(cell.formula);
        if is_error(&cell.value) {
            error_cells += 1;
        }
    }
    let cols = grid
        .extents()
        .map_or(0, |(_, _, _, last_column)| last_column.saturating_add(1));
    Ok(SheetOutline {
        name: sheet.name.clone(),
        rows: rows.len(),
        cols,
        formula_cells,
        error_cells,
    })
}

fn outline(file: &Xls, out: &mut impl fmt::Write) -> Result<()> {
    for sheet in worksheets(file)? {
        let outline = sheet_outline(&sheet)?;
        writeln!(
            out,
            "├── \"{}\" ({} rows × {} cols, {} formula(s), {} error cell(s))",
            outline.name, outline.rows, outline.cols, outline.formula_cells, outline.error_cells
        )
        .map_err(sink)?;
    }
    Ok(())
}

#[derive(Debug, Clone, Default)]
struct Stats {
    sheets: usize,
    total: usize,
    empty: usize,
    formulas: usize,
    errors: usize,
    types: Vec<(&'static str, usize)>,
}

fn stats(file: &Xls) -> Result<Stats> {
    let sheets = worksheets(file)?;
    let mut stats = Stats {
        sheets: sheets.len(),
        ..Stats::default()
    };
    for sheet in &sheets {
        let grid = sheet.grid()?;
        for (_, cell) in grid.cells_in((0, 0, u32::MAX, u32::MAX)) {
            stats.total += 1;
            if matches!(&cell.value, CellValue::Empty) {
                stats.empty += 1;
            }
            stats.formulas += usize::from(cell.formula);
            if is_error(&cell.value) {
                stats.errors += 1;
            }
            let kind = cell_type(&cell.value);
            match stats.types.iter_mut().find(|(seen, _)| *seen == kind) {
                Some((_, count)) => *count += 1,
                None => stats.types.push((kind, 1)),
            }
        }
    }
    stats
        .types
        .sort_by_key(|(_, count)| std::cmp::Reverse(*count));
    Ok(stats)
}

impl Stats {
    fn write(&self, out: &mut impl fmt::Write) -> fmt::Result {
        writeln!(out, "Sheets: {}", self.sheets)?;
        writeln!(out, "Total Cells: {}", self.total)?;
        writeln!(out, "Empty Cells: {}", self.empty)?;
        writeln!(out, "Formula Cells: {}", self.formulas)?;
        writeln!(out, "Error Cells: {}", self.errors)?;
        writeln!(out)?;
        writeln!(out, "Data Type Distribution:")?;
        for (kind, count) in &self.types {
            writeln!(out, "  {kind}: {count}")?;
        }
        Ok(())
    }

    fn json(self) -> StatsJson {
        StatsJson {
            sheets: self.sheets,
            total_cells: self.total,
            empty_cells: self.empty,
            formula_cells: self.formulas,
            error_cells: self.errors,
            data_type_distribution: Distribution(self.types),
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(untagged)]
pub(crate) enum Json {
    Text(TextJson),
    Annotated(Annotated),
    Outline(OutlineJson),
    Stats(StatsJson),
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct TextJson {
    sheets: Vec<TextSheet>,
}

#[derive(Debug, Clone, Serialize)]
struct TextSheet {
    name: String,
    rows: Vec<TextRow>,
}

#[derive(Debug, Clone)]
struct Cells(Vec<(String, String)>);

impl Serialize for Cells {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(self.0.len()))?;
        for (reference, value) in &self.0 {
            map.serialize_entry(reference, value)?;
        }
        map.end()
    }
}

#[derive(Debug, Clone, Serialize)]
struct TextRow {
    row: u32,
    cells: Cells,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct Annotated {
    view: &'static str,
    content: String,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct OutlineJson {
    sheets: Vec<SheetOutline>,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct StatsJson {
    sheets: usize,
    #[serde(rename = "totalCells")]
    total_cells: usize,
    #[serde(rename = "emptyCells")]
    empty_cells: usize,
    #[serde(rename = "formulaCells")]
    formula_cells: usize,
    #[serde(rename = "errorCells")]
    error_cells: usize,
    #[serde(rename = "dataTypeDistribution")]
    data_type_distribution: Distribution,
}

#[derive(Debug, Clone)]
struct Distribution(Vec<(&'static str, usize)>);

impl Serialize for Distribution {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(self.0.len()))?;
        for (kind, count) in &self.0 {
            map.serialize_entry(kind, count)?;
        }
        map.end()
    }
}

fn text_json(file: &Xls, options: &Options) -> Result<TextJson> {
    let clip = options.range.as_ref();
    let mut sheets_json = Vec::new();
    let mut emitted = 0_u32;
    for sheet in worksheets(file)? {
        if clip.is_some_and(|range| !same_sheet(&sheet.name, &range.sheet)) {
            continue;
        }
        let grid = sheet.grid()?;
        let rows = row_indices(&grid);
        let mut selected = Vec::new();
        let mut stop = false;
        visit_rows(&grid, &rows, |row, ordinal, grid| {
            if options
                .start
                .is_some_and(|start| ordinal.saturating_add(1) < start)
            {
                return Ok(ControlFlow::Continue(()));
            }
            if options
                .end
                .is_some_and(|end| ordinal.saturating_add(1) > end)
            {
                return Ok(ControlFlow::Break(()));
            }
            if clip.is_some_and(|range| row < range.first.1 || row > range.last.1) {
                return Ok(ControlFlow::Continue(()));
            }
            if options.max_lines.is_some_and(|max| emitted >= max) {
                stop = true;
                return Ok(ControlFlow::Break(()));
            }
            let cells = grid
                .cells_in((row, 0, row, u32::MAX))
                .filter(|((_, column), cell)| {
                    !matches!(&cell.value, CellValue::Empty)
                        && options.shows_column(*column)
                        && clip
                            .is_none_or(|range| *column >= range.first.0 && *column <= range.last.0)
                })
                .map(|((cell_row, column), cell)| {
                    (
                        Label::cell(column, cell_row).as_str().to_owned(),
                        display_value(&cell.value).into_owned(),
                    )
                })
                .collect();
            selected.push(TextRow {
                row: row + 1,
                cells: Cells(cells),
            });
            emitted = emitted.saturating_add(1);
            Ok(ControlFlow::Continue(()))
        })?;
        sheets_json.push(TextSheet {
            name: sheet.name,
            rows: selected,
        });
        if stop {
            break;
        }
    }
    Ok(TextJson {
        sheets: sheets_json,
    })
}

fn outline_json(file: &Xls) -> Result<OutlineJson> {
    Ok(OutlineJson {
        sheets: worksheets(file)?
            .iter()
            .map(sheet_outline)
            .collect::<Result<Vec<_>>>()?,
    })
}

/// Keep the view's final newline/whitespace behavior aligned with XLSX/XLSB.
struct TrimEnd<'w, W> {
    out: &'w mut W,
    pending: String,
}

impl<'w, W: fmt::Write> TrimEnd<'w, W> {
    fn new(out: &'w mut W) -> Self {
        Self {
            out,
            pending: String::new(),
        }
    }
}

impl<W: fmt::Write> fmt::Write for TrimEnd<'_, W> {
    fn write_str(&mut self, text: &str) -> fmt::Result {
        let kept = text.trim_end();
        if kept.is_empty() {
            self.pending.push_str(text);
            return Ok(());
        }
        if !self.pending.is_empty() {
            self.out.write_str(&self.pending)?;
            self.pending.clear();
        }
        self.out.write_str(kept)?;
        self.pending.push_str(&text[kept.len()..]);
        Ok(())
    }
}

fn is_error(value: &CellValue) -> bool {
    match value {
        CellValue::Error(_) => true,
        CellValue::Formula {
            cached_value: Some(value),
            ..
        } => is_error(value),
        _ => false,
    }
}
