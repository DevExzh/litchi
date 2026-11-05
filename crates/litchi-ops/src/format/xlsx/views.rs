//! XLSX views (record 0010 §6): the text, annotated, outline and stats
//! modes, with OfficeCLI's content and JSON shapes
//! (`ExcelHandler.View.cs`).
//!
//! Views read each worksheet's grid through the leaf's cache and write
//! straight into the caller's sink. The text and annotated views stop at
//! `end` and `max_lines`, so a bounded view reads only the rows it shows past
//! the first touch of each sheet; the outline and stats views count every
//! record. Like OfficeCLI, trailing white space of the whole output is
//! dropped.
//!
//! | Mode | Content |
//! |---|---|
//! | text | `=== Sheet: name ===`, then `[/Sheet/row[N]] A1=v⇥B1=v` per row record; records without content are omitted; `start`/`end` count row records per sheet, `max_lines` all of them |
//! | annotated | `=== Sheet: name ===`, then `  A1: [value] ← =formula` (or the cell type), with `⚠ empty` and `⚠ formula error` |
//! | outline | `├── "name" (R rows × C cols, F formula(s), E error cell(s), T table(s))` per sheet |
//! | stats | sheet, cell, empty, formula and error counts, then the cell-type distribution |
//!
//! OfficeCLI also prints the file name in the outline and counts charts,
//! pivot tables and OLE objects; the handle knows no file name and Litchi
//! does not read those objects yet (spec gaps), so they are left out rather
//! than reported as zero.

use std::fmt;

use serde::Serialize;
use serde::ser::{SerializeMap, Serializer};

use super::Xlsx;
use super::children::row_rect;
use super::node::{Grid, Label, SheetHandle, cell_type, display_text, has_content, sheet_name};
use crate::error::{Code, Error, Result};
use crate::view::mode::{Mode, Options};

/// One worksheet of the workbook, in order: its name and handle. Its grid
/// is read when the view reaches the sheet, so a view that stops early
/// leaves the later worksheets unread.
struct Worksheet<'f> {
    position: usize,
    name: &'f str,
    handle: SheetHandle,
}

impl Worksheet<'_> {
    fn grid(&self) -> Result<Grid> {
        self.handle.grid()
    }
}

/// The worksheets (grid sheets) of the workbook, in tab order; chart and
/// dialog sheets have no rows to view. Reads the catalog only.
fn worksheets(file: &Xlsx) -> Result<Vec<Worksheet<'_>>> {
    let mut sheets = Vec::new();
    for position in 0..file.0.owner.len() {
        let Some(handle) = SheetHandle::at(file, position)? else {
            continue;
        };
        if handle.is_grid() {
            sheets.push(Worksheet {
                position,
                name: sheet_name(file, position),
                handle,
            });
        }
    }
    Ok(sheets)
}

/// Writes the `mode` view of the workbook into `out`.
///
/// # Errors
///
/// `unsupported_capability` for the modes Litchi does not render, the leaf's
/// error for a worksheet that fails to read, and `io_error` when the sink
/// refuses output.
pub(crate) fn view_to<W: fmt::Write>(
    file: &Xlsx,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    file.0.owner.finish(|| {
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

/// The `mode` view of the workbook as OfficeCLI's JSON.
///
/// # Errors
///
/// As [`view_to`].
pub(crate) fn view_json(file: &Xlsx, mode: Mode, options: &Options) -> Result<Json> {
    file.0.owner.finish(|| {
        Ok(match mode {
            Mode::Text => Json::Text(text_json(file, options)?),
            Mode::Annotated => {
                let mut content = String::new();
                view_to(file, mode, options, &mut content)?;
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
    Error::new(Code::IoError, "the view's output sink refused to write")
}

/// The line a truncated text or annotated view ends with.
fn truncation(out: &mut impl fmt::Write, emitted: u32, total: usize) -> fmt::Result {
    writeln!(
        out,
        "... (showed {emitted} rows, {total} total in sheet, use --start/--end to view more)"
    )
}

/// Whether OfficeCLI's sheet name comparison holds: ASCII letters fold, as
/// `OrdinalIgnoreCase` does for sheet names in practice.
fn same_sheet(name: &str, wanted: &str) -> bool {
    crate::path::vocabulary::Case::Caseless.matches(name, wanted)
}

fn text(file: &Xlsx, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let sheets = worksheets(file)?;
    let clip = options.range.as_ref();
    let mut emitted = 0_u32;
    for (index, sheet) in sheets.iter().enumerate() {
        if clip.is_some_and(|clip| !same_sheet(sheet.name, &clip.sheet)) {
            continue;
        }
        writeln!(out, "=== Sheet: {} ===", sheet.name).map_err(sink)?;
        let grid = sheet.grid()?;
        let total = grid.rows()?.len();
        for (row, line) in grid.rows()?.zip(1_u32..) {
            if options.start.is_some_and(|start| line < start) {
                continue;
            }
            if options.end.is_some_and(|end| line > end) {
                break;
            }
            let index_of_row = row.index().get();
            if let Some(clip) = clip {
                if index_of_row < clip.first.1 {
                    continue;
                }
                if index_of_row > clip.last.1 {
                    break;
                }
            }
            if options.max_lines.is_some_and(|max| emitted >= max) {
                truncation(out, emitted, total).map_err(sink)?;
                return Ok(());
            }
            write!(
                out,
                "[/sheet[{}]/row[{}]] ",
                sheet.position + 1,
                index_of_row + 1
            )
            .map_err(sink)?;
            let mut first = true;
            for (at, cell) in grid.cells(row_rect(index_of_row)?)? {
                let column = at.column().get();
                if !has_content(cell)
                    || !options.shows_column(column)
                    || clip.is_some_and(|clip| column < clip.first.0 || column > clip.last.0)
                {
                    continue;
                }
                if !first {
                    out.write_char('\t').map_err(sink)?;
                }
                first = false;
                write!(
                    out,
                    "{}={}",
                    Label::cell(column, at.row().get()).as_str(),
                    display_text(cell)
                )
                .map_err(sink)?;
            }
            out.write_char('\n').map_err(sink)?;
            emitted += 1;
        }
        if index + 1 < sheets.len() {
            out.write_char('\n').map_err(sink)?;
        }
    }
    Ok(())
}

fn annotated(file: &Xlsx, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let mut emitted = 0_u32;
    for sheet in worksheets(file)? {
        writeln!(out, "=== Sheet: {} ===", sheet.name).map_err(sink)?;
        let grid = sheet.grid()?;
        let total = grid.rows()?.len();
        for (row, line) in grid.rows()?.zip(1_u32..) {
            if options.start.is_some_and(|start| line < start) {
                continue;
            }
            if options.end.is_some_and(|end| line > end) {
                break;
            }
            if options.max_lines.is_some_and(|max| emitted >= max) {
                truncation(out, emitted, total).map_err(sink)?;
                return Ok(());
            }
            for (at, cell) in grid.cells(row_rect(row.index().get())?)? {
                let column = at.column().get();
                if !options.shows_column(column) {
                    continue;
                }
                let value = display_text(cell);
                let formula = match cell {
                    litchi_xlsx::Cell::Formula(formula) => Some(formula.text()),
                    _ => None,
                };
                let label = Label::cell(column, at.row().get());
                write!(out, "  {}: [{value}] \u{2190} ", label.as_str()).map_err(sink)?;
                match formula {
                    Some(formula) => write!(out, "={formula}"),
                    None => out.write_str(cell_type(cell)),
                }
                .map_err(sink)?;
                if value.is_empty() && formula.is_none() {
                    out.write_str(" \u{26a0} empty").map_err(sink)?;
                } else if formula.is_some() && is_error_cell(cell, value) {
                    out.write_str(" \u{26a0} formula error").map_err(sink)?;
                }
                out.write_char('\n').map_err(sink)?;
            }
            emitted += 1;
        }
    }
    Ok(())
}

/// Whether a record is an error cell, by OfficeCLI's `IsExcelErrorValue`: an
/// error value, or a display value shaped like one (`#` and an upper-case
/// letter).
pub(super) fn is_error_cell(cell: &litchi_xlsx::Cell, display: &str) -> bool {
    let bytes = display.as_bytes();
    let shaped = bytes.first() == Some(&b'#') && bytes.get(1).is_some_and(u8::is_ascii_uppercase);
    shaped || super::node::is_error(cell)
}

/// One sheet's outline counts.
#[derive(Debug, Clone, Serialize)]
struct SheetOutline {
    name: String,
    rows: usize,
    cols: u32,
    formulas: usize,
    #[serde(rename = "errorCells")]
    error_cells: usize,
    tables: usize,
    /// The sheet kind of a chart, dialog or macro sheet, which has no grid.
    #[serde(skip)]
    kind: Option<&'static str>,
}

fn sheet_outline(file: &Xlsx, position: usize, handle: &SheetHandle) -> Result<SheetOutline> {
    let name = sheet_name(file, position).to_owned();
    if !handle.is_grid() {
        let kind = match handle.kind() {
            litchi_xlsx::WorksheetKind::Chart => "chart sheet",
            litchi_xlsx::WorksheetKind::Dialog => "dialog sheet",
            litchi_xlsx::WorksheetKind::Macro => "macro sheet",
            _ => "sheet without cells",
        };
        return Ok(SheetOutline {
            name,
            rows: 0,
            cols: 0,
            formulas: 0,
            error_cells: 0,
            tables: 0,
            kind: Some(kind),
        });
    }
    let grid = handle.grid()?;
    let rows = grid.rows()?.len();
    let mut formulas = 0;
    let mut error_cells = 0;
    let mut widest = 0_u32;
    if let Some(extent) = grid.extents()?.stored() {
        let mut row = None;
        let mut in_row = 0_u32;
        for (at, cell) in grid.cells(extent)? {
            if row != Some(at.row().get()) {
                row = Some(at.row().get());
                in_row = 0;
            }
            in_row += 1;
            widest = widest.max(in_row);
            if matches!(cell, litchi_xlsx::Cell::Formula(_)) {
                formulas += 1;
            }
            if is_error_cell(cell, display_text(cell)) {
                error_cells += 1;
            }
        }
    }
    // OfficeCLI reads the width from the declared dimension, else from the
    // widest row's cell count.
    let cols = grid
        .extents()?
        .declared()
        .map_or(widest, |declared| declared.end().1);
    Ok(SheetOutline {
        name,
        rows,
        cols,
        formulas,
        error_cells,
        tables: handle.tables()?.len(),
        kind: None,
    })
}

fn sheet_outlines(file: &Xlsx) -> Result<Vec<SheetOutline>> {
    let mut sheets = Vec::new();
    for position in 0..file.0.owner.len() {
        if let Some(handle) = SheetHandle::at(file, position)? {
            sheets.push(sheet_outline(file, position, &handle)?);
        }
    }
    Ok(sheets)
}

fn outline(file: &Xlsx, out: &mut impl fmt::Write) -> Result<()> {
    for sheet in sheet_outlines(file)? {
        write!(out, "\u{251c}\u{2500}\u{2500} \"{}\" (", sheet.name).map_err(sink)?;
        if let Some(kind) = sheet.kind {
            writeln!(out, "{kind})").map_err(sink)?;
            continue;
        }
        write!(out, "{} rows \u{00d7} {} cols", sheet.rows, sheet.cols).map_err(sink)?;
        if sheet.formulas > 0 {
            write!(out, ", {} formula(s)", sheet.formulas).map_err(sink)?;
        }
        if sheet.error_cells > 0 {
            write!(out, ", {} error cell(s)", sheet.error_cells).map_err(sink)?;
        }
        if sheet.tables > 0 {
            write!(out, ", {} table(s)", sheet.tables).map_err(sink)?;
        }
        out.write_str(")\n").map_err(sink)?;
    }
    Ok(())
}

/// The workbook's cell counts.
#[derive(Debug, Clone, Default)]
struct Stats {
    sheets: usize,
    total: usize,
    empty: usize,
    formulas: usize,
    errors: usize,
    /// Cell types in first-seen order, with their counts.
    types: Vec<(&'static str, usize)>,
}

fn stats(file: &Xlsx) -> Result<Stats> {
    let mut stats = Stats::default();
    for sheet in worksheets(file)? {
        stats.sheets += 1;
        let grid = sheet.grid()?;
        let Some(extent) = grid.extents()?.stored() else {
            continue;
        };
        for (_, cell) in grid.cells(extent)? {
            stats.total += 1;
            let display = display_text(cell);
            if display.is_empty() {
                stats.empty += 1;
            }
            if matches!(cell, litchi_xlsx::Cell::Formula(_)) {
                stats.formulas += 1;
            }
            if is_error_cell(cell, display) {
                stats.errors += 1;
            }
            let kind = cell_type(cell);
            match stats.types.iter_mut().find(|(seen, _)| *seen == kind) {
                Some((_, count)) => *count += 1,
                None => stats.types.push((kind, 1)),
            }
        }
    }
    // Descending by count; ties keep first-seen order, as OfficeCLI's stable
    // `OrderByDescending` does.
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

/// A view as OfficeCLI's JSON (`view --json`), without the CLI envelope.
#[derive(Debug, Clone, Serialize)]
#[serde(untagged)]
pub(crate) enum Json {
    /// `{"sheets":[{"name":..,"rows":[{"row":N,"cells":{"A1":"v"}}]}]}`.
    Text(TextJson),
    /// `{"view":"annotated","content":"..."}`.
    Annotated(Annotated),
    /// `{"sheets":[{"name","rows","cols","formulas","errorCells","tables"}]}`.
    Outline(OutlineJson),
    /// `{"sheets","totalCells","emptyCells","formulaCells","errorCells","dataTypeDistribution"}`.
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

#[derive(Debug, Clone, Serialize)]
struct TextRow {
    row: u32,
    cells: Cells,
}

/// The cells of one row as a JSON object, in column order.
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

/// Cell types and their counts as a JSON object, in order.
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

fn text_json(file: &Xlsx, options: &Options) -> Result<TextJson> {
    let clip = options.range.as_ref();
    let mut sheets = Vec::new();
    let mut emitted = 0_u32;
    'sheets: for sheet in worksheets(file)? {
        if clip.is_some_and(|clip| !same_sheet(sheet.name, &clip.sheet)) {
            continue;
        }
        let grid = sheet.grid()?;
        let mut rows = Vec::new();
        for (row, line) in grid.rows()?.zip(1_u32..) {
            if options.start.is_some_and(|start| line < start) {
                continue;
            }
            if options.end.is_some_and(|end| line > end) {
                break;
            }
            let index = row.index().get();
            if let Some(clip) = clip {
                if index < clip.first.1 {
                    continue;
                }
                if index > clip.last.1 {
                    break;
                }
            }
            if options.max_lines.is_some_and(|max| emitted >= max) {
                sheets.push(TextSheet {
                    name: sheet.name.to_owned(),
                    rows,
                });
                break 'sheets;
            }
            let mut cells = Vec::new();
            for (at, cell) in grid.cells(row_rect(index)?)? {
                let column = at.column().get();
                if !has_content(cell)
                    || !options.shows_column(column)
                    || clip.is_some_and(|clip| column < clip.first.0 || column > clip.last.0)
                {
                    continue;
                }
                cells.push((
                    Label::cell(column, at.row().get()).as_str().to_owned(),
                    display_text(cell).to_owned(),
                ));
            }
            rows.push(TextRow {
                row: index + 1,
                cells: Cells(cells),
            });
            emitted += 1;
        }
        sheets.push(TextSheet {
            name: sheet.name.to_owned(),
            rows,
        });
    }
    Ok(TextJson { sheets })
}

fn outline_json(file: &Xlsx) -> Result<OutlineJson> {
    Ok(OutlineJson {
        sheets: sheet_outlines(file)?,
    })
}

/// A sink wrapper that drops the output's trailing white space, as
/// OfficeCLI's views do (`TrimEnd`): white space is held back until text
/// follows it.
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
