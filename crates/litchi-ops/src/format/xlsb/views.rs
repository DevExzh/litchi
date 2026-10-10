//! XLSB workbook views (record 0010 section 6).
//!
//! The projections use the XLSB node bridge directly. A view opens one
//! semantic sheet at a time, walks its sparse typed grid, and leaves later
//! sheets unopened when a row limit ends the operation. In particular, this
//! module does not route through the compatibility Worksheet trait: that
//! route copies a complete worksheet and cannot preserve source errors or
//! the owner-scoped source fence.

use std::fmt;
use std::ops::ControlFlow;

use litchi_core::Position;
use litchi_core::sheet::{Cell as CoreCell, CellValue};
use serde::Serialize;
use serde::ser::{SerializeMap, Serializer};

use super::Xlsb;
use super::children::row_rect;
use super::node::{Grid, Sheet, cell_type, is_error};
use super::owner::SheetHandle;
use crate::error::{Code, Error, Result};
use crate::format::workbook::Label;
use crate::path::vocabulary::Case;
use crate::view::mode::{Mode, Options};

/// A worksheet catalog entry. The name and handle are retained, but opening
/// the handle's grid is deferred until the entry reaches the view cursor.
struct Worksheet {
    position: usize,
    name: String,
    sheet: Sheet,
}

impl Worksheet {
    fn grid(&self) -> Result<Grid> {
        self.sheet.grid().cloned()
    }
}

/// Read the complete semantic tab catalog and retain only tabs with a grid.
///
/// SheetHandle::at reads tab metadata only. In particular, it does not
/// materialize any worksheet body, so a bounded view does not touch later
/// worksheets.
fn worksheets(file: &Xlsb) -> Result<Vec<Worksheet>> {
    let count = file.tab_count()?;
    let mut sheets = Vec::new();
    sheets.try_reserve_exact(count).map_err(|_| {
        Error::new(
            Code::InternalError,
            "could not reserve the XLSB view catalog",
        )
    })?;
    for position in 0..count {
        let Some(handle) = SheetHandle::at(file, position)? else {
            continue;
        };
        if handle.is_grid() {
            let name = handle.name().to_owned();
            sheets.push(Worksheet {
                position,
                name,
                sheet: Sheet::new(Position::new(position), handle),
            });
        }
    }
    Ok(sheets)
}

/// Fence a complete view operation with the owner-scoped source checks.
fn finish<T>(file: &Xlsb, operation: impl FnOnce() -> Result<T>) -> Result<T> {
    file.finish(operation)
}

/// Writes one supported workbook view to a caller-owned sink.
pub(crate) fn view_to<W: fmt::Write>(
    file: &Xlsb,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    finish(file, || {
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

/// Builds the typed JSON form of one supported workbook view.
pub(crate) fn view_json(file: &Xlsb, mode: Mode, options: &Options) -> Result<Json> {
    finish(file, || {
        Ok(match mode {
            Mode::Text => Json::Text(text_json(file, options)?),
            Mode::Annotated => {
                let mut content = String::new();
                {
                    let mut trimmed = TrimEnd::new(&mut content);
                    annotated(file, options, &mut trimmed)?;
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
    Error::new(Code::IoError, "the view's output sink refused to write")
}

fn truncation(out: &mut impl fmt::Write, emitted: u32, total: usize) -> fmt::Result {
    writeln!(
        out,
        "... (showed {emitted} rows, {total} total in sheet, use --start/--end to view more)"
    )
}

fn same_sheet(name: &str, wanted: &str) -> bool {
    Case::Caseless.matches(name, wanted)
}

/// Visit logical row records in order. The row index list is produced by the
/// typed grid and only the sparse cells of the current row are opened.
fn visit_rows(
    grid: &Grid,
    rows: &[u32],
    mut visit: impl FnMut(u32, u32, &Grid) -> Result<ControlFlow<()>>,
) -> Result<()> {
    for (ordinal, &row) in rows.iter().enumerate() {
        let ordinal = u32::try_from(ordinal).unwrap_or(u32::MAX);
        if matches!(visit(row, ordinal, grid)?, ControlFlow::Break(())) {
            break;
        }
    }
    Ok(())
}

/// Return row records from the authoritative grid metadata. The cached-only
/// source fallback deliberately has no row metadata, so derive only the rows
/// represented by its scalar cache and keep formula metadata unavailable.
fn row_indices(grid: &Grid) -> Result<Vec<u32>> {
    if !grid.cached_only() {
        return grid.row_indices();
    }
    let Some(extent) = grid.extents()? else {
        return Ok(Vec::new());
    };
    let mut rows = Vec::new();
    for (at, _) in grid.cells(extent)? {
        if rows.last().copied() != Some(at.0) {
            rows.push(at.0);
        }
    }
    rows.sort_unstable();
    rows.dedup();
    Ok(rows)
}

fn text(file: &Xlsb, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let sheets = worksheets(file)?;
    let clip = options.range.as_ref();
    let mut emitted = 0_u32;

    for (index, sheet) in sheets.iter().enumerate() {
        if clip.is_some_and(|clip| !same_sheet(&sheet.name, &clip.sheet)) {
            continue;
        }
        writeln!(out, "=== Sheet: {} ===", sheet.name).map_err(sink)?;
        let grid = sheet.grid()?;
        let rows = row_indices(&grid)?;
        let total = rows.len();
        let mut stop_all = false;
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
            if clip.is_some_and(|clip| row < clip.first.1 || row > clip.last.1) {
                return Ok(ControlFlow::Continue(()));
            }
            if options.max_lines.is_some_and(|max| emitted >= max) {
                truncation(out, emitted, total).map_err(sink)?;
                stop_all = true;
                return Ok(ControlFlow::Break(()));
            }
            write!(out, "[/sheet[{}]/row[{}]] ", sheet.position + 1, row + 1).map_err(sink)?;
            let mut first = true;
            for (at, cell) in grid.cells(row_rect(row)?)? {
                let (cell_row, column) = at;
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
                    Label::cell(column, cell_row).as_str(),
                    display_string(cell)
                )
                .map_err(sink)?;
            }
            out.write_char('\n').map_err(sink)?;
            emitted = emitted.saturating_add(1);
            Ok(ControlFlow::Continue(()))
        })?;
        if stop_all {
            return Ok(());
        }
        if index + 1 < sheets.len() {
            out.write_char('\n').map_err(sink)?;
        }
    }
    Ok(())
}

fn annotated(file: &Xlsb, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let mut emitted = 0_u32;
    for sheet in worksheets(file)? {
        writeln!(out, "=== Sheet: {} ===", sheet.name).map_err(sink)?;
        let grid = sheet.grid()?;
        let rows = row_indices(&grid)?;
        let total = rows.len();
        let cached_only = grid.cached_only();
        let mut stop_all = false;
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
                truncation(out, emitted, total).map_err(sink)?;
                stop_all = true;
                return Ok(ControlFlow::Break(()));
            }
            for (at, cell) in grid.cells(row_rect(row)?)? {
                let (cell_row, column) = at;
                if !options.shows_column(column) {
                    continue;
                }
                let display = display_string(cell);
                let formula = (!cached_only).then(|| formula_text(cell)).flatten();
                write!(
                    out,
                    "  {}: [{display}] \u{2190} ",
                    Label::cell(column, cell_row).as_str()
                )
                .map_err(sink)?;
                if cached_only {
                    out.write_str("cached value (formula unavailable)")
                } else {
                    match formula {
                        Some(formula) => write!(out, "={formula}"),
                        None => out.write_str(cell_type(cell.value())),
                    }
                }
                .map_err(sink)?;
                if display.is_empty() && formula.is_none() {
                    out.write_str(" \u{26a0} empty").map_err(sink)?;
                } else if !cached_only && formula.is_some() && is_error_cell(cell, &display) {
                    out.write_str(" \u{26a0} formula error").map_err(sink)?;
                }
                out.write_char('\n').map_err(sink)?;
            }
            emitted = emitted.saturating_add(1);
            Ok(ControlFlow::Continue(()))
        })?;
        if stop_all {
            return Ok(());
        }
    }
    Ok(())
}

fn has_content(cell: &litchi_xlsb::Cell) -> bool {
    !matches!(cell.value(), CellValue::Empty)
}

fn formula_text(cell: &litchi_xlsb::Cell) -> Option<&str> {
    match cell.value() {
        CellValue::Formula { formula, .. } => Some(formula),
        _ => None,
    }
}

/// Render the cached formula scalar when it exists. A formula without a
/// cached scalar has no display text, matching the XLSX view; annotated output
/// still exposes its source formula separately.
fn display_string(cell: &litchi_xlsb::Cell) -> String {
    fn render_value(value: &CellValue) -> String {
        match value {
            CellValue::Empty => String::new(),
            CellValue::Bool(value) => {
                if *value {
                    "TRUE".to_owned()
                } else {
                    "FALSE".to_owned()
                }
            },
            CellValue::Int(value) => value.to_string(),
            CellValue::Float(value) | CellValue::DateTime(value) => value.to_string(),
            CellValue::String(value) | CellValue::Error(value) => (*value).to_owned(),
            CellValue::Formula { cached_value, .. } => match cached_value.as_deref() {
                Some(CellValue::Empty) | None => String::new(),
                Some(cached) => render_value(cached),
            },
        }
    }
    render_value(cell.value())
}

fn is_error_cell(cell: &litchi_xlsb::Cell, display: &str) -> bool {
    let bytes = display.as_bytes();
    is_error(cell)
        || (bytes.first() == Some(&b'#') && bytes.get(1).is_some_and(u8::is_ascii_uppercase))
}

#[derive(Debug, Clone, Serialize)]
struct SheetOutline {
    name: String,
    rows: usize,
    cols: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    formulas: Option<usize>,
    #[serde(rename = "errorCells")]
    error_cells: usize,
    tables: usize,
}

fn sheet_outline(sheet: &Worksheet) -> Result<SheetOutline> {
    let grid = sheet.grid()?;
    let rows = row_indices(&grid)?;
    let mut formulas = (!grid.cached_only()).then_some(0);
    let mut error_cells = 0;
    for row in &rows {
        for (_, cell) in grid.cells(row_rect(*row)?)? {
            if let Some(formulas) = formulas.as_mut() {
                if matches!(cell.value(), CellValue::Formula { .. }) {
                    *formulas += 1;
                }
            }
            let display = display_string(cell);
            if is_error_cell(cell, &display) {
                error_cells += 1;
            }
        }
    }
    let cols = grid.extents()?.map_or(0, |extent| extent.end().1);
    Ok(SheetOutline {
        name: sheet.name.clone(),
        rows: rows.len(),
        cols,
        formulas,
        error_cells,
        tables: sheet.sheet.handle.tables()?.len(),
    })
}

fn sheet_outlines(file: &Xlsb) -> Result<Vec<SheetOutline>> {
    let catalog = worksheets(file)?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(catalog.len())
        .map_err(|_| Error::new(Code::InternalError, "could not reserve XLSB outline"))?;
    for sheet in &catalog {
        result.push(sheet_outline(sheet)?);
    }
    Ok(result)
}

fn outline(file: &Xlsb, out: &mut impl fmt::Write) -> Result<()> {
    for sheet in sheet_outlines(file)? {
        write!(out, "\u{251c}\u{2500}\u{2500} \"{}\" (", sheet.name).map_err(sink)?;
        write!(out, "{} rows \u{00d7} {} cols", sheet.rows, sheet.cols).map_err(sink)?;
        if sheet.formulas.is_some_and(|count| count > 0) {
            write!(out, ", {} formula(s)", sheet.formulas.unwrap_or_default()).map_err(sink)?;
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

#[derive(Debug, Clone, Default)]
struct Stats {
    sheets: usize,
    total: usize,
    empty: usize,
    formulas: Option<usize>,
    errors: usize,
    types: Vec<(&'static str, usize)>,
}

fn stats(file: &Xlsb) -> Result<Stats> {
    let catalog = worksheets(file)?;
    let mut stats = Stats {
        formulas: Some(0),
        ..Stats::default()
    };
    for sheet in &catalog {
        stats.sheets += 1;
        let grid = sheet.grid()?;
        let rows = row_indices(&grid)?;
        if grid.cached_only() {
            stats.formulas = None;
        }
        for row in &rows {
            for (_, cell) in grid.cells(row_rect(*row)?)? {
                stats.total += 1;
                let display = display_string(cell);
                if display.is_empty() {
                    stats.empty += 1;
                }
                if matches!(cell.value(), CellValue::Formula { .. })
                    && let Some(formulas) = stats.formulas.as_mut()
                {
                    *formulas += 1;
                }
                if is_error_cell(cell, &display) {
                    stats.errors += 1;
                }
                let kind = cell_type(cell.value());
                match stats.types.iter_mut().find(|(seen, _)| *seen == kind) {
                    Some((_, count)) => *count += 1,
                    None => stats.types.push((kind, 1)),
                }
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
        match self.formulas {
            Some(formulas) => writeln!(out, "Formula Cells: {formulas}")?,
            None => writeln!(out, "Formula Cells: unavailable")?,
        }
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

/// A typed JSON view for dispatch through crate::view::Json.
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

#[derive(Debug, Clone, Serialize)]
struct TextRow {
    row: u32,
    cells: Cells,
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
    #[serde(rename = "formulaCells", skip_serializing_if = "Option::is_none")]
    formula_cells: Option<usize>,
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

fn text_json(file: &Xlsb, options: &Options) -> Result<TextJson> {
    let clip = options.range.as_ref();
    let catalog = worksheets(file)?;
    let mut sheets = Vec::new();
    let mut emitted = 0_u32;

    'sheets: for sheet in &catalog {
        if clip.is_some_and(|clip| !same_sheet(&sheet.name, &clip.sheet)) {
            continue;
        }
        let grid = sheet.grid()?;
        let rows = row_indices(&grid)?;
        let mut selected = Vec::new();
        let mut stop_all = false;
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
            if clip.is_some_and(|clip| row < clip.first.1 || row > clip.last.1) {
                return Ok(ControlFlow::Continue(()));
            }
            if options.max_lines.is_some_and(|max| emitted >= max) {
                stop_all = true;
                return Ok(ControlFlow::Break(()));
            }
            let mut cells = Vec::new();
            for (at, cell) in grid.cells(row_rect(row)?)? {
                let (cell_row, column) = at;
                if !has_content(cell)
                    || !options.shows_column(column)
                    || clip.is_some_and(|clip| column < clip.first.0 || column > clip.last.0)
                {
                    continue;
                }
                cells.push((
                    Label::cell(column, cell_row).as_str().to_owned(),
                    display_string(cell),
                ));
            }
            selected.push(TextRow {
                row: row + 1,
                cells: Cells(cells),
            });
            emitted = emitted.saturating_add(1);
            Ok(ControlFlow::Continue(()))
        })?;
        sheets.push(TextSheet {
            name: sheet.name.clone(),
            rows: selected,
        });
        if stop_all {
            break 'sheets;
        }
    }
    Ok(TextJson { sheets })
}

fn outline_json(file: &Xlsb) -> Result<OutlineJson> {
    Ok(OutlineJson {
        sheets: sheet_outlines(file)?,
    })
}

/// A sink wrapper that drops trailing whitespace, matching the XLSX view.
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
