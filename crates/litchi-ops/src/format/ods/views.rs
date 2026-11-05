//! OfficeCLI-shaped textual and JSON views for ODS workbooks (record
//! 0277), with the layout of the other workbook adapters: `text` lists each
//! content row's cells, `annotated` each content cell with its type,
//! `outline` each sheet's size and formula count, `stats` the cell type
//! distribution.

use std::fmt;
use std::ops::ControlFlow;

use litchi_ods::{CellValue, Sheet};
use serde::Serialize;
use serde::ser::{SerializeMap, Serializer};

use super::Ods;
use super::index::SheetIndex;
use super::node::{cell_type, sheet_index};
use crate::error::{Code, Error, Result};
use crate::format::workbook::Label;
use crate::path::vocabulary::Case;
use crate::view::mode::{Mode, Options};

struct Worksheet<'f> {
    position: usize,
    sheet: &'f Sheet,
    index: &'f SheetIndex,
}

pub(crate) fn view_to<W: fmt::Write>(
    file: &Ods,
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

pub(crate) fn view_json(file: &Ods, mode: Mode, options: &Options) -> Result<Json> {
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
    Error::new(Code::IoError, "the ODS view output sink refused to write")
        .with_format(litchi_core::FileFormat::Ods)
}

fn worksheets(file: &Ods) -> Result<Vec<Worksheet<'_>>> {
    let sheets = file.worksheets()?;
    let mut out = Vec::new();
    out.try_reserve_exact(sheets.len()).map_err(|_| {
        Error::new(
            Code::InternalError,
            "could not reserve the ODS view catalog",
        )
        .with_format(litchi_core::FileFormat::Ods)
    })?;
    for (position, sheet) in sheets.iter().enumerate() {
        out.push(Worksheet {
            position,
            sheet,
            index: sheet_index(file, u32::try_from(position).unwrap_or(u32::MAX))?,
        });
    }
    Ok(out)
}

fn same_sheet(name: &str, wanted: &str) -> bool {
    Case::Caseless.matches(name, wanted)
}

fn visit_rows(
    index: &SheetIndex,
    mut visit: impl FnMut(u32, u32) -> Result<ControlFlow<()>>,
) -> Result<()> {
    for (ordinal, (row, _)) in index.row_entries().iter().enumerate() {
        let ordinal = u32::try_from(ordinal).unwrap_or(u32::MAX);
        if matches!(visit(*row, ordinal)?, ControlFlow::Break(())) {
            break;
        }
    }
    Ok(())
}

fn text(file: &Ods, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let sheets = worksheets(file)?;
    let clip = options.range.as_ref();
    let mut emitted = 0_u32;
    for (sheet_index, sheet) in sheets.iter().enumerate() {
        if clip.is_some_and(|range| !same_sheet(&sheet.sheet.name, &range.sheet)) {
            continue;
        }
        writeln!(out, "=== Sheet: {} ===", sheet.sheet.name).map_err(sink)?;
        let total = sheet.index.row_count();
        let mut stop = false;
        visit_rows(sheet.index, |row, ordinal| {
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
            for entry in sheet.index.row_cells(row) {
                let Some(cell) = entry.cell(sheet.sheet) else {
                    continue;
                };
                if cell.text.is_empty()
                    || !options.shows_column(entry.column)
                    || clip.is_some_and(|range| {
                        entry.column < range.first.0 || entry.column > range.last.0
                    })
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
                    Label::cell(entry.column, entry.row).as_str(),
                    cell.text
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

fn annotated(file: &Ods, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let mut emitted = 0_u32;
    for sheet in worksheets(file)? {
        writeln!(out, "=== Sheet: {} ===", sheet.sheet.name).map_err(sink)?;
        let total = sheet.index.row_count();
        let mut stop = false;
        visit_rows(sheet.index, |row, ordinal| {
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
                stop = true;
                return Ok(ControlFlow::Break(()));
            }
            for entry in sheet.index.row_cells(row) {
                if !options.shows_column(entry.column) {
                    continue;
                }
                let Some(cell) = entry.cell(sheet.sheet) else {
                    continue;
                };
                write!(
                    out,
                    "  {}: [{}] ← ",
                    Label::cell(entry.column, entry.row).as_str(),
                    cell.text
                )
                .map_err(sink)?;
                match &cell.formula {
                    Some(formula) => write!(out, "formula {formula}"),
                    None => out.write_str(cell_type(&cell.value).unwrap_or("Empty")),
                }
                .map_err(sink)?;
                if cell.text.is_empty() {
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

fn sheet_outline(sheet: &Worksheet<'_>) -> SheetOutline {
    let formula_cells = sheet
        .index
        .cells()
        .iter()
        .filter(|entry| {
            entry
                .cell(sheet.sheet)
                .is_some_and(|cell| cell.formula.is_some())
        })
        .count();
    SheetOutline {
        name: sheet.sheet.name.clone(),
        rows: sheet.index.row_count(),
        cols: sheet
            .index
            .last_column()
            .map_or(0, |column| column.saturating_add(1)),
        formula_cells,
        error_cells: 0,
    }
}

fn outline(file: &Ods, out: &mut impl fmt::Write) -> Result<()> {
    for sheet in worksheets(file)? {
        let outline = sheet_outline(&sheet);
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

fn stats(file: &Ods) -> Result<Stats> {
    let sheets = worksheets(file)?;
    let mut stats = Stats {
        sheets: sheets.len(),
        ..Stats::default()
    };
    for sheet in &sheets {
        for entry in sheet.index.cells() {
            let Some(cell) = entry.cell(sheet.sheet) else {
                continue;
            };
            stats.total += 1;
            if matches!(cell.value, CellValue::Empty) {
                stats.empty += 1;
            }
            stats.formulas += usize::from(cell.formula.is_some());
            let kind = cell_type(&cell.value).unwrap_or("Empty");
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

fn text_json(file: &Ods, options: &Options) -> Result<TextJson> {
    let clip = options.range.as_ref();
    let mut sheets_json = Vec::new();
    let mut emitted = 0_u32;
    for sheet in worksheets(file)? {
        if clip.is_some_and(|range| !same_sheet(&sheet.sheet.name, &range.sheet)) {
            continue;
        }
        let mut selected = Vec::new();
        let mut stop = false;
        visit_rows(sheet.index, |row, ordinal| {
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
            let cells = sheet
                .index
                .row_cells(row)
                .iter()
                .filter_map(|entry| entry.cell(sheet.sheet).map(|cell| (entry, cell)))
                .filter(|(entry, cell)| {
                    !cell.text.is_empty()
                        && options.shows_column(entry.column)
                        && clip.is_none_or(|range| {
                            entry.column >= range.first.0 && entry.column <= range.last.0
                        })
                })
                .map(|(entry, cell)| {
                    (
                        Label::cell(entry.column, entry.row).as_str().to_owned(),
                        cell.text.clone(),
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
            name: sheet.sheet.name.clone(),
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

fn outline_json(file: &Ods) -> Result<OutlineJson> {
    Ok(OutlineJson {
        sheets: worksheets(file)?.iter().map(sheet_outline).collect(),
    })
}

/// Keep the view's final newline/whitespace behavior aligned with the other
/// workbook adapters.
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
