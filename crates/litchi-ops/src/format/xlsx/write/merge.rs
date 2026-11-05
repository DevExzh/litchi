//! XLSX template merge (record 0274): placeholders in every text cell of
//! every worksheet, staged in one native workbook edit. Formulas are never
//! touched; a cell keeps its style. A cell showing a formatted shared
//! string is filled run by run and keeps its runs' formatting as its own
//! inline rich string (record 0295).

use litchi_core::FileFormat;
use litchi_xlsx::workbook::WorksheetKind;
use litchi_xlsx::{self as xlsx};

use super::super::Xlsx;
use super::find::CellText;
use super::{Commit, native_error};
use crate::error::{Code, Error, Result};
use crate::find::surface::{self, Run};
use crate::import::MAX_CELL_CHARS;
use crate::merge::{Options, Planner, Record, Summary, normalize_breaks};
use crate::transaction::SemanticLimits;

/// Fill `record` into every text cell. `None` means nothing matched.
pub(crate) fn merge(
    file: &Xlsx,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    file.check_source()?;
    let result = (|| {
        let source = file.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let workbook = source.workbook();
        let mut planner = Planner::new(record, options, limits);
        // (worksheet position, address, new text)
        let mut cells: Vec<(usize, xlsx::Address, CellText)> = Vec::new();
        for sheet in workbook.sheets() {
            if sheet.kind() != WorksheetKind::Worksheet {
                continue;
            }
            let Some(extent) = sheet.stored_extent().map_err(native_error)? else {
                continue;
            };
            for (address, cell) in sheet.cells(extent).map_err(native_error)? {
                let xlsx::cell::Cell::Value(xlsx::cell::Value::Text(text)) = cell else {
                    continue;
                };
                let text = text.as_str();
                if text.contains("{{")
                    && sheet.has_rich_text(address).map_err(native_error)?
                    && let Some(mut rich) = sheet.rich_text(address).map_err(native_error)?
                {
                    // Record 0295: fill the runs, each keeping its
                    // properties, as the cell's own inline rich string.
                    let runs: Vec<Run<'_>> = rich
                        .runs()
                        .iter()
                        .map(|run| Run::new(run.text(), 0))
                        .collect();
                    let changes =
                        surface::merge(&runs, &mut planner, true, FileFormat::Xlsx, limits)?;
                    drop(runs);
                    if !changes.is_empty() {
                        for change in changes {
                            rich.set_run_text(change.run, change.text)
                                .map_err(native_error)?;
                        }
                        if rich.text().chars().count() > MAX_CELL_CHARS {
                            return Err(Error::new(
                                Code::InvalidValue,
                                "a merged cell would be longer than 32,767 characters",
                            )
                            .with_format(FileFormat::Xlsx));
                        }
                        cells.push((sheet.position(), address, CellText::Rich(rich)));
                    }
                    continue;
                }
                let mut planned: Vec<(std::ops::Range<usize>, &str)> = Vec::new();
                planner.text(
                    text,
                    |_| true,
                    &[],
                    |range, value| planned.push((range, value)),
                )?;
                if planned.is_empty() {
                    continue;
                }
                let mut content = String::with_capacity(text.len());
                let mut cursor = 0;
                for (range, value) in &planned {
                    content.push_str(&text[cursor..range.start]);
                    content.push_str(&normalize_breaks(value));
                    cursor = range.end;
                }
                content.push_str(&text[cursor..]);
                if content.chars().count() > MAX_CELL_CHARS {
                    return Err(Error::new(
                        Code::InvalidValue,
                        "a merged cell would be longer than 32,767 characters",
                    )
                    .with_format(FileFormat::Xlsx));
                }
                if content != text {
                    cells.push((sheet.position(), address, CellText::Plain(content)));
                }
            }
        }
        let summary = planner.finish();
        if cells.is_empty() {
            return Ok((None, summary));
        }
        let mut edit = super::edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        let mut index = 0;
        while index < cells.len() {
            let position = cells[index].0;
            let mut worksheet = edit.sheet(position).map_err(native_error)?.ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSX worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsx)
            })?;
            while index < cells.len() && cells[index].0 == position {
                let (_, address, text) = &cells[index];
                match text {
                    CellText::Plain(text) => worksheet
                        .set(*address, xlsx::cell::Value::text(text.as_str()))
                        .map_err(native_error)?,
                    CellText::Rich(rich) => worksheet
                        .set_rich_text(*address, rich.clone())
                        .map_err(native_error)?,
                };
                index += 1;
            }
        }
        let native = edit.commit().map_err(native_error)?;
        Ok((
            Some(Commit::from_native_with_limit(
                file,
                source,
                native,
                limits.max_snapshot_bytes(),
            )?),
            summary,
        ))
    })();
    let freshness = file.check_source();
    freshness.and(result)
}
