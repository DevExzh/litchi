//! XLSX template merge (record 0274): placeholders in every text cell of
//! every worksheet, staged in one native workbook edit. Formulas are never
//! touched; a cell keeps its style.

use litchi_core::FileFormat;
use litchi_xlsx::workbook::WorksheetKind;
use litchi_xlsx::{self as xlsx};

use super::super::Xlsx;
use super::{Commit, native_error};
use crate::error::{Code, Error, Result};
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
        let source = file.editor()?;
        let workbook = source.workbook();
        let mut planner = Planner::new(record, options, limits);
        // (worksheet position, address, new text)
        let mut cells: Vec<(usize, xlsx::Address, String)> = Vec::new();
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
                    cells.push((sheet.position(), address, content));
                }
            }
        }
        let summary = planner.finish();
        if cells.is_empty() {
            return Ok((None, summary));
        }
        let mut edit = source.edit().map_err(native_error)?;
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
                worksheet
                    .set(*address, xlsx::cell::Value::text(text.as_str()))
                    .map_err(native_error)?;
                index += 1;
            }
        }
        let native = edit.commit().map_err(native_error)?;
        Ok((Some(Commit::from_native(file, source, native)?), summary))
    })();
    let freshness = file.check_source();
    freshness.and(result)
}
