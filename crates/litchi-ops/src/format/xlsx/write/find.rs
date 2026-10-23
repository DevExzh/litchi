//! XLSX find/replace (record 0273): `set` with `find=` and `replace=`.
//!
//! As OfficeCLI's `ExcelHandler.FindAndReplace`, the workbook root searches
//! every worksheet and a sheet its own cells; Litchi also accepts a cell, a
//! range, a row or a column. Each stored text value is one search group.
//! Formula cells and non-text values are not searched. A cell showing a
//! formatted shared string is rewritten run by run as its own inline rich
//! string, each run keeping its properties (record 0295); an inline rich
//! string or an item with phonetic runs is refused when it would change. A
//! match spanning runs collapses into the run where it starts. Every
//! changed cell is staged in one native workbook edit.

use std::collections::HashSet;

use litchi_core::FileFormat;
use litchi_xlsx::{self as xlsx, Address, Rect};

use super::super::node::Item;
use super::super::{Nav, Xlsx};
use super::{Commit, native_error, validate_xml_text};
use crate::error::{Code, Error, Result};
use crate::find::{Piece, Scratch, Spec, plan_with};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// The XLSX grid: 1,048,576 rows by 16,384 columns.
const MAX_ROWS: u32 = 1 << 20;
const MAX_COLUMNS: u32 = 1 << 14;

impl Xlsx {
    /// Find and replace text in the cells below `target`, returning the
    /// commit and the match count.
    pub(crate) fn find_replace(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        self.check_source_for_replay()?;
        let result = self
            .find_inner(target, spec, limits)
            .map_err(|error| error.with_format(FileFormat::Xlsx));
        self.check_source_for_replay()?;
        result
    }

    fn find_inner(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        if spec.formats() {
            // OfficeCLI: "Excel only supports 'find' with 'replace'".
            let key = spec
                .format()
                .iter()
                .next()
                .map_or(crate::find::FIND, |(key, _)| key.as_str());
            return Err(Error::new(
                Code::UnsupportedProperty,
                "an XLSX find replaces text and applies no formatting",
            )
            .with_property(key.to_owned()));
        }
        let Some(replacement) = spec.replacement() else {
            return Err(
                Error::new(Code::MissingProperty, "an XLSX find needs `replace`")
                    .with_property(crate::find::REPLACE),
            );
        };
        let scopes = scopes(self, target)?;
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let workbook = source.workbook();
        let matcher = spec.matcher();
        let mut matches = 0usize;
        let mut scratch = Scratch::default();
        let mut staged = 0usize;
        let mut changes: Vec<(usize, Address, CellText)> = Vec::new();
        let mut seen: HashSet<(usize, Address)> = HashSet::new();
        for (position, worksheet) in workbook.sheets().enumerate() {
            let areas: Vec<Rect> = scopes
                .iter()
                .filter(|(sheet, _)| *sheet == position)
                .map(|(_, area)| *area)
                .collect();
            if areas.is_empty() {
                continue;
            }
            let extent = match worksheet.stored_extent() {
                Ok(Some(extent)) => extent,
                // Chart sheets and empty sheets hold no cells.
                Ok(None) | Err(xlsx::Error::NotWorksheet { .. }) => continue,
                Err(error) => return Err(native_error(error)),
            };
            for area in areas {
                let Some(area) = intersect(area, extent) else {
                    continue;
                };
                for (address, cell) in worksheet.cells(area).map_err(native_error)? {
                    let xlsx::cell::Cell::Value(xlsx::cell::Value::Text(text)) = cell else {
                        continue;
                    };
                    if !seen.insert((position, address)) {
                        continue;
                    }
                    // One cell value is one search group of one piece.
                    let planned = plan_with(
                        &mut scratch,
                        &[Piece {
                            text: text.as_str(),
                            group: 0,
                            scoped: true,
                        }],
                        &matcher,
                        Some(replacement),
                        false,
                        limits.max_staged_bytes().saturating_sub(staged),
                    )?;
                    matches = matches.saturating_add(planned.matches);
                    let Some(output) = planned
                        .rewrites
                        .into_iter()
                        .next()
                        .and_then(|rewrite| rewrite.text)
                    else {
                        continue;
                    };
                    let content = if worksheet.has_rich_text(address).map_err(native_error)? {
                        // Record 0295: rewrite the runs of the cell's
                        // shared string, each keeping its properties.
                        let Some(mut rich) = worksheet.rich_text(address).map_err(native_error)?
                        else {
                            return Err(rich_refusal());
                        };
                        let pieces: Vec<Piece<'_>> = rich
                            .runs()
                            .iter()
                            .map(|run| Piece {
                                text: run.text(),
                                group: 0,
                                scoped: true,
                            })
                            .collect();
                        let rewrites = plan_with(
                            &mut scratch,
                            &pieces,
                            &matcher,
                            Some(replacement),
                            false,
                            limits.max_staged_bytes(),
                        )?
                        .rewrites;
                        drop(pieces);
                        for rewrite in rewrites {
                            if let Some(text) = rewrite.text {
                                validate_xml_text(&text, crate::find::REPLACE)?;
                                rich.set_run_text(rewrite.piece, text)
                                    .map_err(native_error)?;
                            }
                        }
                        CellText::Rich(rich)
                    } else {
                        validate_xml_text(&output, crate::find::REPLACE)?;
                        CellText::Plain(output.clone())
                    };
                    staged = staged.saturating_add(output.len());
                    if staged > limits.max_staged_bytes() {
                        return Err(Error::new(
                            Code::LimitExceeded,
                            "find/replace would stage more text than the semantic limits allow",
                        ));
                    }
                    changes.try_reserve(1).map_err(allocation)?;
                    changes.push((position, address, content));
                }
            }
        }
        let mut edit = super::edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        let mut index = 0;
        while index < changes.len() {
            let sheet = changes[index].0;
            let mut worksheet = edit.sheet(sheet).map_err(native_error)?.ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSX worksheet position is out of range",
                )
            })?;
            while let Some((_, address, text)) =
                changes.get(index).filter(|change| change.0 == sheet)
            {
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
            Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())?,
            matches,
        ))
    }
}

/// The new text of a changed cell.
pub(super) enum CellText {
    Plain(String),
    /// The cell's rich runs with new texts (record 0295).
    Rich(xlsx::RichText),
}

/// A rich-text cell whose runs a rewrite cannot keep (an inline rich
/// string, phonetic runs).
pub(super) fn rich_refusal() -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "this XLSX rich-text cell cannot be rewritten run by run",
    )
    .with_property(crate::find::FIND)
    .with_suggestion(
        "an inline rich string or phonetic runs would be lost by a plain value (record 0295 follow-up)",
    )
}

/// The worksheet areas the target nodes search.
fn scopes(file: &Xlsx, target: &TargetArg<'_>) -> Result<Vec<(usize, Rect)>> {
    let nav = Nav::new(file);
    let mut scopes = Vec::new();
    let mut push = |node: super::super::Node<'_>| -> Result<()> {
        let sheet = node.sheet_position();
        let area = match node.item() {
            Item::Workbook => {
                let count = file.worksheet_names()?.len();
                for sheet in 0..count {
                    scopes.push((sheet, whole()?));
                }
                return Ok(());
            },
            Item::Sheet(_) => whole()?,
            Item::Cell(cell) => Rect::single(cell.at),
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
    };
    if let Some(path) = target.path() {
        push(crate::path::resolve::get(&nav, path)?)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(&nav, query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            ));
        }
        for node in walk {
            push(node?)?;
        }
    } else {
        return Err(Error::new(
            Code::InvalidArgument,
            "the mutation target is empty",
        ));
    }
    Ok(scopes)
}

fn whole() -> Result<Rect> {
    rect(0, 0, MAX_ROWS, MAX_COLUMNS)
}

fn rect(row: u32, column: u32, end_row: u32, end_column: u32) -> Result<Rect> {
    let start = Address::at(row, column).map_err(|error| native_error(error.into()))?;
    Rect::new(start, end_row, end_column).map_err(|error| native_error(error.into()))
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
    Rect::new(Address::at(row, column).ok()?, end_row, end_column).ok()
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::from(crate::transaction::Error::Allocation { source })
}
