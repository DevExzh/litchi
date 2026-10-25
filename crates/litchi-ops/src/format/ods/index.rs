//! Lazily built read indexes over the typed ODS worksheets (record 0277).
//!
//! ODF stores a worksheet as physical row and cell *runs*
//! (`table:number-rows-repeated`, `table:number-columns-repeated`), and
//! producers pad every sheet with huge repeated runs of empty, often styled,
//! cells. The unified layer therefore enumerates *content cells*: a logical
//! cell whose run has a value, displayed text, a formula, hyperlinks or a
//! merge span. Padding runs are skipped without expansion; content runs are
//! expanded to their logical coordinates, bounded by [`MAX_CELLS`].
//!
//! Each index is built once per snapshot and sheet, on the first query,
//! view or row walk that needs it; a direct coordinate (`/Sheet1/B3`)
//! resolves against the typed runs without building one.

use std::sync::OnceLock;

use litchi_core::FileFormat;
use litchi_odf_common::XmlStreamLimits;
use litchi_ods::{Cell, CellValue, Merge, Sheet};

use crate::error::{Code, Error, Result};

/// The most content cells one worksheet index holds.
pub(crate) const MAX_CELLS: usize = 4 * 1024 * 1024;

/// One content cell: its logical coordinate and where its run lives.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Entry {
    pub(crate) row: u32,
    pub(crate) column: u32,
    physical_row: u32,
    physical_cell: u32,
}

impl Entry {
    /// The cell run this logical cell belongs to.
    pub(crate) fn cell<'s>(&self, sheet: &'s Sheet) -> Option<&'s Cell> {
        sheet
            .rows
            .get(self.physical_row as usize)?
            .cells
            .get(self.physical_cell as usize)
    }
}

/// One merge anchor: its logical coordinate and span.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Span {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) last_row: u32,
    pub(crate) last_column: u32,
}

/// The content cells of one worksheet in row-major order.
#[derive(Debug, Default)]
pub(crate) struct SheetIndex {
    cells: Vec<Entry>,
    /// Each content row with the offset of its first cell in `cells`.
    rows: Vec<(u32, u32)>,
    /// The columns holding a content cell, ascending.
    columns: Vec<u32>,
    merges: Vec<Span>,
}

impl SheetIndex {
    fn build(sheet: &Sheet) -> Result<Self> {
        let mut index = Self::default();
        let mut row_number = 0_u64;
        for (physical_row, row) in sheet.rows.iter().enumerate() {
            let repeat = row.repeat() as u64;
            if !row.cells.iter().any(is_content) {
                row_number = row_number.saturating_add(repeat);
                continue;
            }
            let physical_row = coordinate(physical_row as u64)?;
            for offset in 0..repeat {
                let logical_row = coordinate(row_number + offset)?;
                let start = index.cells.len();
                let mut column_number = 0_u64;
                for (physical_cell, cell) in row.cells.iter().enumerate() {
                    let cell_repeat = cell.repeat() as u64;
                    if is_content(cell) {
                        let physical_cell = coordinate(physical_cell as u64)?;
                        for cell_offset in 0..cell_repeat {
                            if index.cells.len() >= MAX_CELLS {
                                return Err(too_many());
                            }
                            let logical_column = coordinate(column_number + cell_offset)?;
                            index.cells.push(Entry {
                                row: logical_row,
                                column: logical_column,
                                physical_row,
                                physical_cell,
                            });
                            if let Merge::Span { rows, columns } = cell.merge {
                                index.merges.push(Span {
                                    row: logical_row,
                                    column: logical_column,
                                    last_row: logical_row
                                        .saturating_add(coordinate(rows.get() as u64 - 1)?),
                                    last_column: logical_column
                                        .saturating_add(coordinate(columns.get() as u64 - 1)?),
                                });
                            }
                        }
                    }
                    column_number = column_number.saturating_add(cell_repeat);
                }
                if index.cells.len() > start {
                    index.rows.push((logical_row, coordinate(start as u64)?));
                }
            }
            row_number = row_number.saturating_add(repeat);
        }
        let mut columns: Vec<u32> = index.cells.iter().map(|entry| entry.column).collect();
        columns.sort_unstable();
        columns.dedup();
        index.columns = columns;
        Ok(index)
    }

    /// Every content cell, row-major.
    pub(crate) fn cells(&self) -> &[Entry] {
        &self.cells
    }

    /// Each content row with the offset of its first cell.
    pub(crate) fn row_entries(&self) -> &[(u32, u32)] {
        &self.rows
    }

    /// The number of content rows.
    pub(crate) fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// The content cells of `row`, ascending by column.
    pub(crate) fn row_cells(&self, row: u32) -> &[Entry] {
        let Ok(position) = self.rows.binary_search_by_key(&row, |(number, _)| *number) else {
            return &[];
        };
        let start = self.rows[position].1 as usize;
        let end = self
            .rows
            .get(position + 1)
            .map_or(self.cells.len(), |(_, next)| *next as usize);
        &self.cells[start..end]
    }

    /// The content cells inside the inclusive rectangle
    /// `(first_row, first_column, last_row, last_column)`, row-major.
    pub(crate) fn cells_in(&self, rect: (u32, u32, u32, u32)) -> RectCells<'_> {
        let (first_row, first_column, last_row, last_column) = rect;
        let start = self.rows.partition_point(|(row, _)| *row < first_row);
        RectCells {
            index: self,
            rows: self.rows[start..].iter(),
            current: [].iter(),
            first_column,
            last_row,
            last_column,
        }
    }

    /// The columns holding a content cell, ascending.
    pub(crate) fn columns(&self) -> &[u32] {
        &self.columns
    }

    /// The merge anchors, row-major.
    pub(crate) fn merges(&self) -> &[Span] {
        &self.merges
    }

    /// The last column with content, if any.
    pub(crate) fn last_column(&self) -> Option<u32> {
        self.columns.last().copied()
    }
}

/// The content cells of an index inside one rectangle, row-major.
pub(crate) struct RectCells<'i> {
    index: &'i SheetIndex,
    rows: std::slice::Iter<'i, (u32, u32)>,
    current: std::slice::Iter<'i, Entry>,
    first_column: u32,
    last_row: u32,
    last_column: u32,
}

impl<'i> Iterator for RectCells<'i> {
    type Item = &'i Entry;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            for entry in self.current.by_ref() {
                if entry.column > self.last_column {
                    break;
                }
                if entry.column >= self.first_column {
                    return Some(entry);
                }
            }
            let (row, _) = self.rows.next()?;
            if *row > self.last_row {
                self.rows = [].iter();
                return None;
            }
            self.current = self.index.row_cells(*row).iter();
        }
    }
}

/// Whether a cell run carries content the unified layer enumerates.
pub(crate) fn is_content(cell: &Cell) -> bool {
    !matches!(cell.value, CellValue::Empty)
        || cell.formula.is_some()
        || !cell.text.is_empty()
        || !cell.hyperlinks.is_empty()
        || matches!(cell.merge, Merge::Span { .. })
}

fn coordinate(value: u64) -> Result<u32> {
    u32::try_from(value).map_err(|_| {
        Error::new(
            Code::LimitExceeded,
            "an ODS worksheet coordinate exceeds the addressable grid",
        )
        .with_format(FileFormat::Ods)
    })
}

fn too_many() -> Error {
    Error::new(
        Code::LimitExceeded,
        "an ODS worksheet holds more content cells than one read index admits",
    )
    .with_format(FileFormat::Ods)
    .with_suggestion("address cells directly by coordinate, or split the worksheet")
}

/// The logical end of each physical row run of one worksheet, so a
/// coordinate finds its row by binary search instead of a scan (record
/// 0287). Built only by callers that look up many coordinates of one
/// snapshot (a coalesced batch); one direct lookup scans the runs.
#[derive(Debug)]
pub(crate) struct RowEnds(Box<[usize]>);

impl RowEnds {
    fn build(sheet: &Sheet) -> Self {
        let mut end = 0_usize;
        Self(
            sheet
                .rows
                .iter()
                .map(|row| {
                    end = end.saturating_add(row.repeat());
                    end
                })
                .collect(),
        )
    }

    /// The physical run covering `(row, column)`: what
    /// [`Sheet::cell`] returns, found by binary search over the row runs.
    pub(crate) fn cell<'s>(&self, sheet: &'s Sheet, row: usize, column: usize) -> Option<&'s Cell> {
        // `Sheet::row` takes the first run whose end lies past `row`; the
        // ends never decrease.
        let physical = self.0.partition_point(|&end| end <= row);
        sheet.rows.get(physical)?.cell(column)
    }
}

/// The read caches of one snapshot.
#[derive(Default)]
pub(crate) struct Cache {
    sheets: OnceLock<Box<[OnceLock<std::result::Result<SheetIndex, Error>>]>>,
    rows: OnceLock<Box<[OnceLock<RowEnds>]>>,
    annotations: OnceLock<std::result::Result<litchi_ods::annotations::Snapshot, Error>>,
    /// Per-sheet worksheet features of `content.xml` (record 0327).
    features: OnceLock<Box<[OnceLock<FeatureSlot>]>>,
    /// The named cell styles conditional formats apply (record 0327).
    format_styles: OnceLock<std::result::Result<FormatStyles, Error>>,
}

type FeatureSlot = std::result::Result<litchi_ods::advanced::sheet_features::Features, Error>;
type FormatStyles =
    std::collections::BTreeMap<String, litchi_ods::advanced::sheet_features::FormatStyle>;

impl Cache {
    /// The row-run ends of the worksheet at `position` of `sheets`.
    pub(crate) fn rows<'c>(&'c self, sheets: &[Sheet], position: usize) -> Result<&'c RowEnds> {
        let slots = self.rows.get_or_init(|| {
            (0..sheets.len())
                .map(|_| OnceLock::new())
                .collect::<Vec<_>>()
                .into_boxed_slice()
        });
        let (Some(slot), Some(sheet)) = (slots.get(position), sheets.get(position)) else {
            return Err(Error::new(
                Code::PathNotFound,
                "the worksheet position is outside the workbook",
            )
            .with_format(FileFormat::Ods));
        };
        Ok(slot.get_or_init(|| RowEnds::build(sheet)))
    }

    /// The content index of the worksheet at `position` of `sheets`.
    pub(crate) fn sheet<'c>(&'c self, sheets: &[Sheet], position: usize) -> Result<&'c SheetIndex> {
        let slots = self.sheets.get_or_init(|| {
            (0..sheets.len())
                .map(|_| OnceLock::new())
                .collect::<Vec<_>>()
                .into_boxed_slice()
        });
        let (Some(slot), Some(sheet)) = (slots.get(position), sheets.get(position)) else {
            return Err(Error::new(
                Code::PathNotFound,
                "the worksheet position is outside the workbook",
            )
            .with_format(FileFormat::Ods));
        };
        slot.get_or_init(|| SheetIndex::build(sheet))
            .as_ref()
            .map_err(Clone::clone)
    }

    /// The worksheet features of the sheet at `position` of `count`
    /// sheets, read by `read` once.
    pub(crate) fn features(
        &self,
        count: usize,
        position: usize,
        read: impl FnOnce() -> Result<litchi_ods::advanced::sheet_features::Features>,
    ) -> Result<&litchi_ods::advanced::sheet_features::Features> {
        let slots = self.features.get_or_init(|| {
            (0..count)
                .map(|_| OnceLock::new())
                .collect::<Vec<_>>()
                .into_boxed_slice()
        });
        let slot = slots.get(position).ok_or_else(|| {
            Error::new(
                Code::PathNotFound,
                "the worksheet position is outside the workbook",
            )
            .with_format(FileFormat::Ods)
        })?;
        if let Some(result) = slot.get() {
            return result.as_ref().map_err(Clone::clone);
        }
        let read = read();
        slot.get_or_init(|| read).as_ref().map_err(Clone::clone)
    }

    /// The named cell styles of `styles.xml`, read by `read` once.
    pub(crate) fn format_styles(
        &self,
        read: impl FnOnce() -> Result<FormatStyles>,
    ) -> Result<&FormatStyles> {
        if let Some(result) = self.format_styles.get() {
            return result.as_ref().map_err(Clone::clone);
        }
        let read = read();
        self.format_styles
            .get_or_init(|| read)
            .as_ref()
            .map_err(Clone::clone)
    }

    /// The cell annotations of the snapshot, parsed from `xml` once.
    pub(crate) fn annotations<'c>(
        &'c self,
        xml: impl FnOnce() -> Result<&'c str>,
        limits: &XmlStreamLimits,
    ) -> Result<&'c litchi_ods::annotations::Snapshot> {
        if let Some(result) = self.annotations.get() {
            return result.as_ref().map_err(Clone::clone);
        }
        let parsed = xml().and_then(|xml| {
            litchi_ods::annotations::Snapshot::parse(xml, limits)
                .map_err(|error| Error::from(error).with_format(FileFormat::Ods))
        });
        self.annotations
            .get_or_init(|| parsed)
            .as_ref()
            .map_err(Clone::clone)
    }
}
