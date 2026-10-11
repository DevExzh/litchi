//! The small native-owner bridge used by the legacy XLS read adapter.
//!
//! `litchi_xls` deliberately exposes two worksheet routes: the eager reader
//! implements the generic core worksheet traits, while the positional owner
//! exposes a source-backed cell scanner.  This module turns both routes into
//! one immutable, sparse grid.  The grid is materialized only when a path,
//! query, view or child walk first needs worksheet payload.

use std::collections::BTreeMap;
use std::sync::Arc;

use litchi_core::sheet::{Cell, CellValue, WorkbookTrait};
use litchi_xls::{Package, SheetKind, SheetMetadata, SheetVisibility};

use super::Xls;
use crate::error::{Code, Error, Result};
use crate::path::vocabulary::Case;

/// A worksheet descriptor whose BIFF payload remains unopened.
#[derive(Clone)]
pub(super) struct SheetHandle {
    package: Package,
    /// Position in the complete BoundSheet8 catalog.
    position: usize,
    metadata: SheetMetadata,
}

pub(super) type EagerCells<'a> = Box<dyn Iterator<Item = &'a litchi_xls::Cell> + Send + Sync + 'a>;

impl SheetHandle {
    pub(super) fn at(file: &Xls, position: usize) -> Result<Option<Self>> {
        let workbook = file.owner().workbook();
        let Some(entry) = workbook.sheet(position)? else {
            return Ok(None);
        };
        Ok(Some(Self::from_metadata(file, position, entry)))
    }

    pub(super) fn named(file: &Xls, name: &str, case: Case) -> Result<Option<Self>> {
        let names = file.owner().workbook().sheets()?;
        let mut found: Option<(usize, SheetMetadata)> = None;
        for (position, candidate) in names.into_iter().enumerate() {
            if case.matches(candidate.name(), name) {
                if found.is_some() {
                    return Err(Error::new(
                        Code::InvalidPath,
                        "semantic sheet name selects multiple worksheets",
                    )
                    .with_format(litchi_core::FileFormat::Xls));
                }
                found = Some((position, candidate));
            }
        }
        Ok(found.map(|(position, entry)| Self::from_metadata(file, position, entry)))
    }

    fn from_metadata(file: &Xls, position: usize, entry: SheetMetadata) -> Self {
        Self {
            package: file.owner().clone(),
            position,
            metadata: entry,
        }
    }

    pub(super) const fn position(&self) -> usize {
        self.position
    }

    pub(super) fn name(&self) -> &str {
        self.metadata.name()
    }

    pub(super) fn visibility(&self) -> SheetVisibility {
        self.metadata.visibility()
    }

    pub(super) fn is_grid(&self) -> bool {
        self.metadata.parsed_worksheet_index().is_some()
            && matches!(self.metadata.kind(), SheetKind::WorksheetOrDialog)
    }

    pub(super) fn grid(&self) -> Result<Grid> {
        let worksheet_index = self.metadata.parsed_worksheet_index().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the selected XLS tab has no worksheet grid",
            )
            .with_format(litchi_core::FileFormat::Xls)
        })?;
        Grid::from_package(&self.package, worksheet_index)
    }

    /// Read one cell without materializing the complete sparse worksheet.
    ///
    /// The source-backed owner still validates the worksheet stream while
    /// answering this lookup, but it returns only the selected cell. Eager
    /// readers use their indexed worksheet map directly.
    pub(super) fn cell(&self, at: (u32, u32)) -> Result<Option<CellData>> {
        let worksheet_index = self.metadata.parsed_worksheet_index().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the selected XLS tab has no worksheet grid",
            )
            .with_format(litchi_core::FileFormat::Xls)
        })?;
        if let Some(source) = self.package.source_workbook() {
            return Ok(source
                .cell_by_index(worksheet_index, at.0, at.1)?
                .map(|cell| CellData {
                    value: cell.value().clone(),
                    formula: cell.is_formula(),
                }));
        }
        let reader = self.package.reader_ref()?;
        let worksheet = reader
            .xls_worksheet(worksheet_index)
            .map_err(|error| worksheet_error(error.to_string()))?;
        Ok(worksheet.get_cell(at.0, at.1).map(|cell| CellData {
            value: cell.value().clone(),
            formula: cell.is_formula(),
        }))
    }

    /// Borrow the eager worksheet's native cell iterator when the owner has
    /// already parsed its cells. Returning `None` preserves the source-backed
    /// fallback until a resumable positional cell cursor is available.
    pub(super) fn eager_cells<'f>(&self, file: &'f Xls) -> Result<Option<EagerCells<'f>>> {
        if self.package.source_workbook().is_some() {
            return Ok(None);
        }
        let worksheet_index = self.metadata.parsed_worksheet_index().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the selected XLS tab has no worksheet grid",
            )
            .with_format(litchi_core::FileFormat::Xls)
        })?;
        let reader = file.owner().reader_ref()?;
        let worksheet = reader
            .xls_worksheet(worksheet_index)
            .map_err(|error| worksheet_error(error.to_string()))?;
        Ok(Some(Box::new(worksheet.stored_cells())))
    }
}

/// One owned sparse cell snapshot.  Formula text is not part of the common
/// source-backed native cell projection, so the adapter reports only the
/// scalar value and retains the eager route's formula bit for statistics.
#[derive(Debug, Clone)]
pub(super) struct CellData {
    pub(super) value: CellValue,
    pub(super) formula: bool,
}

/// A lazily materialized worksheet grid.
#[derive(Debug, Clone)]
pub(super) struct Grid {
    cells: Arc<BTreeMap<(u32, u32), CellData>>,
    extent: Option<(u32, u32, u32, u32)>,
}

impl Grid {
    fn from_package(package: &Package, position: usize) -> Result<Self> {
        let mut cells = BTreeMap::new();
        if let Some(source) = package.source_workbook() {
            let worksheet = source
                .worksheet_by_index(position)?
                .ok_or_else(|| missing_sheet(position))?;
            worksheet.visit_cells(|cell| {
                cells.insert(
                    (cell.row(), cell.column()),
                    CellData {
                        value: cell.value().clone(),
                        formula: cell.is_formula(),
                    },
                );
                Ok(())
            })?;
        } else {
            let reader = package.reader_ref()?;
            let worksheet = WorkbookTrait::worksheet_by_index(reader, position)
                .map_err(|error| worksheet_error(error.to_string()))?;
            let mut iterator = worksheet.cells();
            while let Some(cell) = iterator.next() {
                let cell = cell.map_err(|error| worksheet_error(error.to_string()))?;
                cells.insert(
                    (cell.row(), cell.column()),
                    CellData {
                        value: cell.value().clone(),
                        formula: cell.is_formula(),
                    },
                );
            }
        }
        let extent: Option<(u32, u32, u32, u32)> =
            cells.keys().fold(None, |extent, &(row, column)| {
                Some(match extent {
                    None => (row, column, row, column),
                    Some((first_row, first_column, last_row, last_column)) => (
                        first_row.min(row),
                        first_column.min(column),
                        last_row.max(row),
                        last_column.max(column),
                    ),
                })
            });
        Ok(Self {
            cells: Arc::new(cells),
            extent,
        })
    }

    pub(super) fn record(&self, at: (u32, u32)) -> Option<&CellData> {
        self.cells.get(&at)
    }

    pub(super) fn cells_in(
        &self,
        rect: (u32, u32, u32, u32),
    ) -> impl Iterator<Item = ((u32, u32), &CellData)> {
        let (first_row, first_column, last_row, last_column) = rect;
        self.cells
            .range((first_row, 0)..=(last_row, u32::MAX))
            .filter(move |((_, column), _)| *column >= first_column && *column <= last_column)
            .map(|(&(row, column), cell)| ((row, column), cell))
    }

    pub(super) fn extents(&self) -> Option<(u32, u32, u32, u32)> {
        self.extent
    }

    pub(super) fn row_indices(&self) -> Vec<u32> {
        let mut rows = self.cells.keys().map(|(row, _)| *row).collect::<Vec<_>>();
        rows.sort_unstable();
        rows.dedup();
        rows
    }

    pub(super) fn column_ranges(&self) -> Vec<(u32, u32)> {
        let mut columns = self
            .cells
            .keys()
            .map(|(_, column)| *column)
            .collect::<Vec<_>>();
        columns.sort_unstable();
        columns.dedup();
        let mut ranges: Vec<(u32, u32)> = Vec::new();
        for column in columns {
            match ranges.last_mut() {
                Some((_, last)) if column <= last.saturating_add(1) => *last = column,
                _ => ranges.push((column, column)),
            }
        }
        ranges
    }
}

fn missing_sheet(position: usize) -> Error {
    Error::new(
        Code::NotFound,
        format!("XLS worksheet position {position} does not exist"),
    )
    .with_format(litchi_core::FileFormat::Xls)
}

fn worksheet_error(message: String) -> Error {
    Error::new(Code::CorruptFile, message).with_format(litchi_core::FileFormat::Xls)
}
