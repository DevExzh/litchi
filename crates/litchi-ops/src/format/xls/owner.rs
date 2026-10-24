//! The small native-owner bridge used by the legacy XLS read adapter.
//!
//! `litchi_xls` deliberately exposes two worksheet routes: the eager reader
//! implements the generic core worksheet traits, while the positional owner
//! exposes a source-backed cell scanner.  This module turns both routes into
//! one immutable, sparse grid.  The grid is materialized only when a path,
//! query, view or child walk first needs worksheet payload.

use std::sync::Arc;

use litchi_core::sheet::{Cell, CellValue};
use litchi_xls::{Package, SheetKind, SheetMetadata, SheetVisibility};

use super::Xls;
use crate::error::{Code, Error, Result};
use crate::path::vocabulary::Case;

/// A worksheet descriptor whose BIFF payload remains unopened.  It borrows
/// the immutable owner and shares its immutable tab catalog, so resolving a
/// tab copies no names.
#[derive(Clone)]
pub(super) struct SheetHandle<'f> {
    package: &'f Package,
    /// Position in the complete BoundSheet8 catalog.
    position: usize,
    catalog: Arc<[SheetMetadata]>,
}

impl<'f> SheetHandle<'f> {
    pub(super) fn at(file: &'f Xls, position: usize) -> Result<Option<Self>> {
        let catalog = catalog(file)?;
        let Some(_entry) = catalog.get(position) else {
            return Ok(None);
        };
        Ok(Some(Self::from_catalog(file, position, catalog)))
    }

    pub(super) fn named(file: &'f Xls, name: &str, case: Case) -> Result<Option<Self>> {
        let catalog = catalog(file)?;
        let mut found = None;
        for (position, candidate) in catalog.iter().enumerate() {
            if case.matches(candidate.name(), name) {
                if found.is_some() {
                    return Err(Error::new(
                        Code::InvalidPath,
                        "semantic sheet name selects multiple worksheets",
                    )
                    .with_format(litchi_core::FileFormat::Xls));
                }
                found = Some(position);
            }
        }
        Ok(found.map(|position| Self::from_catalog(file, position, catalog)))
    }

    fn from_catalog(file: &'f Xls, position: usize, catalog: Arc<[SheetMetadata]>) -> Self {
        Self {
            package: file.owner(),
            position,
            catalog,
        }
    }

    fn metadata(&self) -> &SheetMetadata {
        &self.catalog[self.position]
    }

    pub(super) const fn position(&self) -> usize {
        self.position
    }

    pub(super) fn name(&self) -> &str {
        self.metadata().name()
    }

    pub(super) fn visibility(&self) -> SheetVisibility {
        self.metadata().visibility()
    }

    pub(super) fn is_grid(&self) -> bool {
        self.metadata().parsed_worksheet_index().is_some()
            && matches!(self.metadata().kind(), SheetKind::WorksheetOrDialog)
    }

    pub(super) fn grid(&self) -> Result<Grid> {
        let worksheet_index = self.metadata().parsed_worksheet_index().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the selected XLS tab has no worksheet grid",
            )
            .with_format(litchi_core::FileFormat::Xls)
        })?;
        Grid::from_package(self.package, worksheet_index)
    }

    /// Read one cell without materializing the complete sparse worksheet.
    ///
    /// The source-backed owner still validates the worksheet stream while
    /// answering this lookup, but it returns only the selected cell. Eager
    /// readers use their indexed worksheet map directly.
    pub(super) fn cell(&self, at: (u32, u32)) -> Result<Option<CellData>> {
        let worksheet_index = self.metadata().parsed_worksheet_index().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the selected XLS tab has no worksheet grid",
            )
            .with_format(litchi_core::FileFormat::Xls)
        })?;
        if let Some(source) = self.package.source_workbook() {
            return source
                .cell_by_index(worksheet_index, at.0, at.1)?
                .map(|cell| CellData::from_source(self.package, worksheet_index, &cell))
                .transpose();
        }
        let reader = self.package.reader_ref()?;
        let worksheet = reader
            .xls_worksheet(worksheet_index)
            .map_err(|error| worksheet_error(error.to_string()))?;
        Ok(worksheet.get_cell(at.0, at.1).map(CellData::from_eager))
    }

    /// The row's `Row` record layout, through the eager reader.
    ///
    /// The positional owner does not project layout records, so a
    /// source-backed workbook materializes its eager reader once (cached by
    /// the native owner) on the first layout read.
    pub(super) fn row_layout(&self, row: u32) -> Result<Option<RowLayout>> {
        let Ok(row) = u16::try_from(row) else {
            return Ok(None);
        };
        let worksheet = self.eager_worksheet()?;
        Ok(worksheet.row_layout(row).map(|layout| RowLayout {
            height_twips: layout.height_twips(),
            hidden: layout.is_hidden(),
            custom_height: layout.has_custom_height(),
            outline_level: layout.outline_level(),
            collapsed: layout.is_collapsed(),
        }))
    }

    /// The `ColInfo` range layout covering one column, through the eager
    /// reader (see [`Self::row_layout`]).
    pub(super) fn column_layout(&self, column: u32) -> Result<Option<ColumnLayout>> {
        let Ok(column) = u8::try_from(column) else {
            return Ok(None);
        };
        let worksheet = self.eager_worksheet()?;
        Ok(worksheet.column_layout(column).map(|layout| ColumnLayout {
            width_256ths: layout.width_256ths(),
            hidden: layout.is_hidden(),
            user_set: layout.is_user_set(),
            outline_level: layout.outline_level(),
            collapsed: layout.is_collapsed(),
        }))
    }

    fn eager_worksheet(&self) -> Result<&'f litchi_xls::Worksheet> {
        let worksheet_index = self.metadata().parsed_worksheet_index().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the selected XLS tab has no worksheet grid",
            )
            .with_format(litchi_core::FileFormat::Xls)
        })?;
        self.package
            .reader_ref()?
            .xls_worksheet(worksheet_index)
            .map_err(|error| worksheet_error(error.to_string()))
    }

    /// The positional owner's cell cursor over this worksheet, with the
    /// worksheet's index (record 0267); `None` for an eager owner, whose
    /// cells are already parsed ([`Self::eager_cells`]), and for a worksheet
    /// whose stored positions do not ascend (its sorted grid is not its
    /// stream order).
    pub(super) fn cell_cursor(&self) -> Result<Option<(litchi_xls::CellCursor, usize)>> {
        let Some(source) = self.package.source_workbook() else {
            return Ok(None);
        };
        let worksheet_index = self.metadata().parsed_worksheet_index().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the selected XLS tab has no worksheet grid",
            )
            .with_format(litchi_core::FileFormat::Xls)
        })?;
        let worksheet = source
            .worksheet_by_index(worksheet_index)?
            .ok_or_else(|| missing_sheet(worksheet_index))?;
        let Some(cursor) = worksheet.cell_cursor()? else {
            return Ok(None);
        };
        // A formula whose text needs the workbook's name tables reads it
        // from the eager worksheet ([`CellData::from_source`]), as the grid
        // does; load it now, so a worksheet the eager reader refuses is
        // refused before any cell, as the grid refuses it.
        if cursor.needs_formula_context() {
            self.eager_worksheet()?;
        }
        Ok(Some((cursor, worksheet_index)))
    }

    /// The owner, for cells a cursor reads.
    pub(super) const fn package(&self) -> &'f Package {
        self.package
    }

    /// Borrow the eager worksheet's native cell iterator when the owner has
    /// already parsed its cells. Returning `None` preserves the source-backed
    /// fallback until a resumable positional cell cursor is available.
    pub(super) fn eager_cells(
        &self,
        file: &'f Xls,
    ) -> Result<Option<litchi_xls::worksheet::StoredCells<'f>>> {
        if self.package.source_workbook().is_some() {
            return Ok(None);
        }
        let worksheet_index = self.metadata().parsed_worksheet_index().ok_or_else(|| {
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
        Ok(Some(worksheet.stored_cells()))
    }
}

/// Use the adapter-local shared catalog only inside the active unified read
/// fence. Outside that scope, retain the package projection's source check.
fn catalog(file: &Xls) -> Result<Arc<[SheetMetadata]>> {
    if crate::path::fence::is_active(file.fence_key())
        && let Some(catalog) = file.cached_sheet_catalog()
    {
        return Ok(catalog);
    }
    file.sheet_catalog()
}

/// One row's layout record.
#[derive(Debug, Clone, Copy)]
pub(crate) struct RowLayout {
    pub(super) height_twips: u16,
    pub(super) hidden: bool,
    pub(super) custom_height: bool,
    pub(super) outline_level: u8,
    pub(super) collapsed: bool,
}

/// One column's covering `ColInfo` layout.
#[derive(Debug, Clone, Copy)]
pub(crate) struct ColumnLayout {
    pub(super) width_256ths: u16,
    pub(super) hidden: bool,
    pub(super) user_set: bool,
    pub(super) outline_level: u8,
    pub(super) collapsed: bool,
}

/// One owned sparse cell snapshot: the semantic (cached) value, the formula
/// bit used by statistics, and the rendered formula text.
#[derive(Debug, Clone)]
pub(super) struct CellData {
    pub(super) value: CellValue,
    pub(super) formula: bool,
    /// Rendered formula text with its leading `=`, when the cell is a formula
    /// whose tokens render.
    pub(super) formula_text: Option<Arc<str>>,
    /// The record's cell format index.
    pub(super) xf: u16,
}

impl CellData {
    pub(super) fn from_eager(cell: &litchi_xls::Cell) -> Self {
        Self {
            value: cell.value().clone(),
            formula: cell.is_formula(),
            formula_text: cell.formula().map(Arc::from),
            xf: cell.xf_index(),
        }
    }

    /// The positional owner renders formulas without the workbook's name and
    /// external-sheet tables.  A formula that needs them falls back to the
    /// eager reader's complete rendering; that compatibility escape
    /// materializes the workbook once and is cached by the native owner.
    pub(super) fn from_source(
        package: &Package,
        worksheet: usize,
        cell: &litchi_xls::SourceBackedCell,
    ) -> Result<Self> {
        let formula_text = match cell.formula() {
            Some(text) => Some(Arc::from(text)),
            None if cell.is_formula() => package
                .reader_ref()?
                .xls_worksheet(worksheet)
                .map_err(|error| worksheet_error(error.to_string()))?
                .get_cell(cell.row(), cell.column())
                .and_then(litchi_xls::Cell::formula)
                .map(Arc::from),
            None => None,
        };
        Ok(Self {
            value: cell.value().clone(),
            formula: cell.is_formula(),
            formula_text,
            xf: cell.xf_index(),
        })
    }
}

/// A lazily materialized worksheet grid: stored cells sorted by
/// `(row, column)`.  A duplicate coordinate keeps its last physical record,
/// as the worksheet projection does.
#[derive(Debug, Clone)]
pub(super) struct Grid {
    cells: Arc<[((u32, u32), CellData)]>,
    extent: Option<(u32, u32, u32, u32)>,
}

impl Grid {
    /// The grid of a node resolved by [`super::Nav::coordinates`]: no cells
    /// are read. Only the write resolvers that use only a node's coordinates
    /// make such nodes, and they do not escape them.
    pub(super) fn coordinates_only() -> Self {
        Self {
            cells: Arc::from(Vec::new()),
            extent: None,
        }
    }

    fn from_package(package: &Package, position: usize) -> Result<Self> {
        let mut cells: Vec<((u32, u32), CellData)> = Vec::new();
        if let Some(source) = package.source_workbook() {
            let worksheet = source
                .worksheet_by_index(position)?
                .ok_or_else(|| missing_sheet(position))?;
            let mut fallback = false;
            worksheet.visit_cells(|cell| {
                fallback |= cell.is_formula() && cell.formula().is_none();
                cells.push((
                    (cell.row(), cell.column()),
                    CellData {
                        value: cell.value().clone(),
                        formula: cell.is_formula(),
                        formula_text: cell.formula().map(Arc::from),
                        xf: cell.xf_index(),
                    },
                ));
                Ok(())
            })?;
            if fallback {
                let reader = package.reader_ref()?;
                let eager = reader
                    .xls_worksheet(position)
                    .map_err(|error| worksheet_error(error.to_string()))?;
                for ((row, column), cell) in &mut cells {
                    if cell.formula && cell.formula_text.is_none() {
                        cell.formula_text = eager
                            .get_cell(*row, *column)
                            .and_then(litchi_xls::Cell::formula)
                            .map(Arc::from);
                    }
                }
            }
        } else {
            let reader = package.reader_ref()?;
            let worksheet = reader
                .xls_worksheet(position)
                .map_err(|error| worksheet_error(error.to_string()))?;
            cells.extend(
                worksheet
                    .stored_cells()
                    .map(|cell| ((cell.row(), cell.column()), CellData::from_eager(cell))),
            );
        }
        // Physical record order is usually row-major already; a stable sort
        // keeps duplicate coordinates in record order so the last one wins.
        if !cells.is_sorted_by(|left, right| left.0 < right.0) {
            cells.sort_by_key(|(at, _)| *at);
            let mut deduplicated: Vec<((u32, u32), CellData)> = Vec::with_capacity(cells.len());
            for entry in cells {
                match deduplicated.last_mut() {
                    Some(last) if last.0 == entry.0 => *last = entry,
                    _ => deduplicated.push(entry),
                }
            }
            cells = deduplicated;
        }
        let extent: Option<(u32, u32, u32, u32)> =
            cells.iter().fold(None, |extent, &((row, column), _)| {
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
            cells: Arc::from(cells),
            extent,
        })
    }

    pub(super) fn record(&self, at: (u32, u32)) -> Option<&CellData> {
        self.cells
            .binary_search_by_key(&at, |(key, _)| *key)
            .ok()
            .map(|index| &self.cells[index].1)
    }

    pub(super) fn cells_in(
        &self,
        rect: (u32, u32, u32, u32),
    ) -> impl Iterator<Item = ((u32, u32), &CellData)> {
        let (first_row, first_column, last_row, last_column) = rect;
        let start = self.cells.partition_point(|(at, _)| *at < (first_row, 0));
        let end = self
            .cells
            .partition_point(|(at, _)| at.0 <= last_row)
            .max(start);
        self.cells[start..end]
            .iter()
            .filter(move |((_, column), _)| *column >= first_column && *column <= last_column)
            .map(|(at, cell)| (*at, cell))
    }

    pub(super) fn extents(&self) -> Option<(u32, u32, u32, u32)> {
        self.extent
    }

    pub(super) fn row_indices(&self) -> Vec<u32> {
        let mut rows = self
            .cells
            .iter()
            .map(|((row, _), _)| *row)
            .collect::<Vec<_>>();
        rows.dedup();
        rows
    }

    pub(super) fn column_ranges(&self) -> Vec<(u32, u32)> {
        let mut columns = self
            .cells
            .iter()
            .map(|((_, column), _)| *column)
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

/// Optional-handle conveniences for node projections.
pub(super) trait LayoutLookup {
    fn and_then_row(self, row: u32) -> Result<Option<RowLayout>>;
    fn and_then_column(self, column: u32) -> Result<Option<ColumnLayout>>;
}

impl LayoutLookup for Option<SheetHandle<'_>> {
    fn and_then_row(self, row: u32) -> Result<Option<RowLayout>> {
        self.map_or(Ok(None), |handle| handle.row_layout(row))
    }

    fn and_then_column(self, column: u32) -> Result<Option<ColumnLayout>> {
        self.map_or(Ok(None), |handle| handle.column_layout(column))
    }
}
