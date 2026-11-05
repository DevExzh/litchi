//! The worksheet, cell and iterator types the workbook handles lend through
//! the generic [`litchi_core::sheet`] traits.
//!
//! [`WorkbookView`](crate::view::WorkbookView) and the format handles
//! (`format::xlsx::Xlsx`, `format::xlsb::Xlsb`, `format::xls::Xls`) name
//! these types as their
//! [`WorkbookTrait`](litchi_core::sheet::WorkbookTrait) and
//! [`Worksheet`](litchi_core::sheet::Worksheet) associated types. Each is an
//! enum over the concrete per-format types, so a read dispatches statically
//! with one match and allocates nothing for dispatch.
//!
//! ```no_run
//! use litchi_core::sheet::{Cell as _, WorkbookTrait, Worksheet as _};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
//! let file = litchi_ops::open("budget.xlsx")?;
//! let workbook = file.workbook()?;
//! let sheet: litchi_ops::view::sheet::Worksheet<'_> =
//!     WorkbookTrait::worksheet_by_index(&workbook, 0)?;
//! for cell in sheet.cells() {
//!     let cell = cell?;
//!     println!("{} = {:?}", cell.coordinate(), cell.value());
//! }
//! # Ok(())
//! # }
//! ```

// With ODS alone the per-format enums are uninhabited: every read refuses
// before reaching them.
#![cfg_attr(
    not(any(feature = "xlsx", feature = "xlsb", feature = "xls")),
    allow(unreachable_code, unused_variables)
)]

use std::borrow::Cow;
#[cfg(feature = "xlsb")]
use std::sync::Arc;

use litchi_core::sheet::{Cell as CoreCell, CellValue, Worksheet as CoreWorksheet};

use crate::Result;

#[cfg(feature = "xlsb")]
use crate::format::xlsb::source;
#[cfg(feature = "xlsx")]
use crate::format::xlsx::sheet as xlsx;

/// Apply `$body` to the format value inside `$value`, an inner enum of this
/// module whose variants are `Xlsx`, `XlsbSource`, `XlsbEager` and `Xls`.
macro_rules! each {
    ($kind:ident, $value:expr, $inner:ident => $body:expr $(, $extra:pat => $other:expr)*) => {
        match $value {
            #[cfg(feature = "xlsx")]
            $kind::Xlsx($inner) => $body,
            #[cfg(feature = "xlsb")]
            $kind::XlsbSource($inner) => $body,
            #[cfg(feature = "xlsb")]
            $kind::XlsbEager($inner) => $body,
            #[cfg(feature = "xls")]
            $kind::Xls($inner) => $body,
            $($extra => $other,)*
        }
    };
}

/// One worksheet of a workbook, borrowed from its handle for `'a`.
pub struct Worksheet<'a>(SheetInner<'a>);

enum SheetInner<'a> {
    #[cfg(feature = "xlsx")]
    Xlsx(xlsx::Sheet),
    #[cfg(feature = "xlsb")]
    XlsbSource(source::SourceSheet<'a>),
    #[cfg(feature = "xlsb")]
    XlsbEager(Arc<litchi_xlsb::Worksheet>),
    #[cfg(feature = "xls")]
    Xls(&'a litchi_xls::Worksheet),
    #[allow(
        dead_code,
        reason = "uninhabited: carries the lifetime when no borrowing format is enabled"
    )]
    Never(std::convert::Infallible, std::marker::PhantomData<&'a ()>),
}

impl<'a> Worksheet<'a> {
    #[cfg(feature = "xlsx")]
    pub(crate) fn xlsx(sheet: xlsx::Sheet) -> Self {
        Self(SheetInner::Xlsx(sheet))
    }

    #[cfg(feature = "xlsb")]
    pub(crate) fn xlsb_source(sheet: source::SourceSheet<'a>) -> Self {
        Self(SheetInner::XlsbSource(sheet))
    }

    #[cfg(feature = "xlsb")]
    pub(crate) fn xlsb_eager(sheet: Arc<litchi_xlsb::Worksheet>) -> Self {
        Self(SheetInner::XlsbEager(sheet))
    }

    #[cfg(feature = "xls")]
    pub(crate) fn xls(sheet: &'a litchi_xls::Worksheet) -> Self {
        Self(SheetInner::Xls(sheet))
    }
}

impl std::fmt::Debug for Worksheet<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Worksheet")
            .field("name", &CoreWorksheet::name(self))
            .finish_non_exhaustive()
    }
}

/// One cell of a [`Worksheet`].
#[derive(Clone)]
pub struct Cell<'a>(CellInner<'a>);

#[derive(Clone)]
enum CellInner<'a> {
    #[cfg(feature = "xlsx")]
    Xlsx(xlsx::XlsxCell),
    #[cfg(feature = "xlsb")]
    XlsbSource(source::FacadeCell),
    #[cfg(feature = "xlsb")]
    XlsbEager(Stored<'a, litchi_xlsb::Cell>),
    #[cfg(feature = "xls")]
    Xls(Stored<'a, litchi_xls::Cell>),
    #[allow(
        dead_code,
        reason = "uninhabited: carries the lifetime when no borrowing format is enabled"
    )]
    Never(std::convert::Infallible, std::marker::PhantomData<&'a ()>),
}

impl std::fmt::Debug for Cell<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Cell")
            .field("row", &CoreCell::row(self))
            .field("column", &CoreCell::column(self))
            .finish_non_exhaustive()
    }
}

impl CoreCell for Cell<'_> {
    #[inline]
    fn column(&self) -> u32 {
        each!(CellInner, &self.0, cell => cell.column(), CellInner::Never(never, _) => match *never {})
    }

    #[inline]
    fn coordinate(&self) -> String {
        each!(CellInner, &self.0, cell => cell.coordinate(), CellInner::Never(never, _) => match *never {})
    }

    #[inline]
    fn is_date(&self) -> bool {
        each!(CellInner, &self.0, cell => cell.is_date(), CellInner::Never(never, _) => match *never {})
    }

    #[inline]
    fn is_empty(&self) -> bool {
        each!(CellInner, &self.0, cell => cell.is_empty(), CellInner::Never(never, _) => match *never {})
    }

    #[inline]
    fn is_formula(&self) -> bool {
        each!(CellInner, &self.0, cell => cell.is_formula(), CellInner::Never(never, _) => match *never {})
    }

    #[inline]
    fn row(&self) -> u32 {
        each!(CellInner, &self.0, cell => cell.row(), CellInner::Never(never, _) => match *never {})
    }

    #[inline]
    fn value(&self) -> &CellValue {
        each!(CellInner, &self.0, cell => cell.value(), CellInner::Never(never, _) => match *never {})
    }
}

/// A cell of an eager BIFF worksheet (XLS, XLSB): a stored cell borrowed
/// from the worksheet, or an absent position, which reads as empty.
///
/// The leaf worksheets lend stored cells by reference and an absent position
/// as an owned empty cell; keeping only the absent position's coordinates
/// keeps every [`Cell`] small (a leaf BIFF cell is several times the size of
/// the other formats' cells).
#[cfg(any(feature = "xlsb", feature = "xls"))]
#[derive(Clone)]
enum Stored<'a, C> {
    Cell(&'a C),
    Absent { row: u32, column: u32 },
}

/// A leaf BIFF cell type that can describe an absent position.
#[cfg(any(feature = "xlsb", feature = "xls"))]
trait Blank: CoreCell + Clone + Sized {
    fn blank(row: u32, column: u32) -> Self;
}

#[cfg(feature = "xlsb")]
impl Blank for litchi_xlsb::Cell {
    fn blank(row: u32, column: u32) -> Self {
        Self::new(row, column, CellValue::Empty)
    }
}

#[cfg(feature = "xls")]
impl Blank for litchi_xls::Cell {
    fn blank(row: u32, column: u32) -> Self {
        Self::new(row, column, CellValue::Empty)
    }
}

#[cfg(any(feature = "xlsb", feature = "xls"))]
impl<'a, C: Blank> Stored<'a, C> {
    /// The leaf's stored-cell lookup at `(row, column)`, which the leaf
    /// trait's `cell` wraps: `None` is an absent position.
    fn at(cell: Option<&'a C>, row: u32, column: u32) -> Self {
        cell.map_or(Self::Absent { row, column }, Self::Cell)
    }

    /// A leaf read: borrowed when stored; an owned cell is the leaf's empty
    /// cell for an absent position.
    fn new(cell: Cow<'a, C>) -> Self {
        match cell {
            Cow::Borrowed(cell) => Self::Cell(cell),
            Cow::Owned(cell) => {
                debug_assert!(cell.is_empty() && !cell.is_formula());
                Self::Absent {
                    row: cell.row(),
                    column: cell.column(),
                }
            },
        }
    }
}

#[cfg(any(feature = "xlsb", feature = "xls"))]
impl<C: Blank> CoreCell for Stored<'_, C> {
    #[inline]
    fn column(&self) -> u32 {
        match self {
            Self::Cell(cell) => cell.column(),
            Self::Absent { column, .. } => *column,
        }
    }

    #[inline]
    fn coordinate(&self) -> String {
        match self {
            Self::Cell(cell) => cell.coordinate(),
            Self::Absent { row, column } => C::blank(*row, *column).coordinate(),
        }
    }

    #[inline]
    fn is_date(&self) -> bool {
        match self {
            Self::Cell(cell) => cell.is_date(),
            Self::Absent { row, column } => C::blank(*row, *column).is_date(),
        }
    }

    #[inline]
    fn is_empty(&self) -> bool {
        match self {
            Self::Cell(cell) => cell.is_empty(),
            Self::Absent { row, column } => C::blank(*row, *column).is_empty(),
        }
    }

    #[inline]
    fn is_formula(&self) -> bool {
        match self {
            Self::Cell(cell) => cell.is_formula(),
            Self::Absent { row, column } => C::blank(*row, *column).is_formula(),
        }
    }

    #[inline]
    fn row(&self) -> u32 {
        match self {
            Self::Cell(cell) => cell.row(),
            Self::Absent { row, .. } => *row,
        }
    }

    #[inline]
    fn value(&self) -> &CellValue {
        match self {
            Self::Cell(cell) => cell.value(),
            Self::Absent { .. } => CellValue::EMPTY,
        }
    }
}

/// The stored cells of a [`Worksheet`]: the iterator
/// [`Worksheet::cells`](CoreWorksheet::cells) returns.
pub struct Cells<'a>(CellsInner<'a>);

enum CellsInner<'a> {
    #[cfg(feature = "xlsx")]
    Xlsx(xlsx::Cells),
    #[cfg(feature = "xlsb")]
    XlsbSource(source::Cells),
    #[cfg(feature = "xlsb")]
    XlsbEager(litchi_xlsb::sheet::StoredCells<'a>),
    #[cfg(feature = "xls")]
    Xls(litchi_xls::worksheet::Cells<'a>),
    #[allow(
        dead_code,
        reason = "uninhabited: carries the lifetime when no borrowing format is enabled"
    )]
    Never(std::convert::Infallible, std::marker::PhantomData<&'a ()>),
}

impl<'a> Iterator for Cells<'a> {
    type Item = Result<Cell<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        let cell: Result<CellInner<'a>> = match &mut self.0 {
            #[cfg(feature = "xlsx")]
            CellsInner::Xlsx(cells) => cells.next()?.map(CellInner::Xlsx),
            #[cfg(feature = "xlsb")]
            CellsInner::XlsbSource(cells) => cells.next()?.map(CellInner::XlsbSource),
            #[cfg(feature = "xlsb")]
            CellsInner::XlsbEager(cells) => cells
                .next()?
                .map(|cell| CellInner::XlsbEager(Stored::new(cell)))
                .map_err(Into::into),
            #[cfg(feature = "xls")]
            CellsInner::Xls(cells) => cells
                .next()?
                .map(|cell| CellInner::Xls(Stored::new(cell)))
                .map_err(Into::into),
            CellsInner::Never(never, _) => match *never {},
        };
        Some(cell.map(Cell))
    }
}

/// The rows of a [`Worksheet`]: the iterator
/// [`Worksheet::rows`](CoreWorksheet::rows) returns.
pub struct Rows<'a>(RowsInner<'a>);

enum RowsInner<'a> {
    #[cfg(feature = "xlsx")]
    Xlsx(xlsx::Rows<'a>),
    #[cfg(feature = "xlsb")]
    XlsbSource(source::Rows<'a, 'a>),
    #[cfg(feature = "xlsb")]
    XlsbEager(litchi_xlsb::sheet::Rows<'a>),
    #[cfg(feature = "xls")]
    Xls(litchi_xls::worksheet::Rows<'a>),
    #[allow(
        dead_code,
        reason = "uninhabited: carries the lifetime when no borrowing format is enabled"
    )]
    Never(std::convert::Infallible, std::marker::PhantomData<&'a ()>),
}

impl<'a> Iterator for Rows<'a> {
    type Item = Result<Cow<'a, [CellValue]>>;

    #[allow(
        clippy::useless_conversion,
        reason = "the XLSX arm's error is already the unified error"
    )]
    fn next(&mut self) -> Option<Self::Item> {
        each!(RowsInner, &mut self.0, rows => rows.next().map(|row| row.map_err(Into::into)), RowsInner::Never(never, _) => match *never {})
    }
}

impl CoreWorksheet for Worksheet<'_> {
    type Error = crate::Error;
    type Cell<'b>
        = Cell<'b>
    where
        Self: 'b;
    type Cells<'b>
        = Cells<'b>
    where
        Self: 'b;
    type Rows<'b>
        = Rows<'b>
    where
        Self: 'b;

    #[inline]
    fn cell(&self, row: u32, column: u32) -> Result<Cell<'_>> {
        let cell = match &self.0 {
            #[cfg(feature = "xlsx")]
            SheetInner::Xlsx(sheet) => CellInner::Xlsx(sheet.cell(row, column)?),
            #[cfg(feature = "xlsb")]
            SheetInner::XlsbSource(sheet) => CellInner::XlsbSource(sheet.cell(row, column)?),
            #[cfg(feature = "xlsb")]
            SheetInner::XlsbEager(sheet) => {
                CellInner::XlsbEager(Stored::at(sheet.get_cell(row, column), row, column))
            },
            #[cfg(feature = "xls")]
            SheetInner::Xls(sheet) => {
                CellInner::Xls(Stored::at(sheet.get_cell(row, column), row, column))
            },
            SheetInner::Never(never, _) => match *never {},
        };
        Ok(Cell(cell))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Cell<'_>> {
        let cell = match &self.0 {
            #[cfg(feature = "xlsx")]
            SheetInner::Xlsx(sheet) => CellInner::Xlsx(sheet.cell_by_coordinate(coordinate)?),
            #[cfg(feature = "xlsb")]
            SheetInner::XlsbSource(sheet) => {
                CellInner::XlsbSource(sheet.cell_by_coordinate(coordinate)?)
            },
            #[cfg(feature = "xlsb")]
            SheetInner::XlsbEager(sheet) => {
                CellInner::XlsbEager(Stored::new(sheet.cell_by_coordinate(coordinate)?))
            },
            #[cfg(feature = "xls")]
            SheetInner::Xls(sheet) => {
                CellInner::Xls(Stored::new(sheet.cell_by_coordinate(coordinate)?))
            },
            SheetInner::Never(never, _) => match *never {},
        };
        Ok(Cell(cell))
    }

    #[allow(
        clippy::useless_conversion,
        reason = "the XLSX arm's error is already the unified error"
    )]
    #[inline]
    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        each!(SheetInner, &self.0, sheet => sheet.cell_value(row, column).map_err(Into::into), SheetInner::Never(never, _) => match *never {})
    }

    fn cells(&self) -> Cells<'_> {
        Cells(match &self.0 {
            #[cfg(feature = "xlsx")]
            SheetInner::Xlsx(sheet) => CellsInner::Xlsx(sheet.cells()),
            #[cfg(feature = "xlsb")]
            SheetInner::XlsbSource(sheet) => CellsInner::XlsbSource(sheet.cells()),
            #[cfg(feature = "xlsb")]
            SheetInner::XlsbEager(sheet) => CellsInner::XlsbEager(sheet.cells()),
            #[cfg(feature = "xls")]
            SheetInner::Xls(sheet) => CellsInner::Xls(sheet.cells()),
            SheetInner::Never(never, _) => match *never {},
        })
    }

    fn column_count(&self) -> usize {
        each!(SheetInner, &self.0, sheet => sheet.column_count(), SheetInner::Never(never, _) => match *never {})
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        each!(SheetInner, &self.0, sheet => sheet.dimensions(), SheetInner::Never(never, _) => match *never {})
    }

    fn name(&self) -> &str {
        each!(SheetInner, &self.0, sheet => sheet.name(), SheetInner::Never(never, _) => match *never {})
    }

    #[allow(
        clippy::useless_conversion,
        reason = "the XLSX arm's error is already the unified error"
    )]
    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        each!(SheetInner, &self.0, sheet => sheet.row(row_idx).map_err(Into::into), SheetInner::Never(never, _) => match *never {})
    }

    fn row_count(&self) -> usize {
        each!(SheetInner, &self.0, sheet => sheet.row_count(), SheetInner::Never(never, _) => match *never {})
    }

    fn rows(&self) -> Rows<'_> {
        Rows(match &self.0 {
            #[cfg(feature = "xlsx")]
            SheetInner::Xlsx(sheet) => RowsInner::Xlsx(sheet.rows()),
            #[cfg(feature = "xlsb")]
            SheetInner::XlsbSource(sheet) => RowsInner::XlsbSource(sheet.rows()),
            #[cfg(feature = "xlsb")]
            SheetInner::XlsbEager(sheet) => RowsInner::XlsbEager(sheet.rows()),
            #[cfg(feature = "xls")]
            SheetInner::Xls(sheet) => RowsInner::Xls(sheet.rows()),
            SheetInner::Never(never, _) => match *never {},
        })
    }
}

/// The worksheets of a workbook, in workbook order: the iterator
/// [`WorkbookTrait::worksheets`](litchi_core::sheet::WorkbookTrait::worksheets)
/// returns.
pub struct Worksheets<'a>(WorksheetsInner<'a>);

pub(crate) enum WorksheetsInner<'a> {
    #[cfg(feature = "xlsx")]
    Xlsx(crate::format::xlsx::Worksheets<'a>),
    #[cfg(feature = "xlsb")]
    XlsbSource(source::SourceWorksheets<'a>),
    #[cfg(feature = "xlsb")]
    XlsbEager(litchi_xlsb::workbook::Worksheets<'a>),
    #[cfg(feature = "xls")]
    Xls(crate::format::xls::Worksheets<'a>),
    /// Yields one error, then ends.
    #[cfg(any(feature = "ods", feature = "xlsb"))]
    Refused(Option<crate::Error>),
    #[allow(
        dead_code,
        reason = "uninhabited: carries the lifetime when no borrowing format is enabled"
    )]
    Never(std::convert::Infallible, std::marker::PhantomData<&'a ()>),
}

impl<'a> Worksheets<'a> {
    pub(crate) const fn new(inner: WorksheetsInner<'a>) -> Self {
        Self(inner)
    }
}

impl<'a> Iterator for Worksheets<'a> {
    type Item = Result<Worksheet<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        let sheet: Result<SheetInner<'a>> = match &mut self.0 {
            #[cfg(feature = "xlsx")]
            WorksheetsInner::Xlsx(sheets) => sheets.next()?.map(SheetInner::Xlsx),
            #[cfg(feature = "xlsb")]
            WorksheetsInner::XlsbSource(sheets) => sheets.next()?.map(SheetInner::XlsbSource),
            #[cfg(feature = "xlsb")]
            WorksheetsInner::XlsbEager(sheets) => sheets
                .next()?
                .map(SheetInner::XlsbEager)
                .map_err(Into::into),
            #[cfg(feature = "xls")]
            WorksheetsInner::Xls(sheets) => sheets.next()?.map(SheetInner::Xls),
            #[cfg(any(feature = "ods", feature = "xlsb"))]
            WorksheetsInner::Refused(error) => return error.take().map(Err),
            WorksheetsInner::Never(never, _) => match *never {},
        };
        Some(sheet.map(Worksheet))
    }
}
