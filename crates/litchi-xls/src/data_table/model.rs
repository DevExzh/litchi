use crate::{XlsError, XlsResult};

use super::invalid;

/// The cell range covered by a data table (BIFF8 `Ref`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDataTableRange {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) first_col: u8,
    pub(super) last_col: u8,
}

impl XlsDataTableRange {
    /// A range; both the first row and the first column are 1-based per the
    /// `Table` record constraints.
    pub fn new(first_row: u16, last_row: u16, first_col: u8, last_col: u8) -> XlsResult<Self> {
        if first_row == 0 || first_col == 0 {
            return Err(invalid("data-table range origin is 1-based"));
        }
        if last_row < first_row || last_col < first_col {
            return Err(invalid("data-table range is reversed"));
        }
        Ok(Self {
            first_row,
            last_row,
            first_col,
            last_col,
        })
    }

    pub const fn first_row(&self) -> u16 {
        self.first_row
    }
    pub const fn last_row(&self) -> u16 {
        self.last_row
    }
    pub const fn first_col(&self) -> u8 {
        self.first_col
    }
    pub const fn last_col(&self) -> u8 {
        self.last_col
    }
}

/// An input cell of a data table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataTableInputCell {
    /// The input cell reference.
    Present {
        /// Row of the input cell.
        row: u16,
        /// Column of the input cell.
        col: u8,
    },
    /// The referenced input cell has been deleted (`fDeleted1`/`fDeleted2`).
    Deleted,
}

impl XlsDataTableInputCell {
    /// Create a present input-cell reference from zero-based raw indices.
    pub fn present(row: u32, col: u16) -> XlsResult<Self> {
        let invalid = || {
            XlsError::InvalidCellReference(format!(
                "data-table input row {row}, column {col} is outside the BIFF8 grid"
            ))
        };
        let row = u16::try_from(row).map_err(|_| invalid())?;
        let col = u8::try_from(col).map_err(|_| invalid())?;
        Ok(Self::Present { row, col })
    }
}

/// The shape of a data table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataTableKind {
    /// One-variable table (`fTbl2` = 0): a single input cell, interpreted as
    /// a row input cell when `row_orientation` is set and as a column input
    /// cell otherwise.
    OneVariable {
        /// The single input cell.
        input: XlsDataTableInputCell,
        /// Raw `(rwInpCol, colInpCol)` pair; undefined for one-variable
        /// tables and preserved verbatim for round-trips.
        ignored_coordinates: (u16, u16),
        /// Raw `fDeleted2` bit; undefined for one-variable tables and
        /// preserved verbatim for round-trips.
        ignored_deleted2: bool,
    },
    /// Two-variable table (`fTbl2` = 1): row and column input cells.
    TwoVariable {
        /// Row input cell.
        row_input: XlsDataTableInputCell,
        /// Column input cell.
        column_input: XlsDataTableInputCell,
    },
}

/// Typed `Table` record content (MS-XLS 2.4.319).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDataTable {
    pub(super) range: XlsDataTableRange,
    pub(super) always_calc: bool,
    /// `fRw`; the one-variable orientation, preserved verbatim for
    /// two-variable tables.
    pub(super) row_orientation: bool,
    pub(super) kind: XlsDataTableKind,
}

impl XlsDataTable {
    /// A one-variable data table.
    pub fn one_variable(
        range: XlsDataTableRange,
        row_orientation: bool,
        input: XlsDataTableInputCell,
    ) -> Self {
        Self {
            range,
            always_calc: false,
            row_orientation,
            kind: XlsDataTableKind::OneVariable {
                input,
                ignored_coordinates: (0, 0),
                ignored_deleted2: false,
            },
        }
    }

    /// A two-variable data table.
    pub fn two_variable(
        range: XlsDataTableRange,
        row_input: XlsDataTableInputCell,
        column_input: XlsDataTableInputCell,
    ) -> Self {
        Self {
            range,
            always_calc: false,
            row_orientation: false,
            kind: XlsDataTableKind::TwoVariable {
                row_input,
                column_input,
            },
        }
    }

    pub const fn range(&self) -> XlsDataTableRange {
        self.range
    }
    pub const fn always_calc(&self) -> bool {
        self.always_calc
    }
    pub fn set_always_calc(&mut self, always_calc: bool) {
        self.always_calc = always_calc;
    }
    /// Whether the single input cell of a one-variable table is a row input
    /// cell (`fRw`); preserved verbatim for two-variable tables.
    pub const fn row_orientation(&self) -> bool {
        self.row_orientation
    }
    pub const fn is_two_variable(&self) -> bool {
        matches!(self.kind, XlsDataTableKind::TwoVariable { .. })
    }
    pub const fn kind(&self) -> &XlsDataTableKind {
        &self.kind
    }
}
