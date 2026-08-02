//! BIFF8 `Table` record (MS-XLS 2.4.319): one- and two-variable what-if data
//! tables.
//!
//! A `Table` record follows the `Formula` record whose token stream begins
//! with `PtgTbl` (MS-XLS 2.5.198.92); the token names the first row and
//! column of the table range.

use super::{XlsError, XlsResult};

/// Record type of the `Table` record.
pub(crate) const TABLE_RECORD_TYPE: u16 = 0x0236;
/// `PtgTbl` token identifier (MS-XLS 2.5.198.92).
pub(crate) const PTG_TBL: u8 = 0x02;
/// Serialized size of a `Table` record payload.
const PAYLOAD_LEN: usize = 16;

// Flag bits of the 2-byte bitfield at offset 6.
const ALWAYS_CALC: u16 = 0x0001;
const ROW_ORIENTATION: u16 = 0x0004;
const TWO_VARIABLE: u16 = 0x0008;
const DELETED1: u16 = 0x0010;
const DELETED2: u16 = 0x0020;
/// Coordinate pair marking a deleted input cell.
const DELETED_COORDINATE: u16 = 0xFFFF;

fn invalid(message: &str) -> XlsError {
    XlsError::InvalidRecord {
        record_type: TABLE_RECORD_TYPE,
        message: message.to_string(),
    }
}

/// The cell range covered by a data table (BIFF8 `Ref`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDataTableRange {
    first_row: u16,
    last_row: u16,
    first_col: u8,
    last_col: u8,
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

    fn decode(row: u16, col: u16, deleted: bool) -> XlsResult<Self> {
        if deleted {
            if row != DELETED_COORDINATE || col != DELETED_COORDINATE {
                return Err(invalid("deleted input cell must carry the -1 coordinates"));
            }
            return Ok(Self::Deleted);
        }
        let col =
            u8::try_from(col).map_err(|_| invalid("input cell column exceeds the BIFF8 grid"))?;
        Ok(Self::Present { row, col })
    }

    fn encode(self) -> (u16, u16, bool) {
        match self {
            Self::Present { row, col } => (row, u16::from(col), false),
            Self::Deleted => (DELETED_COORDINATE, DELETED_COORDINATE, true),
        }
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
    range: XlsDataTableRange,
    always_calc: bool,
    /// `fRw`; the one-variable orientation, preserved verbatim for
    /// two-variable tables.
    row_orientation: bool,
    kind: XlsDataTableKind,
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

    /// Parse a `Table` record payload.
    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let read_u16 = |offset: usize| u16::from_le_bytes([data[offset], data[offset + 1]]);
        let range = XlsDataTableRange::new(read_u16(0), read_u16(2), data[4], data[5])?;
        let flags = read_u16(6);
        let always_calc = flags & ALWAYS_CALC != 0;
        let row_orientation = flags & ROW_ORIENTATION != 0;
        let two_variable = flags & TWO_VARIABLE != 0;
        let input1 =
            XlsDataTableInputCell::decode(read_u16(8), read_u16(10), flags & DELETED1 != 0)?;
        let kind = if two_variable {
            XlsDataTableKind::TwoVariable {
                row_input: input1,
                column_input: XlsDataTableInputCell::decode(
                    read_u16(12),
                    read_u16(14),
                    flags & DELETED2 != 0,
                )?,
            }
        } else {
            XlsDataTableKind::OneVariable {
                input: input1,
                ignored_coordinates: (read_u16(12), read_u16(14)),
                ignored_deleted2: flags & DELETED2 != 0,
            }
        };
        Ok(Self {
            range,
            always_calc,
            row_orientation,
            kind,
        })
    }

    /// Serialize back to a complete `Table` record payload.
    pub(crate) fn to_payload(self) -> Vec<u8> {
        let mut flags = 0u16;
        if self.always_calc {
            flags |= ALWAYS_CALC;
        }
        if self.row_orientation {
            flags |= ROW_ORIENTATION;
        }
        let (input1, input2, deleted2) = match &self.kind {
            XlsDataTableKind::OneVariable {
                input,
                ignored_coordinates,
                ignored_deleted2,
            } => (*input, *ignored_coordinates, *ignored_deleted2),
            XlsDataTableKind::TwoVariable {
                row_input,
                column_input,
            } => {
                let (row, col, deleted) = column_input.encode();
                (*row_input, (row, col), deleted)
            },
        };
        if self.is_two_variable() {
            flags |= TWO_VARIABLE;
        }
        let (row1, col1, deleted1) = input1.encode();
        if deleted1 {
            flags |= DELETED1;
        }
        if deleted2 {
            flags |= DELETED2;
        }
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&self.range.first_row.to_le_bytes());
        payload.extend_from_slice(&self.range.last_row.to_le_bytes());
        payload.push(self.range.first_col);
        payload.push(self.range.last_col);
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&row1.to_le_bytes());
        payload.extend_from_slice(&col1.to_le_bytes());
        payload.extend_from_slice(&input2.0.to_le_bytes());
        payload.extend_from_slice(&input2.1.to_le_bytes());
        payload
    }

    /// The `PtgTbl` token stream of the associated table formula.
    pub(crate) fn ptg_tbl_tokens(&self) -> [u8; 5] {
        let mut tokens = [0; 5];
        tokens[0] = PTG_TBL;
        tokens[1..3].copy_from_slice(&self.range.first_row.to_le_bytes());
        tokens[3..5].copy_from_slice(&u16::from(self.range.first_col).to_le_bytes());
        tokens
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn range() -> XlsDataTableRange {
        XlsDataTableRange::new(2, 8, 3, 5).unwrap()
    }

    #[test]
    fn one_variable_round_trips() {
        let table = XlsDataTable::one_variable(
            range(),
            true,
            XlsDataTableInputCell::Present { row: 0, col: 6 },
        );
        let parsed = XlsDataTable::parse(&table.to_payload()).unwrap();
        assert_eq!(parsed, table);
        assert!(!parsed.is_two_variable());
        assert!(parsed.row_orientation());
        let XlsDataTableKind::OneVariable { input, .. } = parsed.kind() else {
            panic!()
        };
        assert_eq!(*input, XlsDataTableInputCell::Present { row: 0, col: 6 });
    }

    #[test]
    fn two_variable_with_deleted_input_round_trips() {
        let mut table = XlsDataTable::two_variable(
            range(),
            XlsDataTableInputCell::Present { row: 1, col: 2 },
            XlsDataTableInputCell::Deleted,
        );
        table.set_always_calc(true);
        let parsed = XlsDataTable::parse(&table.to_payload()).unwrap();
        assert_eq!(parsed, table);
        assert!(parsed.is_two_variable());
        assert!(parsed.always_calc());
    }

    #[test]
    fn one_variable_preserves_undefined_tail() {
        let mut payload = XlsDataTable::one_variable(
            range(),
            false,
            XlsDataTableInputCell::Present { row: 1, col: 2 },
        )
        .to_payload();
        // Scribble the undefined rwInpCol/colInpCol pair and fDeleted2.
        payload[6] |= 0x20;
        payload[12..14].copy_from_slice(&7u16.to_le_bytes());
        payload[14..16].copy_from_slice(&9u16.to_le_bytes());
        let parsed = XlsDataTable::parse(&payload).unwrap();
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn ptg_tbl_tokens_name_the_range_origin() {
        let table = XlsDataTable::one_variable(range(), false, XlsDataTableInputCell::Deleted);
        assert_eq!(table.ptg_tbl_tokens(), [PTG_TBL, 2, 0, 3, 0]);
    }

    #[test]
    fn rejects_malformed_records() {
        assert!(XlsDataTable::parse(&[0; 15]).is_err());
        assert!(XlsDataTable::parse(&[0; 17]).is_err());
        // Zero-based origin.
        let mut payload = XlsDataTable::one_variable(
            range(),
            false,
            XlsDataTableInputCell::Present { row: 1, col: 2 },
        )
        .to_payload();
        payload[0..2].copy_from_slice(&0u16.to_le_bytes());
        assert!(XlsDataTable::parse(&payload).is_err());
        // Deleted input without the -1 coordinates.
        let mut payload = XlsDataTable::one_variable(
            range(),
            false,
            XlsDataTableInputCell::Present { row: 1, col: 2 },
        )
        .to_payload();
        payload[6] |= 0x10;
        assert!(XlsDataTable::parse(&payload).is_err());
        // A present input column outside the BIFF8 cell grid.
        let mut payload = XlsDataTable::one_variable(
            range(),
            false,
            XlsDataTableInputCell::Present { row: 1, col: 2 },
        )
        .to_payload();
        payload[10..12].copy_from_slice(&256u16.to_le_bytes());
        assert!(XlsDataTable::parse(&payload).is_err());
        // Reversed or zero-based range.
        assert!(XlsDataTableRange::new(5, 2, 3, 3).is_err());
        assert!(XlsDataTableRange::new(0, 2, 3, 3).is_err());
        assert!(XlsDataTableRange::new(1, 2, 0, 3).is_err());
    }
}
