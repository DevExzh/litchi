use crate::{Error, Result};

use super::invalid;
use super::model::{DataTable, DataTableInputCell, DataTableKind, DataTableRange};

/// `PtgTbl` token identifier (MS-XLS 2.5.198.92).
pub(super) const PTG_TBL: u8 = 0x02;
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

impl DataTableInputCell {
    fn decode(row: u16, col: u16, deleted: bool) -> Result<Self> {
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

impl DataTable {
    /// Parse a `Table` record payload.
    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(Error::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let read_u16 = |offset: usize| u16::from_le_bytes([data[offset], data[offset + 1]]);
        let range = DataTableRange::new(read_u16(0), read_u16(2), data[4], data[5])?;
        let flags = read_u16(6);
        let always_calc = flags & ALWAYS_CALC != 0;
        let row_orientation = flags & ROW_ORIENTATION != 0;
        let two_variable = flags & TWO_VARIABLE != 0;
        let input1 = DataTableInputCell::decode(read_u16(8), read_u16(10), flags & DELETED1 != 0)?;
        let kind = if two_variable {
            DataTableKind::TwoVariable {
                row_input: input1,
                column_input: DataTableInputCell::decode(
                    read_u16(12),
                    read_u16(14),
                    flags & DELETED2 != 0,
                )?,
            }
        } else {
            DataTableKind::OneVariable {
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
            DataTableKind::OneVariable {
                input,
                ignored_coordinates,
                ignored_deleted2,
            } => (*input, *ignored_coordinates, *ignored_deleted2),
            DataTableKind::TwoVariable {
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
