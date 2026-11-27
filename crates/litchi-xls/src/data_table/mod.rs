//! BIFF8 `Table` record (MS-XLS 2.4.319): one- and two-variable what-if data
//! tables.
//!
//! A `Table` record follows the `Formula` record whose token stream begins
//! with `PtgTbl` (MS-XLS 2.5.198.92); the token names the first row and
//! column of the table range.

use crate::Error;

mod codec;
mod model;
#[cfg(test)]
mod tests;

/// Record type of the `Table` record.
pub(crate) const TABLE_RECORD_TYPE: u16 = 0x0236;

fn invalid(message: &str) -> Error {
    Error::InvalidRecord {
        record_type: TABLE_RECORD_TYPE,
        message: message.to_string(),
    }
}

pub use model::{DataTable, DataTableInputCell, DataTableKind, DataTableRange};
