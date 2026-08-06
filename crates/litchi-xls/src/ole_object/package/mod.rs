//! CFB/package and workbook transaction layers for XLS OLE objects.

mod transaction;

pub use transaction::Editor;
#[cfg(test)]
pub(super) use transaction::{read_workbook, targets_for_sheets};
