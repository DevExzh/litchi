//! Layered BIFF8 worksheet `INDEX`/`DBCELL` row-block owner.
//!
//! The semantic row-block model is kept separate from the record codecs and
//! workbook-stream collector. This module remains the crate-internal owner
//! and continues to provide the same public XLS facade.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub(crate) const INDEX_RECORD_TYPE: u16 = 0x020B;
pub(crate) const DBCELL_RECORD_TYPE: u16 = 0x00D7;

const BOF_RECORD_TYPE: u16 = 0x0809;
const EOF_RECORD_TYPE: u16 = 0x000A;
const ROW_RECORD_TYPE: u16 = 0x0208;
const DEF_COL_WIDTH_RECORD_TYPE: u16 = 0x0055;
const INDEX_FIXED_LEN: usize = 16;
const MAX_ROW_BLOCKS: usize = 2_048;
const MAX_ROWS_PER_BLOCK: usize = 32;
const MAX_INDEX_PAYLOAD_LEN: usize = INDEX_FIXED_LEN + MAX_ROW_BLOCKS * 4;
const MAX_DBCELL_PAYLOAD_LEN: usize = 4 + MAX_ROWS_PER_BLOCK * 2;

pub(crate) use codec::RowBlockIndexCollector;

pub use model::{DbCellRecord, IndexedRow, RowBlock, RowBlockIndex, WorksheetIndexRecord};
