//! Semantic owner for ODS spreadsheet tracked changes.
//!
//! The historical `crate::model::tracked_changes` path remains the facade.
//! Typed values and validation live in `model`, namespace-aware XML parsing
//! and serialization live in `codec`, and retained regression coverage lives
//! in `tests`.

mod codec;
mod limits;
mod model;
mod transaction;

#[cfg(test)]
mod tests;

const MAX_DEPTH: usize = 256;
const MAX_VALUE_BYTES: usize = 65_536;

pub use limits::Limits;
pub use model::{
    Acceptance, Cell, CellAddress, CellValue, Change, Changes, ContentChange, CutOff, Deletion,
    Dimension, Info, Insertion, Integer, Metadata, Movement, NestedDeletion, PositiveInteger,
    RangeAddress,
};
pub use transaction::{Commit, Patch, Snapshot, Transaction, update};

#[allow(
    clippy::module_name_repetitions,
    reason = "the codec entry points keep their historical element-qualified names"
)]
pub use codec::{parse_tracked_changes, write_tracked_changes};
