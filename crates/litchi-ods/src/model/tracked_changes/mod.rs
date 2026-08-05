//! Semantic owner for ODS spreadsheet tracked changes.
//!
//! The historical `crate::model::tracked_changes` path remains the facade.
//! Typed values and validation live in `model`, namespace-aware XML parsing
//! and serialization live in `codec`, and retained regression coverage lives
//! in `tests`.

mod codec;
mod model;

#[cfg(test)]
mod tests;

use litchi_core::{Error, Result};

const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 1_000_000;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

pub(super) fn append_size(aggregate: &mut usize, amount: usize) -> Result<()> {
    *aggregate = aggregate.checked_add(amount).ok_or_else(|| {
        Error::InvalidFormat("tracked-change aggregate size overflow".to_string())
    })?;
    if *aggregate > MAX_AGGREGATE_BYTES {
        return Err(Error::InvalidFormat(
            "tracked-change metadata exceeds 16 MiB".to_string(),
        ));
    }
    Ok(())
}

pub use model::{
    Acceptance, Cell, CellAddress, CellValue, Change, Changes, ContentChange, CutOff, Deletion,
    Dimension, Info, Insertion, Metadata, Movement, NestedDeletion, RangeAddress,
};

#[allow(
    unused_imports,
    reason = "codec entry points retain the historical crate-internal module path"
)]
pub(crate) use codec::{parse_tracked_changes, write_tracked_changes};
