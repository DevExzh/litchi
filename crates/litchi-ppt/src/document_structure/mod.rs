//! Typed terminal structure of a legacy `PowerPoint` document container.

mod codec;
mod model;
#[cfg_attr(
    not(test),
    expect(
        dead_code,
        reason = "the record-level owner the primary Package::edit drives; its full surface stays exercised by the unit tests"
    )
)]
mod transaction;
mod validation;

pub(crate) use codec::encode_record;
pub use model::{CustomTableStylesPlacement, DocumentStructure, Limits, Master, Slide};
pub(crate) use transaction::Transaction;
pub use transaction::{Revision, Snapshot};

#[cfg(test)]
mod tests;
