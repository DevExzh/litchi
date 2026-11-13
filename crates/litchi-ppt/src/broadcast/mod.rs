//! Inert `PowerPoint` 9 presentation-broadcast metadata from MS-PPT 2.4.17.

mod codec;
mod edit;
mod model;
mod package;
#[cfg_attr(
    not(test),
    expect(
        dead_code,
        reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests"
    )
)]
mod transaction;

#[cfg(test)]
mod tests;

pub use edit::Edit;
pub use model::{Broadcast, BroadcastProperties, Broadcasts, UnknownRecord};
pub use transaction::{Revision, Snapshot};

crate::scoped::family!("broadcast");
