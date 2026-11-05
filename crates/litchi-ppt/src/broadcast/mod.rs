//! Inert `PowerPoint` 9 presentation-broadcast metadata from MS-PPT 2.4.17.

mod codec;
mod edit;
mod model;
mod package;
#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;

#[cfg(test)]
mod tests;

pub use edit::Edit;
#[allow(
    clippy::module_name_repetitions,
    reason = "`BroadcastProperties` is the established public API name mirroring MS-PPT terminology; renaming it would break downstream crates"
)]
pub use model::{Broadcast, BroadcastProperties, Broadcasts, UnknownRecord};
pub use transaction::{Revision, Snapshot};

crate::scoped::family!("broadcast");
