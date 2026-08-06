//! Typed owner for the workbook-global `CompressPictures` record
//! ([MS-XLS] 2.4.55).
//!
//! The semantic payload is bounded to a future-record header, one Boolean,
//! and a retained opaque tail. A detached [`Snapshot`] can retain unrelated
//! BIFF records and apply typed changes through an atomic [`Transaction`].

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

/// BIFF record type for `CompressPictures`.
pub const RECORD_TYPE: u16 = 0x089B;

pub use codec::{parse, write};
pub use model::{Record, Settings, Snapshot, Unknown};
pub use transaction::{Patch, Transaction};
