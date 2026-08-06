//! Layered, typed MS-DOC routing-slip ownership.
//!
//! The facade exposes byte-oriented ANSI fields so parsing never assumes a
//! UTF-8 code page or changes the bytes that Word stored. [`Metadata::parse`]
//! reads the optional table range addressed by FIB index 70; [`parse_bytes`]
//! and [`to_bytes`] operate on one complete `Metadata` payload.
//!
//! The binary model/codec remains below the package seam, while the
//! transaction and package modules own selector-first lifecycle edits over
//! immutable snapshots. Protection remains route policy metadata: it is never
//! conflated with document-level or range-level protection and no caller is
//! authenticated by this passive owner.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_bytes, to_bytes};
pub use model::{DeliveryOption, EditKind, Metadata, NarrowString, Protection, Recipient};
pub use package::{Commit as PackageCommit, Editor, Snapshot as PackageSnapshot};
pub use transaction::{
    Commit, Error as TransactionError, Patch, RecipientSelectionError, RecipientSelector, Snapshot,
    Transaction,
};

/// FIB index of `fcRouteSlip`/`lcbRouteSlip` in `FibRgFcLcb97`.
pub const FIB_INDEX_ROUTE_SLIP: usize = validation::ROUTE_SLIP_FIB_INDEX;
