//! Layered ownership of the inert MS-DOC `SttbfAtnBkmk` table.
//!
//! The table stores one fixed `ATNBE` record for each annotation bookmark.
//! This owner exposes only the opaque tag identity and never resolves comment
//! text, bookmark ranges, or external data. Package edits append a new table
//! payload and update only the FIB range; semantic no-ops return the original
//! CFB bytes unchanged.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_bytes, to_bytes};
pub use model::{Tag, TagId, Tags};
pub use package::{Commit as PackageCommit, Editor, Snapshot as PackageSnapshot};
pub use transaction::{Commit, Error as TransactionError, Patch, Snapshot, Transaction};

/// FIB index of `fcSttbfAtnBkmk`/`lcbSttbfAtnBkmk`.
pub const FIB_INDEX: usize = validation::FIB_INDEX;
