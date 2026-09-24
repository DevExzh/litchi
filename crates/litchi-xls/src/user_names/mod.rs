//! Typed, inert ownership of the BIFF8 `User Names` CFB stream.
//!
//! The owner implements the MS-XLS 2.1.7.17 sequence
//! `CUsr UsrChk CbUsr BCUsrs *UsrInfo`.  Opening a workbook does not read this
//! stream; [`crate::workbook::Workbook::user_names`] reads it on demand and,
//! when the `Revision Log` exists, proves every `UsrInfo.guid` dependency.
//! Editing never acquires shared-workbook locks, merges revisions, or executes
//! any collaboration behavior.

mod codec;
mod edit;
mod model;
mod package;
pub mod publication;

pub use edit::{Commit, Patch, Snapshot, Transaction};
pub use model::{Limits, UserCheck, UserEntry, UserGuid, UserNames};
pub use publication::{
    Commit as PackageCommit, Patch as PackagePatch, Snapshot as PackageSnapshot,
    Transaction as PackageTransaction,
};

/// Exact root-storage stream name required by MS-XLS 2.1.7.17.
pub const USER_NAMES_STREAM_NAME: &str = "User Names";

/// Alias emphasizing the stream owner in APIs that expose several snapshots.
pub type UserNamesSnapshot = Snapshot;
/// Alias emphasizing the stream owner in APIs that expose several transactions.
pub type UserNamesTransaction = Transaction;
/// Alias emphasizing the stream owner in APIs that expose several patches.
pub type UserNamesPatch = Patch;

#[cfg(test)]
mod tests;
