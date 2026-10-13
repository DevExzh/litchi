//! Typed, inert ownership of the BIFF8 `User Names` CFB stream.
//!
//! The owner implements the MS-XLS 2.1.7.17 sequence
//! `CUsr UsrChk CbUsr BCUsrs *UsrInfo`.  Opening a workbook does not read this
//! stream; [`crate::raw::Reader::user_names`] reads it on demand and,
//! when the `Revision Log` exists, proves every `UsrInfo.guid` dependency.
//! Editing never acquires shared-workbook locks, merges revisions, or executes
//! any collaboration behavior.

mod codec;
mod edit;
mod model;
mod package;
pub mod publication;

pub use edit::{Commit, Edit, Patch, Snapshot};
pub use model::{Limits, UserCheck, UserEntry, UserGuid, UserNames};
pub use publication::{
    Commit as PackageCommit, Edit as PackageEdit, Patch as PackagePatch,
    Snapshot as PackageSnapshot,
};

/// Exact root-storage stream name required by MS-XLS 2.1.7.17.
pub const USER_NAMES_STREAM_NAME: &str = "User Names";

/// Alias emphasizing the stream owner in APIs that expose several snapshots.
pub type UserNamesSnapshot = Snapshot;
/// Alias emphasizing the stream owner in APIs that expose several patches.
pub type UserNamesPatch = Patch;

#[cfg(test)]
mod tests;
