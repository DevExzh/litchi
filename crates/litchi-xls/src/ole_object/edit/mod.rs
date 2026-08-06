//! Detached, source-checked OLE-object and form-control publication.
//!
//! The package layer owns BIFF/CFB rewriting. This layer gives callers an
//! immutable typed view, a failure-atomic transaction, and a reversible patch
//! without activating any embedded object or control payload.

mod model;
mod transaction;

pub use model::{Commit, Patch, Snapshot};
pub use transaction::Transaction;
