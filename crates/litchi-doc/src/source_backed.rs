//! Source-backed DOC editing (record 0359a F2, the 0355a §1 shape).
//!
//! ```text
//! source_backed::Package::open(Arc<dyn ReadAt>) / from_bytes / from_shared
//! Package::edit_paragraph(&self, Position) -> paragraph::Edit
//! paragraph::Edit::set_text(&mut self, ..)            // staging only
//! paragraph::Edit::commit(self) -> Result<paragraph::Commit>   // validates, writes nothing
//! paragraph::Commit = litchi_core::edit::Commit<Package, paragraph::Patch, paragraph::Diagnostics>
//! Package::apply_paragraph(&self, &paragraph::Patch) -> Result<paragraph::Commit>
//! Package::write_to(&self, impl Write) / save(path)   // the only output
//! ```
//!
//! A [`Package`] is an immutable positional DOC source admitted for safe
//! same-length Unicode main-story paragraph edits. It does not rebuild a FIB,
//! CLX or any property table: an edit is admitted only when one complete
//! paragraph (including its paragraph mark) is one uncompressed Unicode piece
//! and the replacement has exactly the same UTF-16 width.
//!
//! A commit plans the checked equal-length `WordDocument` splice, reads the
//! edited paragraph back from the composed candidate, and returns a new
//! [`Package`] that reads through the plan's composed source and holds the
//! plan: [`Package::write_to`] and [`Package::save`] stream exactly the
//! planned bytes, and nothing is materialized. An opened package writes its
//! source through a checked identity plan. Edits chain on the committed
//! package.

#![deny(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unwrap_used,
    reason = "the source-backed DOC owner uses checked conversions and propagated typed errors"
)]

mod owner;

pub use owner::{Error, Limits, OpenOptions, Package, Paragraph, Refusal, Result};

pub mod paragraph;
