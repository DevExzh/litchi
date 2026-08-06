//! Transactional edits for standalone Microsoft Graph chart packages.
//!
//! The transaction owns the validated OLE2 package state and only exposes the
//! operation that can be proven at this crate boundary: replacing the one
//! framed Graph chart substream while retaining the workbook prelude and OLE
//! host streams. It is deliberately separate from PPT presentation editing;
//! `[MS-PPT]` external-object and `[MS-ODRAW]` frame records are not changed by
//! this package-level operation.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{PackageEditor, Snapshot};
