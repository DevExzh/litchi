//! Word 2012 paragraph-collapse metadata.
//!
//! The extension is intentionally owned by a focused paragraph module. The
//! semantic value is a closed enum, the XML reader/writer is bounded and
//! loss-preserving. Detached paragraph views change it through
//! [`crate::Paragraph::set_collapsed`]; [`Snapshot`] reads it.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::Collapsed;
pub use transaction::Snapshot;

pub(crate) use codec::append_xml;
