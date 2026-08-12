#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
//! Bounded `WordprocessingML` document-variable model and XML codec.
//!
//! The checked-in `[MS-OE376]` conformance material identifies `docVar` as
//! Part 4 Section 2.15.1.30 and records Word's limits for its `name` and
//! `val` attributes.  This module owns those package-neutral semantics. OPC
//! parts and markup-compatibility preprocessing remain in the host crate.

mod codec;
mod model;
mod transaction;

#[cfg(test)]
mod tests;

pub use codec::parse_variables;
pub use model::{
    MAX_DOCUMENT_VARIABLE_DEPTH, MAX_DOCUMENT_VARIABLE_NAME_CHARS,
    MAX_DOCUMENT_VARIABLE_VALUE_CHARS, MAX_DOCUMENT_VARIABLE_XML_BYTES, MAX_DOCUMENT_VARIABLES,
    Variables,
};
pub use transaction::{Commit, Patch, Snapshot, Transaction};

pub(crate) use codec::{SettingsDialect, ensure_source_backed_rewrite_safe, inspect_source_policy};
