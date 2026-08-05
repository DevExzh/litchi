//! Contextual OLE Property Set metadata facade.
//!
//! The model owns the typed, prefix-free semantic values and edits; the codec
//! owns MS-OLEPS wire parsing, serialization, and CFB integration. Regression
//! coverage remains in the sibling tests module.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::Editor;
pub use model::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Guid, Metadata, SUMMARY_INFORMATION_FMTID,
    Section, Standard, Stream, USER_DEFINED_PROPERTIES_FMTID, Value,
};
#[cfg(test)]
pub(crate) use model::{DEFAULT_CODEPAGE, PID_BEHAVIOR, PID_CODEPAGE};
