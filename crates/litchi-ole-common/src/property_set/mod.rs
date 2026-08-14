//! Shared OLE Property Set metadata grammar and transactional CFB editing.
//!
//! The semantic model follows [MS-OLEPS] and the shared property-set rules in
//! [MS-OSHARED]. Parsing and serialization are lossless within the supported
//! typed grammar, while [`Editor`] stages complete CFB rewrites and publishes
//! them only after every requested mutation has validated successfully.
//!
//! [`PropertySetReader`] is an extension trait because the CFB container owns
//! [`litchi_cfb::OleFile`]. [`SharedPropertySetReader`] is its immutable
//! positional counterpart for [`litchi_cfb::SharedOleFile`]. Import the
//! applicable trait to add property-set readers to an opened compound file.

mod binding;
mod codec;
pub mod document_summary;
mod model;
pub mod summary_information;
pub mod user_defined;

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod tests;

pub use binding::Binding as Standard;
pub use binding::{
    Binding, BindingName, GLOBAL_INFO_FMTID, IMAGE_CONTENTS_FMTID, IMAGE_INFO_FMTID,
};
#[allow(
    clippy::module_name_repetitions,
    reason = "the public extension-trait name distinguishes property-set readers from other readers"
)]
pub use codec::{Editor, PropertySetReader, SharedPropertySetReader};
pub use model::{
    Array, CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Dimension, DocParts, Guid, HeadingPair,
    HeadingPairs, Metadata, PID_DOC_PARTS, PID_HEADING_PAIRS, SUMMARY_INFORMATION_FMTID, Scalar,
    Section, Stream, TextEncoding, USER_DEFINED_PROPERTIES_FMTID, Value, Vector, VersionedStream,
};
#[cfg(test)]
pub(crate) use model::{DEFAULT_CODEPAGE, PID_BEHAVIOR, PID_CODEPAGE};
