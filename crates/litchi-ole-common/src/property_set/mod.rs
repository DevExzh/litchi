//! Shared OLE Property Set metadata grammar and transactional CFB editing.
//!
//! The semantic model follows [MS-OLEPS] and the shared property-set rules in
//! [MS-OSHARED]. Parsing and serialization are lossless within the supported
//! typed grammar, while [`Editor`] stages complete CFB rewrites and publishes
//! them only after every requested mutation has validated successfully.
//!
//! [`PropertySetReader`] is an extension trait because the CFB container owns
//! [`litchi_cfb::OleFile`]. Import the trait to add property-set readers to an
//! opened compound file.

mod binding;
mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use binding::Binding as Standard;
pub use binding::{
    Binding, BindingName, GLOBAL_INFO_FMTID, IMAGE_CONTENTS_FMTID, IMAGE_INFO_FMTID,
};
pub use codec::{Editor, PropertySetReader};
pub use model::{
    Array, CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Dimension, DocParts, Guid, HeadingPair,
    HeadingPairs, Metadata, PID_DOC_PARTS, PID_HEADING_PAIRS, SUMMARY_INFORMATION_FMTID, Scalar,
    Section, Stream, TextEncoding, USER_DEFINED_PROPERTIES_FMTID, Value, Vector, VersionedStream,
};
#[cfg(test)]
pub(crate) use model::{DEFAULT_CODEPAGE, PID_BEHAVIOR, PID_CODEPAGE};
