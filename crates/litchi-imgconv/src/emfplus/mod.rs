//! Safe, zero-copy framing for Enhanced Metafile Format Plus (EMF+).
//!
//! This module deliberately stops at record boundaries.  Record-specific decoding
//! and rendering can be layered on top without duplicating the security-sensitive
//! size, offset, and continuation checks implemented here.

#![allow(
    clippy::missing_errors_doc,
    reason = "all fallible public APIs reject malformed input or configured resource limits"
)]

mod comment;
pub mod objects;
mod parser;
pub mod playback;
pub mod renderer;
pub mod svg;
mod types;

pub use comment::{
    extract_emfplus_comment_body, extract_emfplus_comment_record, try_extract_emfplus_comment_body,
    try_extract_emfplus_comment_record,
};
pub use parser::{
    EmfPlusRecord, EmfPlusRecordIter, EmfPlusStreamValidator, ObjectFragment,
    validate_complete_stream,
};
pub use types::{
    EMFPLUS_COMMENT_IDENTIFIER, EMFPLUS_RECORD_HEADER_SIZE, EMR_COMMENT, MAX_EMFPLUS_OBJECT_SLOTS,
    ObjectId, ObjectRecordFlags, ObjectType, ParserLimits, RecordFlags, RecordHeader, RecordType,
};

#[cfg(test)]
mod corpus_tests;
#[cfg(test)]
mod tests;
