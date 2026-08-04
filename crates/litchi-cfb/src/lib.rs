//! Microsoft Compound File Binary (CFB / OLE2) container parser and writer.
//!
//! This crate provides the CFB substrate consumed by the legacy Office binary
//! format crates (`litchi-doc`, `litchi-ppt`, and `litchi-xls`) and by
//! encrypted OOXML package support.
//!
//! See `[MS-CFB]: Compound File Binary File Format` for the format spec.

#![allow(missing_docs)]
#![allow(
    non_ascii_idents,
    reason = "zerocopy's RawDirectoryEntry derive expansion emits internal identifiers outside this crate's source"
)]

pub mod consts;
mod directory_name;
mod file;
mod metadata;
pub mod writer;

pub use file::{DirectoryEntry, OleError, OleFile, is_ole_file};
pub use metadata::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, OleMetadata, OlePropertySetEditor, PropertySet,
    PropertySetGuid, PropertySetStream, PropertyValue, SUMMARY_INFORMATION_FMTID,
    StandardPropertySet, USER_DEFINED_PROPERTIES_FMTID,
};
pub use writer::OleWriter;

#[cfg(test)]
mod allocation_validation_tests;
#[cfg(test)]
mod directory_validation_tests;
