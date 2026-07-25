//! Microsoft Compound File Binary (CFB / OLE2) container parser and writer.
//!
//! This crate provides the CFB substrate consumed by the legacy Office binary
//! format crates (`litchi-ole` for `.doc`/`.xls`/`.ppt`) and by encrypted
//! OOXML packages (`litchi-ooxml` under its `crypto` feature).
//!
//! See `[MS-CFB]: Compound File Binary File Format` for the format spec.

#![allow(missing_docs)]

pub mod consts;
mod directory_name;
mod file;
mod metadata;
pub mod ovba;
pub mod writer;

pub use file::{DirectoryEntry, OleError, OleFile, is_ole_file};
pub use metadata::{
    DOCUMENT_SUMMARY_INFORMATION_FMTID, OleMetadata, OlePropertySetEditor, PropertySet,
    PropertySetGuid, PropertySetStream, PropertyValue, StandardPropertySet,
    SUMMARY_INFORMATION_FMTID, USER_DEFINED_PROPERTIES_FMTID,
};
pub use writer::OleWriter;

#[cfg(test)]
mod allocation_validation_tests;
#[cfg(test)]
mod directory_validation_tests;
