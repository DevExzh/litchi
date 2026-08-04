//! Shared OpenDocument vocabulary and scalar codecs.
//!
//! This crate contains functionality shared by every ODF document family:
//! constants, spreadsheet coordinates, and lexical data types. Package
//! graphs, encryption, and document-family models remain in `litchi-odf`.

#![forbid(unsafe_code)]

pub mod constants;
pub mod coordinates;
pub mod datatype;
pub mod package;
