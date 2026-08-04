//! Shared OpenDocument vocabulary and scalar codecs.
//!
//! This crate contains functionality shared by every ODF document family:
//! constants, spreadsheet coordinates, namespace/qualified-name vocabulary,
//! and lexical data types.

#![forbid(unsafe_code)]

pub mod annotation;
pub mod constants;
pub mod coordinates;
pub mod core;
pub mod datatype;
pub mod namespace;
pub mod package;
pub mod rdf;
pub mod signature;
