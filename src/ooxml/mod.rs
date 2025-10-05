/// Office Open XML (OOXML) format implementation.
///
/// This module provides parsing and manipulation of Office Open XML documents,
/// including Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files.
///
/// The implementation is based on the Open Packaging Conventions (OPC) and
/// follows the structure of the python-docx library, adapted for Rust with
/// performance optimizations.

pub mod opc;

// Re-export commonly used types
pub use opc::{OpcPackage, PackURI, XmlPart};