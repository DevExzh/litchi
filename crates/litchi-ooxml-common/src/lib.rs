//! Shared OOXML functionality that is independent of DOCX, PPTX, XLSX, and XLSB.

#![forbid(unsafe_code)]

mod error;

pub mod custom;
pub mod custom_xml;
pub mod embedded;
pub mod external_link;
pub mod mce;
pub mod properties;
pub mod relationships;
pub mod ribbon;
pub mod web;
pub mod xml;

pub use error::{Error, Result};
pub use mce::{
    ActiveOffsetLimits, ExpandedName, MceCapabilities, MceError, MceLimits, MceOutput, MceReport,
    active_offsets, process_markup_compatibility, process_ooxml, process_part, process_part_arc,
    process_str,
};
pub use properties::{Keywords, Props};
pub use xml::XmlError;
