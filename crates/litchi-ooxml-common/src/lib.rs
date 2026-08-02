//! Shared OOXML functionality that is independent of DOCX, PPTX, XLSX, and XLSB.

#![forbid(unsafe_code)]

mod error;

pub mod custom;
pub mod custom_xml;
pub mod external_link;
pub mod mce;
pub mod properties;
pub mod xml;

pub use error::{Error, Result};
pub use mce::{
    ExpandedName, MceCapabilities, MceError, MceLimits, MceOutput, MceReport,
    process_markup_compatibility, process_ooxml, process_part, process_part_arc, process_str,
};
pub use properties::DocumentProperties;
pub use xml::XmlError;
