//! Shared OOXML functionality that is independent of DOCX, PPTX, XLSX, and XLSB.

#![forbid(unsafe_code)]

mod error;

pub mod custom;
pub mod custom_xml;
pub mod embedded;
pub mod external_link;
pub mod mce;
#[cfg(feature = "encryption")]
pub mod package_encryption;
pub mod properties;
pub mod relationships;
pub mod ribbon;
pub mod spreadsheet_xml_maps;
#[cfg(feature = "vba-inspection")]
pub mod vba;
pub mod web;
pub mod xml;
pub mod xml_name;

pub use error::{Error, Result};
pub use properties::{Keywords, Props};
pub use xml::XmlError;
