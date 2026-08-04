//! Typed SpreadsheetML data-validation support.
///
/// The owner is layered by responsibility: semantic declarations in
/// model, bounded XML/MCE conversion in codec, and worksheet replacement/transaction
/// operations in package.
mod codec;
mod model;
mod package;

pub use codec::{
    parse_data_validation_collections, validate_data_validation_collections,
    write_data_validation_collections, write_data_validation_core,
    write_data_validation_extensions,
};
pub use model::{
    Collection, Conformance, Formula, ListSource, Range, Source, Sqref, Validation,
    ValidationErrorStyle, ValidationImeMode, ValidationOperator, ValidationType,
};
pub use package::replace_data_validation_collections;

pub(crate) const CORE_URI: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT_URI: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const X14_URI: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
pub(crate) const XM_URI: &str = "http://schemas.microsoft.com/office/excel/2006/main";
pub(crate) const X12AC_URI: &str = "http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac";
pub(crate) const XR_URI: &str = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
pub(crate) const CORE: &[u8] = CORE_URI.as_bytes();
pub(crate) const STRICT: &[u8] = STRICT_URI.as_bytes();
pub(crate) const X14: &[u8] = X14_URI.as_bytes();
pub(crate) const XM: &[u8] = XM_URI.as_bytes();
pub(crate) const X12AC: &[u8] = X12AC_URI.as_bytes();
pub(crate) const XR: &[u8] = XR_URI.as_bytes();
pub(crate) const EXTENSION_URI: &str = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";
pub(crate) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_EVENTS: usize = 1_000_000;
pub(crate) const MAX_NODES: usize = 1_000_000;
pub(crate) const MAX_CAPTURED_COLLECTIONS: usize = 1_024;
pub(crate) const MAX_VALIDATIONS: usize = 65_534;
pub(crate) const MAX_REFERENCES: usize = 32_767;
pub(crate) const MAX_FRAGMENT_BYTES: usize = 8 * 1024 * 1024;
pub(crate) const MAX_FORMULA_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_ATTRIBUTE_BYTES: usize = MAX_FORMULA_BYTES;
pub(crate) const MAX_RETAINED_BYTES: usize = 16 * 1024 * 1024;
