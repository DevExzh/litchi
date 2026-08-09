//! Worksheet data-consolidation settings (`CT_DataConsolidate`).
//!
//! The facade keeps the typed worksheet model separate from its bounded
//! `SpreadsheetML` codec while retaining the historical public API.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::{parse_worksheet_data_consolidation, write_worksheet_data_consolidation};
pub use model::{
    Conformance, DataConsolidation, Function, RangeReference, Reference, ReferenceSource,
    References,
};

pub(crate) const TRANSITIONAL_MAIN: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const TRANSITIONAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const MAX_DATA_REFERENCES: usize = 65_536;
pub(crate) const MAX_XSTRING_CHARS: usize = 32_767;
pub(crate) const MAX_RELATIONSHIP_ID_CHARS: usize = 1_024;

pub(crate) fn invalid(message: impl Into<String>) -> crate::error::Error {
    crate::error::Error::Invalid(message.into())
}
