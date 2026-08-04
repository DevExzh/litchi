//! Compatibility adapter for the canonical XLSX auto-filter codec.
//!
//! The model and bounded SpreadsheetML codec live in `litchi_xlsx`. The
//! wrappers retain the historical host error variants for existing OOXML
//! call sites while re-exporting the owner types.

use crate::error::{OoxmlError, Result};

pub use litchi_xlsx::auto_filter::*;

fn map_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Package(error) => OoxmlError::Opc(error),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::Xml(error.to_string()),
        litchi_xlsx::Error::Common(error) => OoxmlError::Common(error),
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        other => OoxmlError::Xlsx(other),
    }
}

/// Parse a worksheet XML document and return its worksheet-owned auto-filter.
pub fn parse_auto_filter(xml: &[u8]) -> Result<Option<AutoFilterDefinition>> {
    litchi_xlsx::auto_filter::parse_auto_filter(xml).map_err(map_error)
}

/// Parse an isolated auto-filter fragment.
pub fn parse_auto_filter_fragment(xml: &[u8]) -> Result<AutoFilterDefinition> {
    litchi_xlsx::auto_filter::parse_auto_filter_fragment(xml).map_err(map_error)
}

/// Serialize an auto-filter fragment for embedding in a worksheet or cache.
pub fn write_auto_filter_fragment(value: &AutoFilterDefinition) -> Result<Vec<u8>> {
    litchi_xlsx::auto_filter::write_auto_filter_fragment(value).map_err(map_error)
}
