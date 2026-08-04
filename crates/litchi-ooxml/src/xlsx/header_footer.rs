//! Compatibility facade for the canonical XLSX header/footer owner.
//!
//! The worksheet model and bounded SpreadsheetML codec live in
//! litchi_xlsx::header_footer. This module retains the historical
//! litchi_ooxml::xlsx::header_footer path and OoxmlError boundary.

use crate::error::{OoxmlError, Result};

pub use litchi_xlsx::header_footer::{
    HeaderFooterSectionKind, HeaderFooterText, SectionKind, Settings, Text, WorksheetHeaderFooter,
};

/// Parse worksheet header/footer settings through the canonical XLSX owner.
pub fn parse_worksheet_header_footer(xml: &[u8]) -> Result<Option<WorksheetHeaderFooter>> {
    litchi_xlsx::header_footer::parse_worksheet_header_footer(xml).map_err(map_owner_error)
}

fn map_owner_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Package(error) => OoxmlError::Opc(error),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::Xml(error.to_string()),
        litchi_xlsx::Error::Common(error) => OoxmlError::Common(error),
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_xlsx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::Xlsx(other),
    }
}
