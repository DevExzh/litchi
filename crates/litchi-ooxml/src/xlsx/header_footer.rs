//! OOXML package boundary for worksheet header/footer settings.
//!
//! The worksheet model and bounded SpreadsheetML codec live in
//! `litchi_xlsx::header_footer`. This module only maps owner errors into the
//! host error type while the owner supplies the semantic values.

use crate::error::{OoxmlError, Result};

pub use litchi_xlsx::header_footer::{SectionKind, Settings, Text};

/// Parse worksheet header/footer settings through the canonical XLSX owner.
pub fn parse_worksheet_header_footer(xml: &[u8]) -> Result<Option<Settings>> {
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
