//! OOXML adapter for the canonical XLSX conditional-formatting owner.
//!
//! The typed model and bounded XML parser live in
//! [`litchi_xlsx::conditional_formatting`]. These adapters retain the host
//! error boundary used by worksheet and styles parsing.

use crate::error::{OoxmlError, Result};

pub use litchi_xlsx::conditional_formatting::{
    Association, Axis, Color, ColorRole, ColorScale, Component, DataBar, Differential,
    DifferentialRef, Direction, Formatting, IconSet, IconSet14, Icons, Kind, NamedColor,
    NumberFormat, Operator, Payload, Period, Range, Rule, Source, TokenError, Value, ValueKind,
};

/// Parse worksheet conditional-formatting fragments through the canonical
/// owner while preserving the host's historical error variants.
pub(crate) fn parse_conditional_formattings(
    xml: &[u8],
    differential_format_count: usize,
) -> Result<Vec<Formatting>> {
    litchi_xlsx::conditional_formatting::parse_conditional_formattings(
        xml,
        differential_format_count,
    )
    .map_err(map_owner_error)
}

/// Parse differential styles through the canonical owner while preserving the
/// host's historical error variants.
pub(crate) fn parse_differential_formats(xml: &[u8]) -> Result<Vec<Differential>> {
    litchi_xlsx::conditional_formatting::parse_differential_formats(xml).map_err(map_owner_error)
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn owner_invalid_errors_keep_the_historical_host_boundary() {
        let error = parse_conditional_formattings(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A1"><cfRule type="not-a-rule"/></conditionalFormatting></worksheet>"#,
            0,
        )
        .expect_err("invalid rule kind must fail");
        assert!(matches!(error, OoxmlError::InvalidFormat(_)));
    }

    #[test]
    fn owner_xml_errors_keep_the_historical_host_boundary() {
        let error = parse_conditional_formattings(br#"<!DOCTYPE x><worksheet/>"#, 0)
            .expect_err("DTD must fail");
        assert!(matches!(error, OoxmlError::InvalidFormat(_)));
    }
}
