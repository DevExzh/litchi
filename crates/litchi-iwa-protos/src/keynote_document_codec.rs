//! Private-type Buffa projection for the Keynote root document.
//!
//! The generated projection contains only `KN.DocumentArchive.show` and its
//! three scalar `TSP.Reference` fields. The caller must first validate the root
//! document's required opaque base envelope and the uniqueness of the show
//! reference. Unknown bytes remain in the caller-owned source and are never
//! materialized or re-encoded here.

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_document_generated::LitchiIwaProjection as projection;

/// Finite limits already established by the Keynote root wire preflight.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile for one preflighted Keynote root.
    #[must_use]
    pub const fn new(max_message_bytes: usize, recursion_limit: u32) -> Self {
        Self {
            max_message_bytes,
            recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Failure from the private Keynote document projection decoder.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    MissingRequired(&'static str),
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(error),
        }
    }
}

impl DecodeError {
    /// Required schema field absent from the source, when applicable.
    #[must_use]
    pub const fn missing_required(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_) => None,
        }
    }

    const fn missing_required_field(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
        }
    }
}

/// Decode the show identifier from one already-preflighted Keynote root.
///
/// The singular deferred show reference is always accessed, forcing Buffa to
/// validate its complete wire payload and required identifier before the
/// scalar is returned. The generated view and all unknown root fields remain
/// private and borrowed.
pub fn decode_show_identifier(source: &[u8], options: DecodeOptions) -> Result<u64, DecodeError> {
    let view: projection::KeynoteDocumentArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if !view.has_show() {
        return Err(DecodeError::missing_required_field(
            "KN.DocumentArchive.show",
        ));
    }
    let show = view
        .show
        .get()?
        .ok_or_else(|| DecodeError::missing_required_field("KN.DocumentArchive.show"))?;
    if !show.has_identifier() {
        return Err(DecodeError::missing_required_field(
            "TSP.Reference.identifier",
        ));
    }
    Ok(show.identifier)
}

#[cfg(test)]
mod tests {
    use prost::Message as _;

    use super::{DecodeOptions, decode_show_identifier};

    fn decode(source: &[u8]) -> Result<u64, super::DecodeError> {
        decode_show_identifier(source, DecodeOptions::new(source.len(), 2))
    }

    #[test]
    fn opaque_document_super_is_not_decoded() -> Result<(), Box<dyn std::error::Error>> {
        let source = [0x12, 0x02, 0x08, 0x2a, 0x1a, 0x01, 0xff];
        assert_eq!(decode(&source)?, 42);
        Ok(())
    }

    #[test]
    fn canonical_prost_document_matches_the_projection() -> Result<(), Box<dyn std::error::Error>> {
        let source = crate::kn::DocumentArchive {
            super_: crate::tsa::DocumentArchive::default(),
            show: crate::tsp::Reference {
                identifier: 42,
                deprecated_type: Some(7),
                deprecated_is_external: Some(false),
            },
            tables_custom_format_list: None,
        }
        .encode_to_vec();

        assert_eq!(decode(&source)?, 42);
        Ok(())
    }

    #[test]
    fn nested_reference_is_forced_and_required() {
        let Err(error) = decode(&[0x12, 0x00, 0x1a, 0x00]) else {
            panic!("a show reference without its required identifier must fail");
        };
        assert_eq!(error.missing_required(), Some("TSP.Reference.identifier"));
    }

    #[test]
    fn malformed_nested_reference_is_rejected() {
        let Err(error) = decode(&[0x12, 0x01, 0x08, 0x1a, 0x00]) else {
            panic!("the deferred show reference must be decoded");
        };
        assert!(error.missing_required().is_none());
    }

    #[test]
    fn missing_show_is_rejected() {
        let Err(error) = decode(&[0x1a, 0x00]) else {
            panic!("the required show must be present");
        };
        assert_eq!(error.missing_required(), Some("KN.DocumentArchive.show"));
    }
}
