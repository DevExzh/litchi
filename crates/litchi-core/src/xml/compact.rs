//! Typed failures for XML inputs that violate a compact representation policy.

use std::fmt;

/// Why XML could not be accepted by a compact parser.
///
/// The byte offset identifies the first byte at which the parser observed the
/// violation. It is carried by [`crate::Error::XmlCompactness`], keeping this
/// vocabulary independent of any concrete document format or XML backend.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum CompactnessKind {
    /// The input exceeds the parser's configured byte budget.
    InputTooLarge,
    /// XML nesting exceeds the parser's configured depth budget.
    DepthLimit,
    /// A bounded allocation required to represent the XML could not be made.
    AllocationFailed,
    /// Whitespace makes the XML ineligible for the compact representation.
    FormattingWhitespace,
    /// An empty element uses a spaced form that is ineligible for the compact representation.
    SpacedEmptyElement,
    /// The document contains a DTD.
    DocumentType,
    /// The XML is malformed.
    MalformedXml,
}

impl fmt::Display for CompactnessKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputTooLarge => "input is too large",
            Self::DepthLimit => "nesting exceeds the depth limit",
            Self::AllocationFailed => "bounded allocation failed",
            Self::FormattingWhitespace => "formatting whitespace is not compact",
            Self::SpacedEmptyElement => "spaced empty element is not compact",
            Self::DocumentType => "document type declarations are not compact",
            Self::MalformedXml => "malformed XML",
        })
    }
}

#[cfg(test)]
mod tests {
    use super::CompactnessKind;
    use crate::error::Error;

    #[test]
    fn preserves_kind_and_offset_without_format_dependencies() {
        let error = Error::XmlCompactness {
            kind: CompactnessKind::SpacedEmptyElement,
            offset: 19,
        };
        assert!(matches!(
            error,
            Error::XmlCompactness {
                kind: CompactnessKind::SpacedEmptyElement,
                offset: 19,
            }
        ));
    }

    #[test]
    fn kinds_have_stable_concise_messages() {
        assert_eq!(
            CompactnessKind::InputTooLarge.to_string(),
            "input is too large"
        );
        assert_eq!(
            CompactnessKind::DepthLimit.to_string(),
            "nesting exceeds the depth limit"
        );
        assert_eq!(
            CompactnessKind::AllocationFailed.to_string(),
            "bounded allocation failed"
        );
        assert_eq!(
            CompactnessKind::FormattingWhitespace.to_string(),
            "formatting whitespace is not compact"
        );
        assert_eq!(
            CompactnessKind::SpacedEmptyElement.to_string(),
            "spaced empty element is not compact"
        );
        assert_eq!(
            CompactnessKind::DocumentType.to_string(),
            "document type declarations are not compact"
        );
        assert_eq!(CompactnessKind::MalformedXml.to_string(), "malformed XML");
    }
}
