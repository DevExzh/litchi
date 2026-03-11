/// Errors that can occur during OMML parsing
#[derive(Debug)]
pub enum OmmlError {
    XmlError(String),
    ParseError(String),
    InvalidStructure(String),
    ValidationError(String),
    UnsupportedFeature(String),
    EncodingError(String),
    DepthLimitExceeded(usize),
    MalformedElement(String),
    MissingRequiredElement(String),
    InvalidAttribute(String),
    ArenaAllocationError(String),
}

impl std::fmt::Display for OmmlError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            OmmlError::XmlError(msg) => write!(f, "XML parsing error: {}", msg),
            OmmlError::ParseError(msg) => write!(f, "OMML parse error: {}", msg),
            OmmlError::InvalidStructure(msg) => write!(f, "Invalid OMML structure: {}", msg),
            OmmlError::ValidationError(msg) => write!(f, "OMML validation error: {}", msg),
            OmmlError::UnsupportedFeature(msg) => write!(f, "Unsupported OMML feature: {}", msg),
            OmmlError::EncodingError(msg) => write!(f, "Text encoding error: {}", msg),
            OmmlError::DepthLimitExceeded(limit) => {
                write!(f, "XML depth limit exceeded: {}", limit)
            },
            OmmlError::MalformedElement(msg) => write!(f, "Malformed element: {}", msg),
            OmmlError::MissingRequiredElement(msg) => {
                write!(f, "Missing required element: {}", msg)
            },
            OmmlError::InvalidAttribute(msg) => write!(f, "Invalid attribute: {}", msg),
            OmmlError::ArenaAllocationError(msg) => write!(f, "Arena allocation error: {}", msg),
        }
    }
}

impl std::error::Error for OmmlError {}

impl From<std::str::Utf8Error> for OmmlError {
    fn from(err: std::str::Utf8Error) -> Self {
        OmmlError::EncodingError(format!("UTF-8 decoding error: {}", err))
    }
}

impl From<bumpalo::AllocErr> for OmmlError {
    fn from(err: bumpalo::AllocErr) -> Self {
        OmmlError::ArenaAllocationError(format!("Arena allocation failed: {}", err))
    }
}

impl From<quick_xml::Error> for OmmlError {
    fn from(err: quick_xml::Error) -> Self {
        OmmlError::XmlError(format!("Quick XML error: {}", err))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_omml_error_display() {
        let err = OmmlError::XmlError("xml test".to_string());
        assert!(err.to_string().contains("XML parsing error"));
        assert!(err.to_string().contains("xml test"));

        let err = OmmlError::ParseError("parse test".to_string());
        assert!(err.to_string().contains("OMML parse error"));
        assert!(err.to_string().contains("parse test"));

        let err = OmmlError::InvalidStructure("structure test".to_string());
        assert!(err.to_string().contains("Invalid OMML structure"));

        let err = OmmlError::ValidationError("validation test".to_string());
        assert!(err.to_string().contains("OMML validation error"));

        let err = OmmlError::UnsupportedFeature("feature test".to_string());
        assert!(err.to_string().contains("Unsupported OMML feature"));

        let err = OmmlError::EncodingError("encoding test".to_string());
        assert!(err.to_string().contains("Text encoding error"));

        let err = OmmlError::DepthLimitExceeded(100);
        assert!(err.to_string().contains("XML depth limit exceeded"));
        assert!(err.to_string().contains("100"));

        let err = OmmlError::MalformedElement("malformed test".to_string());
        assert!(err.to_string().contains("Malformed element"));

        let err = OmmlError::MissingRequiredElement("missing test".to_string());
        assert!(err.to_string().contains("Missing required element"));

        let err = OmmlError::InvalidAttribute("invalid attr".to_string());
        assert!(err.to_string().contains("Invalid attribute"));

        let err = OmmlError::ArenaAllocationError("arena test".to_string());
        assert!(err.to_string().contains("Arena allocation error"));
    }

    #[test]
    fn test_omml_error_from_utf8_error() {
        // Create an invalid UTF-8 sequence
        let invalid_utf8 = vec![0x80, 0x81, 0x82];
        let utf8_err = std::str::from_utf8(&invalid_utf8).unwrap_err();
        let err: OmmlError = utf8_err.into();
        assert!(matches!(err, OmmlError::EncodingError(_)));
        assert!(err.to_string().contains("UTF-8"));
    }

    #[test]
    fn test_omml_error_debug() {
        let err = OmmlError::ParseError("test".to_string());
        let debug_str = format!("{:?}", err);
        assert!(debug_str.contains("ParseError"));
        assert!(debug_str.contains("test"));
    }
}
