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
            OmmlError::DepthLimitExceeded(limit) => write!(f, "XML depth limit exceeded: {}", limit),
            OmmlError::MalformedElement(msg) => write!(f, "Malformed element: {}", msg),
            OmmlError::MissingRequiredElement(msg) => write!(f, "Missing required element: {}", msg),
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
