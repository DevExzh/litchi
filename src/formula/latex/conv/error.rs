// Error definitions for LaTeX conversion

/// Errors that can occur during LaTeX conversion
#[derive(Debug)]
pub enum LatexError {
    FormatError(String),
    InvalidNode(String),
}

impl std::fmt::Display for LatexError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            LatexError::FormatError(msg) => write!(f, "Format error: {}", msg),
            LatexError::InvalidNode(msg) => write!(f, "Invalid node: {}", msg),
        }
    }
}

impl std::error::Error for LatexError {}
