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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_latex_error_display() {
        let err = LatexError::FormatError("test format".to_string());
        assert!(err.to_string().contains("Format error"));
        assert!(err.to_string().contains("test format"));

        let err = LatexError::InvalidNode("test node".to_string());
        assert!(err.to_string().contains("Invalid node"));
        assert!(err.to_string().contains("test node"));
    }

    #[test]
    fn test_latex_error_debug() {
        let err = LatexError::FormatError("test".to_string());
        let debug_str = format!("{:?}", err);
        assert!(debug_str.contains("FormatError"));
        assert!(debug_str.contains("test"));
    }
}
