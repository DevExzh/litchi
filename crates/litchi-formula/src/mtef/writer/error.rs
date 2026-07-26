//! Errors raised while serializing a formula AST to MTEF
//!
//! The writer never panics: every limit imposed by the binary format (record
//! field widths, matrix dimensions, nesting depth) is reported as one of these
//! variants instead.

/// Error produced when a formula cannot be encoded as MTEF
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum MtefWriteError {
    /// The AST nests deeper than the writer is willing to recurse
    ///
    /// The limit exists so that a pathological (or maliciously constructed)
    /// tree cannot exhaust the stack.
    DepthExceeded {
        /// Maximum nesting depth accepted by the writer
        limit: usize,
    },

    /// A character cannot be represented by a 16-bit MathType character code
    ///
    /// MTEF stores characters as 16-bit MTCode values, so characters outside
    /// the Basic Multilingual Plane have no encoding.
    UnsupportedCharacter(char),

    /// A matrix has more rows or columns than the MATRIX record can describe
    MatrixTooLarge {
        /// Requested number of rows
        rows: usize,
        /// Requested number of columns
        cols: usize,
        /// Largest row or column count the record supports
        limit: usize,
    },

    /// A template variation exceeds the 15 bits available to encode it
    VariationTooLarge(u32),

    /// The encoded equation is larger than the OLE header length field allows
    OutputTooLarge(usize),

    /// A font name contains a NUL byte and cannot be stored as a C string
    InvalidFontName,
}

impl std::fmt::Display for MtefWriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            MtefWriteError::DepthExceeded { limit } => {
                write!(f, "formula nests deeper than {} levels", limit)
            },
            MtefWriteError::UnsupportedCharacter(ch) => {
                write!(
                    f,
                    "character U+{:04X} is not representable in MTEF",
                    *ch as u32
                )
            },
            MtefWriteError::MatrixTooLarge { rows, cols, limit } => write!(
                f,
                "matrix of {}x{} exceeds the MTEF limit of {} rows/columns",
                rows, cols, limit
            ),
            MtefWriteError::VariationTooLarge(variation) => {
                write!(f, "template variation {} is not encodable", variation)
            },
            MtefWriteError::OutputTooLarge(len) => {
                write!(f, "encoded equation of {} bytes is too large", len)
            },
            MtefWriteError::InvalidFontName => {
                write!(f, "font name contains an embedded NUL byte")
            },
        }
    }
}

impl std::error::Error for MtefWriteError {}
