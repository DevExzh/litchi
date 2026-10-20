//! The crate's typed error.

use std::collections::TryReserveError;

use crate::limits::LimitKind;

/// A specialized result for this crate.
pub type Result<T, E = Error> = std::result::Result<T, E>;

/// The text format an error was found in.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Format {
    /// Delimited text: CSV, TSV, PRN or another single-byte delimiter.
    Delimited,
    /// Data Interchange Format.
    Dif,
    /// SYmbolic LinK.
    Sylk,
    /// Fixed-width columns.
    FixedWidth,
}

impl Format {
    /// The lowercase name: `delimited`, `dif`, `sylk` or `fixed-width`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Delimited => "delimited",
            Self::Dif => "dif",
            Self::Sylk => "sylk",
            Self::FixedWidth => "fixed-width",
        }
    }
}

/// Which option an [`Error::InvalidOption`] names.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OptionKind {
    /// The delimiter is not ASCII, is a line break or NUL, or equals the
    /// quote.
    Delimiter,
    /// The quote is not ASCII or is a line break or NUL.
    Quote,
    /// The comment marker is not ASCII, is a line break or NUL, or equals
    /// the delimiter or the quote.
    Comment,
    /// The byte-order mark to write names an encoding the writers do not
    /// produce (UTF-32).
    Bom,
}

/// A content-free failure: positions and figures, never the text read.
///
/// Lines are one-based and count physical lines of the decoded text, so a
/// quoted line break advances them.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// Reading the source or writing the output failed.
    #[error("text workbook I/O failed: {0}")]
    Io(#[from] std::io::Error),
    /// An option cannot be used.
    #[error("the {} option cannot be used", option_name(*.option))]
    InvalidOption {
        /// The option.
        option: OptionKind,
    },
    /// The bytes are not valid in the selected encoding.
    #[error("the source is not valid in its encoding at byte {offset}")]
    Encoding {
        /// Byte offset of the first invalid sequence.
        offset: usize,
    },
    /// A quoted field is not closed before the end of the input.
    #[error("the quoted field starting on line {line} is not closed")]
    UnterminatedQuote {
        /// The line where the field starts.
        line: usize,
    },
    /// A C0 control character other than tab and line breaks, or the
    /// non-character U+FFFE or U+FFFF, where a tokenizer refuses them.
    #[error("line {line} contains a character no spreadsheet cell can store")]
    InvalidCharacter {
        /// The line where it occurs.
        line: usize,
    },
    /// The source does not follow its format's structure.
    #[error("the {} source is malformed at line {line}", .format.as_str())]
    Syntax {
        /// The format.
        format: Format,
        /// The line where the structure breaks (the line count at the end
        /// of the input when a required section is missing).
        line: usize,
    },
    /// A cell coordinate lies outside the spreadsheet grid.
    #[error("the {} source addresses a cell outside the grid at line {line}", .format.as_str())]
    OutOfGrid {
        /// The format.
        format: Format,
        /// The line of the record.
        line: usize,
    },
    /// A coordinate such as `B2` is not valid A1 syntax.
    #[error("the cell coordinate is not valid A1 syntax")]
    InvalidCoordinate,
    /// A zero-based row index is past the last row.
    #[error("row {row} is past the last row")]
    RowOutOfRange {
        /// The zero-based row index.
        row: usize,
    },
    /// No worksheet has the requested name or position; a text workbook has
    /// exactly one.
    #[error("the text workbook has no such worksheet")]
    MissingWorksheet,
    /// A limit was exceeded while reading.
    #[error("the {} limit of {maximum} was exceeded ({observed})", .kind.as_str())]
    Limit {
        /// The dimension.
        kind: LimitKind,
        /// The observed amount (at least the first value past the maximum).
        observed: usize,
        /// The maximum.
        maximum: usize,
    },
    /// A limit value cannot be used.
    #[error("the {} limit cannot be {value} (1..={maximum})", .kind.as_str())]
    InvalidLimit {
        /// The dimension.
        kind: LimitKind,
        /// The refused value.
        value: usize,
        /// The largest accepted value.
        maximum: usize,
    },
    /// An allocation failed.
    #[error("allocation failed")]
    Allocation {
        /// The failed reservation.
        #[source]
        source: TryReserveError,
    },
}

impl From<TryReserveError> for Error {
    fn from(source: TryReserveError) -> Self {
        Self::Allocation { source }
    }
}

/// Classify a text-workbook error for the format-neutral sheet error the
/// formula engine reports. The message is static: no source text, name or
/// position crosses.
impl From<Error> for litchi_core::sheet::Error {
    #[cold]
    #[inline(never)]
    fn from(error: Error) -> Self {
        use litchi_core::edit::RefusalKind;
        let refused = |kind, reason| Self::refused(kind, "read", reason);
        match error {
            Error::Io(_) => Self::Io,
            Error::MissingWorksheet => refused(
                RefusalKind::NotFound,
                "the text workbook has no such worksheet",
            ),
            Error::RowOutOfRange { .. } => {
                refused(RefusalKind::NotFound, "the row is past the last row")
            },
            Error::InvalidCoordinate | Error::InvalidOption { .. } | Error::InvalidLimit { .. } => {
                refused(RefusalKind::InvalidValue, "the request is not valid")
            },
            Error::Limit { .. } | Error::Allocation { .. } => refused(
                RefusalKind::UnsupportedContent,
                "the text source exceeds a configured limit",
            ),
            Error::Encoding { .. }
            | Error::UnterminatedQuote { .. }
            | Error::InvalidCharacter { .. }
            | Error::Syntax { .. }
            | Error::OutOfGrid { .. } => Self::Malformed("the text source is malformed"),
        }
    }
}

const fn option_name(option: OptionKind) -> &'static str {
    match option {
        OptionKind::Delimiter => "delimiter",
        OptionKind::Quote => "quote",
        OptionKind::Comment => "comment",
        OptionKind::Bom => "byte-order mark",
    }
}
