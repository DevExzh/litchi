//! The error of the generic sheet reads and of formula evaluation.

use std::fmt;

use crate::edit::{Refusal, RefusalKind};

/// A failure of a generic sheet read or of a formula evaluation.
///
/// Implementations of [`Worksheet`](super::Worksheet) and
/// [`WorkbookTrait`](super::WorkbookTrait) name their own error type (the
/// `Error` associated type), so a format keeps its typed error end to end.
/// This type is the shared, format-neutral one: the formula engine reports
/// it, and a format error converts into it by classification
/// (`impl From<FormatError> for sheet::Error`).
///
/// It follows the record 0351 §2.4 taxonomy and is content-free by
/// construction: every message is static text, never a sheet name, a cell
/// value or formula text.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed read or evaluation this route refuses: a missing
    /// worksheet or name, an invalid coordinate, an unsupported construct.
    Refused(Refusal),
    /// A bound was exceeded.
    Limit {
        /// The bound.
        kind: LimitKind,
        /// The count that exceeded it.
        observed: usize,
        /// The bound's value.
        maximum: usize,
    },
    /// A replay recipient does not match the patch's recorded base.
    PatchConflict,
    /// The source file changed after it was opened.
    SourceChanged,
    /// Reading the source failed.
    Io,
    /// The source data is malformed.
    Malformed(&'static str),
    /// A broken invariant.
    Internal(&'static str),
}

/// The bound an [`Error::Limit`] reports.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// The chain of formula cells one evaluation follows.
    Depth,
    /// The cells loaded or read.
    Cells,
}

/// The result of a generic sheet read or a formula evaluation; `E` is the
/// implementation's error type.
pub type Result<T, E = Error> = std::result::Result<T, E>;

impl Error {
    /// A refusal of `operation` for `reason` (both static text).
    #[must_use]
    pub const fn refused(kind: RefusalKind, operation: &'static str, reason: &'static str) -> Self {
        Self::Refused(Refusal::new(kind, operation, reason))
    }

    /// The refusal, when this is one.
    #[must_use]
    pub const fn refusal(&self) -> Option<&Refusal> {
        match self {
            Self::Refused(refusal) => Some(refusal),
            _ => None,
        }
    }
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Depth => "formula evaluation depth",
            Self::Cells => "cell count",
        })
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(formatter, "{kind} {observed} exceeds its bound {maximum}"),
            Self::PatchConflict => {
                formatter.write_str("the patch base does not match the recipient")
            },
            Self::SourceChanged => formatter.write_str("the source changed after it was opened"),
            Self::Io => formatter.write_str("reading the source failed"),
            Self::Malformed(reason) => write!(formatter, "malformed source: {reason}"),
            Self::Internal(reason) => write!(formatter, "internal error: {reason}"),
        }
    }
}

impl std::error::Error for Error {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn messages_are_static_and_classified() {
        let refused = Error::refused(
            RefusalKind::NotFound,
            "worksheet_by_name",
            "no such worksheet",
        );
        assert_eq!(
            refused.refusal().map(Refusal::kind),
            Some(RefusalKind::NotFound)
        );
        assert_eq!(
            refused.to_string(),
            "worksheet_by_name refused (not_found): no such worksheet"
        );
        let limit = Error::Limit {
            kind: LimitKind::Depth,
            observed: 9,
            maximum: 8,
        };
        assert_eq!(limit.refusal(), None);
        assert_eq!(
            limit.to_string(),
            "formula evaluation depth 9 exceeds its bound 8"
        );
        assert_eq!(Error::Io.to_string(), "reading the source failed");
    }
}
