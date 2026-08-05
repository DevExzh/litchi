//! Archive-free Numbers document semantics.

use std::fmt;
use std::sync::Arc;

use crate::Sheet;

/// Maximum number of ordered sheets retained by one semantic document.
pub const MAX_SHEETS: usize = 4096;

/// Errors returned while constructing a bounded semantic document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The supplied sheet sequence exceeds the selected bound.
    TooManySheets {
        /// Number of supplied sheets.
        actual: usize,
        /// Maximum accepted sheets.
        limit: usize,
    },
    /// A sheet does not carry its canonical position in the ordered sequence.
    InvalidSheetIndex {
        /// Position occupied by the sheet in the supplied sequence.
        expected: usize,
        /// Index stored by the sheet.
        actual: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::TooManySheets { actual, limit } => write!(
                formatter,
                "Numbers document contains {actual} sheets; maximum is {limit}"
            ),
            Self::InvalidSheetIndex { expected, actual } => write!(
                formatter,
                "Numbers sheet index {actual} is not the expected index {expected}"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result type for bounded Numbers semantic construction.
pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug)]
struct State {
    sheets: Box<[Sheet]>,
}

/// An immutable, archive-free Numbers document snapshot.
///
/// The document owns only semantic [`Sheet`] values. Its hidden state is
/// reference counted so cloning or taking a snapshot never copies the sheet
/// or table storage. Native archives, protobuf values, package entries, and
/// physical object identifiers are intentionally outside this API.
#[derive(Debug, Clone)]
pub struct Document {
    state: Arc<State>,
}

impl Document {
    /// Build a document from sheets in source order.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManySheets`] when the hard semantic bound is
    /// exceeded, or [`Error::InvalidSheetIndex`] when a sheet is not numbered
    /// by its zero-based position in the supplied sequence.
    pub fn from_sheets(sheets: Vec<Sheet>) -> Result<Self> {
        Self::from_sheets_with_max_sheets(sheets, MAX_SHEETS)
    }

    /// Build a document under a caller-selected sheet-count budget.
    ///
    /// The package-independent hard cap [`MAX_SHEETS`] cannot be relaxed by a
    /// caller. The input vector is consumed without rebuilding its sheet
    /// values when construction succeeds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManySheets`] when the supplied count exceeds either
    /// the caller budget or the hard semantic cap, or
    /// [`Error::InvalidSheetIndex`] when a sheet is not numbered by its
    /// zero-based position in the supplied sequence.
    pub fn from_sheets_with_max_sheets(sheets: Vec<Sheet>, max_sheets: usize) -> Result<Self> {
        let limit = max_sheets.min(MAX_SHEETS);
        if sheets.len() > limit {
            return Err(Error::TooManySheets {
                actual: sheets.len(),
                limit,
            });
        }

        for (expected, sheet) in sheets.iter().enumerate() {
            if sheet.index() != expected {
                return Err(Error::InvalidSheetIndex {
                    expected,
                    actual: sheet.index(),
                });
            }
        }

        Ok(Self {
            state: Arc::new(State {
                sheets: sheets.into_boxed_slice(),
            }),
        })
    }

    /// Capture another cheap handle to the same immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow all sheets in stable source order.
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.state.sheets
    }

    /// Select a sheet by checked zero-based position.
    #[must_use]
    pub fn sheet(&self, index: usize) -> Option<&Sheet> {
        self.state.sheets.get(index)
    }

    /// Return the number of semantic sheets.
    #[must_use]
    pub fn sheet_count(&self) -> usize {
        self.state.sheets.len()
    }

    /// Return whether the document contains no semantic sheets.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.state.sheets.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn empty_document_is_a_valid_bounded_snapshot() {
        let document = Document::from_sheets(Vec::new())
            .unwrap_or_else(|error| panic!("empty document should be valid: {error}"));

        assert_send_sync::<Document>();
        assert!(document.is_empty());
        assert_eq!(document.sheet_count(), 0);
        assert!(document.sheets().is_empty());
        assert!(document.sheet(0).is_none());
    }

    #[test]
    fn construction_checks_budget_and_canonical_order() {
        let too_many = Document::from_sheets_with_max_sheets(vec![Sheet::new("Sheet 1", 0)], 0);
        assert!(matches!(
            too_many,
            Err(Error::TooManySheets {
                actual: 1,
                limit: 0,
            })
        ));

        let invalid = Document::from_sheets(vec![Sheet::new("Sheet 2", 1)]);
        assert!(matches!(
            invalid,
            Err(Error::InvalidSheetIndex {
                expected: 0,
                actual: 1,
            })
        ));
    }

    #[test]
    fn clones_share_ordered_semantic_storage() {
        let document =
            Document::from_sheets(vec![Sheet::new("Sheet 1", 0), Sheet::new("Sheet 2", 1)])
                .unwrap_or_else(|error| panic!("document should be valid: {error}"));
        let snapshot = document.snapshot();

        assert!(Arc::ptr_eq(&document.state, &snapshot.state));
        assert_eq!(snapshot.sheet_count(), 2);
        assert_eq!(snapshot.sheet(0).map(Sheet::name), Some("Sheet 1"));
        assert_eq!(snapshot.sheet(1).map(Sheet::name), Some("Sheet 2"));
        assert!(snapshot.sheet(2).is_none());
    }
}
