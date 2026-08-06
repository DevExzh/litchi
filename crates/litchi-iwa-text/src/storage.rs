//! Archive-free rich-text storage values.
//!
//! This module models text and validated byte ranges only. Native storage
//! identifiers, style-table references, protobuf messages, and package
//! traversal stay in the owning IWA adapter.

/// Maximum number of runs retained by one semantic storage.
///
/// The limit protects callers that construct a storage from untrusted range
/// metadata without imposing a limit on the text bytes already owned by the
/// caller. Format adapters apply their own document-wide text budgets.
pub const MAX_RUNS: usize = 1_048_576;

/// Why a semantic storage could not be constructed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The supplied run list exceeds [`MAX_RUNS`].
    TooManyRuns { actual: usize, limit: usize },
    /// A run extends beyond the storage text or its end offset overflows.
    RunOutOfBounds { index: usize },
    /// A run boundary falls inside a UTF-8 code point.
    RunNotOnBoundary { index: usize },
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::TooManyRuns { actual, limit } => write!(
                formatter,
                "text storage contains {actual} runs; maximum is {limit}"
            ),
            Self::RunOutOfBounds { index } => {
                write!(formatter, "text storage run {index} is outside the text")
            },
            Self::RunNotOnBoundary { index } => write!(
                formatter,
                "text storage run {index} does not align to a UTF-8 boundary"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// A checked-by-`Storage` byte range within UTF-8 text.
///
/// A standalone `Run` can describe a range that is not valid for a particular
/// string. Use [`Storage::try_from_parts`] to validate the complete relation
/// before publishing it as semantic storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Run {
    start: usize,
    length: usize,
}

impl Run {
    /// Create a run from a byte offset and byte length.
    #[must_use]
    pub const fn new(start: usize, length: usize) -> Self {
        Self { start, length }
    }

    /// Return the UTF-8 byte offset of this run.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }

    /// Return the UTF-8 byte length of this run.
    #[must_use]
    pub const fn len(self) -> usize {
        self.length
    }

    /// Return the exclusive end offset when addition does not overflow.
    #[must_use]
    pub const fn end(self) -> Option<usize> {
        self.start.checked_add(self.length)
    }

    /// Return whether the run has no bytes.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.length == 0
    }
}

/// An immutable rich-text storage value.
///
/// The text owns one UTF-8 allocation and the run metadata is kept in one
/// compact boxed slice. No field contains a native object ID, style-table ID,
/// protobuf message, archive, or package handle.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Storage {
    text: String,
    runs: Box<[Run]>,
}

impl Storage {
    /// Create an empty storage with no run allocation.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a storage containing one full-text run, unless `text` is empty.
    #[must_use]
    pub fn from_text(text: String) -> Self {
        let runs = if text.is_empty() {
            Box::<[Run]>::default()
        } else {
            Box::new([Run::new(0, text.len())])
        };
        Self { text, runs }
    }

    /// Create storage from text and explicit semantic ranges.
    ///
    /// Empty runs are retained in [`Self::runs`] so valid but unsupported
    /// range metadata is not silently discarded. Iteration skips them because
    /// they do not produce a text fragment.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManyRuns`] when the run list exceeds [`MAX_RUNS`],
    /// [`Error::RunOutOfBounds`] for an overflowing or out-of-text range, or
    /// [`Error::RunNotOnBoundary`] when a range splits a UTF-8 code point.
    pub fn try_from_parts(text: String, runs: Vec<Run>) -> Result<Self, Error> {
        if runs.len() > MAX_RUNS {
            return Err(Error::TooManyRuns {
                actual: runs.len(),
                limit: MAX_RUNS,
            });
        }

        for (index, run) in runs.iter().copied().enumerate() {
            let Some(end) = run.end() else {
                return Err(Error::RunOutOfBounds { index });
            };
            if end > text.len() {
                return Err(Error::RunOutOfBounds { index });
            }
            if text.get(run.start..end).is_none() {
                return Err(Error::RunNotOnBoundary { index });
            }
        }

        Ok(Self {
            text,
            runs: runs.into_boxed_slice(),
        })
    }

    /// Borrow the UTF-8 text without copying it.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Consume the storage and return its owned UTF-8 text.
    ///
    /// The semantic range metadata is intentionally dropped at this boundary;
    /// callers that need it should borrow [`Self::runs`] before consuming the
    /// value. Moving the string avoids a second allocation when an adapter
    /// needs to hand plain text to another owner.
    #[must_use]
    pub fn into_text(self) -> String {
        self.text
    }

    /// Return the UTF-8 byte length of the text.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.text.len()
    }

    /// Return whether the storage contains no text.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.text.is_empty()
    }

    /// Borrow all validated semantic runs, including empty runs.
    #[must_use]
    pub fn runs(&self) -> &[Run] {
        &self.runs
    }

    /// Iterate over non-empty borrowed text fragments in run order.
    pub fn fragments(&self) -> impl Iterator<Item = Fragment<'_>> {
        self.runs.iter().filter_map(|run| {
            if run.is_empty() {
                return None;
            }
            let end = run.end()?;
            self.text
                .get(run.start..end)
                .map(|text| Fragment { text, run })
        })
    }
}

impl AsRef<str> for Storage {
    fn as_ref(&self) -> &str {
        self.text()
    }
}

/// A borrowed non-empty text fragment and its archive-free range metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Fragment<'a> {
    text: &'a str,
    run: &'a Run,
}

impl<'a> Fragment<'a> {
    /// Borrow the fragment text without allocating.
    #[must_use]
    pub const fn text(self) -> &'a str {
        self.text
    }

    /// Borrow the semantic range that produced this fragment.
    #[must_use]
    pub const fn run(self) -> &'a Run {
        self.run
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn from_text_uses_one_full_run_and_no_empty_allocation() {
        let storage = Storage::from_text("Hello, World!".to_owned());
        assert_eq!(storage.text(), "Hello, World!");
        assert_eq!(storage.runs(), [Run::new(0, 13)]);
        assert_eq!(
            storage.fragments().map(Fragment::text).collect::<Vec<_>>(),
            ["Hello, World!"]
        );

        let empty = Storage::from_text(String::new());
        assert!(empty.runs().is_empty());
        assert!(empty.fragments().next().is_none());
    }

    #[test]
    fn fragments_borrow_text_and_keep_only_semantic_ranges() {
        let storage = Storage::try_from_parts(
            "Hello World".to_owned(),
            vec![Run::new(0, 5), Run::new(5, 0), Run::new(6, 5)],
        )
        .unwrap_or_else(|error| panic!("valid ranges should construct: {error}"));

        let fragments = storage.fragments().collect::<Vec<_>>();
        assert_eq!(fragments.len(), 2);
        assert_eq!(fragments[0].text(), "Hello");
        assert_eq!(*fragments[0].run(), Run::new(0, 5));
        assert_eq!(fragments[1].text(), "World");
        assert!(std::ptr::eq(
            fragments[0].text().as_ptr(),
            storage.text().as_ptr()
        ));
        assert!(std::ptr::eq(
            fragments[1].text().as_ptr(),
            storage.text().as_ptr().wrapping_add(6)
        ));
    }

    #[test]
    fn malformed_ranges_return_typed_errors_before_publication() {
        assert_eq!(
            Storage::try_from_parts("é".to_owned(), vec![Run::new(1, 1)]),
            Err(Error::RunNotOnBoundary { index: 0 })
        );
        assert_eq!(
            Storage::try_from_parts("text".to_owned(), vec![Run::new(3, 2)]),
            Err(Error::RunOutOfBounds { index: 0 })
        );
        assert_eq!(
            Storage::try_from_parts("text".to_owned(), vec![Run::new(usize::MAX, 1)]),
            Err(Error::RunOutOfBounds { index: 0 })
        );
    }
}
