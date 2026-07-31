//! Borrowed worksheet row views.

use std::slice;

use litchi_sheet::Row as Index;

/// One stored SpreadsheetML row record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Stored {
    pub(crate) index: Index,
    pub(crate) hidden: bool,
}

impl Stored {
    pub(crate) const fn new(index: Index, hidden: bool) -> Self {
        Self { index, hidden }
    }
}

/// Borrowed view of one logical worksheet row.
///
/// Every checked grid row has a view. [`Self::stored`] distinguishes an
/// explicit SpreadsheetML row record from an implicit default row.
#[derive(Debug, Clone, Copy)]
pub struct Row<'a> {
    index: Index,
    stored: Option<&'a Stored>,
}

impl<'a> Row<'a> {
    pub(crate) const fn new(index: Index, stored: Option<&'a Stored>) -> Self {
        Self { index, stored }
    }

    /// Checked zero-based row coordinate.
    pub const fn index(self) -> Index {
        self.index
    }

    /// Whether the worksheet contains an explicit row record here.
    pub const fn stored(self) -> bool {
        self.stored.is_some()
    }

    /// Whether the row is explicitly hidden.
    pub const fn hidden(self) -> bool {
        match self.stored {
            Some(row) => row.hidden,
            None => false,
        }
    }
}

/// Lazy borrowed traversal of explicit worksheet row records.
#[derive(Debug, Clone)]
pub struct Rows<'a> {
    inner: slice::Iter<'a, Stored>,
}

impl<'a> Rows<'a> {
    pub(crate) fn new(rows: &'a [Stored]) -> Self {
        Self { inner: rows.iter() }
    }
}

impl<'a> Iterator for Rows<'a> {
    type Item = Row<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        self.inner
            .next()
            .map(|stored| Row::new(stored.index, Some(stored)))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.inner.size_hint()
    }
}

impl DoubleEndedIterator for Rows<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.inner
            .next_back()
            .map(|stored| Row::new(stored.index, Some(stored)))
    }
}

impl ExactSizeIterator for Rows<'_> {}
impl std::iter::FusedIterator for Rows<'_> {}
