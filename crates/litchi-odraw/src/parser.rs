//! Concise entry point for checked `OfficeArt` record traversal.

use crate::{Children, Container, Limits, Record, RecordKind, Result};

/// A zero-copy parser over one `OfficeArt` drawing stream.
#[derive(Debug, Clone, Copy)]
pub struct Parser<'data> {
    data: &'data [u8],
    limits: Limits,
}

impl<'data> Parser<'data> {
    /// Borrows an `OfficeArt` drawing stream.
    #[must_use]
    pub const fn new(data: &'data [u8], limits: Limits) -> Self {
        Self { data, limits }
    }

    /// Returns the original drawing bytes.
    #[must_use]
    pub const fn data(&self) -> &'data [u8] {
        self.data
    }

    /// Parses the only top-level record when it is a container.
    ///
    /// An empty stream or a valid leading atom yields `Ok(None)`; malformed data
    /// yields `Err` so absence is never conflated with corruption.
    ///
    /// # Errors
    ///
    /// Returns an error from `Record::parse` if the leading record is malformed
    /// (`Error::TruncatedHeader`, `Error::TruncatedPayload`, or
    /// `Error::ArithmeticOverflow`), or `Error::TrailingData` if bytes follow
    /// the first top-level record.
    pub fn root(&self) -> Result<Option<Container<'data>>> {
        if self.data.is_empty() {
            return Ok(None);
        }
        let (record, consumed) = Record::parse(self.data, 0, self.limits)?;
        if consumed != self.data.len() {
            return Err(crate::Error::TrailingData { offset: consumed });
        }
        if record.is_container() {
            Container::try_new(record).map(Some)
        } else {
            Ok(None)
        }
    }

    /// Lazily parses the top-level record sequence.
    #[must_use]
    pub const fn records(&self) -> Children<'data> {
        Children::new(self.data, self.limits)
    }

    /// Collects every shape container below the root container.
    ///
    /// # Errors
    ///
    /// Returns an error from `Parser::root` if the stream is malformed, or from
    /// `Container::find_recursive` if a descendant record is malformed or the
    /// caller traversal limits are exceeded.
    pub fn shapes(&self, limits: Limits) -> Result<Vec<Record<'data>>> {
        match self.root()? {
            Some(root) => root.find_recursive(RecordKind::SpContainer, limits),
            None => Ok(Vec::new()),
        }
    }

    /// Collects every client-textbox record below the root container.
    ///
    /// # Errors
    ///
    /// Returns an error from `Parser::root` if the stream is malformed, or from
    /// `Container::find_recursive` if a descendant record is malformed or the
    /// caller traversal limits are exceeded.
    pub fn textboxes(&self, limits: Limits) -> Result<Vec<Record<'data>>> {
        match self.root()? {
            Some(root) => root.find_recursive(RecordKind::ClientTextbox, limits),
            None => Ok(Vec::new()),
        }
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

    #[test]
    fn parses_a_root_container() {
        let bytes = [0x0F, 0x00, 0x02, 0xF0, 0, 0, 0, 0];
        let parser = Parser::new(&bytes, Limits::default());
        let root = parser.root().expect("valid stream").expect("container");

        assert_eq!(root.record().kind(), RecordKind::DgContainer);
    }

    #[test]
    fn malformed_root_is_an_error() {
        let parser = Parser::new(&[0x0F, 0x00, 0x02], Limits::default());

        assert!(matches!(
            parser.root(),
            Err(crate::Error::TruncatedHeader { .. })
        ));
    }

    #[test]
    fn rejects_more_than_one_top_level_record() {
        let bytes = [
            0x0F, 0x00, 0x02, 0xF0, 0, 0, 0, 0, 0x0F, 0x00, 0x02, 0xF0, 0, 0, 0, 0,
        ];

        assert!(matches!(
            Parser::new(&bytes, Limits::default()).root(),
            Err(crate::Error::TrailingData { offset: 8 })
        ));
        assert_eq!(Parser::new(&bytes, Limits::default()).records().count(), 2);
    }
}
