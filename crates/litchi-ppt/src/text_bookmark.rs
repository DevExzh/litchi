//! Typed `PowerPoint` text-range bookmark atoms.

use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;

/// A validated `TextBookmarkAtom` linking a text range to summary metadata.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextBookmark {
    pub begin: u32,
    pub end: u32,
    pub id: u32,
}

impl TextBookmark {
    /// Create a validated bookmark over the text range `begin..end`.
    ///
    /// # Errors
    ///
    /// Returns an error if `end` is not greater than `begin`.
    pub fn new(begin: u32, end: u32, id: u32) -> Result<Self> {
        let bookmark = Self { begin, end, id };
        bookmark.validate()?;
        Ok(bookmark)
    }

    /// Parse a complete `TextBookmarkAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header is invalid, the payload is not
    /// exactly 12 bytes, or the bookmark range is empty or reversed.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0
            || record.instance != 0
            || record.record_type_raw != RecordType::TextBookmarkAtom.as_u16()
            || record.data.len() != 12
            || record.data_length != 12
        {
            return corrupted("TextBookmarkAtom has an invalid header or size");
        }
        let bookmark = Self {
            begin: read_u32(record, 0)?,
            end: read_u32(record, 4)?,
            id: read_u32(record, 8)?,
        };
        bookmark.validate()?;
        Ok(bookmark)
    }

    /// Serialize as a parsed `Record`, verifying the canonical encoding round-trips.
    ///
    /// # Errors
    ///
    /// Returns an error if the bookmark range is invalid or the canonical
    /// record bytes cannot be re-parsed exactly.
    pub fn to_record(&self) -> Result<Record> {
        let bytes = self.to_record_bytes()?;
        let (record, end) = Record::parse(&bytes, 0)?;
        if end != bytes.len() {
            return corrupted("canonical TextBookmarkAtom did not consume its bytes");
        }
        Ok(record)
    }

    /// Serialize the complete 20-byte `TextBookmarkAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the bookmark range is empty or reversed.
    pub fn to_record_bytes(&self) -> Result<[u8; 20]> {
        self.validate()?;
        let mut bytes = [0; 20];
        bytes[2..4].copy_from_slice(&RecordType::TextBookmarkAtom.as_u16().to_le_bytes());
        bytes[4..8].copy_from_slice(&12u32.to_le_bytes());
        bytes[8..12].copy_from_slice(&self.begin.to_le_bytes());
        bytes[12..16].copy_from_slice(&self.end.to_le_bytes());
        bytes[16..20].copy_from_slice(&self.id.to_le_bytes());
        Ok(bytes)
    }

    #[must_use]
    pub fn len(&self) -> u32 {
        self.end - self.begin
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        false
    }

    /// Whether the range meets the producer interoperability recommendation.
    #[must_use]
    pub fn is_compatibly_short(&self) -> bool {
        self.len() <= 255
    }

    fn validate(&self) -> Result<()> {
        if self.end <= self.begin {
            return corrupted("TextBookmarkAtom end must be greater than begin");
        }
        Ok(())
    }
}

fn read_u32(record: &Record, offset: usize) -> Result<u32> {
    litchi_core::binary::read_u32_le_at(&record.data, offset)
        .map_err(|_err| Error::Corrupted("TextBookmarkAtom payload is truncated".into()))
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn text_bookmark_roundtrips_and_reports_advisory_length() {
        let bookmark = TextBookmark::new(10, 265, 41).unwrap();
        assert!(bookmark.is_compatibly_short());
        let parsed = TextBookmark::parse(&bookmark.to_record().unwrap()).unwrap();
        assert_eq!(parsed, bookmark);
        let long = TextBookmark::new(10, 266, 42).unwrap();
        assert!(!long.is_compatibly_short());
        assert_eq!(long.len(), 256);
    }

    #[test]
    fn rejects_empty_reversed_and_malformed_atoms() {
        assert!(TextBookmark::new(10, 10, 1).is_err());
        assert!(TextBookmark::new(11, 10, 1).is_err());
        let mut bytes = TextBookmark::new(1, 2, 1)
            .unwrap()
            .to_record_bytes()
            .unwrap();
        bytes[0] = 1;
        let record = Record::parse(&bytes, 0).unwrap().0;
        assert!(TextBookmark::parse(&record).is_err());
    }
}
