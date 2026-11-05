#![allow(
    clippy::cast_possible_truncation,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error to this codec boundary"
)]

//! Shared strings table writer for XLSB

use crate::package::error::{Error, Result};
use crate::raw::Encoder;
use crate::raw::kind;
use std::collections::HashMap;
use std::io::Write;

/// Shared strings table writer
pub(crate) struct SharedStringsWriter {
    strings: Vec<String>,
    string_map: HashMap<String, u32>,
    total_count: u64,
}

impl SharedStringsWriter {
    /// Create a new shared strings writer
    pub(crate) fn new() -> Self {
        SharedStringsWriter {
            strings: Vec::new(),
            string_map: HashMap::new(),
            total_count: 0,
        }
    }

    /// Add a string to the shared strings table
    ///
    /// Returns the index of the string (existing or newly added)
    pub(crate) fn add_string(&mut self, s: String) -> u32 {
        self.total_count = self.total_count.saturating_add(1);
        if let Some(&index) = self.string_map.get(&s) {
            index
        } else {
            let index = self.strings.len() as u32;
            self.string_map.insert(s.clone(), index);
            self.strings.push(s);
            index
        }
    }

    /// Get the count of unique strings
    #[cfg(test)]
    pub(crate) fn len(&self) -> usize {
        self.strings.len()
    }

    /// Total number of string-cell occurrences added, including duplicates.
    #[cfg(test)]
    pub(crate) fn total_count(&self) -> u64 {
        self.total_count
    }

    /// Check if the table is empty
    pub(crate) fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Write shared strings table to binary format
    pub(crate) fn write<W: Write>(&self, writer: &mut Encoder<W>) -> Result<()> {
        // Write BrtBeginSst
        let mut sst_header = Vec::new();
        let mut temp_writer = Encoder::new(&mut sst_header, crate::raw::Limits::DEFAULT);
        let total_count = u32::try_from(self.total_count)
            .map_err(|_| Error::Encoding("shared-string occurrence count exceeds u32"))?;
        let unique_count = u32::try_from(self.strings.len())
            .map_err(|_| Error::Encoding("shared-string unique count exceeds u32"))?;
        if total_count > 0x7FFF_FFFF || unique_count > total_count {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::InvalidValue,
                "build_xlsb",
                "shared-string counts are outside the XLSB range",
            ));
        }
        temp_writer.write_u32(total_count)?;
        temp_writer.write_u32(unique_count)?;

        writer.write_record(kind::BEGIN_SST, &sst_header)?;

        // Write each string
        for string in &self.strings {
            self.write_sst_item(writer, string)?;
        }

        // Write BrtEndSst
        writer.write_record(kind::END_SST, &[])?;

        Ok(())
    }

    /// Write a single SST item
    fn write_sst_item<W: Write>(&self, writer: &mut Encoder<W>, s: &str) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Encoder::new(&mut data, crate::raw::Limits::DEFAULT);

        temp_writer.write_u8(0)?; // Flags (0 for plain text)
        temp_writer.write_wide_string(s)?;

        writer.write_record(kind::SST_ITEM, &data)?;
        Ok(())
    }
}

impl Default for SharedStringsWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_shared_strings_writer_new() {
        let writer = SharedStringsWriter::new();
        assert!(writer.is_empty());
        assert_eq!(writer.len(), 0);
    }

    #[test]
    fn test_shared_strings_writer_default() {
        let writer: SharedStringsWriter = Default::default();
        assert!(writer.is_empty());
        assert_eq!(writer.len(), 0);
    }

    #[test]
    fn test_add_string() {
        let mut writer = SharedStringsWriter::new();
        let idx1 = writer.add_string("Hello".to_string());
        let idx2 = writer.add_string("World".to_string());

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(writer.len(), 2);
        assert!(!writer.is_empty());
    }

    #[test]
    fn test_add_duplicate_string() {
        let mut writer = SharedStringsWriter::new();
        let idx1 = writer.add_string("Test".to_string());
        let idx2 = writer.add_string("Test".to_string());
        let idx3 = writer.add_string("Other".to_string());

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 0); // Same index for duplicate
        assert_eq!(idx3, 1);
        assert_eq!(writer.len(), 2); // Only 2 unique strings
        assert_eq!(writer.total_count(), 3);
    }

    #[test]
    fn test_add_multiple_strings() {
        let mut writer = SharedStringsWriter::new();
        let strings = ["A", "B", "C", "D", "E"];

        for (i, s) in strings.iter().enumerate() {
            let idx = writer.add_string(s.to_string());
            assert_eq!(idx, i as u32);
        }

        assert_eq!(writer.len(), 5);
    }

    #[test]
    fn test_add_empty_string() {
        let mut writer = SharedStringsWriter::new();
        let idx = writer.add_string("".to_string());

        assert_eq!(idx, 0);
        assert_eq!(writer.len(), 1);
    }

    #[test]
    fn test_add_unicode_string() {
        let mut writer = SharedStringsWriter::new();
        let idx = writer.add_string("Hello 世界 🌍".to_string());

        assert_eq!(idx, 0);
        assert_eq!(writer.len(), 1);
    }

    #[test]
    fn writes_total_and_unique_counts() {
        let mut table = SharedStringsWriter::new();
        table.add_string("same".to_string());
        table.add_string("same".to_string());
        table.add_string("other".to_string());

        let mut bytes = Vec::new();
        table
            .write(&mut Encoder::new(&mut bytes, crate::raw::Limits::DEFAULT))
            .unwrap();
        let mut records = crate::raw::Records::new(bytes.as_slice(), crate::raw::Limits::DEFAULT);
        let begin = records.next().unwrap().unwrap();
        assert_eq!(begin.kind(), kind::BEGIN_SST);
        assert_eq!(&begin.payload()[..4], &3u32.to_le_bytes());
        assert_eq!(&begin.payload()[4..], &2u32.to_le_bytes());
    }
}
