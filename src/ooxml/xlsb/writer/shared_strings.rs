//! Shared strings table writer for XLSB

use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::records::record_types;
use crate::ooxml::xlsb::writer::RecordWriter;
use std::collections::HashMap;
use std::io::Write;

/// Shared strings table writer
pub struct MutableSharedStringsWriter {
    strings: Vec<String>,
    string_map: HashMap<String, u32>,
}

impl MutableSharedStringsWriter {
    /// Create a new shared strings writer
    pub fn new() -> Self {
        MutableSharedStringsWriter {
            strings: Vec::new(),
            string_map: HashMap::new(),
        }
    }

    /// Add a string to the shared strings table
    ///
    /// Returns the index of the string (existing or newly added)
    pub fn add_string(&mut self, s: String) -> u32 {
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
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Check if the table is empty
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Write shared strings table to binary format
    pub(crate) fn write<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        // Write BrtBeginSst
        let mut sst_header = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut sst_header);
        temp_writer.write_u32(self.strings.len() as u32)?; // Total unique strings
        temp_writer.write_u32(self.strings.len() as u32)?; // Total string count (same for now)

        writer.write_record(record_types::BEGIN_SST, &sst_header)?;

        // Write each string
        for string in &self.strings {
            self.write_sst_item(writer, string)?;
        }

        // Write BrtEndSst
        writer.write_record(record_types::END_SST, &[])?;

        Ok(())
    }

    /// Write a single SST item
    fn write_sst_item<W: Write>(&self, writer: &mut RecordWriter<W>, s: &str) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        temp_writer.write_u8(0)?; // Flags (0 for plain text)
        temp_writer.write_wide_string(s)?;

        writer.write_record(record_types::SST_ITEM, &data)?;
        Ok(())
    }
}

impl Default for MutableSharedStringsWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_shared_strings_writer_new() {
        let writer = MutableSharedStringsWriter::new();
        assert!(writer.is_empty());
        assert_eq!(writer.len(), 0);
    }

    #[test]
    fn test_shared_strings_writer_default() {
        let writer: MutableSharedStringsWriter = Default::default();
        assert!(writer.is_empty());
        assert_eq!(writer.len(), 0);
    }

    #[test]
    fn test_add_string() {
        let mut writer = MutableSharedStringsWriter::new();
        let idx1 = writer.add_string("Hello".to_string());
        let idx2 = writer.add_string("World".to_string());

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(writer.len(), 2);
        assert!(!writer.is_empty());
    }

    #[test]
    fn test_add_duplicate_string() {
        let mut writer = MutableSharedStringsWriter::new();
        let idx1 = writer.add_string("Test".to_string());
        let idx2 = writer.add_string("Test".to_string());
        let idx3 = writer.add_string("Other".to_string());

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 0); // Same index for duplicate
        assert_eq!(idx3, 1);
        assert_eq!(writer.len(), 2); // Only 2 unique strings
    }

    #[test]
    fn test_add_multiple_strings() {
        let mut writer = MutableSharedStringsWriter::new();
        let strings = vec!["A", "B", "C", "D", "E"];

        for (i, s) in strings.iter().enumerate() {
            let idx = writer.add_string(s.to_string());
            assert_eq!(idx, i as u32);
        }

        assert_eq!(writer.len(), 5);
    }

    #[test]
    fn test_add_empty_string() {
        let mut writer = MutableSharedStringsWriter::new();
        let idx = writer.add_string("".to_string());

        assert_eq!(idx, 0);
        assert_eq!(writer.len(), 1);
    }

    #[test]
    fn test_add_unicode_string() {
        let mut writer = MutableSharedStringsWriter::new();
        let idx = writer.add_string("Hello 世界 🌍".to_string());

        assert_eq!(idx, 0);
        assert_eq!(writer.len(), 1);
    }
}
