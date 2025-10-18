/// File Information Block (FIB) parser for DOC files.
///
/// The FIB is located at the beginning of the WordDocument stream and contains
/// critical information about the document structure, including:
/// - File format version
/// - Which table stream to use (0Table or 1Table)
/// - Pointers to various data structures
/// - Document flags and properties
use super::super::package::{DocError, Result};
use zerocopy::{FromBytes, LE, U16, U32};

/// Minimum FIB size in bytes (the base FIB structure)
const FIB_BASE_SIZE: usize = 32;

/// File Information Block.
///
/// The FIB is the primary metadata structure in a DOC file.
/// It's located at offset 0 in the WordDocument stream.
///
/// # Structure (simplified)
///
/// - Bytes 0-1: wIdent (magic number, should be 0xA5EC)
/// - Bytes 2-3: nFib (version number)
/// - Bytes 10-11: flags (including which table stream to use)
/// - Bytes 32+: Variable length fields pointing to data structures
#[derive(Debug, Clone)]
pub struct FileInformationBlock {
    /// File format version
    nfib: u16,
    /// Flags including encryption, table stream selection, etc.
    flags: u16,
    /// Whether to use 1Table (true) or 0Table (false)
    which_table_stream: bool,
    /// Language ID
    lid: u16,
    /// Complete FIB data for extended parsing
    data: Vec<u8>,
}

impl FileInformationBlock {
    /// Parse a FIB from the WordDocument stream.
    ///
    /// # Arguments
    ///
    /// * `word_document` - The WordDocument stream data
    ///
    /// # Returns
    ///
    /// A parsed FIB or an error if the data is invalid.
    pub fn parse(word_document: &[u8]) -> Result<Self> {
        if word_document.len() < FIB_BASE_SIZE {
            return Err(DocError::Corrupted(
                "WordDocument stream too short for FIB".to_string(),
            ));
        }

        // Read the base FIB fields (little-endian)
        let magic = U16::<LE>::read_from_bytes(&word_document[0..2])
            .map(|v| v.get())
            .unwrap_or(0);
        let nfib = U16::<LE>::read_from_bytes(&word_document[2..4])
            .map(|v| v.get())
            .unwrap_or(0);
        let lid = U16::<LE>::read_from_bytes(&word_document[6..8])
            .map(|v| v.get())
            .unwrap_or(0);
        let flags = U16::<LE>::read_from_bytes(&word_document[10..12])
            .map(|v| v.get())
            .unwrap_or(0);

        // Validate magic number
        if magic != 0xA5EC && magic != 0xA5DC {
            // 0xA5DC for Word 6.0/95, 0xA5EC for Word 97+
            return Err(DocError::InvalidFormat(format!(
                "Invalid FIB magic number: 0x{:04X}",
                magic
            )));
        }

        // Extract which table stream to use (bit 9 of flags at offset 0x0A)
        // This is the fWhichTblStm flag
        let which_table_stream = (flags & 0x0200) != 0;

        // Store the complete FIB data for later parsing of variable fields
        let data = word_document.to_vec();

        Ok(Self {
            nfib,
            flags,
            which_table_stream,
            lid,
            data,
        })
    }

    /// Get the file format version.
    ///
    /// Common values:
    /// - 0x00C1 (193): Word 97 through Word 2003
    /// - 0x0101 (257): Word 2007
    /// - 0x0112 (274): Word 2010+
    #[inline]
    pub fn version(&self) -> u16 {
        self.nfib
    }

    /// Get which table stream to use.
    ///
    /// Returns `true` for "1Table", `false` for "0Table".
    #[inline]
    pub fn which_table_stream(&self) -> bool {
        self.which_table_stream
    }

    /// Check if the document is encrypted.
    ///
    /// Returns `true` if the document requires a password to open.
    ///
    /// Note: This library currently does not support encrypted documents.
    #[inline]
    pub fn is_encrypted(&self) -> bool {
        // fEncrypted flag is bit 8 at offset 0x0A
        (self.flags & 0x0100) != 0
    }

    /// Get the language ID.
    #[inline]
    pub fn language_id(&self) -> u16 {
        self.lid
    }

    /// Get a pointer to a structure in the table stream.
    ///
    /// The FIB contains many pairs of (offset, length) values pointing to
    /// structures in the table stream. This is a helper to extract them.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the field in the FibRgFcLcb array
    ///
    /// # Returns
    ///
    /// A tuple of (offset, length) in bytes, or None if out of bounds.
    pub fn get_table_pointer(&self, index: usize) -> Option<(u32, u32)> {
        // The FibRgFcLcb array starts at offset 154 for Word 97+
        // Each entry is 8 bytes: 4 bytes offset, 4 bytes length
        let base_offset = 154;
        let entry_offset = base_offset + (index * 8);

        if entry_offset + 8 > self.data.len() {
            return None;
        }

        let offset = U32::<LE>::read_from_bytes(&self.data[entry_offset..entry_offset + 4])
            .map(|v| v.get())
            .unwrap_or(0);

        let length = U32::<LE>::read_from_bytes(&self.data[entry_offset + 4..entry_offset + 8])
            .map(|v| v.get())
            .unwrap_or(0);

        Some((offset, length))
    }

    /// Get access to the raw FIB data.
    #[inline]
    pub fn raw_data(&self) -> &[u8] {
        &self.data
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_fib_min_size() {
        let short_data = vec![0u8; 16];
        let result = FileInformationBlock::parse(&short_data);
        assert!(result.is_err());
    }

    #[test]
    fn test_fib_magic_validation() {
        let mut data = vec![0u8; 512];
        // Set invalid magic number
        data[0] = 0xFF;
        data[1] = 0xFF;

        let result = FileInformationBlock::parse(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_fib_valid() {
        let mut data = vec![0u8; 512];
        // Set valid magic number for Word 97+
        data[0] = 0xEC;
        data[1] = 0xA5;
        // Set nFib to Word 97 version
        data[2] = 0xC1;
        data[3] = 0x00;

        let result = FileInformationBlock::parse(&data);
        assert!(result.is_ok());

        let fib = result.unwrap();
        assert_eq!(fib.version(), 0x00C1);
        assert!(!fib.is_encrypted());
    }

    #[test]
    fn test_fib_table_stream_flag() {
        let mut data = vec![0u8; 512];
        data[0] = 0xEC;
        data[1] = 0xA5;
        // Set fWhichTblStm flag (bit 9)
        data[10] = 0x00;
        data[11] = 0x02;

        let fib = FileInformationBlock::parse(&data).unwrap();
        assert!(fib.which_table_stream());
    }
}

