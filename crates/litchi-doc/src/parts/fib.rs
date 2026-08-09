/// File Information Block (FIB) parser for DOC files.
///
/// The FIB is located at the beginning of the `WordDocument` stream and contains
/// critical information about the document structure, including:
/// - File format version
/// - Which table stream to use (0Table or 1Table)
/// - Pointers to various data structures
/// - Document flags and properties
use super::super::package::{Error as PackageError, Result};
use zerocopy::{FromBytes, LE, U16, U32};

/// Minimum FIB size in bytes (the base FIB structure)
const FIB_BASE_SIZE: usize = 32;
/// Offset of `cbRgFcLcb` in Word 97 and newer FIBs.
const TABLE_POINTER_COUNT_OFFSET: usize = 152;
/// Offset of the first `FibRgFcLcb` pair in Word 97 and newer FIBs.
const TABLE_POINTERS_OFFSET: usize = 154;
/// Word 6/95 omit `cbRgFcLcb` and contain this fixed number of pairs.
const LEGACY_TABLE_POINTER_COUNT: usize = 74;
/// One `fc`/`lcb` pair consists of two 32-bit integers.
const TABLE_POINTER_SIZE: usize = 8;
/// Word 8 and newer use the counted `FibRgFcLcb` representation.
const COUNTED_TABLE_POINTER_NFIB: u16 = 0x0069;

/// `nFib` of the first Word 97 binary format, the earliest release MS-DOC
/// describes. Anything lower is a Word 6.0/95 file with a different layout.
pub const WORD_97_NFIB: u16 = 0x00C1;

/// File Information Block.
///
/// The FIB is the primary metadata structure in a DOC file.
/// It's located at offset 0 in the `WordDocument` stream.
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
    /// XOR password verifier or size of the clear table-stream encryption header.
    l_key: u32,
    /// Whether to use 1Table (true) or 0Table (false)
    which_table_stream: bool,
    /// Language ID
    lid: u16,
    /// Complete FIB data for extended parsing
    data: Vec<u8>,
}

impl FileInformationBlock {
    /// Parse a FIB from the `WordDocument` stream.
    ///
    /// # Arguments
    ///
    /// * `word_document` - The `WordDocument` stream data
    ///
    /// # Returns
    ///
    /// A parsed FIB or an error if the data is invalid.
    pub fn parse(word_document: &[u8]) -> Result<Self> {
        Self::parse_at(word_document, 0)
    }

    /// Parse a FIB at an explicit byte offset in the `WordDocument` stream.
    ///
    /// Secondary glossary FIBs are addressed by `FibBase.pnNext * 512`.
    /// The returned object owns only the suffix beginning at the requested FIB,
    /// rather than copying the prefix that precedes it.
    pub fn parse_at(word_document: &[u8], offset: usize) -> Result<Self> {
        let data = word_document.get(offset..).ok_or_else(|| {
            PackageError::Corrupted(format!(
                "FIB offset {offset} is beyond the WordDocument stream"
            ))
        })?;
        if data.len() < FIB_BASE_SIZE {
            return Err(PackageError::Corrupted(
                "WordDocument stream too short for FIB".to_string(),
            ));
        }

        // Read the base FIB fields (little-endian)
        let magic = U16::<LE>::read_from_bytes(&data[0..2]).map_or(0, U16::get);
        let nfib = U16::<LE>::read_from_bytes(&data[2..4]).map_or(0, U16::get);
        let lid = U16::<LE>::read_from_bytes(&data[6..8]).map_or(0, U16::get);
        let flags = U16::<LE>::read_from_bytes(&data[10..12]).map_or(0, U16::get);
        let l_key = U32::<LE>::read_from_bytes(&data[14..18]).map_or(0, U32::get);

        // Validate magic number
        if magic != 0xA5EC && magic != 0xA5DC {
            // 0xA5DC for Word 6.0/95, 0xA5EC for Word 97+
            return Err(PackageError::InvalidFormat(format!(
                "Invalid FIB magic number: 0x{magic:04X}"
            )));
        }

        // Extract which table stream to use (bit 9 of flags at offset 0x0A)
        // This is the fWhichTblStm flag
        let which_table_stream = (flags & 0x0200) != 0;

        // Store the complete FIB data for later parsing of variable fields
        let data = data.to_vec();

        Ok(Self {
            nfib,
            flags,
            l_key,
            which_table_stream,
            lid,
            data,
        })
    }

    /// Get the file format version.
    ///
    /// Common values:
    /// - 0x0065 (101): Word 6.0
    /// - 0x0067 (103): Word 95 (7.0)
    /// - 0x00C1 (193): Word 97 through Word 2003
    /// - 0x0101 (257): Word 2007
    /// - 0x0112 (274): Word 2010+
    ///
    /// Note: All Word versions use 2-byte SPRM opcodes in the binary format,
    /// regardless of the file version. This is consistent with Apache POI's implementation.
    #[inline]
    #[must_use]
    pub fn version(&self) -> u16 {
        self.nfib
    }

    /// Get a human-readable description of the Word version.
    #[must_use]
    pub fn version_name(&self) -> &'static str {
        match self.nfib {
            0x0021 => "Word 1.0",
            0x0045 => "Word 2.0",
            0x0065 => "Word 6.0",
            // LibreOffice treats every nFib through 104 as Word 95 or earlier
            // (sw/source/filter/ww8/ww8scan.cxx), and both 103 and 104 occur.
            0x0067 | 0x0068 => "Word 95 (7.0)",
            0x00C1 => "Word 97",
            0x00D9 => "Word 2000",
            0x0101 => "Word 2002/2003",
            0x010C => "Word 2007",
            0x0112 => "Word 2010",
            0x0113 => "Word 2013",
            _ if self.nfib >= 0x00C1 => "Word 97+",
            _ => "Unknown",
        }
    }

    /// Get which table stream to use.
    ///
    /// Returns `true` for "1Table", `false` for "0Table".
    #[inline]
    #[must_use]
    pub fn which_table_stream(&self) -> bool {
        self.which_table_stream
    }

    /// Check if the document is encrypted.
    ///
    /// Returns `true` if the document requires a password to open.
    ///
    /// Note: This library currently does not support encrypted documents.
    #[inline]
    #[must_use]
    pub fn is_encrypted(&self) -> bool {
        // fEncrypted flag is bit 8 at offset 0x0A
        (self.flags & 0x0100) != 0
    }

    /// Whether this FIB describes an AutoText-only glossary document.
    #[inline]
    #[must_use]
    pub fn is_glossary_document(&self) -> bool {
        // fGlsy is bit 1 of the FibBase flags field.
        (self.flags & 0x0002) != 0
    }

    /// Whether this FIB describes a document template.
    #[inline]
    #[must_use]
    pub fn is_template(&self) -> bool {
        // fDot is bit 0 of the FibBase flags field.
        (self.flags & 0x0001) != 0
    }

    /// Page number of the attached glossary FIB, or zero when none is attached.
    #[inline]
    #[must_use]
    pub fn next_fib_page(&self) -> u16 {
        U16::<LE>::read_from_bytes(&self.data[8..10]).map_or(0, U16::get)
    }

    /// `FibRgLw97.cbMac`, shared by a template and its attached glossary FIB.
    #[inline]
    #[must_use]
    pub fn word_document_size(&self) -> Option<u32> {
        self.data
            .get(64..68)
            .and_then(|bytes| U32::<LE>::read_from_bytes(bytes).ok())
            .map(U32::get)
    }

    /// Check whether the encrypted document uses legacy XOR obfuscation.
    #[inline]
    pub(crate) fn is_obfuscated(&self) -> bool {
        // fObfuscated is bit 15 of the FibBase flags field.
        (self.flags & 0x8000) != 0
    }

    /// Get the size of the clear encryption header at the start of the table stream.
    #[inline]
    pub(crate) fn encryption_header_size(&self) -> u32 {
        self.l_key
    }

    /// Get the 32-bit password verifier used by legacy XOR obfuscation.
    #[inline]
    pub(crate) fn xor_obfuscation_verifier(&self) -> u32 {
        self.l_key
    }

    /// Get the language ID.
    #[inline]
    #[must_use]
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
    /// * `index` - Index of the field in the `FibRgFcLcb` array
    ///
    /// # Returns
    ///
    /// A tuple of (offset, length) in bytes, or `None` if the requested
    /// entry is outside the FIB-declared array or that array is truncated.
    #[must_use]
    pub fn get_table_pointer(&self, index: usize) -> Option<(u32, u32)> {
        let (base_offset, count) = self.table_pointer_layout()?;
        if index >= count {
            return None;
        }
        let entry_offset = base_offset.checked_add(index.checked_mul(TABLE_POINTER_SIZE)?)?;
        let entry = self
            .data
            .get(entry_offset..entry_offset.checked_add(TABLE_POINTER_SIZE)?)?;

        let offset = U32::<LE>::read_from_bytes(&entry[..4]).map_or(0, U32::get);
        let length = U32::<LE>::read_from_bytes(&entry[4..]).map_or(0, U32::get);

        Some((offset, length))
    }

    /// Number of `FibRgFcLcb` offset/length pairs declared by this FIB.
    ///
    /// Returns `None` when the declaration itself or its complete array is
    /// truncated. Consumers must not infer additional pairs from subsequent
    /// bytes in the `WordDocument` stream.
    #[must_use]
    pub fn table_pointer_count(&self) -> Option<usize> {
        self.table_pointer_layout().map(|(_, count)| count)
    }

    fn table_pointer_layout(&self) -> Option<(usize, usize)> {
        let (offset, count) = if self.nfib < COUNTED_TABLE_POINTER_NFIB {
            (TABLE_POINTER_COUNT_OFFSET, LEGACY_TABLE_POINTER_COUNT)
        } else {
            let count_bytes = self
                .data
                .get(TABLE_POINTER_COUNT_OFFSET..TABLE_POINTERS_OFFSET)?;
            let count = usize::from(U16::<LE>::read_from_bytes(count_bytes).map_or(0, U16::get));
            (TABLE_POINTERS_OFFSET, count)
        };
        let byte_len = count.checked_mul(TABLE_POINTER_SIZE)?;
        let end = offset.checked_add(byte_len)?;
        (end <= self.data.len()).then_some((offset, count))
    }

    /// Minimum byte extent occupied by this FIB through its declared pointer array.
    pub(crate) fn minimum_serialized_size(&self) -> Option<usize> {
        let (offset, count) = self.table_pointer_layout()?;
        offset.checked_add(count.checked_mul(TABLE_POINTER_SIZE)?)
    }

    /// Get access to the raw FIB data.
    #[inline]
    #[must_use]
    pub fn raw_data(&self) -> &[u8] {
        &self.data
    }

    /// Get character count for a subdocument.
    ///
    /// Character counts are stored in the `FibRgLw97` structure.
    /// `FibRgLw97` starts at offset 64 (after FibBase=32, csw=2, FibRgW97=28, cslw=2).
    /// Within `FibRgLw97`, character counts start at offset 0xC (after cbMac, reserved1, reserved2).
    /// Each count is a 4-byte little-endian integer.
    ///
    /// # Arguments
    ///
    /// * `index` - Subdocument index:
    ///   - 0: ccpText (main document) at FibRgLw97+0xC = FIB offset 0x4C (76)
    ///   - 1: ccpFtn (footnotes) at FibRgLw97+0x10 = FIB offset 0x50 (80)
    ///   - 2: ccpHdd (headers/footers) at FibRgLw97+0x14 = FIB offset 0x54 (84)
    ///   - 3: ccpMcr (macros) at FibRgLw97+0x18 = FIB offset 0x58 (88)
    ///   - 4: ccpAtn (annotations/comments) at FibRgLw97+0x1C = FIB offset 0x5C (92)
    ///   - 5: ccpEdn (endnotes) at FibRgLw97+0x20 = FIB offset 0x60 (96)
    ///   - 6: ccpTxbx (text boxes) at FibRgLw97+0x24 = FIB offset 0x64 (100)
    ///   - 7: ccpHdrTxbx (header text boxes) at FibRgLw97+0x28 = FIB offset 0x68 (104)
    ///
    /// # Returns
    ///
    /// Character count, or 0 if out of bounds
    fn get_character_count(&self, index: usize) -> u32 {
        // FibRgLw97 starts at offset 64, character counts start at +0xC
        let offset = 64 + 0xC + (index * 4);
        if offset + 4 > self.data.len() {
            return 0;
        }
        U32::<LE>::read_from_bytes(&self.data[offset..offset + 4]).map_or(0, U32::get)
    }

    /// End CP of the concatenated set of all document parts.
    #[must_use]
    pub fn get_document_parts_end(&self) -> Option<u32> {
        (0..8).try_fold(0u32, |total, index| {
            total.checked_add(self.get_character_count(index))
        })
    }

    /// Get the main document character position range.
    ///
    /// Returns (`start_cp`, `end_cp`) for the main document text.
    #[must_use]
    pub fn get_main_doc_range(&self) -> (u32, u32) {
        let ccp_text = self.get_character_count(0);
        (0, ccp_text)
    }

    /// Get the footnote subdocument character position range.
    ///
    /// Returns `Some((start_cp`, `end_cp`)) if footnotes exist, None otherwise.
    #[must_use]
    pub fn get_footnote_range(&self) -> Option<(u32, u32)> {
        let base = self.get_character_count(0);
        let ccp_ftn = self.get_character_count(1);
        if ccp_ftn > 0 {
            Some((base, base + ccp_ftn))
        } else {
            None
        }
    }

    /// Get the header/footer subdocument character position range.
    ///
    /// Returns `Some((start_cp`, `end_cp`)) if headers/footers exist, None otherwise.
    #[must_use]
    pub fn get_header_range(&self) -> Option<(u32, u32)> {
        let base = self.get_character_count(0) + self.get_character_count(1);
        let ccp_hdd = self.get_character_count(2);
        if ccp_hdd > 0 {
            Some((base, base + ccp_hdd))
        } else {
            None
        }
    }

    /// Get the annotations/comments subdocument character position range.
    ///
    /// Returns `Some((start_cp`, `end_cp`)) if comments exist, None otherwise.
    #[must_use]
    pub fn get_comment_range(&self) -> Option<(u32, u32)> {
        let base = self.get_character_count(0)
            + self.get_character_count(1)
            + self.get_character_count(2)
            + self.get_character_count(3); // Skip macros
        let ccp_atn = self.get_character_count(4);
        if ccp_atn > 0 {
            Some((base, base + ccp_atn))
        } else {
            None
        }
    }

    /// Get the endnotes subdocument character position range.
    ///
    /// Returns `Some((start_cp`, `end_cp`)) if endnotes exist, None otherwise.
    #[must_use]
    pub fn get_endnote_range(&self) -> Option<(u32, u32)> {
        let base = self.get_character_count(0)
            + self.get_character_count(1)
            + self.get_character_count(2)
            + self.get_character_count(3)
            + self.get_character_count(4);
        let ccp_edn = self.get_character_count(5);
        if ccp_edn > 0 {
            Some((base, base + ccp_edn))
        } else {
            None
        }
    }

    /// Get the text box subdocument character position range.
    ///
    /// Returns `Some((start_cp`, `end_cp`)) if text boxes exist, None otherwise.
    #[must_use]
    pub fn get_textbox_range(&self) -> Option<(u32, u32)> {
        let base = self.get_character_count(0)
            + self.get_character_count(1)
            + self.get_character_count(2)
            + self.get_character_count(3)
            + self.get_character_count(4)
            + self.get_character_count(5);
        let ccp_txbx = self.get_character_count(6);
        if ccp_txbx > 0 {
            Some((base, base + ccp_txbx))
        } else {
            None
        }
    }

    /// Get the header text box subdocument character position range.
    ///
    /// Returns `Some((start_cp`, `end_cp`)) if header text boxes exist, None otherwise.
    #[must_use]
    pub fn get_header_textbox_range(&self) -> Option<(u32, u32)> {
        let base = self.get_character_count(0)
            + self.get_character_count(1)
            + self.get_character_count(2)
            + self.get_character_count(3)
            + self.get_character_count(4)
            + self.get_character_count(5)
            + self.get_character_count(6);
        let ccp_hdr_txbx = self.get_character_count(7);
        if ccp_hdr_txbx > 0 {
            Some((base, base + ccp_hdr_txbx))
        } else {
            None
        }
    }

    /// Get all subdocument ranges that exist in this document.
    ///
    /// Returns a vector of (name, `start_cp`, `end_cp`) tuples for all non-empty subdocuments.
    #[must_use]
    pub fn get_all_subdoc_ranges(&self) -> Vec<(&'static str, u32, u32)> {
        let mut ranges = Vec::new();

        let (start, end) = self.get_main_doc_range();
        if end > start {
            ranges.push(("Main Document", start, end));
        }

        if let Some((start, end)) = self.get_footnote_range() {
            ranges.push(("Footnotes", start, end));
        }

        if let Some((start, end)) = self.get_header_range() {
            ranges.push(("Headers/Footers", start, end));
        }

        if let Some((start, end)) = self.get_comment_range() {
            ranges.push(("Comments", start, end));
        }

        if let Some((start, end)) = self.get_endnote_range() {
            ranges.push(("Endnotes", start, end));
        }

        if let Some((start, end)) = self.get_textbox_range() {
            ranges.push(("Text Boxes", start, end));
        }

        if let Some((start, end)) = self.get_header_textbox_range() {
            ranges.push(("Header Text Boxes", start, end));
        }

        ranges
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
        data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        // Set fWhichTblStm flag (bit 9)
        data[10] = 0x00;
        data[11] = 0x02;

        let fib = FileInformationBlock::parse(&data).unwrap();
        assert!(fib.which_table_stream());
    }

    #[test]
    fn table_pointers_stop_at_the_declared_count() {
        let mut data = vec![0u8; TABLE_POINTERS_OFFSET + 2 * TABLE_POINTER_SIZE];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        data[TABLE_POINTER_COUNT_OFFSET..TABLE_POINTERS_OFFSET]
            .copy_from_slice(&1u16.to_le_bytes());
        data[TABLE_POINTERS_OFFSET..TABLE_POINTERS_OFFSET + 4]
            .copy_from_slice(&17u32.to_le_bytes());
        data[TABLE_POINTERS_OFFSET + 4..TABLE_POINTERS_OFFSET + 8]
            .copy_from_slice(&23u32.to_le_bytes());
        // These bytes follow the declared array and must never be interpreted
        // as a second pointer pair.
        data[TABLE_POINTERS_OFFSET + 8..TABLE_POINTERS_OFFSET + 12]
            .copy_from_slice(&99u32.to_le_bytes());
        data[TABLE_POINTERS_OFFSET + 12..TABLE_POINTERS_OFFSET + 16]
            .copy_from_slice(&101u32.to_le_bytes());

        let fib = FileInformationBlock::parse(&data).unwrap();
        assert_eq!(fib.table_pointer_count(), Some(1));
        assert_eq!(fib.get_table_pointer(0), Some((17, 23)));
        assert_eq!(fib.get_table_pointer(1), None);
    }

    #[test]
    fn truncated_declared_table_pointer_array_is_unavailable() {
        let mut data = vec![0u8; TABLE_POINTERS_OFFSET + TABLE_POINTER_SIZE];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        data[TABLE_POINTER_COUNT_OFFSET..TABLE_POINTERS_OFFSET]
            .copy_from_slice(&2u16.to_le_bytes());

        let fib = FileInformationBlock::parse(&data).unwrap();
        assert_eq!(fib.table_pointer_count(), None);
        assert_eq!(fib.get_table_pointer(0), None);
    }

    #[test]
    fn word_95_uses_its_fixed_uncounted_pointer_array() {
        let mut data =
            vec![0u8; TABLE_POINTER_COUNT_OFFSET + LEGACY_TABLE_POINTER_COUNT * TABLE_POINTER_SIZE];
        data[0..2].copy_from_slice(&0xA5DCu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x0067u16.to_le_bytes());
        data[TABLE_POINTER_COUNT_OFFSET..TABLE_POINTER_COUNT_OFFSET + 4]
            .copy_from_slice(&31u32.to_le_bytes());
        data[TABLE_POINTER_COUNT_OFFSET + 4..TABLE_POINTER_COUNT_OFFSET + 8]
            .copy_from_slice(&37u32.to_le_bytes());

        let fib = FileInformationBlock::parse(&data).unwrap();
        assert_eq!(fib.table_pointer_count(), Some(LEGACY_TABLE_POINTER_COUNT));
        assert_eq!(fib.get_table_pointer(0), Some((31, 37)));
    }
}
