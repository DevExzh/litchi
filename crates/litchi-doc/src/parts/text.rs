/// Text extraction from DOC files.
///
/// This module handles extracting text from the binary structures in DOC files.
/// Text in DOC files is stored in a complex way:
/// - The actual text bytes are in the `WordDocument` stream
/// - A "Piece Table" (CLX structure) in the Table stream maps character positions to file positions
/// - Text can be in either 8-bit (Windows-1252) or 16-bit (UTF-16LE) format
use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use litchi_core::binary::{read_u16_le, read_u32_le};
use std::sync::Arc;

/// Size of a `PieceDescriptor` in bytes (8 bytes as per Apache POI)
pub const PIECE_DESCRIPTOR_SIZE: usize = 8;

/// Bytes per UTF-16LE code unit, for sizing a non-ANSI piece's text run.
const UTF16_CODE_UNIT_BYTES: usize = 2;

/// CLX (Compound Line Extension) parsing utilities.
///
/// Based on Apache POI's `PlexOfCps` implementation, the CLX structure is a
/// Property List with Character Positions (PLCF) that contains:
/// - 4-byte count of entries
/// - For each entry: 4 bytes (CP start) + 4 bytes (CP end) + 8 bytes (`PieceDescriptor`)
/// - Total: (count + 1) * 4 + count * 8 bytes
///
/// Text extractor for DOC files.
///
/// Handles the complex text extraction process from DOC binary structures.
pub struct TextExtractor {
    /// The extracted text
    text: Arc<String>,
    /// UTF-16 CP boundary to UTF-8 byte-offset mapping.
    cp_to_byte: Arc<[usize]>,
}

impl TextExtractor {
    /// Create a new `TextExtractor` and extract text.
    ///
    /// # Arguments
    ///
    /// * `fib` - The File Information Block
    /// * `word_document` - The `WordDocument` stream
    /// * `table_stream` - The Table stream (0Table or 1Table)
    pub fn new(
        fib: &FileInformationBlock,
        word_document: &[u8],
        table_stream: &[u8],
    ) -> Result<Self> {
        // Extract text using the piece table
        let utf16 = Self::extract_text_from_pieces(fib, word_document, table_stream)?;
        let text = Arc::new(String::from_utf16_lossy(&utf16));
        let cp_to_byte = Arc::from(Self::build_cp_to_byte_map(&text).into_boxed_slice());

        Ok(Self { text, cp_to_byte })
    }

    /// Extract all text from the document.
    pub fn extract_all_text(&self) -> Result<String> {
        Ok(self.text.as_ref().clone())
    }

    /// Get a reference to the full extracted text.
    #[inline]
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Share the decoded text with semantic readers without copying it.
    pub(crate) fn text_shared(&self) -> Arc<String> {
        Arc::clone(&self.text)
    }

    /// Share the UTF-16 CP boundary map with semantic readers without rebuilding it.
    pub(crate) fn cp_to_byte_shared(&self) -> Arc<[usize]> {
        Arc::clone(&self.cp_to_byte)
    }

    /// Extract text for a specific character position (CP) range.
    ///
    /// This is used for subdocument text extraction (headers, footnotes, etc.)
    /// where the text occupies a known CP range within the full document text.
    ///
    /// # Arguments
    ///
    /// * `start_cp` - Starting character position (inclusive)
    /// * `end_cp` - Ending character position (exclusive)
    ///
    /// # Returns
    ///
    /// Text for the given UTF-16 CP range, or an empty string if out of bounds.
    ///
    /// MS-DOC ranges must not split a UTF-16 surrogate pair. Such malformed ranges return an
    /// empty slice because Rust strings cannot represent an isolated surrogate code unit.
    #[must_use]
    pub fn text_at_range(&self, start_cp: u32, end_cp: u32) -> &str {
        let start = usize::try_from(start_cp)
            .unwrap_or(usize::MAX)
            .min(self.cp_to_byte.len().saturating_sub(1));
        let end = usize::try_from(end_cp)
            .unwrap_or(usize::MAX)
            .min(self.cp_to_byte.len().saturating_sub(1));
        if start >= end
            || (start > 0 && self.cp_to_byte[start] == self.cp_to_byte[start - 1])
            || (end > 0 && self.cp_to_byte[end] == self.cp_to_byte[end - 1])
        {
            return "";
        }
        &self.text[self.cp_to_byte[start]..self.cp_to_byte[end]]
    }

    /// Whether a UTF-16 character position maps to a Rust string boundary.
    pub(crate) fn is_cp_boundary(&self, cp: u32) -> bool {
        let Ok(index) = usize::try_from(cp) else {
            return false;
        };
        index < self.cp_to_byte.len()
            && (index == 0 || self.cp_to_byte[index] != self.cp_to_byte[index - 1])
    }

    fn build_cp_to_byte_map(text: &str) -> Vec<usize> {
        let mut cp_to_byte = Vec::with_capacity(text.encode_utf16().count() + 1);
        for (offset, ch) in text.char_indices() {
            for _ in 0..ch.len_utf16() {
                cp_to_byte.push(offset);
            }
        }
        cp_to_byte.push(text.len());
        cp_to_byte
    }

    /// Extract text using the piece table (CLX structure).
    ///
    /// The piece table maps character positions in the document to file positions
    /// in the `WordDocument` stream. This allows Word to efficiently handle insertions
    /// and deletions without rewriting the entire file.
    fn extract_text_from_pieces(
        fib: &FileInformationBlock,
        word_document: &[u8],
        table_stream: &[u8],
    ) -> Result<Vec<u16>> {
        // Get the CLX (piece table) location from FIB
        // CLX is at FibRgFcLcb index 33 (fcClx, lcbClx) according to Apache POI's FIBFieldHandler
        let (clx_offset, clx_length) = fib
            .get_table_pointer(33)
            .ok_or_else(|| PackageError::Corrupted("CLX pointer not found in FIB".to_string()))?;

        if clx_length == 0 {
            // No piece table means text starts at offset 0x200 or 0x800
            // This is a simplified doc or Word 6.0 format
            return Self::extract_text_simple(word_document);
        }

        // Extract the CLX from the table stream
        let clx_offset = clx_offset as usize;
        let clx_length = clx_length as usize;

        if clx_offset >= table_stream.len() {
            return Err(PackageError::Corrupted(format!(
                "CLX offset {} is beyond table stream length {}",
                clx_offset,
                table_stream.len()
            )));
        }

        if clx_offset + clx_length > table_stream.len() {
            return Err(PackageError::Corrupted(format!(
                "CLX extends beyond table stream: offset={}, length={}, stream_len={}",
                clx_offset,
                clx_length,
                table_stream.len()
            )));
        }

        let clx_data = &table_stream[clx_offset..clx_offset + clx_length];

        // Try to parse the piece table from CLX
        match Self::parse_piece_table(clx_data, word_document) {
            Ok(text) if !text.is_empty() => Ok(text),
            _ => {
                // If CLX parsing fails or returns empty, fall back to simple text extraction
                Self::extract_text_simple(word_document)
            },
        }
    }

    /// Parse the piece table from CLX data using Apache POI's `ComplexFileTable` logic.
    ///
    /// The CLX structure follows Apache POI's `ComplexFileTable` format:
    /// - Optional GRPPR L sections (type 0x01) for fast-saved files
    /// - `TEXT_PIECE_TABLE_TYPE` marker (0x02)
    /// - 4-byte size of the piece table data
    /// - The piece table data itself (`PlexOfCps` structure)
    fn parse_piece_table(clx_data: &[u8], word_document: &[u8]) -> Result<Vec<u16>> {
        let mut offset = 0;

        // Skip GRPPR L sections (type 0x01) until we find the piece table
        while offset < clx_data.len() {
            if offset >= clx_data.len() {
                return Err(PackageError::Corrupted(
                    "Unexpected end of CLX data".to_string(),
                ));
            }

            let section_type = clx_data[offset];
            offset += 1;

            match section_type {
                0x01 => {
                    // GRPPR L section - skip it
                    if offset + 2 > clx_data.len() {
                        return Err(PackageError::Corrupted(
                            "GRPPR L section truncated".to_string(),
                        ));
                    }

                    let size = read_u16_le(clx_data, offset).unwrap_or(0) as usize;
                    offset += 2 + size;
                },
                0x02 => {
                    // TEXT_PIECE_TABLE_TYPE - this is the piece table
                    if offset + 4 > clx_data.len() {
                        return Err(PackageError::Corrupted(
                            "Piece table size field truncated".to_string(),
                        ));
                    }

                    let piece_table_size = read_u32_le(clx_data, offset).unwrap_or(0) as usize;
                    offset += 4;

                    if offset + piece_table_size > clx_data.len() {
                        return Err(PackageError::Corrupted(
                            "Piece table data truncated".to_string(),
                        ));
                    }

                    let piece_table_data = &clx_data[offset..offset + piece_table_size];

                    // Parse the piece table using PlexOfCps logic
                    let pieces = Self::parse_plex_of_cps(piece_table_data)?;

                    // Extract text from the parsed pieces
                    return Self::extract_text_from_piece_descriptors(&pieces, word_document);
                },
                0x14 => {
                    // Document Properties Descriptor - contains document-wide properties
                    if offset + 2 > clx_data.len() {
                        return Err(PackageError::Corrupted(
                            "Document Properties section truncated".to_string(),
                        ));
                    }

                    let size = read_u16_le(clx_data, offset).unwrap_or(0) as usize;
                    offset += 2 + size;
                },
                _ => {
                    // For unknown section types, try to skip them gracefully
                    if offset + 2 <= clx_data.len() {
                        let size = read_u16_le(clx_data, offset).unwrap_or(0) as usize;
                        offset += 2 + size;
                    } else {
                        return Err(PackageError::Corrupted(format!(
                            "Unexpected CLX section type 0x{section_type:02X} at end of data"
                        )));
                    }
                },
            }
        }

        // If we reach here, no piece table was found in the CLX
        // Return empty string to trigger fallback
        Ok(Vec::new())
    }

    /// Parse a `PlexOfCps` structure (Property List with Character Positions).
    ///
    /// Based on Apache POI's `PlexOfCps` implementation:
    /// Format: [CP0] [CP1] ... [`CP_n`] [Struct0] [Struct1] ... [Struct_{n-1}]
    /// - For n pieces, there are (n+1) CPs (4 bytes each)
    /// - Followed by n `PieceDescriptor` structs (8 bytes each)
    /// - Number of pieces: n = (size - 4) / (4 + 8)
    fn parse_plex_of_cps(plex_data: &[u8]) -> Result<Vec<PieceDescriptor>> {
        if plex_data.len() < 4 {
            return Ok(Vec::new());
        }

        // Calculate number of pieces using Apache POI's formula:
        // _iMac = (cb - 4) / (4 + cbStruct)
        let num_pieces = (plex_data.len() - 4) / (4 + PIECE_DESCRIPTOR_SIZE);

        if num_pieces == 0 {
            return Ok(Vec::new());
        }

        // Validate size
        let expected_size = 4 + num_pieces * (4 + PIECE_DESCRIPTOR_SIZE);
        if plex_data.len() < expected_size {
            return Err(PackageError::Corrupted(format!(
                "PlexOfCps truncated: expected {} bytes, got {}",
                expected_size,
                plex_data.len()
            )));
        }

        let mut pieces = Vec::with_capacity(num_pieces);

        // Read all CPs first (they're at the beginning)
        let mut cps = Vec::with_capacity(num_pieces + 1);
        for i in 0..=num_pieces {
            let offset = i * 4;
            let cp = read_u32_le(plex_data, offset).unwrap_or(0);
            cps.push(cp);
        }

        // Now read all PieceDescriptors (they're after all the CPs)
        let struct_offset = (num_pieces + 1) * 4;
        for i in 0..num_pieces {
            let offset = struct_offset + i * PIECE_DESCRIPTOR_SIZE;

            if offset + PIECE_DESCRIPTOR_SIZE > plex_data.len() {
                return Err(PackageError::Corrupted(format!(
                    "PieceDescriptor {i} truncated"
                )));
            }

            let piece_data = &plex_data[offset..offset + PIECE_DESCRIPTOR_SIZE];

            // Parse PieceDescriptor (matches Apache POI's PieceDescriptor constructor)
            let _descriptor = read_u16_le(piece_data, 0).unwrap_or(0);
            let mut fc = read_u32_le(piece_data, 2).unwrap_or(0);
            let _prm = read_u16_le(piece_data, 6).unwrap_or(0);

            // Extract encoding information from fc (bit 30 indicates encoding)
            // From Apache POI PieceDescriptor.java lines 69-76:
            // If bit 30 is clear, this is Unicode (UTF-16LE)
            // If bit 30 is set, this is compressed ANSI (Windows-1252)
            let is_ansi = (fc & 0x40000000) != 0;
            if is_ansi {
                fc &= !0x40000000; // Clear the encoding bit
                fc /= 2; // Adjust for ANSI (1 byte per character vs 2 bytes for Unicode)
            }

            pieces.push(PieceDescriptor {
                cp_start: cps[i],
                cp_end: cps[i + 1],
                file_pos: fc as usize,
                is_ansi,
            });
        }

        Ok(pieces)
    }

    /// Extract text from piece descriptors.
    ///
    /// Based on Apache POI's `TextPieceTable` logic, each piece descriptor maps
    /// a range of character positions (CP) to a location in the `WordDocument` stream.
    fn extract_text_from_piece_descriptors(
        pieces: &[PieceDescriptor],
        word_document: &[u8],
    ) -> Result<Vec<u16>> {
        let mut text = Vec::new();

        for piece in pieces {
            // Character positions in a piece table ascend, but a corrupt or
            // hostile table can violate that. A descending range is skipped
            // rather than allowed to wrap the subtraction.
            let Some(char_count) = piece.cp_end.checked_sub(piece.cp_start) else {
                continue;
            };
            let char_count = char_count as usize;

            if char_count == 0 {
                continue; // Empty piece
            }

            // Calculate text size in bytes based on encoding
            let byte_count = if piece.is_ansi {
                char_count // 1 byte per character for ANSI
            } else {
                char_count.saturating_mul(UTF16_CODE_UNIT_BYTES)
            };

            // Read the text data from the WordDocument stream. A piece that
            // starts past the end contributes nothing; one that merely extends
            // past it contributes the bytes that are present.
            let start = piece.file_pos;
            if start >= word_document.len() {
                continue;
            }
            let end = start
                .checked_add(byte_count)
                .unwrap_or(word_document.len())
                .min(word_document.len());
            let text_data = &word_document[start..end];

            if piece.is_ansi {
                // 8-bit ANSI text (Windows-1252)
                for &byte in text_data {
                    let mut encoded = [0u16; 2];
                    let units = windows_1252_to_char(byte).encode_utf16(&mut encoded);
                    text.extend_from_slice(units);
                }
            } else {
                // 16-bit Unicode (UTF-16LE)
                // Make sure we have complete UTF-16LE pairs
                let utf16_data = if text_data.len().is_multiple_of(2) {
                    text_data
                } else {
                    &text_data[..text_data.len() & !1] // Truncate to even length
                };

                for chunk in utf16_data.chunks_exact(2) {
                    let code_unit = read_u16_le(chunk, 0).unwrap_or(0);
                    text.push(code_unit);
                }
            }
        }

        Ok(text)
    }

    /// Extract text in a simple way (for Word 6.0 or simplified docs).
    ///
    /// In older or simplified DOC files, text may start at a fixed offset
    /// without a piece table.
    fn extract_text_simple(word_document: &[u8]) -> Result<Vec<u16>> {
        // Text typically starts at 0x200 (512) or 0x800 (2048)
        let start_offset = 0x200;

        if word_document.len() <= start_offset {
            return Ok(Vec::new());
        }

        // Try to extract as Windows-1252
        let mut text = Vec::new();
        for &byte in &word_document[start_offset..] {
            // Stop at null terminator or control characters
            if byte == 0 {
                break;
            }
            let mut encoded = [0u16; 2];
            let units = windows_1252_to_char(byte).encode_utf16(&mut encoded);
            text.extend_from_slice(units);
        }

        Ok(text)
    }
}

/// A piece descriptor.
///
/// Maps a range of character positions to a location in the `WordDocument` stream.
#[derive(Debug, Clone)]
struct PieceDescriptor {
    /// Starting character position
    cp_start: u32,
    /// Ending character position
    cp_end: u32,
    /// File position in `WordDocument` stream
    file_pos: usize,
    /// Whether text is ANSI (true) or Unicode (false)
    is_ansi: bool,
}

/// Convert a Windows-1252 byte to a Unicode character.
///
/// Windows-1252 is mostly compatible with ISO-8859-1, but has additional
/// printable characters in the 0x80-0x9F range.
fn windows_1252_to_char(byte: u8) -> char {
    match byte {
        0x80 => '€',
        0x82 => '‚',
        0x83 => 'ƒ',
        0x84 => '„',
        0x85 => '…',
        0x86 => '†',
        0x87 => '‡',
        0x88 => 'ˆ',
        0x89 => '‰',
        0x8A => 'Š',
        0x8B => '‹',
        0x8C => 'Œ',
        0x8E => 'Ž',
        0x91 => '\u{2018}', // Left single quotation mark
        0x92 => '\u{2019}', // Right single quotation mark
        0x93 => '"',
        0x94 => '"',
        0x95 => '•',
        0x96 => '–',
        0x97 => '—',
        0x98 => '˜',
        0x99 => '™',
        0x9A => 'š',
        0x9B => '›',
        0x9C => 'œ',
        0x9E => 'ž',
        0x9F => 'Ÿ',
        _ => byte as char, // For 0x00-0x7F and 0xA0-0xFF, direct conversion works
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_windows_1252_conversion() {
        assert_eq!(windows_1252_to_char(0x41), 'A');
        assert_eq!(windows_1252_to_char(0x80), '€');
        assert_eq!(windows_1252_to_char(0x93), '"');
        assert_eq!(windows_1252_to_char(0x94), '"');
    }

    #[test]
    fn test_clx_parsing_structure() {
        // Test that the CLX structure constants are correct
        assert_eq!(PIECE_DESCRIPTOR_SIZE, 8);

        // Test basic parsing of a minimal PlexOfCps structure
        // Format: [CP0] [CP1] [Struct0]
        // For 1 piece: 2 CPs (8 bytes) + 1 PieceDescriptor (8 bytes) = 16 bytes
        let minimal_plex = [
            // CPs (2 of them for 1 piece)
            0x00, 0x00, 0x00, 0x00, // cp_start = 0
            0x10, 0x00, 0x00, 0x00, // cp_end = 16 (16 characters)
            // PieceDescriptor (8 bytes)
            0x00, 0x00, // descriptor
            0x00, 0x00, 0x00, 0x00, // fc = 0 (Unicode at position 0)
            0x00, 0x00, // prm
        ];

        // Should parse successfully
        let result = TextExtractor::parse_plex_of_cps(&minimal_plex);
        assert!(result.is_ok());
        let pieces = result.unwrap();
        assert_eq!(pieces.len(), 1);
        assert_eq!(pieces[0].cp_start, 0);
        assert_eq!(pieces[0].cp_end, 16);
        assert_eq!(pieces[0].file_pos, 0);
        assert!(!pieces[0].is_ansi); // Bit 30 not set = Unicode
    }

    #[test]
    fn unicode_piece_preserves_surrogate_pairs_in_cp_domain() {
        let utf16 = "A😀B".encode_utf16().collect::<Vec<_>>();
        let word_document = utf16
            .iter()
            .flat_map(|unit| unit.to_le_bytes())
            .collect::<Vec<_>>();
        let pieces = [PieceDescriptor {
            cp_start: 0,
            cp_end: 4,
            file_pos: 0,
            is_ansi: false,
        }];

        let extracted =
            TextExtractor::extract_text_from_piece_descriptors(&pieces, &word_document).unwrap();
        assert_eq!(extracted, utf16);
        assert_eq!(String::from_utf16(&extracted).unwrap(), "A😀B");
    }

    #[test]
    fn text_ranges_are_utf16_code_unit_positions() {
        let utf16 = "A😀B".encode_utf16().collect::<Vec<_>>();
        let extractor = TextExtractor {
            text: Arc::new(String::from_utf16(&utf16).unwrap()),
            cp_to_byte: Arc::from(TextExtractor::build_cp_to_byte_map("A😀B").into_boxed_slice()),
        };

        assert_eq!(extractor.text(), "A😀B");
        assert_eq!(extractor.text_at_range(0, 1), "A");
        assert_eq!(extractor.text_at_range(1, 3), "😀");
        assert_eq!(extractor.text_at_range(3, 4), "B");
        assert_eq!(extractor.text_at_range(1, 2), "");
        assert_eq!(extractor.text_at_range(2, 3), "");
        assert_eq!(extractor.text_at_range(99, 100), "");
    }

    #[test]
    fn shared_text_and_cp_map_retain_one_decoded_source() {
        let extractor = TextExtractor {
            text: Arc::new("A😀B".to_string()),
            cp_to_byte: Arc::from(TextExtractor::build_cp_to_byte_map("A😀B").into_boxed_slice()),
        };

        let text = extractor.text_shared();
        let cp_to_byte = extractor.cp_to_byte_shared();
        assert!(Arc::ptr_eq(&text, &extractor.text));
        assert!(Arc::ptr_eq(&cp_to_byte, &extractor.cp_to_byte));
        assert_eq!(text.as_str(), "A😀B");
        assert_eq!(&*cp_to_byte, &[0, 1, 1, 5]);
    }
}
