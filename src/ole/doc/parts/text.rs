/// Text extraction from DOC files.
///
/// This module handles extracting text from the binary structures in DOC files.
/// Text in DOC files is stored in a complex way:
/// - The actual text bytes are in the WordDocument stream
/// - A "Piece Table" (CLX structure) in the Table stream maps character positions to file positions
/// - Text can be in either 8-bit (Windows-1252) or 16-bit (UTF-16LE) format
use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;

/// Text extractor for DOC files.
///
/// Handles the complex text extraction process from DOC binary structures.
pub struct TextExtractor {
    /// The extracted text
    text: String,
}

impl TextExtractor {
    /// Create a new TextExtractor and extract text.
    ///
    /// # Arguments
    ///
    /// * `fib` - The File Information Block
    /// * `word_document` - The WordDocument stream
    /// * `table_stream` - The Table stream (0Table or 1Table)
    pub fn new(
        fib: &FileInformationBlock,
        word_document: &[u8],
        table_stream: &[u8],
    ) -> Result<Self> {
        // Extract text using the piece table
        let text = Self::extract_text_from_pieces(fib, word_document, table_stream)?;

        Ok(Self { text })
    }

    /// Extract all text from the document.
    pub fn extract_all_text(&self) -> Result<String> {
        Ok(self.text.clone())
    }

    /// Extract text using the piece table (CLX structure).
    ///
    /// The piece table maps character positions in the document to file positions
    /// in the WordDocument stream. This allows Word to efficiently handle insertions
    /// and deletions without rewriting the entire file.
    fn extract_text_from_pieces(
        fib: &FileInformationBlock,
        word_document: &[u8],
        table_stream: &[u8],
    ) -> Result<String> {
        // Get the CLX (piece table) location from FIB
        // CLX is at FibRgFcLcb index 1 (fcClx, lcbClx)
        let (clx_offset, clx_length) = fib
            .get_table_pointer(1)
            .ok_or_else(|| DocError::Corrupted("CLX pointer not found in FIB".to_string()))?;

        if clx_length == 0 {
            // No piece table means text starts at offset 0x200 or 0x800
            // This is a simplified doc or Word 6.0 format
            return Self::extract_text_simple(word_document);
        }

        // Extract the CLX from the table stream
        let clx_offset = clx_offset as usize;
        let clx_length = clx_length as usize;

        if clx_offset + clx_length > table_stream.len() {
            return Err(DocError::Corrupted("CLX extends beyond table stream".to_string()));
        }

        let clx_data = &table_stream[clx_offset..clx_offset + clx_length];

        // Parse the piece table from CLX
        Self::parse_piece_table(clx_data, word_document)
    }

    /// Parse the piece table and extract text.
    ///
    /// The CLX structure contains:
    /// - One or more PRC structures (property records) - we can skip these
    /// - A PCD structure (piece descriptor array) containing the actual pieces
    fn parse_piece_table(clx_data: &[u8], word_document: &[u8]) -> Result<String> {
        let mut offset = 0;
        let mut pieces = Vec::new();

        // Skip PRC structures (type 0x01) until we find PCD (type 0x02)
        while offset < clx_data.len() {
            if offset + 1 > clx_data.len() {
                break;
            }

            let clxt = clx_data[offset];
            offset += 1;

            if clxt == 0x02 {
                // This is the PCD (piece descriptor array)
                if offset + 4 > clx_data.len() {
                    return Err(DocError::Corrupted("PCD structure truncated".to_string()));
                }

                // Read the size of the PCD
                let pcd_size = u32::from_le_bytes([
                    clx_data[offset],
                    clx_data[offset + 1],
                    clx_data[offset + 2],
                    clx_data[offset + 3],
                ]) as usize;
                offset += 4;

                if offset + pcd_size > clx_data.len() {
                    return Err(DocError::Corrupted("PCD extends beyond CLX".to_string()));
                }

                let pcd_data = &clx_data[offset..offset + pcd_size];

                // Parse the pieces
                pieces = Self::parse_pcd(pcd_data)?;
                break;
            } else if clxt == 0x01 {
                // PRC - skip it
                if offset + 2 > clx_data.len() {
                    return Err(DocError::Corrupted("PRC structure truncated".to_string()));
                }

                let prc_size = u16::from_le_bytes([clx_data[offset], clx_data[offset + 1]]) as usize;
                offset += 2 + prc_size;
            } else {
                return Err(DocError::Corrupted(format!("Unknown CLX type: 0x{:02X}", clxt)));
            }
        }

        if pieces.is_empty() {
            return Err(DocError::Corrupted("No pieces found in CLX".to_string()));
        }

        // Extract text from pieces
        Self::extract_text_from_piece_descriptors(&pieces, word_document)
    }

    /// Parse the PCD (piece descriptor array).
    ///
    /// PCD format:
    /// - Array of CP (character positions) - n+1 entries of 4 bytes each
    /// - Array of PieceDescriptors - n entries of 8 bytes each
    fn parse_pcd(pcd_data: &[u8]) -> Result<Vec<PieceDescriptor>> {
        // Number of pieces = (length - 4) / 12
        // The first (n+1)*4 bytes are character positions
        // The next n*8 bytes are piece descriptors
        let num_pieces = (pcd_data.len() - 4) / 12;

        if num_pieces == 0 {
            return Ok(Vec::new());
        }

        let mut pieces = Vec::with_capacity(num_pieces);

        for i in 0..num_pieces {
            let cp_offset = i * 4;
            let pd_offset = (num_pieces + 1) * 4 + i * 8;

            if cp_offset + 8 > pcd_data.len() || pd_offset + 8 > pcd_data.len() {
                break;
            }

            // Read CP (character position)
            let cp_start = u32::from_le_bytes([
                pcd_data[cp_offset],
                pcd_data[cp_offset + 1],
                pcd_data[cp_offset + 2],
                pcd_data[cp_offset + 3],
            ]);

            let cp_end = u32::from_le_bytes([
                pcd_data[cp_offset + 4],
                pcd_data[cp_offset + 5],
                pcd_data[cp_offset + 6],
                pcd_data[cp_offset + 7],
            ]);

            // Read piece descriptor
            // Bytes 2-5: file position (with encoding flag in bit 30)
            let fc = u32::from_le_bytes([
                pcd_data[pd_offset + 2],
                pcd_data[pd_offset + 3],
                pcd_data[pd_offset + 4],
                pcd_data[pd_offset + 5],
            ]);

            // Bit 30 of fc determines encoding:
            // 0 = 16-bit Unicode (UTF-16LE)
            // 1 = 8-bit ANSI (Windows-1252)
            let is_ansi = (fc & 0x40000000) != 0;
            let file_pos = (fc & 0x3FFFFFFF) as usize;

            // Adjust file position for ANSI text
            let actual_file_pos = if is_ansi { file_pos / 2 } else { file_pos };

            pieces.push(PieceDescriptor {
                cp_start,
                cp_end,
                file_pos: actual_file_pos,
                is_ansi,
            });
        }

        Ok(pieces)
    }

    /// Extract text from piece descriptors.
    fn extract_text_from_piece_descriptors(
        pieces: &[PieceDescriptor],
        word_document: &[u8],
    ) -> Result<String> {
        let mut text = String::new();

        for piece in pieces {
            let char_count = (piece.cp_end - piece.cp_start) as usize;

            if piece.is_ansi {
                // 8-bit ANSI text (Windows-1252)
                let start = piece.file_pos;
                let end = start + char_count;

                if end > word_document.len() {
                    continue;
                }

                // Decode Windows-1252 to UTF-8
                for &byte in &word_document[start..end] {
                    text.push(windows_1252_to_char(byte));
                }
            } else {
                // 16-bit Unicode (UTF-16LE)
                let start = piece.file_pos;
                let end = start + (char_count * 2);

                if end > word_document.len() {
                    continue;
                }

                // Decode UTF-16LE
                let utf16_data = &word_document[start..end];
                let mut utf16_chars = Vec::with_capacity(char_count);

                for chunk in utf16_data.chunks_exact(2) {
                    let code_unit = u16::from_le_bytes([chunk[0], chunk[1]]);
                    utf16_chars.push(code_unit);
                }

                text.push_str(&String::from_utf16_lossy(&utf16_chars));
            }
        }

        Ok(text)
    }

    /// Extract text in a simple way (for Word 6.0 or simplified docs).
    ///
    /// In older or simplified DOC files, text may start at a fixed offset
    /// without a piece table.
    fn extract_text_simple(word_document: &[u8]) -> Result<String> {
        // Text typically starts at 0x200 (512) or 0x800 (2048)
        let start_offset = 0x200;

        if word_document.len() <= start_offset {
            return Ok(String::new());
        }

        // Try to extract as Windows-1252
        let mut text = String::new();
        for &byte in &word_document[start_offset..] {
            // Stop at null terminator or control characters
            if byte == 0 {
                break;
            }
            text.push(windows_1252_to_char(byte));
        }

        Ok(text)
    }
}

/// A piece descriptor.
///
/// Maps a range of character positions to a location in the WordDocument stream.
#[derive(Debug, Clone)]
struct PieceDescriptor {
    /// Starting character position
    cp_start: u32,
    /// Ending character position
    cp_end: u32,
    /// File position in WordDocument stream
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
}

