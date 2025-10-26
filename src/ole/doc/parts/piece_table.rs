/// Piece Table parser for DOC files.
///
/// Based on Apache POI's ComplexFileTable and TextPieceTable.
/// The piece table maps Character Positions (CP) to File Characters (FC)
/// in the WordDocument stream, handling text stored in different locations.
///
/// References:
/// - org.apache.poi.hwpf.model.ComplexFileTable
/// - org.apache.poi.hwpf.model.TextPieceTable
/// - org.apache.poi.hwpf.model.TextPiece
/// - org.apache.poi.hwpf.model.PieceDescriptor
/// - [MS-DOC] 2.4.1 Clx (Complex file information)
/// - [MS-DOC] 2.9.179 Pcd (Piece Descriptor)
use crate::common::binary::{read_u16_le, read_u32_le};
use crate::ole::plcf::PlcfParser;

/// A text piece - maps a range of CPs to an FC in the WordDocument stream.
///
/// Based on Apache POI's TextPiece.
#[derive(Debug, Clone)]
pub struct TextPiece {
    /// Start character position (CP)
    pub cp_start: u32,
    /// End character position (CP)
    pub cp_end: u32,
    /// File character position (FC) - byte offset in WordDocument stream
    pub fc: u32,
    /// Whether the text is Unicode (true) or single-byte (false)
    pub is_unicode: bool,
}

impl TextPiece {
    /// Get the length in characters.
    #[inline]
    pub fn length(&self) -> u32 {
        self.cp_end - self.cp_start
    }

    /// Convert a CP within this piece to an FC.
    pub fn cp_to_fc(&self, cp: u32) -> Option<u32> {
        if cp < self.cp_start || cp > self.cp_end {
            return None;
        }

        let offset = cp - self.cp_start;
        let byte_offset = if self.is_unicode {
            offset * 2 // Unicode is 2 bytes per character
        } else {
            offset // Single-byte encoding
        };

        Some(self.fc + byte_offset)
    }

    /// Convert an FC to a CP within this piece.
    pub fn fc_to_cp(&self, fc: u32) -> Option<u32> {
        if fc < self.fc {
            return None;
        }

        let byte_offset = fc - self.fc;
        let char_offset = if self.is_unicode {
            byte_offset / 2
        } else {
            byte_offset
        };

        let cp = self.cp_start + char_offset;
        if cp > self.cp_end { None } else { Some(cp) }
    }
}

/// Piece Table - manages the mapping between CP and FC.
///
/// Based on Apache POI's TextPieceTable.
#[derive(Debug, Clone)]
pub struct PieceTable {
    /// All text pieces, sorted by CP
    pieces: Vec<TextPiece>,
}

impl PieceTable {
    /// Parse a piece table from CLX (Complex file information) data.
    ///
    /// Based on Apache POI's ComplexFileTable.parse().
    ///
    /// # Arguments
    ///
    /// * `clx_data` - The CLX data from table stream (at fcClx/lcbClx in FIB)
    ///
    /// # Returns
    ///
    /// Parsed piece table or None if invalid
    pub fn parse(clx_data: &[u8]) -> Option<Self> {
        if clx_data.is_empty() {
            return None;
        }

        let mut offset = 0;

        // CLX structure (from POI's ComplexFileTable.java):
        // - RgPrc (array of Prc - property modifiers) - type 0x01, can be multiple
        // - Pcdt (Piece Descriptor table) - type 0x02
        //
        // POI line 54: while (tableStream[offset] == GRPPRL_TYPE)
        // where GRPPRL_TYPE = 1, TEXT_PIECE_TABLE_TYPE = 2

        // Skip RgPrc entries (type 0x01 only!)
        while offset < clx_data.len() && clx_data[offset] == 0x01 {
            offset += 1;
            if offset + 2 > clx_data.len() {
                return None;
            }
            // Read size as SHORT (2 bytes) - POI line 56
            let size = read_u16_le(clx_data, offset).unwrap_or(0) as usize;
            offset += 2;

            if offset + size > clx_data.len() {
                return None;
            }
            offset += size;
        }

        // Now we should be at the Pcdt marker (0x02)
        if offset >= clx_data.len() || clx_data[offset] != 0x02 {
            return None;
        }

        // Skip the 0x02 marker - POI line 70: ++offset
        offset += 1;

        if offset + 4 > clx_data.len() {
            return None;
        }

        // Read lcb as INT (4 bytes) - POI line 70
        let lcb = read_u32_le(clx_data, offset).unwrap_or(0) as usize;
        offset += 4;

        if offset + lcb > clx_data.len() {
            return None;
        }

        let plcpcd_data = &clx_data[offset..offset + lcb];

        // Parse PlcPcd using PLCF parser
        // Each Pcd is 8 bytes (according to [MS-DOC])
        let plcf = PlcfParser::parse(plcpcd_data, 8)?;

        let mut pieces = Vec::new();

        // Extract TextPieces from PlcPcd
        for i in 0..plcf.count() {
            let (cp_start, cp_end) = plcf.range(i)?;
            let pcd_data = plcf.property(i)?;

            if pcd_data.len() < 8 {
                continue;
            }

            // Parse Pcd (Piece Descriptor)
            // Bytes 0-1: flags (bit 6 = fNoParaLast, others reserved)
            // Bytes 2-5: fc (File Character position)
            // Bytes 6-7: prm (Property modifier - for paragraph/character properties)

            let fc_raw = read_u32_le(pcd_data, 2).unwrap_or(0);

            // FC encoding (from POI's PieceDescriptor.java):
            // - Bit 30 (0x40000000): if CLEAR (0), text is Unicode (UTF-16LE)
            // - If SET (1), text is single-byte (ANSI/codepage) and fc must be divided by 2
            // This is the actual file offset in the WordDocument stream
            let is_unicode = (fc_raw & 0x40000000) == 0;
            let mut fc = fc_raw & 0x3FFFFFFF; // Clear bit 30

            // For non-Unicode text, divide fc by 2 (POI line 74-75)
            if !is_unicode {
                fc /= 2;
            }

            pieces.push(TextPiece {
                cp_start,
                cp_end,
                fc,
                is_unicode,
            });
        }

        // Sort pieces by CP (should already be sorted, but ensure it)
        pieces.sort_by_key(|p| p.cp_start);

        Some(Self { pieces })
    }

    /// Get all text pieces.
    #[inline]
    pub fn pieces(&self) -> &[TextPiece] {
        &self.pieces
    }

    /// Find the text piece containing a given CP.
    pub fn piece_for_cp(&self, cp: u32) -> Option<&TextPiece> {
        // Binary search for efficiency
        self.pieces
            .iter()
            .find(|piece| cp >= piece.cp_start && cp < piece.cp_end)
    }

    /// Convert a CP to an FC.
    pub fn cp_to_fc(&self, cp: u32) -> Option<u32> {
        let piece = self.piece_for_cp(cp)?;
        piece.cp_to_fc(cp)
    }

    /// Convert an FC to a CP.
    pub fn fc_to_cp(&self, fc: u32) -> Option<u32> {
        // Linear search through pieces
        // Could be optimized with a secondary index if needed
        for piece in &self.pieces {
            if let Some(cp) = piece.fc_to_cp(fc) {
                return Some(cp);
            }
        }
        None
    }

    /// Get the total number of characters (last CP).
    pub fn total_cps(&self) -> u32 {
        self.pieces.last().map(|p| p.cp_end).unwrap_or(0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_piece_cp_to_fc() {
        let piece = TextPiece {
            cp_start: 100,
            cp_end: 200,
            fc: 500,
            is_unicode: true,
        };

        // CP 100 -> FC 500
        assert_eq!(piece.cp_to_fc(100), Some(500));
        // CP 150 -> FC 500 + (150-100)*2 = 600
        assert_eq!(piece.cp_to_fc(150), Some(600));
        // CP 200 -> FC 500 + (200-100)*2 = 700
        assert_eq!(piece.cp_to_fc(200), Some(700));
        // CP outside range
        assert_eq!(piece.cp_to_fc(50), None);
        assert_eq!(piece.cp_to_fc(250), None);
    }

    #[test]
    fn test_text_piece_fc_to_cp() {
        let piece = TextPiece {
            cp_start: 100,
            cp_end: 200,
            fc: 500,
            is_unicode: false, // Single-byte
        };

        // FC 500 -> CP 100
        assert_eq!(piece.fc_to_cp(500), Some(100));
        // FC 550 -> CP 150
        assert_eq!(piece.fc_to_cp(550), Some(150));
        // FC 600 -> CP 200
        assert_eq!(piece.fc_to_cp(600), Some(200));
        // FC outside range
        assert_eq!(piece.fc_to_cp(400), None);
        assert_eq!(piece.fc_to_cp(700), None);
    }
}
