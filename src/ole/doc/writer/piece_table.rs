//! Piece Table generation for DOC files
//!
//! The Piece Table (Clx structure) maps character positions in the document
//! to actual file locations in the text stream. This allows for efficient
//! editing without rewriting the entire file.
//!
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's PieceDescriptor.

/// Error type for DOC operations
pub type DocError = std::io::Error;

/// Piece descriptor - maps a range of characters to a file location
#[derive(Debug, Clone)]
pub struct Piece {
    /// Character position start in document
    pub cp_start: u32,
    /// Character position end in document
    pub cp_end: u32,
    /// File character position in text stream
    pub fc: u32,
    /// Is Unicode text (true) or ANSI (false)
    pub is_unicode: bool,
}

impl Piece {
    /// Create a new piece
    pub fn new(cp_start: u32, cp_end: u32, fc: u32, is_unicode: bool) -> Self {
        Self {
            cp_start,
            cp_end,
            fc,
            is_unicode,
        }
    }

    /// Get the length of this piece in characters
    pub fn len(&self) -> u32 {
        self.cp_end - self.cp_start
    }

    /// Check if piece is empty
    pub fn is_empty(&self) -> bool {
        self.cp_end <= self.cp_start
    }
}

/// Piece Table builder
#[derive(Debug)]
pub struct PieceTableBuilder {
    pieces: Vec<Piece>,
}

impl PieceTableBuilder {
    /// Create a new piece table builder
    pub fn new() -> Self {
        Self { pieces: Vec::new() }
    }

    /// Add a piece to the table
    pub fn add_piece(&mut self, piece: Piece) {
        self.pieces.push(piece);
    }

    /// Generate the piece table (Clx structure) as bytes
    ///
    /// # Returns
    ///
    /// Clx structure containing piece descriptors
    ///
    /// Based on Apache POI's ComplexFileTable.writeTo() (line 96-106)
    pub fn generate(&self) -> Result<Vec<u8>, DocError> {
        let mut clx = Vec::new();

        // Prc (piece table header) - POI line 98
        // Apache POI uses TEXT_PIECE_TABLE_TYPE = 2 (ComplexFileTable.java line 38-39, 98)
        clx.push(0x02); // TEXT_PIECE_TABLE_TYPE

        // Calculate size
        let piece_count = self.pieces.len();
        let piece_table_size = (piece_count + 1) * 4 + piece_count * 8;

        // CRITICAL FIX: POI writes size as 4-byte INT (line 102-104), not 2-byte short!
        clx.extend_from_slice(&(piece_table_size as u32).to_le_bytes());

        // PlcPcd (piece table)
        // First, write CP positions
        for piece in &self.pieces {
            clx.extend_from_slice(&piece.cp_start.to_le_bytes());
        }
        // Write final CP
        if let Some(last) = self.pieces.last() {
            clx.extend_from_slice(&last.cp_end.to_le_bytes());
        } else {
            clx.extend_from_slice(&0u32.to_le_bytes());
        }

        // Write piece descriptors (PCD)
        // Based on Apache POI's PieceDescriptor.toByteArray() lines 118-133
        for piece in &self.pieces {
            // Encode FC position according to POI logic
            let mut fc_encoded = piece.fc;
            if !piece.is_unicode {
                // CRITICAL: For ANSI text, multiply FC by 2 (POI line 122)
                fc_encoded *= 2;
                fc_encoded |= 0x40000000; // Bit 30: compressed (ANSI) text
            }

            // PCD structure (8 bytes): descriptor(2) + fc(4) + prm(2)
            clx.extend_from_slice(&0u16.to_le_bytes()); // descriptor (unused)
            clx.extend_from_slice(&fc_encoded.to_le_bytes()); // fcCompressed
            clx.extend_from_slice(&0u16.to_le_bytes()); // prm (unused)
        }

        Ok(clx)
    }

    /// Get the number of pieces
    pub fn piece_count(&self) -> usize {
        self.pieces.len()
    }
}

impl Default for PieceTableBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_piece_creation() {
        let piece = Piece::new(0, 100, 0, true);
        assert_eq!(piece.len(), 100);
        assert!(!piece.is_empty());
    }

    #[test]
    fn test_piece_table_generation() {
        let mut builder = PieceTableBuilder::new();
        builder.add_piece(Piece::new(0, 50, 0, true));
        builder.add_piece(Piece::new(50, 100, 100, true));

        let clx = builder.generate().unwrap();
        assert!(!clx.is_empty());
        assert_eq!(clx[0], 0x02); // Type: TEXT_PIECE_TABLE_TYPE (Apache POI)
    }
}
