use crate::plcf::Plcf;
/// Piece Table parser for DOC files.
///
/// Based on Apache POI's `ComplexFileTable` and `TextPieceTable`.
/// The piece table maps Character Positions (CP) to File Characters (FC)
/// in the `WordDocument` stream, handling text stored in different locations.
///
/// References:
/// - org.apache.poi.hwpf.model.ComplexFileTable
/// - org.apache.poi.hwpf.model.TextPieceTable
/// - org.apache.poi.hwpf.model.TextPiece
/// - org.apache.poi.hwpf.model.PieceDescriptor
/// - [MS-DOC] 2.4.1 Clx (Complex file information)
/// - [MS-DOC] 2.9.179 Pcd (Piece Descriptor)
use litchi_core::binary::{read_u16_le, read_u32_le};

/// Bytes per character in a Unicode (UTF-16LE) text piece.
const UNICODE_CHAR_BYTES: u32 = 2;

/// Bytes per character in a single-byte (ANSI) text piece.
const ANSI_CHAR_BYTES: u32 = 1;

/// A text piece - maps a range of CPs to an FC in the `WordDocument` stream.
///
/// Based on Apache POI's `TextPiece`.
#[derive(Debug, Clone)]
pub struct TextPiece {
    /// Start character position (CP)
    pub cp_start: u32,
    /// End character position (CP)
    pub cp_end: u32,
    /// File character position (FC) - byte offset in `WordDocument` stream
    pub fc: u32,
    /// Whether the text is Unicode (true) or single-byte (false)
    pub is_unicode: bool,
    /// Additional fast-save SPRMs resolved from `Pcd.Prm`.
    property_modifier: Vec<u8>,
}

impl TextPiece {
    /// Get the length in characters.
    ///
    /// A piece table's character positions ascend, but a corrupt or hostile
    /// table can violate that; such a piece reports no length rather than
    /// wrapping the subtraction.
    #[inline]
    #[must_use]
    pub fn length(&self) -> u32 {
        self.cp_end.saturating_sub(self.cp_start)
    }

    /// Convert a CP within this piece to an FC.
    #[must_use]
    pub fn cp_to_fc(&self, cp: u32) -> Option<u32> {
        if cp < self.cp_start || cp > self.cp_end {
            return None;
        }

        let offset = cp.checked_sub(self.cp_start)?;
        let byte_offset = if self.is_unicode {
            offset.checked_mul(UNICODE_CHAR_BYTES)? // Unicode is 2 bytes per character
        } else {
            offset // Single-byte encoding
        };

        self.fc.checked_add(byte_offset)
    }

    /// Convert an FC to a CP within this piece.
    #[must_use]
    pub fn fc_to_cp(&self, fc: u32) -> Option<u32> {
        if fc < self.fc {
            return None;
        }

        let byte_offset = fc - self.fc;
        let char_offset = if self.is_unicode {
            byte_offset / UNICODE_CHAR_BYTES
        } else {
            byte_offset
        };

        let cp = self.cp_start.checked_add(char_offset)?;
        if cp > self.cp_end { None } else { Some(cp) }
    }

    /// Exclusive FC boundary for this piece.
    #[inline]
    fn fc_end(&self) -> u32 {
        let bytes_per_character = if self.is_unicode {
            UNICODE_CHAR_BYTES
        } else {
            ANSI_CHAR_BYTES
        };
        self.fc
            .saturating_add(self.length().saturating_mul(bytes_per_character))
    }

    /// Additional property modifiers applied after FKP properties.
    #[inline]
    #[must_use]
    pub fn property_modifier(&self) -> &[u8] {
        &self.property_modifier
    }
}

/// Piece Table - manages the mapping between CP and FC.
///
/// Based on Apache POI's `TextPieceTable`.
#[derive(Debug, Clone)]
pub struct PieceTable {
    /// All text pieces, sorted by CP
    pieces: Vec<TextPiece>,
    /// The same pieces ordered by physical FC, with a prefix maximum end.
    ///
    /// Fast-save documents may place logical pieces out of physical order and
    /// may contain overlapping physical intervals. The prefix maximum lets a
    /// range lookup skip every earlier piece that cannot overlap without
    /// assuming that physical intervals are disjoint.
    physical_index: Vec<PhysicalPiece>,
}

#[derive(Debug, Clone)]
struct PhysicalPiece {
    piece_index: usize,
    prefix_max_end_fc: u32,
}

impl PieceTable {
    /// Parse a piece table from CLX (Complex file information) data.
    ///
    /// Based on Apache POI's `ComplexFileTable.parse()`.
    ///
    /// # Arguments
    ///
    /// * `clx_data` - The CLX data from table stream (at fcClx/lcbClx in FIB)
    ///
    /// # Returns
    ///
    /// Parsed piece table or None if invalid
    #[must_use]
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

        // Parse RgPrc entries (type 0x01) so Prm1 references can be resolved.
        let mut property_groups = Vec::new();
        while offset < clx_data.len() && clx_data[offset] == 0x01 {
            offset += 1;
            if offset + 2 > clx_data.len() {
                return None;
            }
            // Read size as SHORT (2 bytes) - POI line 56
            let size = read_u16_le(clx_data, offset).unwrap_or(0) as usize;
            offset += 2;

            if size > 0x3fa2 || offset + size > clx_data.len() {
                return None;
            }
            property_groups.push(clx_data[offset..offset + size].to_vec());
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
        let plcf = Plcf::parse(plcpcd_data, 8)?;

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

            let prm = read_u16_le(pcd_data, 6).unwrap_or(0);
            let property_modifier = Self::resolve_property_modifier(prm, &property_groups)?;

            pieces.push(TextPiece {
                cp_start,
                cp_end,
                fc,
                is_unicode,
                property_modifier,
            });
        }

        // Sort pieces by CP (should already be sorted, but ensure it)
        pieces.sort_by_key(|p| p.cp_start);

        let physical_index = Self::build_physical_index(&pieces);

        Some(Self {
            pieces,
            physical_index,
        })
    }

    /// Get all text pieces.
    #[inline]
    #[must_use]
    pub fn pieces(&self) -> &[TextPiece] {
        &self.pieces
    }

    fn build_physical_index(pieces: &[TextPiece]) -> Vec<PhysicalPiece> {
        let mut physical_index: Vec<_> = (0..pieces.len())
            .map(|piece_index| PhysicalPiece {
                piece_index,
                prefix_max_end_fc: 0,
            })
            .collect();
        physical_index.sort_unstable_by_key(|entry| {
            let piece = &pieces[entry.piece_index];
            (
                piece.fc,
                piece.fc_end(),
                piece.cp_start,
                piece.cp_end,
                entry.piece_index,
            )
        });

        let mut prefix_max_end_fc = 0;
        for entry in &mut physical_index {
            prefix_max_end_fc = prefix_max_end_fc.max(pieces[entry.piece_index].fc_end());
            entry.prefix_max_end_fc = prefix_max_end_fc;
        }
        physical_index
    }

    /// Find the text piece containing a given CP.
    #[must_use]
    pub fn piece_for_cp(&self, cp: u32) -> Option<&TextPiece> {
        // Binary search for efficiency
        self.pieces
            .iter()
            .find(|piece| cp >= piece.cp_start && cp < piece.cp_end)
    }

    fn resolve_property_modifier(prm: u16, property_groups: &[Vec<u8>]) -> Option<Vec<u8>> {
        if prm & 1 != 0 {
            return property_groups.get(usize::from(prm >> 1)).cloned();
        }

        let isprm = ((prm & 0x00fe) >> 1) as u8;
        let value = (prm >> 8) as u8;
        if isprm == 0 && value == 0 {
            return Some(Vec::new());
        }
        let opcode = Self::prm0_opcode(isprm)?;
        let mut grpprl = Vec::with_capacity(3);
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.push(value);
        Some(grpprl)
    }

    /// Expand the compact `Prm0.isprm` table from [MS-DOC] 2.9.215.
    fn prm0_opcode(isprm: u8) -> Option<u16> {
        Some(match isprm {
            0x00 => 0x2879, // sprmCLbcCRJ
            0x04 => 0x2602, // sprmPIncLvl
            0x05 => 0x2403, // sprmPJc
            0x07 => 0x2405, // sprmPFKeep
            0x08 => 0x2406, // sprmPFKeepFollow
            0x09 => 0x2407, // sprmPFPageBreakBefore
            0x0c => 0x260a, // sprmPIlvl
            0x0d => 0x2466, // sprmPFMirrorIndents
            0x0e => 0x240c, // sprmPFNoLineNumb
            0x0f => 0x2467, // sprmPTtwo
            0x18 => 0x2416, // sprmPFInTable
            0x19 => 0x2417, // sprmPFTtp
            0x1d => 0x261b, // sprmPPc
            0x25 => 0x2423, // sprmPWr
            0x2c => 0x242a, // sprmPFNoAutoHyph
            0x32 => 0x2430, // sprmPFLocked
            0x33 => 0x2431, // sprmPFWidowControl
            0x35 => 0x2433, // sprmPFKinsoku
            0x36 => 0x2434, // sprmPFWordWrap
            0x37 => 0x2435, // sprmPFOverflowPunct
            0x38 => 0x2436, // sprmPFTopLinePunct
            0x39 => 0x2437, // sprmPFAutoSpaceDE
            0x3a => 0x2438, // sprmPFAutoSpaceDN
            0x41 => 0x0800, // sprmCFRMarkDel
            0x42 => 0x0801, // sprmCFRMarkIns
            0x43 => 0x0802, // sprmCFFldVanish
            0x47 => 0x0806, // sprmCFData
            0x4b => 0x080a, // sprmCFOle2
            0x4d => 0x2a0c, // sprmCHighlight
            0x4e => 0x0858, // sprmCFEmboss
            0x4f => 0x2859, // sprmCSfxText
            0x50 => 0x0811, // sprmCFWebHidden
            0x51 => 0x0818, // sprmCFSpecVanish
            0x53 => 0x2a33, // sprmCPlain
            0x55 => 0x0835, // sprmCFBold
            0x56 => 0x0836, // sprmCFItalic
            0x57 => 0x0837, // sprmCFStrike
            0x58 => 0x0838, // sprmCFOutline
            0x59 => 0x0839, // sprmCFShadow
            0x5a => 0x083a, // sprmCFSmallCaps
            0x5b => 0x083b, // sprmCFCaps
            0x5c => 0x083c, // sprmCFVanish
            0x5e => 0x2a3e, // sprmCKul
            0x62 => 0x2a42, // sprmCIco
            0x68 => 0x2a48, // sprmCIss
            0x73 => 0x2a53, // sprmCFDStrike
            0x74 => 0x0854, // sprmCFImprint
            0x75 => 0x0855, // sprmCFSpec
            0x76 => 0x0856, // sprmCFObj
            0x78 => 0x2640, // sprmPOutLvl
            0x7b => 0x2a90, // sprmCFSdtVanish
            0x7c => 0x2a86, // sprmCNeedFontFixup
            0x7e => 0x2443, // sprmPFNumRMIns
            _ => return None,
        })
    }

    /// Convert a CP to an FC.
    #[must_use]
    pub fn cp_to_fc(&self, cp: u32) -> Option<u32> {
        let piece = self.piece_for_cp(cp)?;
        piece.cp_to_fc(cp)
    }

    /// Convert an FC to a CP.
    #[must_use]
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

    /// Convert a physical FC interval into all logical CP intervals it intersects.
    ///
    /// Fast-saved documents can store logically adjacent text pieces in disjoint
    /// physical locations. A single FKP range therefore can map to more than one
    /// CP interval; returning every intersection avoids inventing CPs from raw FCs.
    #[must_use]
    pub fn fc_range_to_cp_ranges(&self, start_fc: u32, end_fc: u32) -> Vec<(u32, u32)> {
        if start_fc >= end_fc {
            return Vec::new();
        }

        let first_possible = self
            .physical_index
            .partition_point(|entry| entry.prefix_max_end_fc <= start_fc);
        let after_last_possible = self
            .physical_index
            .partition_point(|entry| self.pieces[entry.piece_index].fc < end_fc);

        let mut ranges = Vec::new();
        for entry in &self.physical_index[first_possible..after_last_possible] {
            let piece = &self.pieces[entry.piece_index];
            let intersection_start = start_fc.max(piece.fc);
            let intersection_end = end_fc.min(piece.fc_end());
            if intersection_start >= intersection_end {
                continue;
            }

            let Some(start_cp) = piece.fc_to_cp(intersection_start) else {
                continue;
            };
            let Some(end_cp) = piece.fc_to_cp(intersection_end) else {
                continue;
            };
            if start_cp < end_cp {
                ranges.push((start_cp, end_cp));
            }
        }

        ranges.sort_unstable();
        ranges
    }

    /// Get the total number of characters (last CP).
    #[must_use]
    pub fn total_cps(&self) -> u32 {
        self.pieces.last().map_or(0, |p| p.cp_end)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table_from_pieces(mut pieces: Vec<TextPiece>) -> PieceTable {
        pieces.sort_by_key(|piece| piece.cp_start);
        let physical_index = PieceTable::build_physical_index(&pieces);
        PieceTable {
            pieces,
            physical_index,
        }
    }

    fn scalar_fc_range_to_cp_ranges(
        table: &PieceTable,
        start_fc: u32,
        end_fc: u32,
    ) -> Vec<(u32, u32)> {
        if start_fc >= end_fc {
            return Vec::new();
        }

        let mut ranges = Vec::new();
        for piece in &table.pieces {
            let intersection_start = start_fc.max(piece.fc);
            let intersection_end = end_fc.min(piece.fc_end());
            if intersection_start >= intersection_end {
                continue;
            }
            let Some(start_cp) = piece.fc_to_cp(intersection_start) else {
                continue;
            };
            let Some(end_cp) = piece.fc_to_cp(intersection_end) else {
                continue;
            };
            if start_cp < end_cp {
                ranges.push((start_cp, end_cp));
            }
        }
        ranges.sort_unstable();
        ranges
    }

    #[test]
    fn test_text_piece_cp_to_fc() {
        let piece = TextPiece {
            cp_start: 100,
            cp_end: 200,
            fc: 500,
            is_unicode: true,
            property_modifier: Vec::new(),
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
            property_modifier: Vec::new(),
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

    #[test]
    fn maps_fkp_ranges_across_discontiguous_text_pieces() {
        let table = table_from_pieces(vec![
            TextPiece {
                cp_start: 0,
                cp_end: 3,
                fc: 100,
                is_unicode: false,
                property_modifier: Vec::new(),
            },
            TextPiece {
                cp_start: 3,
                cp_end: 5,
                fc: 200,
                is_unicode: true,
                property_modifier: Vec::new(),
            },
        ]);

        assert_eq!(table.fc_range_to_cp_ranges(101, 204), vec![(1, 3), (3, 5)]);
        assert!(table.fc_range_to_cp_ranges(150, 180).is_empty());
    }

    #[test]
    fn indexed_ranges_match_scalar_for_overlapping_and_numeric_edge_intervals() {
        let table = table_from_pieces(vec![
            TextPiece {
                cp_start: 0,
                cp_end: 10,
                fc: 100,
                is_unicode: false,
                property_modifier: Vec::new(),
            },
            TextPiece {
                cp_start: 10,
                cp_end: 20,
                fc: 105,
                is_unicode: false,
                property_modifier: Vec::new(),
            },
            TextPiece {
                cp_start: 20,
                cp_end: 25,
                fc: 102,
                is_unicode: true,
                property_modifier: Vec::new(),
            },
            TextPiece {
                cp_start: 25,
                cp_end: 27,
                fc: u32::MAX - 2,
                is_unicode: true,
                property_modifier: Vec::new(),
            },
            TextPiece {
                cp_start: 27,
                cp_end: 27,
                fc: 0,
                is_unicode: false,
                property_modifier: Vec::new(),
            },
        ]);

        for (start_fc, end_fc) in [
            (0, 0),
            (0, 1),
            (101, 112),
            (103, 104),
            (104, 113),
            (115, 1_000),
            (u32::MAX - 3, u32::MAX),
        ] {
            assert_eq!(
                table.fc_range_to_cp_ranges(start_fc, end_fc),
                scalar_fc_range_to_cp_ranges(&table, start_fc, end_fc),
                "query [{start_fc}, {end_fc})",
            );
        }
    }

    #[test]
    fn indexed_ranges_match_scalar_for_adversarial_piece_orders() {
        let mut state = 0x91e1_0da5_u32;
        let mut next = || {
            state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
            state
        };
        let mut cp_start = 0_u32;
        let mut pieces = Vec::new();
        for index in 0..256_u32 {
            let length = next() % 32 + 1;
            let fc = if index.is_multiple_of(17) {
                128
            } else {
                next() % 2_048
            };
            pieces.push(TextPiece {
                cp_start,
                cp_end: cp_start + length,
                fc,
                is_unicode: next() & 1 == 0,
                property_modifier: Vec::new(),
            });
            cp_start += length;
        }
        let table = table_from_pieces(pieces);

        for _ in 0..1_024 {
            let left = next() % 2_500;
            let right = next() % 2_500;
            let start_fc = left.min(right);
            let end_fc = left.max(right);
            assert_eq!(
                table.fc_range_to_cp_ranges(start_fc, end_fc),
                scalar_fc_range_to_cp_ranges(&table, start_fc, end_fc),
                "query [{start_fc}, {end_fc})",
            );
        }
    }

    #[test]
    fn resolves_simple_and_complex_piece_modifiers() {
        let simple_bold = (1u16 << 8) | (0x55u16 << 1);
        assert_eq!(
            PieceTable::resolve_property_modifier(simple_bold, &[]),
            Some(vec![0x35, 0x08, 0x01])
        );

        let groups = vec![vec![0x03, 0x24, 0x02]];
        assert_eq!(
            PieceTable::resolve_property_modifier(1, &groups),
            Some(groups[0].clone())
        );
        assert!(PieceTable::resolve_property_modifier(3, &groups).is_none());
    }

    /// A piece table's character positions ascend, but a corrupt or hostile
    /// table can invert them. `length` must report nothing rather than wrap,
    /// which in a debug build panicked and in a release build produced a
    /// nonsense length near `u32::MAX`.
    #[test]
    fn inverted_character_positions_report_no_length() {
        let piece = TextPiece {
            cp_start: 200,
            cp_end: 100,
            fc: 0,
            is_unicode: true,
            property_modifier: Vec::new(),
        };
        assert_eq!(piece.length(), 0);
    }

    /// Converting a position must not overflow the file offset either.
    #[test]
    fn conversions_near_the_numeric_limit_report_none_rather_than_wrap() {
        let piece = TextPiece {
            cp_start: 0,
            cp_end: u32::MAX,
            fc: u32::MAX - 1,
            is_unicode: true,
            property_modifier: Vec::new(),
        };
        // `fc + offset * 2` would wrap for any non-trivial offset.
        assert_eq!(piece.cp_to_fc(u32::MAX / 2), None);
        // The start still resolves, so the guard is not over-broad.
        assert_eq!(piece.cp_to_fc(0), Some(u32::MAX - 1));
    }

    /// `fc_to_cp` adds the character offset to `cp_start`; that must not wrap.
    #[test]
    fn fc_to_cp_does_not_wrap_past_the_numeric_limit() {
        let piece = TextPiece {
            cp_start: u32::MAX - 1,
            cp_end: u32::MAX,
            fc: 0,
            is_unicode: false,
            property_modifier: Vec::new(),
        };
        assert_eq!(piece.fc_to_cp(10), None);
        assert_eq!(piece.fc_to_cp(0), Some(u32::MAX - 1));
    }
}
