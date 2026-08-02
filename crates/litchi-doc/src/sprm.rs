/// Shared SPRM (Single Property Modifier) parsing.
///
/// SPRMs are variable-length records used in both DOC and PPT formats
/// to modify properties. This module provides common SPRM parsing logic
/// based on Apache POI's SPRM handling.
/// SPRM operation types based on size code (from POI's SprmOperation).
use litchi_core::binary::{read_i16_le, read_u16_le, read_u32_le};
use smallvec::SmallVec;

/// Error returned when a `grpprl` is not an exact sequence of SPRMs.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The input ended partway through a two-byte opcode.
    #[error("truncated SPRM opcode at byte {at}: {remaining} byte(s) remain")]
    Opcode {
        /// Byte offset of the incomplete opcode.
        at: usize,
        /// Bytes available from `at` to the end of the input.
        remaining: usize,
    },
    /// A variable-length SPRM has an incomplete length field.
    #[error("truncated length for SPRM 0x{opcode:04X} at byte {at}")]
    Length {
        /// Byte offset of the incomplete length field.
        at: usize,
        /// Opcode whose length could not be read.
        opcode: u16,
    },
    /// A variable-length SPRM used a forbidden zero length.
    #[error("zero length for SPRM 0x{opcode:04X} at byte {at}")]
    ZeroLength {
        /// Byte offset of the invalid length field.
        at: usize,
        /// Opcode whose length was invalid.
        opcode: u16,
    },
    /// Calculating a variable record extent overflowed `usize`.
    #[error("SPRM 0x{opcode:04X} extent overflows at byte {at}")]
    Overflow {
        /// Byte offset where the extent calculation began.
        at: usize,
        /// Opcode whose extent overflowed.
        opcode: u16,
    },
    /// The input ended before the declared operand was complete.
    #[error(
        "truncated operand for SPRM 0x{opcode:04X} at byte {at}: expected {expected} byte(s), found {remaining}"
    )]
    Operand {
        /// Byte offset of the operand.
        at: usize,
        /// Opcode whose operand was incomplete.
        opcode: u16,
        /// Declared operand size.
        expected: usize,
        /// Operand bytes still available.
        remaining: usize,
    },
}

/// Result of strict SPRM parsing.
pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SprmOperation {
    /// Size code 0 - toggle (no operand)
    Toggle,
    /// Size code 1 - 1 byte operand
    Byte,
    /// Size code 2 - 2 byte operand
    Word,
    /// Size code 3 - 4 byte operand
    DWord,
    /// Size code 4 - 2 byte operand
    Word2,
    /// Size code 5 - 2 byte operand
    Word3,
    /// Size code 6 - variable length operand
    Variable,
    /// Size code 7 - 3 byte operand
    ThreeByte,
}

impl From<u8> for SprmOperation {
    fn from(size_code: u8) -> Self {
        match size_code {
            0 => SprmOperation::Toggle,
            1 => SprmOperation::Byte,
            2 => SprmOperation::Word,
            3 => SprmOperation::DWord,
            4 => SprmOperation::Word2,
            5 => SprmOperation::Word3,
            6 => SprmOperation::Variable,
            7 => SprmOperation::ThreeByte,
            _ => unreachable!(),
        }
    }
}

/// An SPRM (Single Property Modifier).
///
/// Based on Apache POI's SprmBuffer and related classes.
///
/// **Performance:** Uses `SmallVec` with 8-byte inline storage for operands
/// to eliminate heap allocations for common cases (most operands are 1-4 bytes).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sprm {
    /// SPRM opcode
    pub opcode: u16,
    /// SPRM operation type
    pub operation: SprmOperation,
    /// SPRM operand data (inline for operands ≤8 bytes to avoid heap allocation)
    pub operand: SmallVec<[u8; 8]>,
    /// Offset in original grpprl array (for accessing raw data)
    pub offset: usize,
    /// Total size of this SPRM in bytes (opcode + size byte if any + operand)
    pub size: usize,
}

impl Sprm {
    /// Get the operand as a byte.
    #[inline]
    pub fn operand_byte(&self) -> Option<u8> {
        self.operand.first().copied()
    }

    /// Get the operand as a word (u16).
    #[inline]
    pub fn operand_word(&self) -> Option<u16> {
        read_u16_le(&self.operand, 0).ok()
    }

    /// Get the operand as a signed word (i16).
    #[inline]
    pub fn operand_i16(&self) -> Option<i16> {
        read_i16_le(&self.operand, 0).ok()
    }

    /// Get the operand as a dword (u32).
    #[inline]
    pub fn operand_dword(&self) -> Option<u32> {
        read_u32_le(&self.operand, 0).ok()
    }

    /// Get the operand as raw bytes.
    #[inline]
    pub fn operand_bytes(&self) -> &[u8] {
        &self.operand
    }
}

/// Parse SPRMs from a byte array (grpprl - group of SPRMs).
///
/// Based on Apache POI's SprmBuffer.findSprms() and SprmIterator.
///
/// **Important:** Apache POI always uses 2-byte SPRM opcodes for all Word versions,
/// including Word 6/7. This is the standard format used by Microsoft Word.
///
/// # Arguments
///
/// * `grpprl` - The byte array containing SPRMs
///
/// # Returns
///
/// A vector containing the exact parsed sequence, or a typed error at the
/// first malformed record. Valid prefixes are never returned as success.
pub fn parse_sprms(grpprl: &[u8]) -> Result<Vec<Sprm>> {
    parse_sprms_two_byte(grpprl)
}

/// Parse SPRMs using 2-byte opcodes (Word 97+).
///
/// **Performance optimizations:**
/// - Pre-allocates vector capacity based on input size estimate
/// - Uses SmallVec for operands to avoid heap allocations (most are ≤8 bytes)
/// - Extracts size code masking constant for better code generation
fn parse_sprms_two_byte(grpprl: &[u8]) -> Result<Vec<Sprm>> {
    // Pre-allocate: estimate ~4 bytes per SPRM on average (2 opcode + 1-2 operand)
    // This significantly reduces reallocation overhead for large grpprl arrays
    let estimated_capacity = (grpprl.len() / 4).max(8);
    let mut sprms = Vec::with_capacity(estimated_capacity);

    // Extract constant for better optimization
    const SIZECODE_MASK: u16 = 0xe000;
    const SIZECODE_SHIFT: u16 = 13;

    let mut offset = 0;

    while offset < grpprl.len() {
        let sprm_start = offset; // Track start of SPRM for offset field

        // Read SPRM opcode (2 bytes in Word 97+)
        let remaining = grpprl.len() - offset;
        if remaining < 2 {
            return Err(Error::Opcode {
                at: offset,
                remaining,
            });
        }
        let opcode_end = offset + 2;
        let opcode = u16::from_le_bytes([grpprl[offset], grpprl[offset + 1]]);
        offset = opcode_end;

        // Extract size code from opcode (bits 13-15, POI's BITFIELD_SIZECODE = 0xe000)
        let size_code = ((opcode & SIZECODE_MASK) >> SIZECODE_SHIFT) as u8;
        let operation = SprmOperation::from(size_code);

        // Determine operand size based on size code (matching POI's initSize method)
        // From POI SprmOperation.initSize():
        //   case 0: case 1: return 3;  // 2 byte opcode + 1 byte operand
        //   case 2: case 4: case 5: return 4;  // 2 byte opcode + 2 byte operand
        //   case 3: return 6;  // 2 byte opcode + 4 byte operand
        //   case 6: variable length
        //   case 7: return 5;  // 2 byte opcode + 3 byte operand
        let operand_size = match size_code {
            0 | 1 => 1,     // 1 byte operand
            2 | 4 | 5 => 2, // 2 byte operand
            3 => 4,         // 4 byte operand
            6 => {
                // Variable operands have a length prefix that is not part of
                // the operand exposed to property decoders.
                if opcode == 0xc615 {
                    let encoded_size = *grpprl
                        .get(offset)
                        .ok_or(Error::Length { at: offset, opcode })?
                        as usize;
                    if encoded_size == 0 {
                        return Err(Error::ZeroLength { at: offset, opcode });
                    }
                    let size = if encoded_size == 255 {
                        let delete_count_offset = offset
                            .checked_add(1)
                            .ok_or(Error::Overflow { at: offset, opcode })?;
                        let delete_count = *grpprl
                            .get(delete_count_offset)
                            .ok_or(Error::Length { at: offset, opcode })?
                            as usize;
                        let deleted_bytes = 4usize
                            .checked_mul(delete_count)
                            .ok_or(Error::Overflow { at: offset, opcode })?;
                        let add_count_offset = offset
                            .checked_add(2)
                            .and_then(|value| value.checked_add(deleted_bytes))
                            .ok_or(Error::Overflow { at: offset, opcode })?;
                        let add_count = *grpprl
                            .get(add_count_offset)
                            .ok_or(Error::Length { at: offset, opcode })?
                            as usize;
                        let added_bytes = 3usize
                            .checked_mul(add_count)
                            .ok_or(Error::Overflow { at: offset, opcode })?;
                        2usize
                            .checked_add(deleted_bytes)
                            .and_then(|value| value.checked_add(added_bytes))
                            .ok_or(Error::Overflow { at: offset, opcode })?
                    } else {
                        encoded_size
                    };
                    offset = offset
                        .checked_add(1)
                        .ok_or(Error::Overflow { at: offset, opcode })?;
                    size
                } else if opcode == 0xd608 {
                    let length_end = offset
                        .checked_add(2)
                        .ok_or(Error::Overflow { at: offset, opcode })?;
                    if length_end > grpprl.len() {
                        return Err(Error::Length { at: offset, opcode });
                    }
                    // For the two long encodings, cb includes one byte beyond
                    // the actual operand (matching MS-DOC and POI's size rule).
                    let encoded_size =
                        u16::from_le_bytes([grpprl[offset], grpprl[offset + 1]]) as usize;
                    if encoded_size == 0 {
                        return Err(Error::ZeroLength { at: offset, opcode });
                    }
                    offset = length_end;
                    encoded_size - 1
                } else {
                    let size = *grpprl
                        .get(offset)
                        .ok_or(Error::Length { at: offset, opcode })?
                        as usize;
                    offset = offset
                        .checked_add(1)
                        .ok_or(Error::Overflow { at: offset, opcode })?;
                    size
                }
            },
            7 => 3, // 3 byte operand
            _ => unreachable!(),
        };

        // Read operand data
        let operand_end = offset
            .checked_add(operand_size)
            .ok_or(Error::Overflow { at: offset, opcode })?;
        if operand_end > grpprl.len() {
            return Err(Error::Operand {
                at: offset,
                opcode,
                expected: operand_size,
                remaining: grpprl.len() - offset,
            });
        }

        // Use SmallVec::from_slice for efficient inline storage (no heap allocation for ≤8 bytes)
        let operand = SmallVec::from_slice(&grpprl[offset..operand_end]);
        offset = operand_end;

        let total_size = offset - sprm_start; // Total size including opcode

        sprms.push(Sprm {
            opcode,
            operation,
            operand,
            offset: sprm_start,
            size: total_size,
        });
    }

    Ok(sprms)
}

/// Find a specific SPRM by opcode in a list of SPRMs.
///
/// # Arguments
///
/// * `sprms` - The list of SPRMs to search
/// * `opcode` - The SPRM opcode to find
///
/// # Returns
///
/// Reference to the first matching SPRM, or None if not found
#[inline]
pub fn find_sprm(sprms: &[Sprm], opcode: u16) -> Option<&Sprm> {
    sprms.iter().find(|sprm| sprm.opcode == opcode)
}

/// Get a boolean value from an SPRM operand.
///
/// Based on Apache POI's SPRM boolean handling.
#[inline]
pub fn get_bool_from_sprm(sprm: &Sprm) -> bool {
    sprm.operand_byte().unwrap_or(0) != 0
}

/// Get an integer value from an SPRM operand.
#[inline]
pub fn get_int_from_sprm(sprm: &Sprm) -> Option<i32> {
    match sprm.operation {
        SprmOperation::Byte | SprmOperation::Toggle => sprm.operand_byte().map(|b| b as i32),
        SprmOperation::Word | SprmOperation::Word2 | SprmOperation::Word3 => {
            sprm.operand_i16().map(|w| w as i32)
        },
        SprmOperation::DWord => sprm.operand_dword().map(|d| d as i32),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sprm_operation_from() {
        assert_eq!(SprmOperation::from(0), SprmOperation::Toggle);
        assert_eq!(SprmOperation::from(1), SprmOperation::Byte);
        assert_eq!(SprmOperation::from(2), SprmOperation::Word);
        assert_eq!(SprmOperation::from(4), SprmOperation::Word2);
        assert_eq!(SprmOperation::from(5), SprmOperation::Word3);
    }

    #[test]
    fn test_parse_sprms() {
        // Create a simple SPRM buffer
        // SPRM 1: opcode 0x0835 (bold, byte operand), operand = 0x01
        // SPRM 2: opcode 0x4A43 (font size, word operand), operand = 0x0018 (24 = 12pt)
        let grpprl = vec![
            0x35, 0x08, // opcode 0x0835 (operation type = 1, byte)
            0x01, // operand = 1 (true)
            0x43, 0x4A, // opcode 0x4A43 (operation type = 2, word)
            0x18, 0x00, // operand = 24
        ];

        let sprms = parse_sprms(&grpprl).unwrap();
        assert_eq!(sprms.len(), 2);

        // Verify the opcodes were correctly parsed (little-endian)
        assert_eq!(sprms[0].opcode, 0x0835); // Bold
        assert_eq!(sprms[1].opcode, 0x4A43); // Font size (0x43, 0x4A bytes → 0x4A43 LE)
    }

    #[test]
    fn variable_sprm_excludes_length_and_preserves_following_sprm() {
        let grpprl = [
            0x71, 0xCA, // sprmCShd
            0x03, // three operand bytes
            0xAA, 0xBB, 0xCC, // operand
            0x35, 0x08, 0x01, // following sprmCFBold
        ];
        let sprms = parse_sprms(&grpprl).unwrap();
        assert_eq!(sprms.len(), 2);
        assert_eq!(sprms[0].operand.as_slice(), &[0xAA, 0xBB, 0xCC]);
        assert_eq!(sprms[0].size, 6);
        assert_eq!(sprms[1].opcode, 0x0835);
        assert_eq!(sprms[1].offset, 6);
    }

    #[test]
    fn long_variable_sprm_decodes_adjusted_word_length() {
        let grpprl = [
            0x08, 0xD6, // sprmTDefTable
            0x04, 0x00, // cb includes one extra byte
            0x02, 0x10, 0x20, // three operand bytes
            0x35, 0x08, 0x01, // following sprmCFBold
        ];
        let sprms = parse_sprms(&grpprl).unwrap();
        assert_eq!(sprms.len(), 2);
        assert_eq!(sprms[0].operand.as_slice(), &[0x02, 0x10, 0x20]);
        assert_eq!(sprms[0].size, 7);
        assert_eq!(sprms[1].offset, 7);
    }

    #[test]
    fn tab_change_sprm_handles_byte_and_extended_lengths() {
        let normal = [0x15, 0xC6, 2, 0, 0, 0x35, 0x08, 1];
        let sprms = parse_sprms(&normal).unwrap();
        assert_eq!(sprms.len(), 2);
        assert_eq!(sprms[0].operand.as_slice(), &[0, 0]);
        assert_eq!(sprms[1].opcode, 0x0835);

        let extended = [
            0x15, 0xC6, 255, // special computed length
            1, 100, 0, 25, 0, // one delete center and close distance
            1, 200, 0, 0, // one added left tab
            0x35, 0x08, 1,
        ];
        let sprms = parse_sprms(&extended).unwrap();
        assert_eq!(sprms.len(), 2);
        assert_eq!(sprms[0].operand.as_slice(), &extended[3..12]);
        assert_eq!(sprms[1].opcode, 0x0835);
        assert_eq!(sprms[1].offset, 12);
    }

    #[test]
    fn rejects_truncated_opcode_without_returning_a_prefix() {
        let grpprl = [
            0x35, 0x08, 0x01, // complete sprmCFBold
            0x43, // first byte of the next opcode
        ];

        assert_eq!(
            parse_sprms(&grpprl),
            Err(Error::Opcode {
                at: 3,
                remaining: 1,
            })
        );
    }

    #[test]
    fn rejects_truncated_fixed_and_variable_operands() {
        assert_eq!(
            parse_sprms(&[0x43, 0x4A, 0x18]),
            Err(Error::Operand {
                at: 2,
                opcode: 0x4A43,
                expected: 2,
                remaining: 1,
            })
        );
        assert_eq!(
            parse_sprms(&[0x08, 0xD6, 0x02, 0x00]),
            Err(Error::Operand {
                at: 4,
                opcode: 0xD608,
                expected: 1,
                remaining: 0,
            })
        );
    }

    #[test]
    fn rejects_invalid_variable_lengths() {
        assert_eq!(
            parse_sprms(&[0x08, 0xD6, 0x00, 0x00]),
            Err(Error::ZeroLength {
                at: 2,
                opcode: 0xD608,
            })
        );
        assert_eq!(
            parse_sprms(&[0x15, 0xC6, 0xFF, 0x01]),
            Err(Error::Length {
                at: 2,
                opcode: 0xC615,
            })
        );
    }

    #[test]
    fn test_find_sprm() {
        let sprms = vec![
            Sprm {
                opcode: 0x0835,
                operation: SprmOperation::Byte,
                operand: SmallVec::from_slice(&[1]),
                offset: 0,
                size: 3,
            },
            Sprm {
                opcode: 0x4A43,
                operation: SprmOperation::Word,
                operand: SmallVec::from_slice(&[24, 0]),
                offset: 3,
                size: 4,
            },
        ];

        assert!(find_sprm(&sprms, 0x0835).is_some());
        assert!(find_sprm(&sprms, 0x4A43).is_some());
        assert!(find_sprm(&sprms, 0xFFFF).is_none());
    }

    #[test]
    fn test_get_bool_from_sprm() {
        let sprm = Sprm {
            opcode: 0x0835,
            operation: SprmOperation::Byte,
            operand: SmallVec::from_slice(&[1]),
            offset: 0,
            size: 3,
        };
        assert!(get_bool_from_sprm(&sprm));

        let sprm_false = Sprm {
            opcode: 0x0835,
            operation: SprmOperation::Byte,
            operand: SmallVec::from_slice(&[0]),
            offset: 0,
            size: 3,
        };
        assert!(!get_bool_from_sprm(&sprm_false));
    }
}
