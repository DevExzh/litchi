/// Shared SPRM (Single Property Modifier) parsing.
///
/// SPRMs are variable-length records used in both DOC and PPT formats
/// to modify properties. This module provides common SPRM parsing logic
/// based on Apache POI's SPRM handling.
/// SPRM operation types based on size code (from POI's SprmOperation).
use crate::common::binary::{read_i16_le, read_u16_le, read_u32_le};
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
#[derive(Debug, Clone)]
pub struct Sprm {
    /// SPRM opcode
    pub opcode: u16,
    /// SPRM operation type
    pub operation: SprmOperation,
    /// SPRM operand data
    pub operand: Vec<u8>,
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
/// A vector of parsed SPRMs
pub fn parse_sprms(grpprl: &[u8]) -> Vec<Sprm> {
    parse_sprms_two_byte(grpprl)
}

/// Parse SPRMs using 2-byte opcodes (Word 97+).
fn parse_sprms_two_byte(grpprl: &[u8]) -> Vec<Sprm> {
    let mut sprms = Vec::new();
    let mut offset = 0;

    while offset + 2 <= grpprl.len() {
        // Read SPRM opcode (2 bytes in Word 97+)
        let opcode = read_u16_le(grpprl, offset).unwrap_or(0);

        offset += 2;

        // Extract size code from opcode (bits 13-15, POI's BITFIELD_SIZECODE = 0xe000)
        let size_code = ((opcode & 0xe000) >> 13) as u8;
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
                // Variable length - read size from first byte (or 2 bytes for long SPRMs)
                if offset + 1 < grpprl.len() {
                    // Check if this is a long SPRM (SPRM_LONG_PARAGRAPH or SPRM_LONG_TABLE)
                    if opcode == 0xc615 || opcode == 0xd608 {
                        // Long SPRM - operand size in next 2 bytes
                        if offset + 3 <= grpprl.len() {
                            read_u16_le(grpprl, offset).unwrap_or(0) as usize
                        } else {
                            break;
                        }
                    } else {
                        // Regular variable SPRM - size in first byte
                        grpprl[offset] as usize
                    }
                } else {
                    break;
                }
            },
            7 => 3, // 3 byte operand
            _ => unreachable!(),
        };

        // Read operand data
        if offset + operand_size > grpprl.len() {
            break;
        }

        let operand = grpprl[offset..offset + operand_size].to_vec();
        offset += operand_size;

        sprms.push(Sprm {
            opcode,
            operation,
            operand,
        });
    }

    sprms
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

        let sprms = parse_sprms(&grpprl);
        assert_eq!(sprms.len(), 2);

        // Verify the opcodes were correctly parsed (little-endian)
        assert_eq!(sprms[0].opcode, 0x0835); // Bold
        assert_eq!(sprms[1].opcode, 0x4A43); // Font size (0x43, 0x4A bytes → 0x4A43 LE)
    }

    #[test]
    fn test_find_sprm() {
        let sprms = vec![
            Sprm {
                opcode: 0x0835,
                operation: SprmOperation::Byte,
                operand: vec![1],
            },
            Sprm {
                opcode: 0x4A43,
                operation: SprmOperation::Word,
                operand: vec![24, 0],
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
            operand: vec![1],
        };
        assert!(get_bool_from_sprm(&sprm));

        let sprm_false = Sprm {
            opcode: 0x0835,
            operation: SprmOperation::Byte,
            operand: vec![0],
        };
        assert!(!get_bool_from_sprm(&sprm_false));
    }
}
