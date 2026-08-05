//! Bitfield decoding helpers for SPRM opcodes.

/// Extract SPRM type from opcode (bits 10-12).
///
/// Returns:
/// - 1: PAP (Paragraph Properties)
/// - 2: CHP (Character Properties)
/// - 3: PIC (Picture Properties)
/// - 4: SEP (Section Properties)
/// - 5: TAP (Table Properties)
#[inline]
pub fn get_sprm_type(opcode: u16) -> u8 {
    ((opcode >> 10) & 0x07) as u8
}

/// Extract SPRM operation code from opcode (bits 0-8).
#[inline]
pub fn get_sprm_operation(opcode: u16) -> u16 {
    opcode & 0x01FF
}

/// Extract SPRM size code from opcode (bits 13-15).
///
/// Returns:
/// - 0, 1: 1-byte operand
/// - 2, 4, 5: 2-byte operand
/// - 3: 4-byte operand
/// - 6: Variable length
/// - 7: 3-byte operand
#[inline]
pub fn get_sprm_size_code(opcode: u16) -> u8 {
    ((opcode >> 13) & 0x07) as u8
}

/// Check if SPRM is a "special" operation (bit 9).
#[inline]
pub fn is_sprm_special(opcode: u16) -> bool {
    (opcode & 0x0200) != 0
}
