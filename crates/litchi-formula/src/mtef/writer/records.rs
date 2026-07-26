//! Low-level MTEF 5 record encoders
//!
//! Every function here writes one record (or one primitive field) into the
//! caller's buffer. MTEF 5 records share a common shape:
//!
//! ```text
//! tag:u8  options:u8  [nudge]  <record specific fields>  [object list]  END
//! ```
//!
//! Up to MTEF 4 the options travelled in the high nibble of the tag byte; MTEF 5
//! gives them a byte of their own, which is the form this module emits.
//!
//! All multi-byte fields are little endian.

use super::error::MtefWriteError;
use crate::mtef::constants::*;

/// Options byte written for records that need no attributes
pub(super) const NO_OPTIONS: u8 = 0;

// ---------------------------------------------------------------------------
// Primitives
// ---------------------------------------------------------------------------

/// Append a single byte
#[inline]
pub(super) fn write_u8(out: &mut Vec<u8>, value: u8) {
    out.push(value);
}

/// Append a little-endian 16-bit value
#[inline]
pub(super) fn write_u16(out: &mut Vec<u8>, value: u16) {
    out.extend_from_slice(&value.to_le_bytes());
}

/// Append a little-endian 32-bit value
#[inline]
pub(super) fn write_u32(out: &mut Vec<u8>, value: u32) {
    out.extend_from_slice(&value.to_le_bytes());
}

/// Append a NUL-terminated string
pub(super) fn write_cstr(out: &mut Vec<u8>, value: &str) -> Result<(), MtefWriteError> {
    if value.as_bytes().contains(&0) {
        return Err(MtefWriteError::InvalidFontName);
    }
    out.extend_from_slice(value.as_bytes());
    out.push(0);
    Ok(())
}

// ---------------------------------------------------------------------------
// Shared record parts
// ---------------------------------------------------------------------------

/// Append an END record, which terminates an object list
#[inline]
pub(super) fn write_end(out: &mut Vec<u8>) {
    write_u8(out, END);
}

/// Append a nudge (fine positioning offset) as it follows an options byte
///
/// Offsets that fit in a byte are stored directly; anything else uses the
/// escaped 16-bit form introduced by the `128, 128` sentinel pair.
pub(super) fn write_nudge(out: &mut Vec<u8>, x: i16, y: i16) {
    let fits = |value: i16| (0..i16::from(NUDGE_ESCAPE)).contains(&value);
    if fits(x) && fits(y) {
        write_u8(out, x as u8);
        write_u8(out, y as u8);
    } else {
        write_u8(out, NUDGE_ESCAPE);
        write_u8(out, NUDGE_ESCAPE);
        write_u16(out, x as u16);
        write_u16(out, y as u16);
    }
}

// ---------------------------------------------------------------------------
// Size records
// ---------------------------------------------------------------------------

/// Append a FULL size record, which selects the equation's base size
///
/// FULL through SUBSYM are single-byte records: the tag alone carries the size
/// level, so no options byte follows.
#[inline]
pub(super) fn write_size_full(out: &mut Vec<u8>) {
    write_u8(out, FULL);
}

// ---------------------------------------------------------------------------
// Structural records
// ---------------------------------------------------------------------------

/// Start a LINE record; the caller writes the contents and then [`write_end`]
#[inline]
pub(super) fn write_line_start(out: &mut Vec<u8>) {
    write_u8(out, LINE);
    write_u8(out, NO_OPTIONS);
}

/// Start a PILE record (a vertical stack of LINE records)
#[inline]
pub(super) fn write_pile_start(out: &mut Vec<u8>, halign: u8, valign: u8) {
    write_u8(out, PILE);
    write_u8(out, NO_OPTIONS);
    write_u8(out, halign);
    write_u8(out, valign);
}

/// Start a MATRIX record
///
/// The row and column partitions are two-bit fields, one per grid line, so a
/// matrix with `rows` rows needs `ceil(2 * (rows + 1) / 8)` partition bytes.
/// Zero selects the default (solid) partition everywhere.
pub(super) fn write_matrix_start(
    out: &mut Vec<u8>,
    rows: usize,
    cols: usize,
) -> Result<(), MtefWriteError> {
    let limit = max_matrix_dimension();
    if rows > limit || cols > limit {
        return Err(MtefWriteError::MatrixTooLarge { rows, cols, limit });
    }

    write_u8(out, MATRIX);
    write_u8(out, NO_OPTIONS);
    write_u8(out, VALIGN_BASELINE);
    write_u8(out, JUSTIFY_CENTER);
    write_u8(out, JUSTIFY_CENTER);
    write_u8(out, rows as u8);
    write_u8(out, cols as u8);

    out.resize(out.len() + partition_bytes(rows) + partition_bytes(cols), 0);
    Ok(())
}

/// Number of partition bytes needed to describe `count` rows or columns
fn partition_bytes(count: usize) -> usize {
    (MATRIX_PARTITION_BITS * (count + 1)).div_ceil(8)
}

/// Largest row or column count whose partition fits the record's fixed array
fn max_matrix_dimension() -> usize {
    (MATRIX_PARTITION_BYTES * 8) / MATRIX_PARTITION_BITS - 1
}

/// Start a TMPL record; the caller writes one LINE per slot and then [`write_end`]
///
/// Variations below `0x80` occupy a single byte. Larger values are split into a
/// low 7-bit half (flagged with the high bit) followed by the remaining bits.
pub(super) fn write_template_start(
    out: &mut Vec<u8>,
    selector: u8,
    variation: u16,
) -> Result<(), MtefWriteError> {
    if variation > VARIATION_MAX {
        return Err(MtefWriteError::VariationTooLarge(u32::from(variation)));
    }

    write_u8(out, TMPL);
    write_u8(out, NO_OPTIONS);
    write_u8(out, selector);
    if variation <= VARIATION_LOW_MASK {
        write_u8(out, variation as u8);
    } else {
        write_u8(
            out,
            (variation & VARIATION_LOW_MASK) as u8 | VARIATION_EXTENDED_FLAG,
        );
        write_u8(out, (variation >> VARIATION_LOW_BITS) as u8);
    }
    write_u8(out, NO_OPTIONS); // template-specific options
    Ok(())
}

// ---------------------------------------------------------------------------
// Leaf records
// ---------------------------------------------------------------------------

/// Append a CHAR record holding a 16-bit MathType character code
///
/// Returns the offset of the record's options byte so that a following
/// embellishment list can set [`CHAR_EMBELL`] on it.
pub(super) fn write_char(out: &mut Vec<u8>, typeface: u8, mtcode: u16) -> usize {
    write_u8(out, CHAR);
    let options_offset = out.len();
    write_u8(out, NO_OPTIONS);
    write_u8(out, typeface);
    write_u16(out, mtcode);
    options_offset
}

/// Append a CHAR record carrying an explicit font position instead of an MTCode
///
/// MathType uses this form for glyphs that only exist in a particular font: the
/// MTCode is suppressed and the 8- or 16-bit index into the font takes over.
pub(super) fn write_char_encoded(out: &mut Vec<u8>, typeface: u8, position: FontPosition) -> usize {
    let flag = match position {
        FontPosition::Bits8(_) => CHAR_ENC_CHAR_8,
        FontPosition::Bits16(_) => CHAR_ENC_CHAR_16,
    };

    write_u8(out, CHAR);
    let options_offset = out.len();
    write_u8(out, CHAR_ENC_NO_MTCODE | flag);
    write_u8(out, typeface);
    match position {
        FontPosition::Bits8(value) => write_u8(out, value),
        FontPosition::Bits16(value) => write_u16(out, value),
    }
    options_offset
}

/// Font-relative character position carried by a CHAR record
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FontPosition {
    /// Eight-bit index into the character's font
    Bits8(u8),
    /// Sixteen-bit index into the character's font
    Bits16(u16),
}

/// Mark a previously written CHAR record as carrying an embellishment list
#[inline]
pub(super) fn set_char_embellished(out: &mut [u8], options_offset: usize) {
    if let Some(options) = out.get_mut(options_offset) {
        *options |= CHAR_EMBELL;
    }
}

/// Append an EMBELL record
///
/// Embellishment lists follow the CHAR record they decorate and are terminated
/// by an END record.
#[inline]
pub(super) fn write_embell(out: &mut Vec<u8>, embellishment: u8) {
    write_u8(out, EMBELL);
    write_u8(out, NO_OPTIONS);
    write_u8(out, embellishment);
}

/// Append a FONT record, which selects a named font for following characters
pub(super) fn write_font(
    out: &mut Vec<u8>,
    typeface: u8,
    style: u8,
    name: &str,
) -> Result<(), MtefWriteError> {
    write_u8(out, FONT);
    write_u8(out, typeface);
    write_u8(out, style);
    write_cstr(out, name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn partition_sizes_match_the_reader() {
        // The reader derives the same counts from rows/cols when parsing.
        assert_eq!(partition_bytes(1), 1);
        assert_eq!(partition_bytes(3), 1);
        assert_eq!(partition_bytes(4), 2);
        assert_eq!(max_matrix_dimension(), 63);
    }

    #[test]
    fn short_variations_use_one_byte() {
        let mut out = Vec::new();
        write_template_start(&mut out, TMPL_FRACT, TV_FRACT_BAR).expect("encodable");
        assert_eq!(out, vec![TMPL, 0, TMPL_FRACT, 0, 0]);
    }

    #[test]
    fn long_variations_use_the_escaped_form() {
        let mut out = Vec::new();
        write_template_start(&mut out, TMPL_INTERVAL, 0x0123).expect("encodable");
        assert_eq!(out, vec![TMPL, 0, TMPL_INTERVAL, 0x23 | 0x80, 0x02, 0]);
    }

    #[test]
    fn oversized_variations_are_rejected() {
        let mut out = Vec::new();
        let err = write_template_start(&mut out, TMPL_INTERVAL, 0xFFFF).expect_err("too large");
        assert_eq!(err, MtefWriteError::VariationTooLarge(0xFFFF));
    }

    #[test]
    fn oversized_matrices_are_rejected() {
        let mut out = Vec::new();
        let err = write_matrix_start(&mut out, 64, 1).expect_err("too large");
        assert!(matches!(
            err,
            MtefWriteError::MatrixTooLarge { limit: 63, .. }
        ));
    }

    #[test]
    fn characters_can_carry_a_font_position_instead_of_an_mtcode() {
        let mut out = Vec::new();
        write_char_encoded(&mut out, TYPEFACE_SYMBOL, FontPosition::Bits8(0xA9));
        assert_eq!(
            out,
            vec![
                CHAR,
                CHAR_ENC_NO_MTCODE | CHAR_ENC_CHAR_8,
                TYPEFACE_SYMBOL,
                0xA9
            ]
        );

        let mut out = Vec::new();
        write_char_encoded(&mut out, TYPEFACE_MTEXTRA, FontPosition::Bits16(0x1234));
        assert_eq!(
            out,
            vec![
                CHAR,
                CHAR_ENC_NO_MTCODE | CHAR_ENC_CHAR_16,
                TYPEFACE_MTEXTRA,
                0x34,
                0x12
            ]
        );
    }

    #[test]
    fn embellished_characters_flag_their_options_byte() {
        let mut out = Vec::new();
        let options = write_char(&mut out, TYPEFACE_VARIABLE, u16::from(b'x'));
        set_char_embellished(&mut out, options);
        write_embell(&mut out, EMB_HAT);
        write_end(&mut out);

        assert_eq!(
            out,
            vec![
                CHAR,
                CHAR_EMBELL,
                TYPEFACE_VARIABLE,
                b'x',
                0,
                EMBELL,
                0,
                EMB_HAT,
                END
            ]
        );
    }

    #[test]
    fn small_nudges_use_two_bytes() {
        let mut out = Vec::new();
        write_nudge(&mut out, 3, 4);
        assert_eq!(out, vec![3, 4]);
    }

    #[test]
    fn large_nudges_use_the_escaped_form() {
        let mut out = Vec::new();
        write_nudge(&mut out, -1, 300);
        assert_eq!(out, vec![128, 128, 0xFF, 0xFF, 0x2C, 0x01]);
    }

    #[test]
    fn font_names_reject_embedded_nul() {
        let mut out = Vec::new();
        assert_eq!(
            write_font(&mut out, TYPEFACE_TEXT, 0, "bad\0name"),
            Err(MtefWriteError::InvalidFontName)
        );
    }
}
