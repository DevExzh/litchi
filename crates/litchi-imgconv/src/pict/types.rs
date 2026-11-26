//! Common data types used by the PICT reader and renderer.

/// QuickDraw rectangle: `top`, `left`, `bottom`, `right` in that order on
/// disk, big-endian.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct PictRect {
    pub(super) top: i16,
    pub(super) left: i16,
    pub(super) bottom: i16,
    pub(super) right: i16,
}

impl PictRect {
    /// Reads a rectangle at `offset`.
    pub(super) fn read(data: &[u8], offset: usize) -> Option<Self> {
        Some(Self {
            top: read_i16(data, offset)?,
            left: read_i16(data, offset.checked_add(2)?)?,
            bottom: read_i16(data, offset.checked_add(4)?)?,
            right: read_i16(data, offset.checked_add(6)?)?,
        })
    }

    /// Width of the rectangle (negative when reversed).
    #[inline]
    pub(super) fn width(&self) -> i32 {
        i32::from(self.right) - i32::from(self.left)
    }

    /// Height of the rectangle (negative when reversed).
    #[inline]
    pub(super) fn height(&self) -> i32 {
        i32::from(self.bottom) - i32::from(self.top)
    }

    /// Returns whether the rectangle encloses at least one pixel.
    #[inline]
    pub(super) fn is_nonempty(&self) -> bool {
        self.right > self.left && self.bottom > self.top
    }
}

pub(super) fn read_u16(data: &[u8], offset: usize) -> Option<u16> {
    let bytes = data.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_be_bytes([bytes[0], bytes[1]]))
}

pub(super) fn read_i16(data: &[u8], offset: usize) -> Option<i16> {
    read_u16(data, offset).map(|value| value as i16)
}

pub(super) fn read_u32(data: &[u8], offset: usize) -> Option<u32> {
    let bytes = data.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_be_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

pub(super) fn read_i32(data: &[u8], offset: usize) -> Option<i32> {
    read_u32(data, offset).map(|value| value as i32)
}

/// Reads a QuickDraw `Fixed` (16.16) value.
pub(super) fn read_fixed(data: &[u8], offset: usize) -> Option<f64> {
    read_i32(data, offset).map(|value| f64::from(value) / 65536.0)
}

/// Decodes Mac OS Roman text, the encoding of QuickDraw text and font names.
pub(super) fn mac_roman_to_string(bytes: &[u8]) -> String {
    bytes.iter().map(|&byte| mac_roman_char(byte)).collect()
}

/// Maps one Mac OS Roman byte to Unicode, following Apple's published
/// `ROMAN.TXT` mapping (0xDB is the euro sign since Mac OS 8.5).
fn mac_roman_char(byte: u8) -> char {
    const HIGH: [char; 128] = [
        // 0x80
        'Ä', 'Å', 'Ç', 'É', 'Ñ', 'Ö', 'Ü', 'á', 'à', 'â', 'ä', 'ã', 'å', 'ç', 'é', 'è',
        // 0x90
        'ê', 'ë', 'í', 'ì', 'î', 'ï', 'ñ', 'ó', 'ò', 'ô', 'ö', 'õ', 'ú', 'ù', 'û', 'ü',
        // 0xA0
        '†', '°', '¢', '£', '§', '•', '¶', 'ß', '®', '©', '™', '´', '¨', '≠', 'Æ', 'Ø',
        // 0xB0
        '∞', '±', '≤', '≥', '¥', 'µ', '∂', '∑', '∏', 'π', '∫', 'ª', 'º', 'Ω', 'æ', 'ø',
        // 0xC0
        '¿', '¡', '¬', '√', 'ƒ', '≈', '∆', '«', '»', '…', '\u{a0}', 'À', 'Ã', 'Õ', 'Œ', 'œ',
        // 0xD0
        '–', '—', '“', '”', '‘', '’', '÷', '◊', 'ÿ', 'Ÿ', '⁄', '€', '‹', '›', 'ﬁ', 'ﬂ',
        // 0xE0
        '‡', '·', '‚', '„', '‰', 'Â', 'Ê', 'Á', 'Ë', 'È', 'Í', 'Î', 'Ï', 'Ì', 'Ó', 'Ô',
        // 0xF0
        '\u{f8ff}', 'Ò', 'Ú', 'Û', 'Ù', 'ı', 'ˆ', '˜', '¯', '˘', '˙', '˚', '¸', '˝', '˛', 'ˇ',
    ];
    if byte < 0x80 {
        char::from(byte)
    } else {
        HIGH[usize::from(byte - 0x80)]
    }
}

/// Decodes text in the Symbol font's encoding, used by pictures that draw
/// Greek letters and mathematical symbols with the Symbol font. Codes the
/// encoding leaves undefined fall back to Mac OS Roman.
pub(super) fn symbol_to_string(bytes: &[u8]) -> String {
    bytes.iter().map(|&byte| symbol_char(byte)).collect()
}

/// Maps one Symbol-encoded byte to Unicode, following Adobe's Symbol
/// encoding as used by the Macintosh Symbol font.
fn symbol_char(byte: u8) -> char {
    const ASCII: [char; 95] = [
        // 0x20
        ' ', '!', '∀', '#', '∃', '%', '&', '∋', '(', ')', '∗', '+', ',', '−', '.', '/',
        // 0x30
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', ':', ';', '<', '=', '>', '?',
        // 0x40
        '≅', 'Α', 'Β', 'Χ', 'Δ', 'Ε', 'Φ', 'Γ', 'Η', 'Ι', 'ϑ', 'Κ', 'Λ', 'Μ', 'Ν', 'Ο',
        // 0x50
        'Π', 'Θ', 'Ρ', 'Σ', 'Τ', 'Υ', 'ς', 'Ω', 'Ξ', 'Ψ', 'Ζ', '[', '∴', ']', '⊥', '_',
        // 0x60
        '‾', 'α', 'β', 'χ', 'δ', 'ε', 'φ', 'γ', 'η', 'ι', 'ϕ', 'κ', 'λ', 'μ', 'ν', 'ο',
        // 0x70
        'π', 'θ', 'ρ', 'σ', 'τ', 'υ', 'ϖ', 'ω', 'ξ', 'ψ', 'ζ', '{', '|', '}', '∼',
    ];
    const HIGH: [char; 96] = [
        // 0xA0
        '€', 'ϒ', '′', '≤', '⁄', '∞', 'ƒ', '♣', '♦', '♥', '♠', '↔', '←', '↑', '→', '↓',
        // 0xB0
        '°', '±', '″', '≥', '×', '∝', '∂', '•', '÷', '≠', '≡', '≈', '…', '⏐', '⎯', '↵',
        // 0xC0
        'ℵ', 'ℑ', 'ℜ', '℘', '⊗', '⊕', '∅', '∩', '∪', '⊃', '⊇', '⊄', '⊂', '⊆', '∈', '∉',
        // 0xD0
        '∠', '∇', '®', '©', '™', '∏', '√', '⋅', '¬', '∧', '∨', '⇔', '⇐', '⇑', '⇒', '⇓',
        // 0xE0
        '◊', '〈', '®', '©', '™', '∑', '⎛', '⎜', '⎝', '⎡', '⎢', '⎣', '⎧', '⎨', '⎩', '⎪',
        // 0xF0
        '\u{f8ff}', '〉', '∫', '⌠', '⎮', '⌡', '⎞', '⎟', '⎠', '⎤', '⎥', '⎦', '⎫', '⎬', '⎭', ' ',
    ];
    match byte {
        0x20..=0x7E => ASCII[usize::from(byte - 0x20)],
        0xA0..=0xFF => HIGH[usize::from(byte - 0xA0)],
        _ => mac_roman_char(byte),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn mac_roman_high_half_maps_to_unicode() {
        assert_eq!(mac_roman_to_string(b"abc"), "abc");
        assert_eq!(mac_roman_to_string(&[0x80, 0xc9, 0xde]), "Ä…ﬁ");
        assert_eq!(mac_roman_to_string(&[0xb7, 0xb8, 0xca, 0xff]), "∑∏\u{a0}ˇ");
    }

    #[test]
    fn symbol_encoding_maps_greek_and_operators() {
        assert_eq!(symbol_to_string(b"abp"), "αβπ");
        assert_eq!(symbol_to_string(&[0xAE, 0xB4, 0xE5, 0x6D]), "→×∑μ");
    }

    #[test]
    fn rect_extents_distinguish_empty_and_reversed() {
        let rect = PictRect {
            top: 4,
            left: 3,
            bottom: 10,
            right: 9,
        };
        assert_eq!((rect.width(), rect.height()), (6, 6));
        assert!(rect.is_nonempty());
        let reversed = PictRect {
            top: 4,
            left: 9,
            bottom: 10,
            right: 3,
        };
        assert!(!reversed.is_nonempty());
        assert_eq!(
            PictRect::read(&[0, 1, 0, 2, 0, 3, 0, 4], 0).unwrap().right,
            4
        );
        assert!(PictRect::read(&[0; 7], 0).is_none());
    }
}
