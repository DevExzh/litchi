/// Character Properties (CHP) parser for DOC files.
///
/// CHP structures define character-level formatting such as:
/// - Font properties (bold, italic, underline, strikethrough)
/// - Font size and name
/// - Text color and highlighting
/// - Superscript/subscript
use super::super::package::Result;
use super::super::super::binary::{read_u16_le, read_u32_le};

/// Character Properties structure.
///
/// Contains formatting information for a run of text.
#[derive(Debug, Clone, Default)]
pub struct CharacterProperties {
    /// Bold text
    pub is_bold: Option<bool>,
    /// Italic text
    pub is_italic: Option<bool>,
    /// Underline style
    pub underline: UnderlineStyle,
    /// Strikethrough
    pub is_strikethrough: Option<bool>,
    /// Font size in half-points (e.g., 24 = 12pt)
    pub font_size: Option<u16>,
    /// Font index in font table
    pub font_index: Option<u16>,
    /// Text color (RGB)
    pub color: Option<(u8, u8, u8)>,
    /// Highlight color
    pub highlight: Option<HighlightColor>,
    /// Superscript/subscript
    pub vertical_position: VerticalPosition,
    /// Small caps
    pub is_small_caps: Option<bool>,
    /// All caps
    pub is_all_caps: Option<bool>,
    /// Hidden text
    pub is_hidden: Option<bool>,
    /// OLE2 object flag (SPRM_FOLE2 = 0x080A)
    pub is_ole2: bool,
    /// Picture offset for embedded objects
    pub pic_offset: Option<u32>,
}

/// Underline styles supported in DOC format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum UnderlineStyle {
    /// No underline
    #[default]
    None,
    /// Single underline
    Single,
    /// Double underline
    Double,
    /// Dotted underline
    Dotted,
    /// Dashed underline
    Dashed,
    /// Wavy underline
    Wavy,
    /// Thick underline
    Thick,
    /// Word-only underline (skip spaces)
    WordsOnly,
}

/// Highlight colors available in DOC format.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HighlightColor {
    None,
    Black,
    Blue,
    Cyan,
    Green,
    Magenta,
    Red,
    Yellow,
    White,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
}

/// Vertical text position (superscript/subscript).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum VerticalPosition {
    /// Normal position
    #[default]
    Normal,
    /// Superscript
    Superscript,
    /// Subscript
    Subscript,
}

impl CharacterProperties {
    /// Create a new CharacterProperties with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse character properties from SPRM (Single Property Modifier) data.
    ///
    /// SPRMs are variable-length records that modify properties.
    /// Format: opcode (1 or 2 bytes) + operand (variable length)
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (property modifications)
    pub fn from_sprm(grpprl: &[u8]) -> Result<Self> {
        let mut chp = Self::default();
        let mut offset = 0;
        
        static mut CALL_COUNT: usize = 0;
        unsafe {
            CALL_COUNT += 1;
            if CALL_COUNT <= 5 {
                eprintln!("DEBUG: CharacterProperties::from_sprm called with {} bytes", grpprl.len());
            }
        }

        while offset < grpprl.len() {
            if offset + 1 > grpprl.len() {
                break;
            }

            // Read SPRM opcode (can be 1 or 2 bytes depending on Word version)
            let sprm = read_u16_le(&grpprl, offset).unwrap_or(0);
            offset += 2;
            
            // Debug: Log every SPRM encountered
            static mut SPRM_COUNT: usize = 0;
            unsafe {
                SPRM_COUNT += 1;
                if SPRM_COUNT <= 50 {
                    eprintln!("DEBUG: SPRM opcode: 0x{:04X}", sprm);
                }
            }

            // Parse SPRM based on opcode
            match sprm {
                // Bold (sprmCFBold)
                0x0835 | 0x0085 => {
                    if offset < grpprl.len() {
                        chp.is_bold = Some(grpprl[offset] != 0);
                        offset += 1;
                    }
                }
                // Italic (sprmCFItalic)
                0x0836 | 0x0086 => {
                    if offset < grpprl.len() {
                        chp.is_italic = Some(grpprl[offset] != 0);
                        offset += 1;
                    }
                }
                // Underline (sprmCKul)
                0x2A3E | 0x003E => {
                    if offset < grpprl.len() {
                        chp.underline = match grpprl[offset] {
                            0 => UnderlineStyle::None,
                            1 => UnderlineStyle::Single,
                            2 => UnderlineStyle::WordsOnly,
                            3 => UnderlineStyle::Double,
                            4 => UnderlineStyle::Dotted,
                            6 => UnderlineStyle::Thick,
                            7 => UnderlineStyle::Dashed,
                            11 => UnderlineStyle::Wavy,
                            _ => UnderlineStyle::Single,
                        };
                        offset += 1;
                    }
                }
                // Strikethrough (sprmCFStrike)
                0x0837 | 0x0087 => {
                    if offset < grpprl.len() {
                        chp.is_strikethrough = Some(grpprl[offset] != 0);
                        offset += 1;
                    }
                }
                // Font size in half-points (sprmCHps)
                0x4A43 | 0x0043 => {
                    if offset + 1 < grpprl.len() {
                        chp.font_size = Some(read_u16_le(&grpprl, offset).unwrap_or(0));
                        offset += 2;
                    }
                }
                // Font (sprmCRgFtc0) - ASCII font
                0x4A4F | 0x004F => {
                    if offset + 1 < grpprl.len() {
                        chp.font_index = Some(read_u16_le(&grpprl, offset).unwrap_or(0));
                        offset += 2;
                    }
                }
                // Small caps (sprmCFSmallCaps)
                0x0838 | 0x0088 => {
                    if offset < grpprl.len() {
                        chp.is_small_caps = Some(grpprl[offset] != 0);
                        offset += 1;
                    }
                }
                // All caps (sprmCFCaps)
                0x0839 | 0x0089 => {
                    if offset < grpprl.len() {
                        chp.is_all_caps = Some(grpprl[offset] != 0);
                        offset += 1;
                    }
                }
                // Hidden (sprmCFVanish)
                0x083A | 0x008A => {
                    if offset < grpprl.len() {
                        chp.is_hidden = Some(grpprl[offset] != 0);
                        offset += 1;
                    }
                }
                // Color (sprmCIco)
                0x2A42 | 0x0042 => {
                    if offset < grpprl.len() {
                        // Standard colors (0-16)
                        let color_index = grpprl[offset];
                        chp.color = match color_index {
                            1 => Some((0, 0, 0)),       // Black
                            2 => Some((0, 0, 255)),     // Blue
                            3 => Some((0, 255, 255)),   // Cyan
                            4 => Some((0, 255, 0)),     // Green
                            5 => Some((255, 0, 255)),   // Magenta
                            6 => Some((255, 0, 0)),     // Red
                            7 => Some((255, 255, 0)),   // Yellow
                            8 => Some((255, 255, 255)), // White
                            _ => None,
                        };
                        offset += 1;
                    }
                }
                // Highlight color (sprmCHighlight)
                0x2A0C | 0x000C => {
                    if offset < grpprl.len() {
                        chp.highlight = match grpprl[offset] {
                            0 => Some(HighlightColor::None),
                            1 => Some(HighlightColor::Black),
                            2 => Some(HighlightColor::Blue),
                            3 => Some(HighlightColor::Cyan),
                            4 => Some(HighlightColor::Green),
                            5 => Some(HighlightColor::Magenta),
                            6 => Some(HighlightColor::Red),
                            7 => Some(HighlightColor::Yellow),
                            8 => Some(HighlightColor::White),
                            9 => Some(HighlightColor::DarkBlue),
                            10 => Some(HighlightColor::DarkCyan),
                            11 => Some(HighlightColor::DarkGreen),
                            12 => Some(HighlightColor::DarkMagenta),
                            13 => Some(HighlightColor::DarkRed),
                            14 => Some(HighlightColor::DarkYellow),
                            15 => Some(HighlightColor::DarkGray),
                            16 => Some(HighlightColor::LightGray),
                            _ => None,
                        };
                        offset += 1;
                    }
                }
                // Superscript/subscript (sprmCIss)
                0x2A3F | 0x003F => {
                    if offset < grpprl.len() {
                        chp.vertical_position = match grpprl[offset] {
                            1 => VerticalPosition::Superscript,
                            2 => VerticalPosition::Subscript,
                            _ => VerticalPosition::Normal,
                        };
                        offset += 1;
                    }
                }
                // OLE2 object flag (SPRM_FOLE2)
                0x080A => {
                    if offset < grpprl.len() {
                        let operand = grpprl[offset];
                        chp.is_ole2 = operand != 0;
                        eprintln!("DEBUG: Found SPRM_FOLE2, operand=0x{:02X}, is_ole2={}", operand, chp.is_ole2);
                        offset += 1;
                    }
                }
                // Object location/pic offset (SPRM_OBJLOCATION = 0x680E)
                0x680E => {
                    if offset + 3 < grpprl.len() {
                        chp.pic_offset = Some(read_u32_le(&grpprl, offset).unwrap_or(0));
                        eprintln!("DEBUG: Found SPRM_OBJLOCATION, pic_offset={:?}", chp.pic_offset);
                        offset += 4;
                    }
                }
                // Unknown SPRM - skip based on size
                _ => {
                    // SPRMs have different sizes based on their type
                    // This is a simplified approach - real implementation would need full SPRM table
                    let size = Self::get_sprm_size(sprm);
                    offset += size;
                }
            }
        }

        Ok(chp)
    }

    /// Get the size of an SPRM operand based on its opcode.
    ///
    /// This is a simplified version. Full implementation would have complete SPRM tables.
    fn get_sprm_size(sprm: u16) -> usize {
        // Extract type from SPRM (bits 0-2 in newer versions)
        let sprm_type = sprm & 0x07;
        match sprm_type {
            0 | 1 => 1,  // 1-byte operand
            2 | 4 | 5 => 2,  // 2-byte operand
            3 => 4,      // 4-byte operand
            6 => {
                // Variable length - would need to read length from data
                // Default to 1 for safety
                1
            }
            7 => 3,      // 3-byte operand
            _ => 1,
        }
    }

    /// Check if any formatting is applied.
    pub fn has_formatting(&self) -> bool {
        self.is_bold.is_some()
            || self.is_italic.is_some()
            || self.underline != UnderlineStyle::None
            || self.is_strikethrough.is_some()
            || self.font_size.is_some()
            || self.color.is_some()
            || self.highlight.is_some()
            || self.vertical_position != VerticalPosition::Normal
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_chp() {
        let chp = CharacterProperties::new();
        assert_eq!(chp.is_bold, None);
        assert_eq!(chp.is_italic, None);
        assert_eq!(chp.underline, UnderlineStyle::None);
        assert!(!chp.has_formatting());
    }

    #[test]
    fn test_underline_style() {
        let single = UnderlineStyle::Single;
        let double = UnderlineStyle::Double;
        assert_ne!(single, double);
        assert_eq!(single, UnderlineStyle::Single);
    }

    #[test]
    fn test_vertical_position() {
        let normal = VerticalPosition::Normal;
        let super_pos = VerticalPosition::Superscript;
        assert_ne!(normal, super_pos);
    }
}

