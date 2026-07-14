/// Character Properties (CHP) parser for DOC files.
///
/// CHP structures define character-level formatting such as:
/// - Font properties (bold, italic, underline, strikethrough)
/// - Font size and name
/// - Text color and highlighting
/// - Superscript/subscript
/// - Embedded objects and pictures
///
/// Based on Apache POI's CharacterSprmUncompressor and CharacterProperties.
use super::super::package::Result;
use crate::sprm::{Sprm, parse_sprms};
use crate::sprm_operations::*;

/// Character Properties structure.
///
/// Contains formatting information for a run of text.
/// Based on Apache POI's CharacterProperties implementation.
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
    /// Double strikethrough
    pub is_double_strikethrough: Option<bool>,
    /// Font size in half-points (e.g., 24 = 12pt)
    pub font_size: Option<u16>,
    /// Font index in font table (ASCII characters)
    pub font_index: Option<u16>,
    /// Font index for Far East characters
    pub font_index_fe: Option<u16>,
    /// Font index for other characters
    pub font_index_other: Option<u16>,
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
    /// OLE2 object flag
    pub is_ole2: bool,
    /// Object flag (fObj)
    pub is_obj: bool,
    /// Special character flag (fSpec)
    pub is_spec: bool,
    /// Data flag (fData) - if true, pic_offset points to NilPICFAndBinData, not picture
    pub is_data: bool,
    /// Picture offset for embedded objects (fc in Data stream)
    pub pic_offset: Option<u32>,
    /// Object offset (fcObj)
    pub obj_offset: Option<u32>,
    /// Outline (hollow)
    pub is_outline: Option<bool>,
    /// Shadow
    pub is_shadow: Option<bool>,
    /// Embossed
    pub is_emboss: Option<bool>,
    /// Imprinted (engraved)
    pub is_imprint: Option<bool>,
    /// Character spacing in twips
    pub char_spacing: Option<i16>,
    /// Kerning in half-points
    pub kerning: Option<u16>,
    /// Character scale percentage
    pub char_scale: Option<u16>,
    /// Language ID
    pub language_id: Option<u16>,
    /// Far East language ID.
    pub language_id_fe: Option<u16>,
    /// Whether this run uses bidirectional/complex-script formatting.
    pub is_bidi: Option<bool>,
    /// Complex-script bold setting.
    pub is_bold_bidi: Option<bool>,
    /// Complex-script italic setting.
    pub is_italic_bidi: Option<bool>,
    /// Complex-script font index.
    pub font_index_bidi: Option<u16>,
    /// Complex-script language ID.
    pub language_id_bidi: Option<u16>,
    /// Complex-script indexed text color.
    pub color_index_bidi: Option<u16>,
    /// Complex-script font size in half-points.
    pub font_size_bidi: Option<u16>,
    /// Font/language bias for characters shared by multiple scripts.
    pub script_hint: Option<CharacterScriptHint>,
    /// Whether spelling and grammar proofing excludes this run.
    pub is_no_proof: Option<bool>,
    /// Style index (istd)
    pub style_index: Option<u16>,
    /// Vanish (hidden)
    pub is_vanish: Option<bool>,
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
    /// Dash-dot underline
    DashDot,
    /// Dash-dot-dot underline
    DashDotDot,
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

/// Script bias stored by `sprmCIdctHint`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterScriptHint {
    /// Bias toward non-Far-East properties.
    Default,
    /// Bias toward Far East properties.
    FarEast,
    /// Bias toward complex-script properties.
    ComplexScript,
    /// A reserved value retained for forward compatibility.
    Reserved(u8),
}

impl From<u8> for CharacterScriptHint {
    fn from(value: u8) -> Self {
        match value {
            0 => Self::Default,
            1 => Self::FarEast,
            2 => Self::ComplexScript,
            value => Self::Reserved(value),
        }
    }
}

// Re-export common VerticalPosition type
pub use litchi_core::VerticalPosition;

impl CharacterProperties {
    /// Create a new CharacterProperties with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse character properties from SPRM (Single Property Modifier) data.
    ///
    /// SPRMs are variable-length records that modify properties.
    /// Format: 2-byte opcode + variable-length operand
    ///
    /// Based on Apache POI's CharacterSprmUncompressor.
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (property modifications)
    pub fn from_sprm(grpprl: &[u8]) -> Result<Self> {
        let mut chp = Self::default();
        let sprms = parse_sprms(grpprl);

        for sprm in &sprms {
            if get_sprm_type(sprm.opcode) == 2 {
                Self::apply_sprm(&mut chp, sprm);
            }
        }

        Ok(chp)
    }

    /// Apply a single SPRM operation to character properties.
    ///
    /// Based on Apache POI's CharacterSprmUncompressor.unCompressCHPOperation().
    ///
    /// # Arguments
    ///
    /// * `chp` - The character properties to modify
    /// * `sprm` - The SPRM operation to apply
    fn apply_sprm(chp: &mut CharacterProperties, sprm: &Sprm) {
        // Extract operation code (bits 0-8 of opcode)
        let operation = get_sprm_operation(sprm.opcode);

        match operation {
            // Operation 0x00: sprmCFRMarkDel - Mark deleted revision
            0x00 => {
                // Not commonly used in basic text extraction
            },
            // Operation 0x01: sprmCFRMark - Mark revision
            0x01 => {
                // Not commonly used in basic text extraction
            },
            // Operation 0x02: sprmCFFldVanish - Field vanish flag
            0x02 => {
                // Not commonly used in basic text extraction
            },
            // Operation 0x03: sprmCPicLocation - Picture/object location
            0x03 => {
                if let Some(fc) = sprm.operand_dword() {
                    chp.pic_offset = Some(fc);
                    chp.is_spec = true;
                }
            },
            // Operation 0x04: sprmCIbstRMark - Revision mark author
            0x04 => {
                // Not commonly used in basic text extraction
            },
            // Operation 0x05: sprmCDttmRMark - Revision mark date/time
            0x05 => {
                // Not commonly used in basic text extraction
            },
            // Operation 0x06: sprmCFData - Data flag
            0x06 => {
                // Data field flag
                debug_assert_eq!(sprm.size, 3);
                if let Some(val) = sprm.operand_byte() {
                    chp.is_data = val != 0;
                }
            },
            // Operation 0x07: sprmCIdslRMark - Revision mark ID
            0x07 => {
                // Not commonly used in basic text extraction
            },
            // Operation 0x08: sprmCChs - Complex character set
            0x08 => {
                // Complex character set handling
            },
            // Operation 0x09: sprmCSymbol - Symbol character
            0x09 => {
                chp.is_spec = true;
                // Symbol character - would need font and character code
            },
            // Operation 0x0A: sprmCFOle2 - OLE2 object flag
            0x0A => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_ole2 = val != 0;
                }
            },
            // Operation 0x0C: sprmCIcoHighlight - Highlight color
            0x0C => {
                if let Some(val) = sprm.operand_byte() {
                    chp.highlight = match val {
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
                }
            },
            // Operation 0x0E: sprmCObjLocation - Object location
            0x0E => {
                if let Some(fc) = sprm.operand_dword() {
                    chp.obj_offset = Some(fc);
                }
            },
            // Operations 0x11-0x2F: Various flags and properties
            0x11 => {
                // sprmCFWebHidden - Web hidden
            },
            0x15 => {
                // sprmCRsidProp - Revision save ID property
            },
            0x16 => {
                // sprmCRsidText - Revision save ID text
            },
            0x17 => {
                // sprmCRsidRMDel - Revision save ID deletion
            },
            0x18 => {
                // sprmCFSpecVanish - Special vanish
            },
            0x1A => {
                // sprmCFMathPr - Math properties
            },
            // Operation 0x30: sprmCIstd - Style index
            0x30 => {
                if let Some(istd) = sprm.operand_word() {
                    chp.style_index = Some(istd);
                }
            },
            // Operation 0x31: sprmCIstdPermute - Style permutation
            0x31 => {
                // Style permutation for fast saves
            },
            // Operation 0x32: sprmCDefault - Reset to default
            0x32 => {
                // Reset formatting to defaults
                chp.is_bold = Some(false);
                chp.is_italic = Some(false);
                chp.is_outline = Some(false);
                chp.is_strikethrough = Some(false);
                chp.is_shadow = Some(false);
                chp.is_small_caps = Some(false);
                chp.is_all_caps = Some(false);
                chp.is_vanish = Some(false);
                chp.underline = UnderlineStyle::None;
                chp.color = None;
            },
            // Operation 0x33: sprmCPlain - Plain text (reset all)
            0x33 => {
                // Reset to plain - preserve fSpec
                let preserve_spec = chp.is_spec;
                *chp = Self::default();
                chp.is_spec = preserve_spec;
            },
            // Operation 0x34: sprmCKcd - Keyboard code
            0x34 => {
                // Keyboard code - not commonly used
            },
            // Operation 0x35: sprmCFBold - Bold
            0x35 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_bold = Some(Self::get_toggle_value(val, chp.is_bold));
                }
            },
            // Operation 0x36: sprmCFItalic - Italic
            0x36 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_italic = Some(Self::get_toggle_value(val, chp.is_italic));
                }
            },
            // Operation 0x37: sprmCFStrike - Strikethrough
            0x37 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_strikethrough = Some(Self::get_toggle_value(val, chp.is_strikethrough));
                }
            },
            // Operation 0x38: sprmCFOutline - Outline
            0x38 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_outline = Some(Self::get_toggle_value(val, chp.is_outline));
                }
            },
            // Operation 0x39: sprmCFShadow - Shadow
            0x39 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_shadow = Some(Self::get_toggle_value(val, chp.is_shadow));
                }
            },
            // Operation 0x3A: sprmCFSmallCaps - Small caps
            0x3A => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_small_caps = Some(Self::get_toggle_value(val, chp.is_small_caps));
                }
            },
            // Operation 0x3B: sprmCFCaps - All caps
            0x3B => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_all_caps = Some(Self::get_toggle_value(val, chp.is_all_caps));
                }
            },
            // Operation 0x3C: sprmCFVanish - Hidden
            0x3C => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_vanish = Some(Self::get_toggle_value(val, chp.is_vanish));
                }
            },
            // Operation 0x3D: sprmCFtcDefault - Default font
            0x3D => {
                if let Some(ftc) = sprm.operand_word() {
                    chp.font_index = Some(ftc);
                }
            },
            // Operation 0x3E: sprmCKul - Underline style
            0x3E => {
                if let Some(val) = sprm.operand_byte() {
                    chp.underline = match val {
                        0 => UnderlineStyle::None,
                        1 => UnderlineStyle::Single,
                        2 => UnderlineStyle::WordsOnly,
                        3 => UnderlineStyle::Double,
                        4 => UnderlineStyle::Dotted,
                        5 => UnderlineStyle::Thick, // Hidden - POI maps to Thick
                        6 => UnderlineStyle::Dashed,
                        7 => UnderlineStyle::DashDot,
                        8 => UnderlineStyle::DashDotDot,
                        9 => UnderlineStyle::Wavy,
                        10 => UnderlineStyle::Thick,
                        11 => UnderlineStyle::Thick, // DottedHeavy - map to Thick
                        _ => UnderlineStyle::Single,
                    };
                }
            },
            // Operation 0x3F: sprmCSizePos - Size and position (complex)
            0x3F => {
                if let Some(operand) = sprm.operand_dword() {
                    let hps = operand & 0xFF;
                    if hps != 0 {
                        chp.font_size = Some(hps as u16);
                    }

                    let c_inc = ((operand & 0xFF00) >> 8) as i8;
                    let c_inc = c_inc >> 1;
                    if c_inc != 0 {
                        let current = chp.font_size.unwrap_or(24);
                        chp.font_size = Some((current as i32 + c_inc as i32 * 2).max(2) as u16);
                    }

                    let hps_pos = ((operand & 0xFF0000) >> 16) as i8;
                    if hps_pos != -128_i8 {
                        // Set position
                    }
                }
            },
            // Operation 0x40: sprmCDxaSpace - Character spacing
            0x40 => {
                if let Some(val) = sprm.operand_i16() {
                    chp.char_spacing = Some(val);
                }
            },
            // Operation 0x41: sprmCLid - Language ID
            0x41 => {
                if let Some(lid) = sprm.operand_word() {
                    chp.language_id = Some(lid);
                }
            },
            // Operation 0x42: sprmCIco - Text color index
            0x42 => {
                if let Some(color_index) = sprm.operand_byte() {
                    chp.color = match color_index {
                        0 => None,                   // Auto
                        1 => Some((0, 0, 0)),        // Black
                        2 => Some((0, 0, 255)),      // Blue
                        3 => Some((0, 255, 255)),    // Cyan
                        4 => Some((0, 255, 0)),      // Green
                        5 => Some((255, 0, 255)),    // Magenta
                        6 => Some((255, 0, 0)),      // Red
                        7 => Some((255, 255, 0)),    // Yellow
                        8 => Some((255, 255, 255)),  // White
                        9 => Some((0, 0, 128)),      // Dark Blue
                        10 => Some((0, 128, 128)),   // Dark Cyan
                        11 => Some((0, 128, 0)),     // Dark Green
                        12 => Some((128, 0, 128)),   // Dark Magenta
                        13 => Some((128, 0, 0)),     // Dark Red
                        14 => Some((128, 128, 0)),   // Dark Yellow
                        15 => Some((128, 128, 128)), // Dark Gray
                        16 => Some((192, 192, 192)), // Light Gray
                        _ => None,
                    };
                }
            },
            // Operation 0x43: sprmCHps - Font size in half-points
            0x43 => {
                if let Some(hps) = sprm.operand_word() {
                    chp.font_size = Some(hps);
                }
            },
            // Operation 0x44: sprmCHpsInc - Font size increment
            0x44 => {
                if let Some(inc) = sprm.operand_byte() {
                    let current = chp.font_size.unwrap_or(24);
                    chp.font_size = Some((current as i32 + inc as i32 * 2).max(2) as u16);
                }
            },
            // Operation 0x45: sprmCHpsPos - Superscript/subscript position
            0x45 => {
                if let Some(_pos) = sprm.operand_i16() {
                    // Position in half-points
                }
            },
            // Operation 0x46: sprmCHpsPosAdj - Position adjustment
            0x46 => {
                // Position adjustment
            },
            // Operation 0x47: sprmCMajority - Majority formatting
            0x47 => {
                // Complex majority formatting - not commonly used
            },
            // Operation 0x48: sprmCIss - Superscript/subscript
            0x48 => {
                if let Some(iss) = sprm.operand_byte() {
                    chp.vertical_position = match iss {
                        0 => VerticalPosition::Normal,
                        1 => VerticalPosition::Superscript,
                        2 => VerticalPosition::Subscript,
                        _ => VerticalPosition::Normal,
                    };
                }
            },
            // Operation 0x49: sprmCHpsNew50 - Font size (Word 6.0)
            0x49 => {
                if let Some(hps) = sprm.operand_word() {
                    chp.font_size = Some(hps);
                }
            },
            // Operation 0x4A: sprmCHpsInc1 - Font size increment
            0x4A => {
                if let Some(inc) = sprm.operand_i16() {
                    let current = chp.font_size.unwrap_or(24);
                    chp.font_size = Some((current as i32 + inc as i32).max(8) as u16);
                }
            },
            // Operation 0x4B: sprmCHpsKern - Kerning
            0x4B => {
                if let Some(kern) = sprm.operand_word() {
                    chp.kerning = Some(kern);
                }
            },
            // Operation 0x4C: sprmCMajority50 - Majority formatting (Word 6.0)
            0x4C => {
                // Complex majority formatting
            },
            // Operation 0x4D: sprmCHpsMul - Font size multiplier
            0x4D => {
                if let Some(multiplier) = sprm.operand_word() {
                    let percentage = multiplier as f32 / 100.0;
                    let current = chp.font_size.unwrap_or(24);
                    let add = (percentage * current as f32) as i32;
                    chp.font_size = Some((current as i32 + add) as u16);
                }
            },
            // Operation 0x4E: sprmCHresi - Hyphenation
            0x4E => {
                // Hyphenation information
            },
            // Operation 0x4F: sprmCRgFtc0 - Font for ASCII
            0x4F => {
                if let Some(ftc) = sprm.operand_word() {
                    chp.font_index = Some(ftc);
                }
            },
            // Operation 0x50: sprmCRgFtc1 - Font for Far East
            0x50 => {
                if let Some(ftc) = sprm.operand_word() {
                    chp.font_index_fe = Some(ftc);
                }
            },
            // Operation 0x51: sprmCRgFtc2 - Font for other
            0x51 => {
                if let Some(ftc) = sprm.operand_word() {
                    chp.font_index_other = Some(ftc);
                }
            },
            // Operation 0x52: sprmCCharScale - Character scale
            0x52 => {
                if let Some(scale) = sprm.operand_word() {
                    chp.char_scale = Some(scale);
                }
            },
            // Operation 0x53: sprmCFDStrike - Double strikethrough
            0x53 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_double_strikethrough =
                        Some(Self::get_toggle_value(val, chp.is_double_strikethrough));
                }
            },
            // Operation 0x54: sprmCFImprint - Imprint
            0x54 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_imprint = Some(val != 0);
                }
            },
            // Operation 0x55: sprmCFSpec - Special character flag
            0x55 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_spec = val != 0;
                }
            },
            // Operation 0x56: sprmCFObj - Object flag
            0x56 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_obj = val != 0;
                }
            },
            // Operation 0x57: sprmCPropRMark - Property revision mark
            0x57 => {
                // Revision mark properties
            },
            // Operation 0x58: sprmCFEmboss - Emboss
            0x58 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_emboss = Some(val != 0);
                }
            },
            // Operation 0x59: sprmCSfxtText - Text animation
            0x59 => {
                // Text animation effect
            },
            // Operation 0x5A: sprmCFBiDi - Complex-script formatting.
            0x5A => {
                if let Some(value) = sprm.operand_byte() {
                    chp.is_bidi = Some(Self::get_toggle_value(value, chp.is_bidi));
                }
            },
            // Operation 0x5C: sprmCFBoldBi - Complex-script bold.
            0x5C => {
                if let Some(value) = sprm.operand_byte() {
                    chp.is_bold_bidi = Some(Self::get_toggle_value(value, chp.is_bold_bidi));
                }
            },
            // Operation 0x5D: sprmCFItalicBi - Complex-script italic.
            0x5D => {
                if let Some(value) = sprm.operand_byte() {
                    chp.is_italic_bidi = Some(Self::get_toggle_value(value, chp.is_italic_bidi));
                }
            },
            // Operation 0x5E: sprmCFtcBi - Complex-script font.
            0x5E => chp.font_index_bidi = sprm.operand_word(),
            // Operation 0x5F: sprmCLidBi - Complex-script language.
            0x5F => chp.language_id_bidi = sprm.operand_word(),
            // Operation 0x60: sprmCIcoBi - Complex-script indexed color.
            0x60 => chp.color_index_bidi = sprm.operand_word(),
            // Operation 0x61: sprmCHpsBi - Complex-script size.
            0x61 => chp.font_size_bidi = sprm.operand_word(),
            // Operations 0x6D and 0x73: legacy and current default language.
            0x6D | 0x73 => chp.language_id = sprm.operand_word(),
            // Operations 0x6E and 0x74: legacy and current Far East language.
            0x6E | 0x74 => chp.language_id_fe = sprm.operand_word(),
            // Operation 0x6F: sprmCIdctHint - Script bias.
            0x6F => chp.script_hint = sprm.operand_byte().map(CharacterScriptHint::from),
            // Operation 0x70: sprmCCv - Color value (RGB)
            0x70 => {
                if let Some(cv) = sprm.operand_dword() {
                    // Extract RGB from COLORREF (0x00BBGGRR)
                    let r = (cv & 0xFF) as u8;
                    let g = ((cv >> 8) & 0xFF) as u8;
                    let b = ((cv >> 16) & 0xFF) as u8;
                    chp.color = Some((r, g, b));
                }
            },
            // Operation 0x75: sprmCFNoProof - Exclude from proofing.
            0x75 => {
                if let Some(value) = sprm.operand_byte() {
                    chp.is_no_proof = Some(Self::get_toggle_value(value, chp.is_no_proof));
                }
            },
            // Remaining border, shading, and revision SPRMs.
            0x5B | 0x62..=0x6C | 0x71 | 0x72 => {
                // Retained for future structured border/revision metadata.
            },
            // Default: Unknown or unsupported SPRM
            _ => {
                // Silently ignore unknown SPRMs
            },
        }
    }

    /// Get toggle value from SPRM operand.
    ///
    /// Based on Apache POI's getCHPFlag method.
    ///
    /// # Arguments
    ///
    /// * `operand` - The SPRM operand byte
    /// * `old_val` - The previous value
    ///
    /// # Returns
    ///
    /// The new boolean value based on the toggle logic:
    /// - 0: false
    /// - 1: true
    /// - 0x80: preserve old value
    /// - 0x81: toggle old value
    fn get_toggle_value(operand: u8, old_val: Option<bool>) -> bool {
        match operand {
            0 => false,
            1 => true,
            0x80 => old_val.unwrap_or(false),
            0x81 => !old_val.unwrap_or(false),
            _ => false,
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
            || self.is_bidi.is_some()
            || self.is_bold_bidi.is_some()
            || self.is_italic_bidi.is_some()
            || self.font_index_bidi.is_some()
            || self.language_id_bidi.is_some()
            || self.color_index_bidi.is_some()
            || self.font_size_bidi.is_some()
            || self.language_id_fe.is_some()
            || self.script_hint.is_some()
            || self.is_no_proof.is_some()
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

    #[test]
    fn test_toggle_value() {
        // Test basic values
        assert!(!CharacterProperties::get_toggle_value(0, None));
        assert!(CharacterProperties::get_toggle_value(1, None));

        // Test preserve old value
        assert!(CharacterProperties::get_toggle_value(0x80, Some(true)));
        assert!(!CharacterProperties::get_toggle_value(0x80, Some(false)));

        // Test toggle old value
        assert!(!CharacterProperties::get_toggle_value(0x81, Some(true)));
        assert!(CharacterProperties::get_toggle_value(0x81, Some(false)));
    }

    #[test]
    fn ignores_non_character_sprms_in_mixed_piece_modifier() {
        let properties = CharacterProperties::from_sprm(&[
            0x03, 0x24, 0x02, // paragraph justification
            0x35, 0x08, 0x01, // character bold
        ])
        .unwrap();
        assert_eq!(properties.is_bold, Some(true));
        assert_eq!(properties.is_strikethrough, None);
    }

    #[test]
    fn parses_complex_script_language_and_proofing_sprms() {
        let properties = CharacterProperties::from_sprm(&[
            0x5A, 0x08, 0x01, // sprmCFBiDi
            0x5C, 0x08, 0x01, // sprmCFBoldBi
            0x5C, 0x08, 0x81, // toggle complex-script bold back off
            0x5D, 0x08, 0x01, // sprmCFItalicBi
            0x5E, 0x4A, 0x34, 0x12, // sprmCFtcBi
            0x5F, 0x48, 0x01, 0x04, // sprmCLidBi
            0x60, 0x4A, 0x0D, 0x00, // sprmCIcoBi
            0x61, 0x4A, 0x1C, 0x00, // sprmCHpsBi
            0x6D, 0x48, 0x09, 0x04, // sprmCRgLid0_80
            0x6E, 0x48, 0x11, 0x04, // sprmCRgLid1_80
            0x73, 0x48, 0x0C, 0x04, // sprmCRgLid0 supersedes legacy
            0x74, 0x48, 0x12, 0x04, // sprmCRgLid1 supersedes legacy
            0x6F, 0x28, 0x02, // sprmCIdctHint
            0x75, 0x08, 0x01, // sprmCFNoProof
        ])
        .unwrap();

        assert_eq!(properties.is_bidi, Some(true));
        assert_eq!(properties.is_bold_bidi, Some(false));
        assert_eq!(properties.is_italic_bidi, Some(true));
        assert_eq!(properties.font_index_bidi, Some(0x1234));
        assert_eq!(properties.language_id_bidi, Some(0x0401));
        assert_eq!(properties.color_index_bidi, Some(13));
        assert_eq!(properties.font_size_bidi, Some(28));
        assert_eq!(properties.language_id, Some(0x040C));
        assert_eq!(properties.language_id_fe, Some(0x0412));
        assert_eq!(
            properties.script_hint,
            Some(CharacterScriptHint::ComplexScript)
        );
        assert_eq!(properties.is_no_proof, Some(true));
        assert!(properties.has_formatting());
    }

    #[test]
    fn preserves_reserved_script_hint_values() {
        let properties = CharacterProperties::from_sprm(&[0x6F, 0x28, 0xFF]).unwrap();
        assert_eq!(
            properties.script_hint,
            Some(CharacterScriptHint::Reserved(0xFF))
        );
    }
}
