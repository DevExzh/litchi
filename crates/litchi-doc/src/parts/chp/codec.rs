use super::super::super::package::{Error as PackageError, Result};
use super::super::tap::TableStyleCondition;
use super::{
    CharacterBorder, CharacterBorderStyle, CharacterColor, CharacterConditionalFormatting,
    CharacterPosition, CharacterProperties, CharacterScriptHint, CharacterShading,
    CharacterShadingPattern, DisplayFieldRevisionProperties, HighlightColor, HresiOperand,
    HyphenationMode, TextEffect, UnderlineStyle, VerticalPosition,
};
use crate::sprm::{Sprm, parse_sprms};
use crate::sprm_operations::{SPRM_C_CNF, get_sprm_operation, get_sprm_type};

impl HyphenationMode {
    const fn raw(self) -> u8 {
        match self {
            Self::Normal => 1,
            Self::AddBefore => 2,
            Self::ChangeBefore => 3,
            Self::DeleteBefore => 4,
            Self::ChangeAfter => 5,
            Self::DeleteAndChange => 6,
        }
    }

    fn from_raw(value: u8) -> Result<Self> {
        match value {
            1 => Ok(Self::Normal),
            2 => Ok(Self::AddBefore),
            3 => Ok(Self::ChangeBefore),
            4 => Ok(Self::DeleteBefore),
            5 => Ok(Self::ChangeAfter),
            6 => Ok(Self::DeleteAndChange),
            _ => Err(PackageError::Corrupted(format!(
                "sprmCHresi has invalid Hres mode {value}"
            ))),
        }
    }
}

impl HresiOperand {
    fn from_bytes(mode: u8, replacement_character: u8) -> Result<Self> {
        let mode = HyphenationMode::from_raw(mode)?;
        if mode == HyphenationMode::Normal {
            if replacement_character != 0 {
                return Err(PackageError::Corrupted(
                    "normal sprmCHresi requires ChHres 0x00".to_string(),
                ));
            }
            Ok(Self::normal())
        } else {
            Self::with_character(mode, replacement_character)
        }
    }

    pub(crate) fn bytes(self) -> [u8; 2] {
        [self.mode().raw(), self.replacement_character().unwrap_or(0)]
    }
}

impl CharacterColor {
    fn from_colorref(data: &[u8]) -> Option<Self> {
        let [red, green, blue, automatic, ..] = data else {
            return None;
        };
        if *automatic == 0xFF {
            Some(Self::Automatic)
        } else {
            Some(Self::Rgb(*red, *green, *blue))
        }
    }

    fn from_palette_index(index: u8) -> Self {
        match index {
            0 => Self::Automatic,
            1 => Self::Rgb(0, 0, 0),
            2 => Self::Rgb(0, 0, 255),
            3 => Self::Rgb(0, 255, 255),
            4 => Self::Rgb(0, 255, 0),
            5 => Self::Rgb(255, 0, 255),
            6 => Self::Rgb(255, 0, 0),
            7 => Self::Rgb(255, 255, 0),
            8 => Self::Rgb(255, 255, 255),
            9 => Self::Rgb(0, 0, 128),
            10 => Self::Rgb(0, 128, 128),
            11 => Self::Rgb(0, 128, 0),
            12 => Self::Rgb(128, 0, 128),
            13 => Self::Rgb(128, 0, 0),
            14 => Self::Rgb(128, 128, 0),
            15 => Self::Rgb(128, 128, 128),
            16 => Self::Rgb(192, 192, 192),
            value => Self::ReservedPaletteIndex(value),
        }
    }
}

impl CharacterProperties {
    /// Parse character properties from SPRM (Single Property Modifier) data.
    ///
    /// SPRMs are variable-length records that modify properties.
    /// Format: 2-byte opcode + variable-length operand
    ///
    /// Based on Apache POI's `CharacterSprmUncompressor`.
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (property modifications)
    pub fn from_sprm(grpprl: &[u8]) -> Result<Self> {
        let mut chp = Self::default();
        let sprms = parse_sprms(grpprl)?;
        let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
        if consumed != grpprl.len() {
            return Err(PackageError::Corrupted(
                "CHP grpprl does not contain a whole number of SPRMs".to_string(),
            ));
        }

        for sprm in &sprms {
            if get_sprm_type(sprm.opcode) == 2 {
                Self::apply_sprm(&mut chp, sprm)?;
            }
        }

        Ok(chp)
    }

    /// Apply a single SPRM operation to character properties.
    ///
    /// Based on Apache POI's `CharacterSprmUncompressor.unCompressCHPOperation()`.
    ///
    /// # Arguments
    ///
    /// * `chp` - The character properties to modify
    /// * `sprm` - The SPRM operation to apply
    pub(crate) fn apply_sprm(chp: &mut CharacterProperties, sprm: &Sprm) -> Result<()> {
        // Extract operation code (bits 0-8 of opcode)
        let operation = get_sprm_operation(sprm.opcode);

        match operation {
            // Operation 0x00: sprmCFRMarkDel - Mark deleted revision
            0x00 => {
                chp.is_revision_deleted = Some(Self::revision_flag(sprm, "sprmCFRMarkDel")?);
            },
            // Operation 0x01: sprmCFRMark - Mark revision
            0x01 => {
                chp.is_revision_inserted = Some(Self::revision_flag(sprm, "sprmCFRMark")?);
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
                chp.revision_author_index = Some(Self::revision_author(sprm, "sprmCIbstRMark")?);
            },
            // Operation 0x05: sprmCDttmRMark - Revision mark date/time
            0x05 => {
                chp.revision_timestamp = Some(sprm.operand_dword().ok_or_else(|| {
                    PackageError::Corrupted("sprmCDttmRMark is missing its DTTM".to_string())
                })?);
            },
            // Operation 0x06: sprmCFData - Data flag
            0x06 => {
                // Data field flag
                debug_assert_eq!(sprm.size, 3);
                if let Some(val) = sprm.operand_byte() {
                    chp.is_data = val != 0;
                }
            },
            // Operation 0x07: sprmCIdslRMark - Revision edit reason.
            0x07 => {
                chp.revision_id = Some(Self::revision_reason(sprm, "sprmCIdslRMark")?);
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
                chp.formatting_revision_save_id = Some(sprm.operand_dword().ok_or_else(|| {
                    PackageError::Corrupted("sprmCRsidProp is missing its RSID".to_string())
                })?);
            },
            0x16 => {
                // sprmCRsidText - Revision save ID text
                chp.insertion_revision_save_id = Some(sprm.operand_dword().ok_or_else(|| {
                    PackageError::Corrupted("sprmCRsidText is missing its RSID".to_string())
                })?);
            },
            0x17 => {
                // sprmCRsidRMDel - Revision save ID deletion
                chp.deletion_revision_save_id = Some(sprm.operand_dword().ok_or_else(|| {
                    PackageError::Corrupted("sprmCRsidRMDel is missing its RSID".to_string())
                })?);
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
                        chp.font_size =
                            Some((i32::from(current) + i32::from(c_inc) * 2).max(2) as u16);
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
                    chp.font_size = Some((i32::from(current) + i32::from(inc) * 2).max(2) as u16);
                }
            },
            // Operation 0x45: sprmCHpsPos - Superscript/subscript position
            0x45 => {
                let position = sprm.operand_i16().ok_or_else(|| {
                    PackageError::Corrupted("sprmCHpsPos is missing its signed operand".to_string())
                })?;
                chp.position = CharacterPosition::new(position)?;
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
                    chp.font_size = Some((i32::from(current) + i32::from(inc)).max(8) as u16);
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
                    let percentage = f32::from(multiplier) / 100.0;
                    let current = chp.font_size.unwrap_or(24);
                    let add = (percentage * f32::from(current)) as i32;
                    chp.font_size = Some((i32::from(current) + add) as u16);
                }
            },
            // Operation 0x4E: sprmCHresi - Hyphenation
            0x4E => {
                let operand = sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted("sprmCHresi is missing its HresiOperand".to_string())
                })?;
                chp.hyphenation = HresiOperand::from_bytes(operand as u8, (operand >> 8) as u8)?;
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
            // Operations 0x57 and 0x89: sprmCPropRMark90 / sprmCPropRMark.
            0x57 | 0x89 => Self::apply_property_revision(chp, sprm)?,
            // Operation 0x58: sprmCFEmboss - Emboss
            0x58 => {
                if let Some(val) = sprm.operand_byte() {
                    chp.is_emboss = Some(val != 0);
                }
            },
            // Operation 0x59: sprmCSfxtText - Text animation
            0x59 => {
                let effect = sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmCSfxText is missing its byte operand".to_string())
                })?;
                chp.text_effect = TextEffect::try_from(effect)?;
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
            // Operation 0x62: sprmCDispFldRMark.
            0x62 => chp.display_field_revision = Some(Self::parse_display_field_revision(sprm)?),
            // Operation 0x65: sprmCBrc80 - palette-based character border.
            0x65 => chp.border = Self::parse_border(sprm.operand_bytes(), true),
            // Operation 0x66: sprmCShd80 - palette-based character shading.
            0x66 => chp.shading = Self::parse_shading80(sprm.operand_bytes()),
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
            // Operation 0x71: sprmCShd - RGB character shading.
            0x71 => chp.shading = Self::parse_shading(sprm.operand_bytes()),
            // Operation 0x72: sprmCBrc - RGB character border.
            0x72 => chp.border = Self::parse_border(sprm.operand_bytes(), false),
            // Operation 0x75: sprmCFNoProof - Exclude from proofing.
            0x75 => {
                if let Some(value) = sprm.operand_byte() {
                    chp.is_no_proof = Some(Self::get_toggle_value(value, chp.is_no_proof));
                }
            },
            // Operation 0x77: sprmCCvUl - Underline COLORREF.
            0x77 => chp.underline_color = CharacterColor::from_colorref(sprm.operand_bytes()),
            // Operation 0x82: sprmCFComplexScripts - Force complex-script formatting.
            0x82 => {
                if let Some(value) = sprm.operand_byte() {
                    chp.is_complex_scripts =
                        Some(Self::get_toggle_value(value, chp.is_complex_scripts));
                }
            },
            // Operation 0x83: sprmCWall - Preserve pre-revision character properties.
            0x83 => {
                let enabled = Self::strict_bool8(sprm, "sprmCWall")?;
                chp.preserved_properties_for_revision = if enabled {
                    let mut previous = chp.clone();
                    previous.properties_preserved_for_revision = false;
                    previous.preserved_properties_for_revision = None;
                    Some(Box::new(previous))
                } else {
                    None
                };
                chp.properties_preserved_for_revision = enabled;
            },
            // Operation 0x85: sprmCCnf - Conditional table-style character formatting.
            0x85 => chp
                .conditional_formats
                .push(Self::parse_conditional_formatting(sprm)?),
            // Remaining border, shading, and revision SPRMs.
            0x63 => {
                chp.deletion_author_index = Some(Self::revision_author(sprm, "sprmCIbstRMarkDel")?);
            },
            0x64 => {
                chp.deletion_timestamp = Some(sprm.operand_dword().ok_or_else(|| {
                    PackageError::Corrupted("sprmCDttmRMarkDel is missing its DTTM".to_string())
                })?);
            },
            0x67 => {
                chp.deletion_revision_id = Some(Self::revision_reason(sprm, "sprmCIdslRMarkDel")?);
            },
            0x5B | 0x68..=0x6C => {
                // Retained for future structured border/revision metadata.
            },
            // Default: Unknown or unsupported SPRM
            _ => {
                // Silently ignore unknown SPRMs
            },
        }
        Ok(())
    }

    fn revision_flag(sprm: &Sprm, name: &str) -> Result<bool> {
        match sprm.operand_byte() {
            Some(0) => Ok(false),
            Some(1) => Ok(true),
            _ => Err(PackageError::Corrupted(format!(
                "{name} must contain a Boolean8 value"
            ))),
        }
    }

    fn strict_bool8(sprm: &Sprm, name: &str) -> Result<bool> {
        match sprm.operand_byte() {
            Some(0) => Ok(false),
            Some(1) => Ok(true),
            _ => Err(PackageError::Corrupted(format!(
                "{name} must contain a Boolean8 value"
            ))),
        }
    }

    fn parse_conditional_formatting(sprm: &Sprm) -> Result<CharacterConditionalFormatting> {
        let operand = sprm.operand_bytes();
        if operand.len() < 2 {
            return Err(PackageError::Corrupted(
                "sprmCCnf must contain a 2-byte condition".to_string(),
            ));
        }
        let code = u16::from_le_bytes([operand[0], operand[1]]);
        let condition = TableStyleCondition::from_code(code).ok_or_else(|| {
            PackageError::Corrupted(format!("sprmCCnf contains invalid condition {code:#06x}"))
        })?;
        let raw_grpprl = operand[2..].to_vec();
        let nested = parse_sprms(&raw_grpprl)?;
        let consumed = nested
            .last()
            .map_or(0, |nested| nested.offset + nested.size);
        if consumed != raw_grpprl.len() {
            return Err(PackageError::Corrupted(
                "sprmCCnf nested grpprl is truncated".to_string(),
            ));
        }
        if nested.iter().any(|nested| nested.opcode == SPRM_C_CNF) {
            return Err(PackageError::Corrupted(
                "sprmCCnf cannot be nested inside another sprmCCnf".to_string(),
            ));
        }
        if nested
            .iter()
            .any(|nested| get_sprm_type(nested.opcode) != 2)
        {
            return Err(PackageError::Corrupted(
                "sprmCCnf can contain only character SPRMs".to_string(),
            ));
        }
        let properties = Box::new(Self::from_sprm(&raw_grpprl)?);
        Ok(CharacterConditionalFormatting {
            condition,
            properties,
            raw_grpprl,
        })
    }

    fn revision_author(sprm: &Sprm, name: &str) -> Result<u16> {
        let value = sprm.operand_i16().ok_or_else(|| {
            PackageError::Corrupted(format!("{name} is missing its author index"))
        })?;
        u16::try_from(value)
            .map_err(|_| PackageError::Corrupted(format!("{name} author index is negative")))
    }

    fn revision_reason(sprm: &Sprm, name: &str) -> Result<u16> {
        let value = sprm
            .operand_word()
            .ok_or_else(|| PackageError::Corrupted(format!("{name} is missing its reason code")))?;
        if value > super::super::super::revision::RevisionReason::MAX_VALUE {
            return Err(PackageError::Corrupted(format!(
                "{name} contains an undefined reason code"
            )));
        }
        Ok(value)
    }

    fn apply_property_revision(chp: &mut CharacterProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 7 {
            return Err(PackageError::Corrupted(
                "sprmCPropRMark operand must contain exactly 7 bytes".to_string(),
            ));
        }
        chp.has_formatting_revision = Some(match operand[0] {
            0 => false,
            1 => true,
            _ => {
                return Err(PackageError::Corrupted(
                    "sprmCPropRMark must begin with a Boolean8 value".to_string(),
                ));
            },
        });
        let author = i16::from_le_bytes([operand[1], operand[2]]);
        chp.formatting_revision_author_index = Some(u16::try_from(author).map_err(|_| {
            PackageError::Corrupted("sprmCPropRMark author index is negative".to_string())
        })?);
        chp.formatting_revision_timestamp = Some(u32::from_le_bytes([
            operand[3], operand[4], operand[5], operand[6],
        ]));
        Ok(())
    }

    fn parse_display_field_revision(sprm: &Sprm) -> Result<DisplayFieldRevisionProperties> {
        let operand = sprm.operand_bytes();
        if operand.len() != 39 {
            return Err(PackageError::Corrupted(
                "sprmCDispFldRMark operand must contain exactly 39 bytes".to_string(),
            ));
        }
        let active = operand[0] != 0;
        let author_index = u16::from_le_bytes([operand[1], operand[2]]);
        let timestamp = u32::from_le_bytes([operand[3], operand[4], operand[5], operand[6]]);
        let string_length = usize::from(u16::from_le_bytes([operand[7], operand[8]]));
        if string_length > 15 {
            return Err(PackageError::Corrupted(
                "LISTNUM previous result exceeds its 15-code-unit XST".to_string(),
            ));
        }
        let units = (0..string_length)
            .map(|index| {
                let offset = 9 + index * 2;
                u16::from_le_bytes([operand[offset], operand[offset + 1]])
            })
            .collect::<Vec<_>>();
        let previous_result = String::from_utf16(&units).map_err(|_| {
            PackageError::Corrupted("LISTNUM previous result is invalid UTF-16".to_string())
        })?;
        Ok(DisplayFieldRevisionProperties {
            active,
            author_index,
            timestamp,
            previous_result,
        })
    }

    fn parse_border(data: &[u8], palette_color: bool) -> Option<CharacterBorder> {
        let (color, width, style, flags) = if palette_color {
            let [width, style, color, flags, ..] = data else {
                return None;
            };
            (
                CharacterColor::from_palette_index(*color),
                *width,
                *style,
                u16::from(*flags),
            )
        } else {
            let [color @ .., width, style, flags_low, flags_high] = data else {
                return None;
            };
            if color.len() != 4 {
                return None;
            }
            (
                CharacterColor::from_colorref(color)?,
                *width,
                *style,
                u16::from_le_bytes([*flags_low, *flags_high]),
            )
        };

        if style == 0 || style == 0xFF {
            return None;
        }

        Some(CharacterBorder {
            color,
            width,
            style: CharacterBorderStyle::from(style),
            spacing: (flags & 0x1F) as u8,
            has_shadow: flags & 0x20 != 0,
            has_frame: flags & 0x40 != 0,
        })
    }

    fn parse_shading(data: &[u8]) -> Option<CharacterShading> {
        let [colors @ .., pattern_low, pattern_high] = data else {
            return None;
        };
        if colors.len() != 8 {
            return None;
        }

        let foreground_color = CharacterColor::from_colorref(&colors[..4])?;
        let background_color = CharacterColor::from_colorref(&colors[4..])?;
        let pattern =
            CharacterShadingPattern::from(u16::from_le_bytes([*pattern_low, *pattern_high]));
        if matches!(pattern, CharacterShadingPattern::Nil)
            || (matches!(pattern, CharacterShadingPattern::Clear)
                && foreground_color == CharacterColor::Automatic
                && background_color == CharacterColor::Automatic)
        {
            return None;
        }

        Some(CharacterShading {
            foreground_color,
            background_color,
            pattern,
        })
    }

    fn parse_shading80(data: &[u8]) -> Option<CharacterShading> {
        let [low, high] = data else {
            return None;
        };
        let value = u16::from_le_bytes([*low, *high]);
        let foreground_index = (value & 0x1F) as u8;
        let background_index = ((value >> 5) & 0x1F) as u8;
        let pattern_value = (value >> 10) & 0x3F;

        if (foreground_index == 0x1F && background_index == 0x1F && pattern_value == 0x3F)
            || (foreground_index == 0 && background_index == 0 && pattern_value == 0)
        {
            return None;
        }

        Some(CharacterShading {
            foreground_color: CharacterColor::from_palette_index(foreground_index),
            background_color: CharacterColor::from_palette_index(background_index),
            pattern: CharacterShadingPattern::from(pattern_value),
        })
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
    pub(super) fn get_toggle_value(operand: u8, old_val: Option<bool>) -> bool {
        match operand {
            0 => false,
            1 => true,
            0x80 => old_val.unwrap_or(false),
            0x81 => !old_val.unwrap_or(false),
            _ => false,
        }
    }
}
