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
use super::super::package::{DocError, Result};
use super::tap::TableStyleCondition;
use crate::sprm::{Sprm, parse_sprms};
use crate::sprm_operations::*;

/// Character Properties structure.
///
/// Contains formatting information for a run of text.
/// Based on Apache POI's CharacterProperties implementation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
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
    /// Color of the text underline.
    pub underline_color: Option<CharacterColor>,
    /// Border drawn on all four sides of this character run.
    pub border: Option<CharacterBorder>,
    /// Background shading for this character run.
    pub shading: Option<CharacterShading>,
    /// Highlight color
    pub highlight: Option<HighlightColor>,
    /// Superscript/subscript
    pub vertical_position: VerticalPosition,
    /// Vertical offset relative to the normal baseline, in signed half-points.
    pub position: CharacterPosition,
    /// Word-breaking behavior used when this run is hyphenated.
    pub hyphenation: HresiOperand,
    /// Animated text effect applied to this run.
    pub text_effect: TextEffect,
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
    /// Whether complex-script formatting is forced for this run.
    pub is_complex_scripts: Option<bool>,
    /// Style index (istd)
    pub style_index: Option<u16>,
    /// Whether character properties before a tracked change are preserved.
    pub properties_preserved_for_revision: bool,
    /// Character state immediately before the active `sprmCWall` boundary.
    pub preserved_properties_for_revision: Option<Box<CharacterProperties>>,
    /// Conditional character formatting definitions carried by a table style.
    pub conditional_formats: Vec<CharacterConditionalFormatting>,
    /// Vanish (hidden)
    pub is_vanish: Option<bool>,
    /// Whether this run is marked as an inserted revision.
    pub is_revision_inserted: Option<bool>,
    /// Insertion revision author index in `SttbfRMark`.
    pub revision_author_index: Option<u16>,
    /// Packed insertion revision DTTM.
    pub revision_timestamp: Option<u32>,
    /// Insertion or modification reason code (`sprmCIdslRMark`).
    pub revision_id: Option<u16>,
    /// Character-formatting revision-save ID.
    pub formatting_revision_save_id: Option<u32>,
    /// Insertion revision-save ID.
    pub insertion_revision_save_id: Option<u32>,
    /// Whether this run is marked as a deleted revision.
    pub is_revision_deleted: Option<bool>,
    /// Deletion revision author index in `SttbfRMark`.
    pub deletion_author_index: Option<u16>,
    /// Packed deletion revision DTTM.
    pub deletion_timestamp: Option<u32>,
    /// Deletion reason code (`sprmCIdslRMarkDel`).
    pub deletion_revision_id: Option<u16>,
    /// Deletion revision-save ID.
    pub deletion_revision_save_id: Option<u32>,
    /// Whether this run has a tracked character-formatting change.
    pub has_formatting_revision: Option<bool>,
    /// Formatting revision author index in `SttbfRMark`.
    pub formatting_revision_author_index: Option<u16>,
    /// Packed formatting revision DTTM.
    pub formatting_revision_timestamp: Option<u32>,
    /// Revision metadata for a LISTNUM display-field result.
    pub display_field_revision: Option<DisplayFieldRevisionProperties>,
}

/// Vertical text position in signed half-points (`sprmCHpsPos`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CharacterPosition(i16);

impl CharacterPosition {
    /// Normal baseline position.
    pub const NORMAL: Self = Self(0);

    /// Construct a validated vertical position in half-points.
    pub fn new(half_points: i16) -> Result<Self> {
        if !(-3168..=3168).contains(&half_points) {
            return Err(DocError::Corrupted(format!(
                "sprmCHpsPos value {half_points} is outside -3168..=3168"
            )));
        }
        Ok(Self(half_points))
    }

    /// Return the signed half-point offset.
    pub const fn half_points(self) -> i16 {
        self.0
    }
}

/// Word-breaking method stored in an [`HresiOperand`].
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum HyphenationMode {
    /// Insert a hyphen and continue on the next line.
    #[default]
    Normal,
    /// Add the replacement character before the hyphen.
    AddBefore,
    /// Change the character before the hyphen.
    ChangeBefore,
    /// Delete the character before the hyphen.
    DeleteBefore,
    /// Change the character after the hyphen.
    ChangeAfter,
    /// Delete two characters before the hyphen and replace them.
    DeleteAndChange,
}

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
            _ => Err(DocError::Corrupted(format!(
                "sprmCHresi has invalid Hres mode {value}"
            ))),
        }
    }
}

/// Validated two-byte `HresiOperand` used by `sprmCHresi`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct HresiOperand {
    mode: HyphenationMode,
    replacement_character: Option<u8>,
}

impl HresiOperand {
    /// Normal word-breaking, whose dependent `ChHres` byte is zero.
    pub const fn normal() -> Self {
        Self {
            mode: HyphenationMode::Normal,
            replacement_character: None,
        }
    }

    /// Construct a non-normal word-breaking operand with a printable ASCII byte.
    pub fn with_character(mode: HyphenationMode, replacement_character: u8) -> Result<Self> {
        if mode == HyphenationMode::Normal {
            return Err(DocError::Corrupted(
                "normal HresiOperand cannot have a replacement character".to_string(),
            ));
        }
        if !replacement_character.is_ascii_graphic() && replacement_character != b' ' {
            return Err(DocError::Corrupted(format!(
                "sprmCHresi ChHres byte 0x{replacement_character:02X} is not printable ASCII"
            )));
        }
        Ok(Self {
            mode,
            replacement_character: Some(replacement_character),
        })
    }

    fn from_bytes(mode: u8, replacement_character: u8) -> Result<Self> {
        let mode = HyphenationMode::from_raw(mode)?;
        if mode == HyphenationMode::Normal {
            if replacement_character != 0 {
                return Err(DocError::Corrupted(
                    "normal sprmCHresi requires ChHres 0x00".to_string(),
                ));
            }
            Ok(Self::normal())
        } else {
            Self::with_character(mode, replacement_character)
        }
    }

    pub(crate) fn bytes(self) -> [u8; 2] {
        [self.mode.raw(), self.replacement_character.unwrap_or(0)]
    }

    /// Return the word-breaking mode.
    pub const fn mode(self) -> HyphenationMode {
        self.mode
    }

    /// Return the dependent printable ASCII byte, or `None` for normal breaking.
    pub const fn replacement_character(self) -> Option<u8> {
        self.replacement_character
    }
}

impl Default for HresiOperand {
    fn default() -> Self {
        Self::normal()
    }
}

/// Animated text effect stored by `sprmCSfxText`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum TextEffect {
    /// No text effect.
    #[default]
    None = 0,
    /// Las Vegas Lights.
    LasVegasLights = 1,
    /// Blinking background.
    BlinkingBackground = 2,
    /// Sparkle Text.
    SparkleText = 3,
    /// Marching Black Ants.
    MarchingBlackAnts = 4,
    /// Marching Red Ants.
    MarchingRedAnts = 5,
    /// Shimmer.
    Shimmer = 6,
}

impl TryFrom<u8> for TextEffect {
    type Error = DocError;

    fn try_from(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::LasVegasLights),
            2 => Ok(Self::BlinkingBackground),
            3 => Ok(Self::SparkleText),
            4 => Ok(Self::MarchingBlackAnts),
            5 => Ok(Self::MarchingRedAnts),
            6 => Ok(Self::Shimmer),
            _ => Err(DocError::Corrupted(format!(
                "sprmCSfxText has invalid text effect {value}"
            ))),
        }
    }
}

impl From<TextEffect> for u8 {
    fn from(value: TextEffect) -> Self {
        value as u8
    }
}

/// Conditional character formatting carried by `sprmCCnf` in a table style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CharacterConditionalFormatting {
    /// Table location or band for which the nested properties apply.
    pub condition: TableStyleCondition,
    /// Typed character properties decoded from the nested grpprl.
    pub properties: Box<CharacterProperties>,
    /// Exact nested grpprl retained for lossless preservation.
    pub raw_grpprl: Vec<u8>,
}

/// Parsed `DispFldRmOperand` state for a LISTNUM display-field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DisplayFieldRevisionProperties {
    /// Whether the field result contains a revision.
    pub active: bool,
    /// Revision author index in `SttbfRMark`.
    pub author_index: u16,
    /// Packed revision DTTM.
    pub timestamp: u32,
    /// Previous LISTNUM field result.
    pub previous_result: String,
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

/// A color stored in a legacy Word character property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterColor {
    /// The application chooses a context-appropriate color.
    Automatic,
    /// An explicit red, green, and blue color.
    Rgb(u8, u8, u8),
    /// An invalid or future palette index retained without interpretation.
    ReservedPaletteIndex(u8),
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

/// Border formatting applied around a character run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CharacterBorder {
    /// Border color, including the context-dependent automatic color.
    pub color: CharacterColor,
    /// Border width in eighths of a point for ordinary character borders.
    pub width: u8,
    /// Border line style.
    pub style: CharacterBorderStyle,
    /// Distance between the border and text, in points.
    pub spacing: u8,
    /// Whether Word adds a shadow to the border.
    pub has_shadow: bool,
    /// Whether Word reverses the border to create a frame effect.
    pub has_frame: bool,
}

/// Line styles supported by MS-DOC character borders.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterBorderStyle {
    Single,
    Double,
    Thin,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
    /// An invalid or future style retained without coercion.
    Reserved(u8),
}

impl From<u8> for CharacterBorderStyle {
    fn from(value: u8) -> Self {
        match value {
            0x01 => Self::Single,
            0x03 => Self::Double,
            0x05 => Self::Thin,
            0x06 => Self::Dotted,
            0x07 => Self::Dashed,
            0x08 => Self::DotDash,
            0x09 => Self::DotDotDash,
            0x0A => Self::Triple,
            0x0B => Self::ThinThickSmallGap,
            0x0C => Self::ThickThinSmallGap,
            0x0D => Self::ThinThickThinSmallGap,
            0x0E => Self::ThinThickMediumGap,
            0x0F => Self::ThickThinMediumGap,
            0x10 => Self::ThinThickThinMediumGap,
            0x11 => Self::ThinThickLargeGap,
            0x12 => Self::ThickThinLargeGap,
            0x13 => Self::ThinThickThinLargeGap,
            0x14 => Self::Wave,
            0x15 => Self::DoubleWave,
            0x16 => Self::DashSmallGap,
            0x17 => Self::DashDotStroked,
            0x18 => Self::ThreeDEmboss,
            0x19 => Self::ThreeDEngrave,
            0x1A => Self::Outset,
            0x1B => Self::Inset,
            value => Self::Reserved(value),
        }
    }
}

/// Shading applied behind a character run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CharacterShading {
    pub foreground_color: CharacterColor,
    pub background_color: CharacterColor,
    pub pattern: CharacterShadingPattern,
}

/// Shading patterns defined by the MS-DOC `Ipat` enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterShadingPattern {
    Clear,
    Solid,
    Percent5,
    Percent10,
    Percent20,
    Percent25,
    Percent30,
    Percent40,
    Percent50,
    Percent60,
    Percent70,
    Percent75,
    Percent80,
    Percent90,
    DarkHorizontal,
    DarkVertical,
    DarkForwardDiagonal,
    DarkBackwardDiagonal,
    DarkCross,
    DarkDiagonalCross,
    Horizontal,
    Vertical,
    ForwardDiagonal,
    BackwardDiagonal,
    Cross,
    DiagonalCross,
    Percent2_5,
    Percent7_5,
    Percent12_5,
    Percent15,
    Percent17_5,
    Percent22_5,
    Percent27_5,
    Percent32_5,
    Percent35,
    Percent37_5,
    Percent42_5,
    Percent45,
    Percent47_5,
    Percent52_5,
    Percent55,
    Percent57_5,
    Percent62_5,
    Percent65,
    Percent67_5,
    Percent72_5,
    Percent77_5,
    Percent82_5,
    Percent85,
    Percent87_5,
    Percent92_5,
    Percent95,
    Percent97_5,
    Nil,
    /// An invalid or future pattern retained without coercion.
    Reserved(u16),
}

impl From<u16> for CharacterShadingPattern {
    fn from(value: u16) -> Self {
        match value {
            0x0000 => Self::Clear,
            0x0001 => Self::Solid,
            0x0002 => Self::Percent5,
            0x0003 => Self::Percent10,
            0x0004 => Self::Percent20,
            0x0005 => Self::Percent25,
            0x0006 => Self::Percent30,
            0x0007 => Self::Percent40,
            0x0008 => Self::Percent50,
            0x0009 => Self::Percent60,
            0x000A => Self::Percent70,
            0x000B => Self::Percent75,
            0x000C => Self::Percent80,
            0x000D => Self::Percent90,
            0x000E => Self::DarkHorizontal,
            0x000F => Self::DarkVertical,
            0x0010 => Self::DarkForwardDiagonal,
            0x0011 => Self::DarkBackwardDiagonal,
            0x0012 => Self::DarkCross,
            0x0013 => Self::DarkDiagonalCross,
            0x0014 => Self::Horizontal,
            0x0015 => Self::Vertical,
            0x0016 => Self::ForwardDiagonal,
            0x0017 => Self::BackwardDiagonal,
            0x0018 => Self::Cross,
            0x0019 => Self::DiagonalCross,
            0x0023 => Self::Percent2_5,
            0x0024 => Self::Percent7_5,
            0x0025 => Self::Percent12_5,
            0x0026 => Self::Percent15,
            0x0027 => Self::Percent17_5,
            0x0028 => Self::Percent22_5,
            0x0029 => Self::Percent27_5,
            0x002A => Self::Percent32_5,
            0x002B => Self::Percent35,
            0x002C => Self::Percent37_5,
            0x002D => Self::Percent42_5,
            0x002E => Self::Percent45,
            0x002F => Self::Percent47_5,
            0x0030 => Self::Percent52_5,
            0x0031 => Self::Percent55,
            0x0032 => Self::Percent57_5,
            0x0033 => Self::Percent62_5,
            0x0034 => Self::Percent65,
            0x0035 => Self::Percent67_5,
            0x0036 => Self::Percent72_5,
            0x0037 => Self::Percent77_5,
            0x0038 => Self::Percent82_5,
            0x0039 => Self::Percent85,
            0x003A => Self::Percent87_5,
            0x003B => Self::Percent92_5,
            0x003C => Self::Percent95,
            0x003D => Self::Percent97_5,
            0xFFFF => Self::Nil,
            value => Self::Reserved(value),
        }
    }
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
        let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
        if consumed != grpprl.len() {
            return Err(DocError::Corrupted(
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
    /// Based on Apache POI's CharacterSprmUncompressor.unCompressCHPOperation().
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
                    DocError::Corrupted("sprmCDttmRMark is missing its DTTM".to_string())
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
                    DocError::Corrupted("sprmCRsidProp is missing its RSID".to_string())
                })?);
            },
            0x16 => {
                // sprmCRsidText - Revision save ID text
                chp.insertion_revision_save_id = Some(sprm.operand_dword().ok_or_else(|| {
                    DocError::Corrupted("sprmCRsidText is missing its RSID".to_string())
                })?);
            },
            0x17 => {
                // sprmCRsidRMDel - Revision save ID deletion
                chp.deletion_revision_save_id = Some(sprm.operand_dword().ok_or_else(|| {
                    DocError::Corrupted("sprmCRsidRMDel is missing its RSID".to_string())
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
                let position = sprm.operand_i16().ok_or_else(|| {
                    DocError::Corrupted("sprmCHpsPos is missing its signed operand".to_string())
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
                let operand = sprm.operand_word().ok_or_else(|| {
                    DocError::Corrupted("sprmCHresi is missing its HresiOperand".to_string())
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
                    DocError::Corrupted("sprmCSfxText is missing its byte operand".to_string())
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
                    DocError::Corrupted("sprmCDttmRMarkDel is missing its DTTM".to_string())
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
            _ => Err(DocError::Corrupted(format!(
                "{name} must contain a Boolean8 value"
            ))),
        }
    }

    fn strict_bool8(sprm: &Sprm, name: &str) -> Result<bool> {
        match sprm.operand_byte() {
            Some(0) => Ok(false),
            Some(1) => Ok(true),
            _ => Err(DocError::Corrupted(format!(
                "{name} must contain a Boolean8 value"
            ))),
        }
    }

    fn parse_conditional_formatting(sprm: &Sprm) -> Result<CharacterConditionalFormatting> {
        let operand = sprm.operand_bytes();
        if operand.len() < 2 {
            return Err(DocError::Corrupted(
                "sprmCCnf must contain a 2-byte condition".to_string(),
            ));
        }
        let code = u16::from_le_bytes([operand[0], operand[1]]);
        let condition = TableStyleCondition::from_code(code).ok_or_else(|| {
            DocError::Corrupted(format!("sprmCCnf contains invalid condition {code:#06x}"))
        })?;
        let raw_grpprl = operand[2..].to_vec();
        let nested = parse_sprms(&raw_grpprl);
        let consumed = nested
            .last()
            .map_or(0, |nested| nested.offset + nested.size);
        if consumed != raw_grpprl.len() {
            return Err(DocError::Corrupted(
                "sprmCCnf nested grpprl is truncated".to_string(),
            ));
        }
        if nested.iter().any(|nested| nested.opcode == SPRM_C_CNF) {
            return Err(DocError::Corrupted(
                "sprmCCnf cannot be nested inside another sprmCCnf".to_string(),
            ));
        }
        if nested
            .iter()
            .any(|nested| get_sprm_type(nested.opcode) != 2)
        {
            return Err(DocError::Corrupted(
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
        let value = sprm
            .operand_i16()
            .ok_or_else(|| DocError::Corrupted(format!("{name} is missing its author index")))?;
        u16::try_from(value)
            .map_err(|_| DocError::Corrupted(format!("{name} author index is negative")))
    }

    fn revision_reason(sprm: &Sprm, name: &str) -> Result<u16> {
        let value = sprm
            .operand_word()
            .ok_or_else(|| DocError::Corrupted(format!("{name} is missing its reason code")))?;
        if value > super::super::revision::RevisionReason::MAX_VALUE {
            return Err(DocError::Corrupted(format!(
                "{name} contains an undefined reason code"
            )));
        }
        Ok(value)
    }

    fn apply_property_revision(chp: &mut CharacterProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 7 {
            return Err(DocError::Corrupted(
                "sprmCPropRMark operand must contain exactly 7 bytes".to_string(),
            ));
        }
        chp.has_formatting_revision = Some(match operand[0] {
            0 => false,
            1 => true,
            _ => {
                return Err(DocError::Corrupted(
                    "sprmCPropRMark must begin with a Boolean8 value".to_string(),
                ));
            },
        });
        let author = i16::from_le_bytes([operand[1], operand[2]]);
        chp.formatting_revision_author_index = Some(u16::try_from(author).map_err(|_| {
            DocError::Corrupted("sprmCPropRMark author index is negative".to_string())
        })?);
        chp.formatting_revision_timestamp = Some(u32::from_le_bytes([
            operand[3], operand[4], operand[5], operand[6],
        ]));
        Ok(())
    }

    fn parse_display_field_revision(sprm: &Sprm) -> Result<DisplayFieldRevisionProperties> {
        let operand = sprm.operand_bytes();
        if operand.len() != 39 {
            return Err(DocError::Corrupted(
                "sprmCDispFldRMark operand must contain exactly 39 bytes".to_string(),
            ));
        }
        let active = operand[0] != 0;
        let author_index = u16::from_le_bytes([operand[1], operand[2]]);
        let timestamp = u32::from_le_bytes([operand[3], operand[4], operand[5], operand[6]]);
        let string_length = usize::from(u16::from_le_bytes([operand[7], operand[8]]));
        if string_length > 15 {
            return Err(DocError::Corrupted(
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
            DocError::Corrupted("LISTNUM previous result is invalid UTF-16".to_string())
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
            || self.underline_color.is_some()
            || self.border.is_some()
            || self.shading.is_some()
            || self.highlight.is_some()
            || self.vertical_position != VerticalPosition::Normal
            || self.position != CharacterPosition::NORMAL
            || self.hyphenation != HresiOperand::normal()
            || self.text_effect != TextEffect::None
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
            || self.is_complex_scripts.is_some()
            || self.properties_preserved_for_revision
            || !self.conditional_formats.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_insertion_and_deletion_revision_sprms() {
        let timestamp =
            30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
        let mut grpprl = Vec::new();
        for (opcode, operand) in [
            (SPRM_C_F_RMARK, vec![1]),
            (SPRM_C_IBST_RMARK, 1u16.to_le_bytes().to_vec()),
            (SPRM_C_DTTM_RMARK, timestamp.to_le_bytes().to_vec()),
            (SPRM_C_IDSL_RMARK, 42u16.to_le_bytes().to_vec()),
            (SPRM_C_RSID_PROP, 0x11223344u32.to_le_bytes().to_vec()),
            (SPRM_C_RSID_TEXT, 0x55667788u32.to_le_bytes().to_vec()),
            (SPRM_C_F_RMARK_DEL, vec![1]),
            (SPRM_C_IBST_RMARK_DEL, 0u16.to_le_bytes().to_vec()),
            (SPRM_C_DTTM_RMARK_DEL, 0u32.to_le_bytes().to_vec()),
            (SPRM_C_IDSL_RMARK_DEL, 7u16.to_le_bytes().to_vec()),
            (SPRM_C_RSID_RM_DEL, 0x99AABBCCu32.to_le_bytes().to_vec()),
        ] {
            grpprl.extend_from_slice(&opcode.to_le_bytes());
            grpprl.extend_from_slice(&operand);
        }
        let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.is_revision_inserted, Some(true));
        assert_eq!(properties.revision_author_index, Some(1));
        assert_eq!(properties.revision_timestamp, Some(timestamp));
        assert_eq!(properties.revision_id, Some(42));
        assert_eq!(properties.formatting_revision_save_id, Some(0x11223344));
        assert_eq!(properties.insertion_revision_save_id, Some(0x55667788));
        assert_eq!(properties.is_revision_deleted, Some(true));
        assert_eq!(properties.deletion_author_index, Some(0));
        assert_eq!(properties.deletion_timestamp, Some(0));
        assert_eq!(properties.deletion_revision_id, Some(7));
        assert_eq!(properties.deletion_revision_save_id, Some(0x99AABBCC));

        let mut malformed = Vec::new();
        malformed.extend_from_slice(&SPRM_C_IBST_RMARK.to_le_bytes());
        malformed.extend_from_slice(&(-1i16).to_le_bytes());
        assert!(CharacterProperties::from_sprm(&malformed).is_err());

        let mut undefined_reason = SPRM_C_IDSL_RMARK.to_le_bytes().to_vec();
        undefined_reason.extend_from_slice(&0x002Cu16.to_le_bytes());
        assert!(CharacterProperties::from_sprm(&undefined_reason).is_err());
    }

    #[test]
    fn parses_both_character_formatting_revision_sprms_strictly() {
        let timestamp =
            30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
        for opcode in [SPRM_C_PROP_RMARK90, SPRM_C_PROP_RMARK_CURRENT] {
            let mut grpprl = opcode.to_le_bytes().to_vec();
            grpprl.push(7);
            grpprl.push(1);
            grpprl.extend_from_slice(&2i16.to_le_bytes());
            grpprl.extend_from_slice(&timestamp.to_le_bytes());
            let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
            assert_eq!(properties.has_formatting_revision, Some(true));
            assert_eq!(properties.formatting_revision_author_index, Some(2));
            assert_eq!(properties.formatting_revision_timestamp, Some(timestamp));
        }

        for operand in [
            vec![2, 0, 0, 0, 0, 0, 0],
            vec![1, 0xFF, 0xFF, 0, 0, 0, 0],
            vec![1, 0, 0, 0, 0, 0],
        ] {
            let mut grpprl = SPRM_C_PROP_RMARK_CURRENT.to_le_bytes().to_vec();
            grpprl.push(operand.len() as u8);
            grpprl.extend_from_slice(&operand);
            assert!(CharacterProperties::from_sprm(&grpprl).is_err());
        }
    }

    #[test]
    fn parses_display_field_revision_strictly() {
        let timestamp =
            30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
        let mut operand = [0u8; 39];
        operand[0] = 2; // Any nonzero value means active.
        operand[1..3].copy_from_slice(&2u16.to_le_bytes());
        operand[3..7].copy_from_slice(&timestamp.to_le_bytes());
        operand[7..9].copy_from_slice(&3u16.to_le_bytes());
        for (index, unit) in "12.".encode_utf16().enumerate() {
            let offset = 9 + index * 2;
            operand[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
        }
        let mut grpprl = SPRM_C_DISP_FLD_RMARK.to_le_bytes().to_vec();
        grpprl.push(39);
        grpprl.extend_from_slice(&operand);
        let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
        let revision = properties.display_field_revision.unwrap();
        assert!(revision.active);
        assert_eq!(revision.author_index, 2);
        assert_eq!(revision.timestamp, timestamp);
        assert_eq!(revision.previous_result, "12.");

        for length in [16u8, 38] {
            let mut invalid = operand.to_vec();
            if length == 16 {
                invalid[7..9].copy_from_slice(&16u16.to_le_bytes());
            } else {
                invalid.truncate(38);
            }
            let mut grpprl = SPRM_C_DISP_FLD_RMARK.to_le_bytes().to_vec();
            grpprl.push(length);
            grpprl.extend_from_slice(&invalid);
            assert!(CharacterProperties::from_sprm(&grpprl).is_err());
        }
    }

    #[test]
    fn test_default_chp() {
        let chp = CharacterProperties::new();
        assert_eq!(chp.is_bold, None);
        assert_eq!(chp.is_italic, None);
        assert_eq!(chp.underline, UnderlineStyle::None);
        assert!(!chp.has_formatting());
    }

    #[test]
    fn preserves_ordered_character_revision_state() {
        let mut grpprl = SPRM_C_F_BOLD.to_le_bytes().to_vec();
        grpprl.push(1);
        grpprl.extend_from_slice(&SPRM_C_WALL.to_le_bytes());
        grpprl.push(1);
        grpprl.extend_from_slice(&SPRM_C_F_ITALIC.to_le_bytes());
        grpprl.push(1);

        let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.is_bold, Some(true));
        assert_eq!(properties.is_italic, Some(true));
        assert!(properties.properties_preserved_for_revision);
        let previous = properties.preserved_properties_for_revision.unwrap();
        assert_eq!(previous.is_bold, Some(true));
        assert_eq!(previous.is_italic, None);

        grpprl.extend_from_slice(&SPRM_C_WALL.to_le_bytes());
        grpprl.push(0);
        let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
        assert!(!properties.properties_preserved_for_revision);
        assert!(properties.preserved_properties_for_revision.is_none());

        let invalid = [SPRM_C_WALL.to_le_bytes().as_slice(), &[2]].concat();
        assert!(CharacterProperties::from_sprm(&invalid).is_err());
    }

    #[test]
    fn parses_conditional_table_style_character_formatting_strictly() {
        let wrap = |condition: u16, nested: &[u8]| {
            let mut grpprl = SPRM_C_CNF.to_le_bytes().to_vec();
            grpprl.push((nested.len() + 2) as u8);
            grpprl.extend_from_slice(&condition.to_le_bytes());
            grpprl.extend_from_slice(nested);
            grpprl
        };
        let nested = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
        let properties = CharacterProperties::from_sprm(&wrap(0x0008, &nested)).unwrap();
        let conditional = &properties.conditional_formats[0];
        assert_eq!(conditional.condition, TableStyleCondition::LastColumn);
        assert_eq!(conditional.raw_grpprl, nested);
        assert_eq!(conditional.properties.is_bold, Some(true));

        let recursive = wrap(0x0002, &[]);
        let paragraph = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
        let truncated = SPRM_C_F_BOLD.to_le_bytes();
        for invalid in [
            [SPRM_C_CNF.to_le_bytes().as_slice(), &[0]].concat(),
            wrap(0x0003, &[]),
            wrap(0x0001, &recursive),
            wrap(0x0001, &paragraph),
            wrap(0x0001, &truncated),
        ] {
            assert!(CharacterProperties::from_sprm(&invalid).is_err());
        }
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

    #[test]
    fn parses_palette_character_border_and_shading() {
        let properties = CharacterProperties::from_sprm(&[
            0x65, 0x68, // sprmCBrc80
            0x10, 0x14, 0x06, 0x65, // 2pt wave, red, 5pt, shadow and frame
            0x66, 0x48, // sprmCShd80
            0xE6, 0x04, // solid red foreground on yellow
        ])
        .unwrap();

        assert_eq!(
            properties.border,
            Some(CharacterBorder {
                color: CharacterColor::Rgb(255, 0, 0),
                width: 0x10,
                style: CharacterBorderStyle::Wave,
                spacing: 5,
                has_shadow: true,
                has_frame: true,
            })
        );
        assert_eq!(
            properties.shading,
            Some(CharacterShading {
                foreground_color: CharacterColor::Rgb(255, 0, 0),
                background_color: CharacterColor::Rgb(255, 255, 0),
                pattern: CharacterShadingPattern::Solid,
            })
        );
    }

    #[test]
    fn parses_rgb_character_border_shading_and_underline() {
        let properties = CharacterProperties::from_sprm(&[
            0x71, 0xCA, // sprmCShd
            0x0A, // SHDOperand byte count
            0x12, 0x34, 0x56, 0x00, // foreground COLORREF
            0x00, 0x00, 0x00, 0xFF, // automatic background COLORREF
            0x25, 0x00, // 12.5 percent pattern
            0x72, 0xCA, // sprmCBrc
            0x08, // BrcOperand byte count
            0xAA, 0xBB, 0xCC, 0x00, // border COLORREF
            0x10, 0x18, 0x43, 0x00, // width, emboss, 3pt, frame
            0x77, 0x68, // sprmCCvUl
            0x01, 0x02, 0x03, 0x00, // underline COLORREF
            0x82, 0x08, 0x01, // sprmCFComplexScripts
        ])
        .unwrap();

        assert_eq!(
            properties.shading,
            Some(CharacterShading {
                foreground_color: CharacterColor::Rgb(0x12, 0x34, 0x56),
                background_color: CharacterColor::Automatic,
                pattern: CharacterShadingPattern::Percent12_5,
            })
        );
        assert_eq!(
            properties.border,
            Some(CharacterBorder {
                color: CharacterColor::Rgb(0xAA, 0xBB, 0xCC),
                width: 0x10,
                style: CharacterBorderStyle::ThreeDEmboss,
                spacing: 3,
                has_shadow: false,
                has_frame: true,
            })
        );
        assert_eq!(
            properties.underline_color,
            Some(CharacterColor::Rgb(1, 2, 3))
        );
        assert_eq!(properties.is_complex_scripts, Some(true));
        assert!(properties.has_formatting());
    }

    #[test]
    fn preserves_reserved_character_format_values() {
        let properties = CharacterProperties::from_sprm(&[
            0x65, 0x68, // sprmCBrc80
            0x02, 0x02, 0x11, 0x00, // reserved style and palette index
            0x71, 0xCA, // sprmCShd
            0x0A, 0x01, 0x02, 0x03, 0x00, 0x04, 0x05, 0x06, 0x00, 0x1A, 0x00,
        ])
        .unwrap();

        let border = properties.border.unwrap();
        assert_eq!(border.style, CharacterBorderStyle::Reserved(0x02));
        assert_eq!(border.color, CharacterColor::ReservedPaletteIndex(0x11));
        assert_eq!(
            properties.shading.unwrap().pattern,
            CharacterShadingPattern::Reserved(0x001A)
        );
    }
}

#[cfg(test)]
mod chpx_position_hresi_effect_tests {
    use super::*;

    fn append(grpprl: &mut Vec<u8>, opcode: u16, operand: &[u8]) {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.extend_from_slice(operand);
    }

    #[test]
    fn decodes_all_values_boundaries_defaults_and_later_wins() {
        let modes = [
            (HyphenationMode::Normal, [1, 0]),
            (HyphenationMode::AddBefore, [2, b'A']),
            (HyphenationMode::ChangeBefore, [3, b'B']),
            (HyphenationMode::DeleteBefore, [4, b' ']),
            (HyphenationMode::ChangeAfter, [5, b'Y']),
            (HyphenationMode::DeleteAndChange, [6, b'Z']),
        ];
        for (mode, bytes) in modes {
            let mut grpprl = Vec::new();
            append(&mut grpprl, SPRM_C_HRESI, &bytes);
            let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
            assert_eq!(properties.hyphenation.mode(), mode);
            assert_eq!(properties.hyphenation.bytes(), bytes);
        }
        for (raw, effect) in [
            (0, TextEffect::None),
            (1, TextEffect::LasVegasLights),
            (2, TextEffect::BlinkingBackground),
            (3, TextEffect::SparkleText),
            (4, TextEffect::MarchingBlackAnts),
            (5, TextEffect::MarchingRedAnts),
            (6, TextEffect::Shimmer),
        ] {
            let mut grpprl = Vec::new();
            append(&mut grpprl, SPRM_C_SFXT_TEXT, &[raw]);
            assert_eq!(
                CharacterProperties::from_sprm(&grpprl).unwrap().text_effect,
                effect
            );
        }

        let mut grpprl = Vec::new();
        append(&mut grpprl, SPRM_C_HPS_POS, &(-3168i16).to_le_bytes());
        append(&mut grpprl, SPRM_C_HPS_POS, &3168i16.to_le_bytes());
        append(&mut grpprl, SPRM_C_HRESI, &[1, 0]);
        append(&mut grpprl, SPRM_C_HRESI, &[5, b'Q']);
        append(&mut grpprl, SPRM_C_SFXT_TEXT, &[1]);
        append(&mut grpprl, SPRM_C_SFXT_TEXT, &[6]);
        let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.position.half_points(), 3168);
        assert_eq!(properties.hyphenation.mode(), HyphenationMode::ChangeAfter);
        assert_eq!(properties.hyphenation.replacement_character(), Some(b'Q'));
        assert_eq!(properties.text_effect, TextEffect::Shimmer);
        assert!(properties.has_formatting());

        let defaults = CharacterProperties::default();
        assert_eq!(defaults.position, CharacterPosition::NORMAL);
        assert_eq!(defaults.hyphenation, HresiOperand::normal());
        assert_eq!(defaults.text_effect, TextEffect::None);
    }

    #[test]
    fn rejects_out_of_range_and_dependent_operands() {
        assert!(CharacterPosition::new(-3169).is_err());
        assert!(CharacterPosition::new(3169).is_err());
        assert!(CharacterPosition::new(-3168).is_ok());
        assert!(CharacterPosition::new(3168).is_ok());
        assert!(HresiOperand::with_character(HyphenationMode::Normal, b'A').is_err());
        for byte in [0x00, 0x1F, 0x7F, 0x80, 0xFF] {
            assert!(HresiOperand::with_character(HyphenationMode::AddBefore, byte).is_err());
        }

        for (opcode, operand) in [
            (SPRM_C_HPS_POS, 3169i16.to_le_bytes()),
            (SPRM_C_HRESI, [0, 0]),
            (SPRM_C_HRESI, [7, b'A']),
            (SPRM_C_HRESI, [1, b'A']),
            (SPRM_C_HRESI, [2, 0]),
        ] {
            let mut grpprl = Vec::new();
            append(&mut grpprl, opcode, &operand);
            assert!(CharacterProperties::from_sprm(&grpprl).is_err());
        }
        let mut grpprl = Vec::new();
        append(&mut grpprl, SPRM_C_SFXT_TEXT, &[7]);
        assert!(CharacterProperties::from_sprm(&grpprl).is_err());
    }
}
